#!/usr/bin/env python3
"""
inventory.py - read-only preflight scanner for a document rebranding batch.

Scans a tree of .docx files and, for each one, works out:
  * where the logo lives and how it is anchored
  * whether its display box matches the image's true aspect ratio
  * the Dok-Nr / version / Bearbeiter / Freigeber / Pruefer / Freigabedatum values
  * whether the Freigabedatum is a date content control (w:sdt) with a stored
    w:fullDate that would need syncing as well as the visible text
  * whether a changelog table exists and can be extended
  * a structural fingerprint, so near-identical files can be grouped

It then DRY-RUNS the patch and reports would_patch = YES / NO per file.
Nothing is ever modified: every archive is opened read-only.

Usage:
    python3 inventory.py --input /path/to/docs --out inventory.csv --errors errors.log
    python3 inventory.py --input . --shapes        # print fingerprint summary only
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import struct
import sys
import traceback
import unicodedata
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V = "urn:schemas-microsoft-com:vml"

ASPECT_TOLERANCE = 0.02  # 2% - beyond this the logo is being visibly stretched

# Metadata labels we care about, in normalised form.
LABELS = {
    "doknrversionnr": "dok_nr",
    "thematik": "thematik",
    "bearbeiter": "bearbeiter",
    "freigeber": "freigeber",
    "prozessverantwortung": "prozessverantwortung",
    "unterkategorie": "unterkategorie",
    "prufer": "pruefer",
    "freigabedatum": "freigabedatum",
}


# --------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------

def norm_label(s: str) -> str:
    """Normalise a table label: lowercase, strip accents and punctuation."""
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def image_size(blob: bytes):
    """Return (w, h) for PNG/JPEG/GIF without requiring Pillow."""
    try:
        if blob[:8] == b"\x89PNG\r\n\x1a\n":
            w, h = struct.unpack(">II", blob[16:24])
            return int(w), int(h)
        if blob[:3] == b"\xff\xd8\xff":
            i = 2
            while i < len(blob) - 9:
                if blob[i] != 0xFF:
                    i += 1
                    continue
                marker = blob[i + 1]
                if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
                    i += 2
                    continue
                seglen = struct.unpack(">H", blob[i + 2:i + 4])[0]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6,
                              0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    h, w = struct.unpack(">HH", blob[i + 5:i + 9])
                    return int(w), int(h)
                i += 2 + seglen
        if blob[:6] in (b"GIF87a", b"GIF89a"):
            w, h = struct.unpack("<HH", blob[6:10])
            return int(w), int(h)
    except Exception:
        pass
    return None


def _direct(node, wanted: str):
    """
    Collect descendants tagged `wanted` that belong to `node` directly, seeing
    through wrappers such as w:sdt / w:sdtContent but never descending into a
    nested table. Word wraps dropdown and date cells in content controls, so a
    plain findall() silently misses them and shifts every later column left.
    """
    found = []

    def walk(n):
        for child in n:
            tag = child.tag
            if tag == wanted:
                found.append(child)
            elif tag == f"{{{W}}}tbl":
                continue
            else:
                walk(child)

    walk(node)
    return found


def cell_text(tc) -> str:
    """Text of a cell, excluding any nested table's text."""
    parts = []

    def walk(n):
        for child in n:
            if child.tag == f"{{{W}}}tbl":
                continue
            if child.tag == f"{{{W}}}t":
                parts.append(child.text or "")
            walk(child)

    walk(tc)
    return "".join(parts)


def grid_span(tc) -> int:
    """How many grid columns this cell occupies (horizontal merge)."""
    for pr in tc:
        if pr.tag == f"{{{W}}}tcPr":
            for el in pr:
                if el.tag == f"{{{W}}}gridSpan":
                    try:
                        return max(1, int(el.get(f"{{{W}}}val", "1")))
                    except (TypeError, ValueError):
                        return 1
    return 1


def parse_tables(root):
    """
    Every table as (element, rows), where rows is a list of lists of cell text.
    Cells are expanded by gridSpan so that column indices line up between the
    label row and the value row beneath it.
    """
    tables = []
    for tbl in root.iter(f"{{{W}}}tbl"):
        rows = []
        for tr in _direct(tbl, f"{{{W}}}tr"):
            row = []
            for tc in _direct(tr, f"{{{W}}}tc"):
                text = cell_text(tc)
                span = grid_span(tc)
                row.append(text)
                row.extend([""] * (span - 1))
            rows.append(row)
        if rows:
            tables.append((tbl, rows))
    return tables


def match_label(text: str):
    """
    Map a cell's text to a known metadata label.
    Returns (key, exact). Falls back to a prefix match so that a label cell
    polluted by a stray character - e.g. 'Freigeber28', where a digit of the
    document number was typed into the header cell - is still recognised.
    """
    n = norm_label(text)
    if not n:
        return None, False
    if n in LABELS:
        return LABELS[n], True
    candidates = [(lbl, key) for lbl, key in LABELS.items() if n.startswith(lbl)]
    if candidates:
        lbl, key = max(candidates, key=lambda pair: len(pair[0]))
        return key, False
    return None, False


def label_lookup(rows):
    """Map label -> value taken from the cell directly below the label cell."""
    found = {}
    for i, row in enumerate(rows):
        for j, text in enumerate(row):
            key, _exact = match_label(text)
            if key and i + 1 < len(rows) and j < len(rows[i + 1]):
                found.setdefault(key, rows[i + 1][j].strip())
    return found


def label_lookup_flags(rows):
    """As label_lookup, but also returns the labels matched only by prefix."""
    found, fuzzy = {}, []
    for i, row in enumerate(rows):
        for j, text in enumerate(row):
            key, exact = match_label(text)
            if key and i + 1 < len(rows) and j < len(rows[i + 1]):
                if key not in found:
                    found[key] = rows[i + 1][j].strip()
                    if not exact:
                        fuzzy.append(f"{key}<-{text.strip()!r}")
    return found, fuzzy


# --------------------------------------------------------------------------
# per-file inspection
# --------------------------------------------------------------------------

class Report(dict):
    """A flat row of results, plus accumulated blocking reasons."""

    def __init__(self, path: Path, root: Path):
        super().__init__()
        try:
            rel = str(path.relative_to(root))
        except ValueError:
            rel = str(path)
        self["file"] = rel
        self["status"] = "OK"
        self["error"] = ""
        self["blockers"] = []
        self["notes"] = []

    def block(self, reason):
        self["blockers"].append(reason)

    def note(self, reason):
        self["notes"].append(reason)


def rels_for(z, part: str) -> dict:
    """rId -> target for a given part, resolved relative to word/."""
    name = f"{os.path.dirname(part)}/_rels/{os.path.basename(part)}.rels"
    out = {}
    try:
        root = ET.fromstring(z.read(name))
    except (KeyError, ET.ParseError):
        return out
    ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    for rel in root.findall(f"{ns}Relationship"):
        tgt = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            continue
        tgt = re.sub(r"^\.\./", "", tgt)
        if not tgt.startswith("word/") and not tgt.startswith("/"):
            tgt = "word/" + tgt.lstrip("/")
        out[rel.get("Id")] = tgt.lstrip("/")
    return out


def images_in(z, part: str):
    """
    Every image reference in a part.
    Returns list of dicts: rid, target, kind (anchor|inline|vml), cx, cy.
    """
    out = []
    try:
        root = ET.fromstring(z.read(part))
    except (KeyError, ET.ParseError):
        return out
    rels = rels_for(z, part)

    for drawing in root.iter(f"{{{W}}}drawing"):
        for kind_tag, kind in ((f"{{{WP}}}anchor", "anchor"), (f"{{{WP}}}inline", "inline")):
            for node in drawing.iter(kind_tag):
                ext = node.find(f"{{{WP}}}extent")
                cx = int(ext.get("cx")) if ext is not None and ext.get("cx") else None
                cy = int(ext.get("cy")) if ext is not None and ext.get("cy") else None
                rid = None
                for blip in node.iter(f"{{{A}}}blip"):
                    rid = blip.get(f"{{{R}}}embed") or blip.get(f"{{{R}}}link")
                    break
                out.append({"rid": rid, "target": rels.get(rid), "kind": kind,
                            "cx": cx, "cy": cy})

    for pict in root.iter(f"{{{W}}}pict"):
        for imagedata in pict.iter(f"{{{V}}}imagedata"):
            rid = imagedata.get(f"{{{R}}}id")
            out.append({"rid": rid, "target": rels.get(rid), "kind": "vml",
                        "cx": None, "cy": None})
    return out


def inspect(path: Path, root: Path) -> Report:
    rep = Report(path, root)

    if path.name.startswith("~$"):
        rep["status"] = "SKIP"
        rep["error"] = "Word lock file"
        return rep

    try:
        with zipfile.ZipFile(path, "r") as z:
            names = set(z.namelist())

            if "word/document.xml" not in names:
                rep["status"] = "ERROR"
                rep["error"] = "no word/document.xml (not a Word file, or .doc renamed)"
                return rep
            if "word/settings.xml" in names:
                s = z.read("word/settings.xml")
                if b"documentProtection" in s and b"enforcement=\"1\"" in s.replace(b"'", b'"'):
                    rep.note("document protection enabled")

            headers = sorted(n for n in names if re.match(r"word/header\d+\.xml$", n))
            footers = sorted(n for n in names if re.match(r"word/footer\d+\.xml$", n))
            rep["n_headers"] = len(headers)
            rep["n_footers"] = len(footers)

            # ---- locate the logo -------------------------------------------------
            hdr_imgs, body_imgs = [], []
            for h in headers:
                for im in images_in(z, h):
                    im["part"] = h
                    hdr_imgs.append(im)
            for im in images_in(z, "word/document.xml"):
                im["part"] = "word/document.xml"
                body_imgs.append(im)

            hdr_targets = {im["target"] for im in hdr_imgs if im["target"]}
            body_targets = {im["target"] for im in body_imgs if im["target"]}

            rep["n_header_images"] = len(hdr_imgs)
            rep["n_body_images"] = len(body_imgs)
            rep["logo_kinds"] = ",".join(sorted({im["kind"] for im in hdr_imgs})) or ""

            if not hdr_targets:
                if body_targets:
                    rep.block("no image in header (logo appears to be in the body)")
                else:
                    rep.block("no image found in header")
                logo = None
            elif len(hdr_targets) > 1:
                rep.block(f"{len(hdr_targets)} distinct images in header - ambiguous which is the logo")
                logo = None
            else:
                logo = next(iter(hdr_targets))

            rep["logo_part"] = logo or ""
            rep["logo_shared_with_body"] = "YES" if logo and logo in body_targets else "NO"
            if logo and logo in body_targets:
                rep.block("logo media part is also used in the body - overwriting bytes would alter a body image")

            # ---- geometry / distortion ------------------------------------------
            if logo:
                try:
                    blob = z.read(logo)
                except KeyError:
                    blob = b""
                    rep.block(f"media part missing from archive: {logo}")
                dims = image_size(blob) if blob else None
                if dims:
                    rep["img_px"] = f"{dims[0]}x{dims[1]}"
                    rep["img_aspect"] = round(dims[0] / dims[1], 4)
                else:
                    rep["img_px"] = ""
                    rep["img_aspect"] = ""
                    rep.note("could not read image dimensions (EMF/WMF or unsupported format)")

                boxes = [(im["cx"], im["cy"]) for im in hdr_imgs
                         if im["target"] == logo and im["cx"] and im["cy"]]
                rep["n_logo_placements"] = len(boxes)
                if not boxes:
                    if rep["logo_kinds"] == "vml":
                        rep.block("logo uses legacy VML markup - needs separate handling")
                    else:
                        rep.block("no wp:extent found for the logo")
                else:
                    uniq = set(boxes)
                    rep["extent"] = ";".join(f"{cx}x{cy}" for cx, cy in sorted(uniq))
                    if len(uniq) > 1:
                        rep.note(f"{len(uniq)} different logo box sizes across headers")
                    cx, cy = boxes[0]
                    box_aspect = cx / cy
                    rep["box_aspect"] = round(box_aspect, 4)
                    if dims:
                        dev = abs(box_aspect - dims[0] / dims[1]) / (dims[0] / dims[1])
                        rep["aspect_dev_pct"] = round(dev * 100, 1)
                        rep["distorted"] = "YES" if dev > ASPECT_TOLERANCE else "NO"

            # ---- metadata table ---------------------------------------------------
            meta, meta_part = {}, ""
            for part in headers + ["word/document.xml"] + footers:
                try:
                    r = ET.fromstring(z.read(part))
                except (KeyError, ET.ParseError):
                    continue
                for _tbl, rows in parse_tables(r):
                    got, fuzzy = label_lookup_flags(rows)
                    if "bearbeiter" in got and "freigeber" in got:
                        meta, meta_part = got, part
                        if fuzzy:
                            rep["label_anomalies"] = "; ".join(fuzzy)
                            rep.note("label cell contains stray text: " + "; ".join(fuzzy))
                        break
                if meta:
                    break

            # count every metadata table, not just the first
            n_meta, doks = 0, set()
            for part in headers + ["word/document.xml"] + footers:
                try:
                    r2 = ET.fromstring(z.read(part))
                except (KeyError, ET.ParseError):
                    continue
                for _t, rws in parse_tables(r2):
                    g = label_lookup(rws)
                    if "bearbeiter" in g and "freigeber" in g:
                        n_meta += 1
                        if g.get("dok_nr"):
                            doks.add(g["dok_nr"].split("/")[0].strip())
            rep["n_meta_tables"] = n_meta
            if len(doks) > 1:
                rep.note("conflicting Dok-Nr across metadata tables: " +
                         ", ".join(sorted(doks)))

            rep["meta_part"] = meta_part
            for key in ("dok_nr", "bearbeiter", "freigeber", "pruefer", "freigabedatum"):
                rep[key] = meta.get(key, "")
            if not meta:
                rep.block("metadata table not found (no Bearbeiter/Freigeber labels)")
            else:
                for required in ("bearbeiter", "freigeber", "pruefer", "freigabedatum"):
                    if required not in meta:
                        rep.block(f"metadata table missing '{required}' column")

            # ---- date content control --------------------------------------------
            rep["date_is_sdt"] = "NO"
            rep["sdt_fulldate"] = ""
            if meta_part:
                try:
                    raw = z.read(meta_part).decode("utf-8", "replace")
                except KeyError:
                    raw = ""
                fd = re.findall(r'w:fullDate="([^"]+)"', raw)
                if fd:
                    rep["date_is_sdt"] = "YES"
                    rep["sdt_fulldate"] = ";".join(sorted(set(fd)))
                    rep.note("Freigabedatum is a date content control - w:fullDate must be synced too")
                # stray whitespace-only runs (the '10 spaces' defect)
                spacers = re.findall(r'<w:r>(?:(?!</w:r>).)*?<w:t xml:space="preserve">(\s+)</w:t>',
                                     raw, re.S)
                rep["spacer_runs"] = len(spacers)
                if spacers:
                    rep.note(f"{len(spacers)} whitespace-only run(s) in the metadata part")

            # ---- changelog --------------------------------------------------------
            rep["changelog_rows"] = 0
            rep["changelog_last"] = ""
            rep["changelog_action"] = ""
            try:
                droot = ET.fromstring(z.read("word/document.xml"))
                for _tbl, rows in parse_tables(droot):
                    vrows = [r for r in rows if r and r[0].strip().lower().startswith("version")]
                    if vrows:
                        rep["changelog_rows"] = len(vrows)
                        rep["changelog_last"] = " | ".join(x.strip() for x in vrows[-1])[:80]
                        break
            except (KeyError, ET.ParseError) as e:
                rep.block(f"cannot parse document.xml: {e.__class__.__name__}")
            if not rep["changelog_rows"]:
                rep["changelog_action"] = "create"
                rep.note("no changelog table - one will be seeded from the current version/date")
            else:
                rep["changelog_action"] = "extend"

            # ---- fingerprint ------------------------------------------------------
            rep["fingerprint"] = "|".join([
                f"h{rep.get('n_headers', 0)}",
                f"img{rep.get('n_header_images', 0)}",
                rep.get("logo_kinds", "") or "noimg",
                "meta:" + (os.path.basename(meta_part) if meta_part else "none"),
                "sdt" if rep["date_is_sdt"] == "YES" else "plain",
                "cl" if rep["changelog_rows"] else "nocl",
            ])

    except zipfile.BadZipFile:
        rep["status"] = "ERROR"
        rep["error"] = "not a valid zip (legacy .doc, corrupt, or password-protected)"
        return rep
    except Exception as e:
        rep["status"] = "ERROR"
        rep["error"] = f"{e.__class__.__name__}: {e}"
        rep["traceback"] = traceback.format_exc()
        return rep

    if rep["blockers"]:
        rep["status"] = "NEEDS_REVIEW"
    return rep


# --------------------------------------------------------------------------
# driver
# --------------------------------------------------------------------------

COLUMNS = [
    "file", "status", "would_patch", "blocking_reasons", "notes", "error",
    "fingerprint", "n_headers", "n_footers", "n_header_images", "n_body_images",
    "logo_part", "logo_kinds", "logo_shared_with_body", "n_logo_placements",
    "img_px", "img_aspect", "extent", "box_aspect", "aspect_dev_pct", "distorted",
    "meta_part", "n_meta_tables", "label_anomalies", "dok_nr", "bearbeiter", "freigeber", "pruefer", "freigabedatum",
    "date_is_sdt", "sdt_fulldate", "spacer_runs",
    "changelog_rows", "changelog_action", "changelog_last",
]


def main(argv=None):
    ap = argparse.ArgumentParser(description="Read-only preflight scan for the docx rebranding batch.")
    ap.add_argument("--input", required=True, help="folder to scan (recursively)")
    ap.add_argument("--out", default="inventory.csv", help="CSV report path")
    ap.add_argument("--errors", default="errors.log", help="log of files needing attention")
    ap.add_argument("--shapes", action="store_true", help="also print a fingerprint summary")
    args = ap.parse_args(argv)

    root = Path(args.input).resolve()
    if not root.is_dir():
        print(f"error: {root} is not a directory", file=sys.stderr)
        return 2

    files = sorted(p for p in root.rglob("*") if p.suffix.lower() == ".docx" and p.is_file())
    if not files:
        print(f"No .docx files found under {root}", file=sys.stderr)
        return 1

    reports = []
    for i, p in enumerate(files, 1):
        try:
            rep = inspect(p, root)
        except Exception as e:  # last-resort guard: one bad file must never stop the run
            rep = Report(p, root)
            rep["status"] = "ERROR"
            rep["error"] = f"uncaught {e.__class__.__name__}: {e}"
            rep["traceback"] = traceback.format_exc()
        rep["would_patch"] = "YES" if rep["status"] == "OK" else "NO"
        rep["blocking_reasons"] = "; ".join(rep["blockers"])
        rep["notes"] = "; ".join(rep["notes"])
        reports.append(rep)
        print(f"\r  scanned {i}/{len(files)}", end="", file=sys.stderr, flush=True)
    print("", file=sys.stderr)

    with open(args.out, "w", newline="", encoding="utf-8-sig") as fh:
        wtr = csv.DictWriter(fh, fieldnames=COLUMNS, extrasaction="ignore")
        wtr.writeheader()
        for rep in reports:
            wtr.writerow(rep)

    problems = [r for r in reports if r["would_patch"] == "NO"]
    with open(args.errors, "w", encoding="utf-8") as fh:
        for r in problems:
            fh.write(f"{r['file']}\n    status : {r['status']}\n")
            if r["error"]:
                fh.write(f"    error  : {r['error']}\n")
            for b in r["blockers"]:
                fh.write(f"    blocker: {b}\n")
            if r.get("traceback"):
                fh.write("    " + r["traceback"].replace("\n", "\n    ") + "\n")
            fh.write("\n")

    ok = sum(1 for r in reports if r["would_patch"] == "YES")
    print(f"\nScanned {len(reports)} file(s)")
    print(f"  patchable      : {ok}")
    print(f"  needs review   : {sum(1 for r in reports if r['status'] == 'NEEDS_REVIEW')}")
    print(f"  errors/skipped : {sum(1 for r in reports if r['status'] in ('ERROR', 'SKIP'))}")
    print(f"  distorted logo : {sum(1 for r in reports if r.get('distorted') == 'YES')}")
    print(f"  date is w:sdt  : {sum(1 for r in reports if r.get('date_is_sdt') == 'YES')}")
    print(f"\nreport -> {args.out}")
    if problems:
        print(f"files needing attention -> {args.errors}")

    if args.shapes:
        print("\nStructural shapes:")
        for fp, n in Counter(r.get("fingerprint", "?") for r in reports).most_common():
            print(f"  {n:5d}  {fp}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
