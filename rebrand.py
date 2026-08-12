#!/usr/bin/env python3
"""
rebrand.py - batch-apply a document rebrand to a tree of .docx files.

For every file it will:
  1. run the inventory preflight; anything not patchable is logged and skipped
  2. add the new logo as a fresh media part and repoint the header relationships
     (never overwrites shared image bytes, so body screenshots are untouched)
  3. resize the logo box to the image's true aspect ratio, keeping the old right
     edge and centring it in the vertical band the old logo occupied
  4. update Bearbeiter / Freigeber / Pruefer / Freigabedatum and bump the version
  5. sync the w:fullDate of a date content control with the visible text
  6. delete whitespace-only "spacer" runs in the metadata cells
  7. append a changelog row with freshly generated w14:paraId values
  8. re-open the result and verify it before keeping it

Originals are opened read-only and never modified; output goes to a separate
tree. If any step fails the partial file is discarded and the name is logged.

    python3 rebrand.py --input ./docs --output ./docs_rebranded \\
        --logo new-logo.jpg \\
        --bearbeiter "Editor Name" --freigeber "Approver Name" --pruefer "Reviewer Name" \\
        --date 30.07.2026 --changelog "Anpassung Rebranding"

    python3 rebrand.py --input ./docs --output ./out --logo l.jpg --dry-run
"""

from __future__ import annotations

import argparse
import copy
import csv
import os
import random
import re
import shutil
import sys
import tempfile
import traceback
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

import inventory as inv
from inventory import W, R, A, WP, image_size, norm_label, LABELS

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
IMG_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

EMU_PER_DXA = 635


# --------------------------------------------------------------------------
# XML load / save that does not disturb Word's namespace declarations
# --------------------------------------------------------------------------

def _register(blob: bytes):
    """Register every prefix a part declares, including the default namespace."""
    head = blob[:4000].decode("utf-8", "replace")
    for prefix, uri in re.findall(r'xmlns:([\w\-]+)="([^"]+)"', head):
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass
    m = re.search(r'<[A-Za-z_][\w\-.:]*[^>]*?\sxmlns="([^"]+)"', head)
    if m:
        try:
            ET.register_namespace("", m.group(1))
        except ValueError:
            pass


def load_xml(blob: bytes):
    """Parse a part, first registering every prefix it declares."""
    _register(blob)
    return ET.fromstring(blob)


def dump_xml(root, original: bytes) -> bytes:
    """
    Serialise, then restore any xmlns declarations ElementTree dropped.
    Word's mc:Ignorable lists prefixes that may appear nowhere else in the
    part; if their declarations vanish the file triggers a repair prompt.
    Namespaces are re-registered here because register_namespace is global
    state and another part may have overwritten it since load time.
    """
    _register(original)
    body = ET.tostring(root, encoding="unicode")

    def start_tag(text: str) -> str:
        m = re.search(r"<[A-Za-z_][\w\-.:]*(?:\s[^>]*?)?/?>", text, re.S)
        return m.group(0) if m else ""

    orig_tag = start_tag(original.decode("utf-8", "replace"))
    new_tag = start_tag(body)
    if orig_tag and new_tag:
        have = set(re.findall(r'xmlns:([\w\-]+)=', new_tag))
        missing = [f'xmlns:{p}="{u}"' for p, u in
                   re.findall(r'xmlns:([\w\-]+)="([^"]+)"', orig_tag) if p not in have]
        if missing:
            patched = new_tag[:-1].rstrip()
            if patched.endswith("/"):
                patched = patched[:-1].rstrip()
                closing = "/>"
            else:
                closing = ">"
            patched = patched + " " + " ".join(missing) + closing
            body = body.replace(new_tag, patched, 1)

    decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    return decl + body.encode("utf-8")


# --------------------------------------------------------------------------
# element-level table helpers (see through w:sdt, honour gridSpan)
# --------------------------------------------------------------------------

def parent_map(root):
    return {c: p for p in root.iter() for c in p}


def table_grid(tbl):
    """Rows of w:tc elements, expanded by gridSpan so columns line up."""
    rows = []
    for tr in inv._direct(tbl, f"{{{W}}}tr"):
        row = []
        for tc in inv._direct(tr, f"{{{W}}}tc"):
            row.append(tc)
            row.extend([tc] * (inv.grid_span(tc) - 1))
        rows.append((tr, row))
    return rows


def texts_in(tc):
    """w:t elements belonging to this cell, excluding nested tables."""
    out = []

    def walk(n):
        for child in n:
            if child.tag == f"{{{W}}}tbl":
                continue
            if child.tag == f"{{{W}}}t":
                out.append(child)
            walk(child)

    walk(tc)
    return out


def drop_spacer_runs(tc) -> int:
    """Remove runs whose entire text is whitespace (the hand-typed indents)."""
    pm = parent_map(tc)
    removed = 0
    for run in list(tc.iter(f"{{{W}}}r")):
        ts = [t for t in run.iter(f"{{{W}}}t")]
        if not ts:
            continue
        joined = "".join(t.text or "" for t in ts)
        if joined and joined.strip() == "":
            parent = pm.get(run)
            if parent is not None:
                parent.remove(run)
                removed += 1
    return removed


def set_cell_text(tc, value: str):
    """
    Replace a cell's visible text with `value`, preserving the first run's
    formatting. If the cell holds a content control the text inside the
    control is what gets written, so the control stays intact.
    """
    drop_spacer_runs(tc)

    sdt_targets = []
    for sdt in tc.iter(f"{{{W}}}sdt"):
        for content in sdt.iter(f"{{{W}}}sdtContent"):
            sdt_targets.extend(t for t in content.iter(f"{{{W}}}t"))
    targets = sdt_targets or texts_in(tc)

    if not targets:  # empty cell - build a minimal run
        paras = [p for p in tc.iter(f"{{{W}}}p")]
        para = paras[0] if paras else ET.SubElement(tc, f"{{{W}}}p")
        run = ET.SubElement(para, f"{{{W}}}r")
        t = ET.SubElement(run, f"{{{W}}}t")
        t.text = value
        return

    targets[0].text = value
    if value != value.strip():
        targets[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in targets[1:]:
        extra.text = ""


def sync_sdt_date(tc, iso_date: str) -> bool:
    """Point a date content control's stored value at the new date."""
    changed = False
    for d in tc.iter(f"{{{W}}}date"):
        d.set(f"{{{W}}}fullDate", iso_date)
        changed = True
    return changed


def find_meta_tables(root):
    """
    Every metadata table in a part. A document can carry one per header - a
    first-page header and a continuation header each have their own - so
    stopping at the first match leaves the others stale.
    """
    out = []
    for tbl in root.iter(f"{{{W}}}tbl"):
        rows = table_grid(tbl)
        cells = {}
        for i, (_tr, row) in enumerate(rows):
            for j, tc in enumerate(row):
                key, _exact = inv.match_label(inv.cell_text(tc))
                if key and i + 1 < len(rows) and j < len(rows[i + 1][1]):
                    cells.setdefault(key, rows[i + 1][1][j])
        if "bearbeiter" in cells and "freigeber" in cells:
            out.append((tbl, cells))
    return out


def find_meta_table(root):
    """First metadata table only - kept for the verification pass."""
    found = find_meta_tables(root)
    return found[0] if found else (None, {})


def meta_part_order(parts, headers):
    """
    Order the parts by what the reader actually sees: the first-page header
    before the default header before the even-page header. Filename order
    (header1, header2, header3) has no relationship to display precedence.
    """
    rels, order = {}, []
    try:
        rroot = ET.fromstring(parts["word/_rels/document.xml.rels"])
        ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        for rel in rroot.findall(f"{ns}Relationship"):
            rels[rel.get("Id")] = "word/" + rel.get("Target", "").lstrip("/")
    except (KeyError, ET.ParseError):
        pass

    try:
        droot = ET.fromstring(parts["word/document.xml"])
        for sect in droot.iter(f"{{{W}}}sectPr"):
            refs = {}
            for ref in sect.findall(f"{{{W}}}headerReference"):
                refs[ref.get(f"{{{W}}}type")] = rels.get(ref.get(f"{{{R}}}id"))
            for kind in ("first", "default", "even"):
                target = refs.get(kind)
                if target and target not in order:
                    order.append(target)
    except (KeyError, ET.ParseError):
        pass

    for h in headers:
        if h not in order:
            order.append(h)
    order.append("word/document.xml")
    order.extend(sorted(n for n in parts if re.match(r"word/footer\d+\.xml$", n)))
    return order


def apply_meta(cells, opts, iso_date) -> dict:
    """Write the new values into one metadata table; report what it held."""
    info = {"fields": 0, "sdt": False, "old_version": "", "old_date": "",
            "new_version": "", "dok_nr": ""}

    if "bearbeiter" in cells:
        set_cell_text(cells["bearbeiter"], opts.bearbeiter)
        info["fields"] += 1
    if "freigeber" in cells:
        set_cell_text(cells["freigeber"], opts.freigeber)
        info["fields"] += 1
    if "pruefer" in cells:
        set_cell_text(cells["pruefer"], opts.pruefer)
        info["fields"] += 1
    if "freigabedatum" in cells:
        cell = cells["freigabedatum"]
        info["old_date"] = inv.cell_text(cell).strip()
        info["sdt"] = sync_sdt_date(cell, iso_date)
        set_cell_text(cell, opts.date)
        info["fields"] += 1
    if "dok_nr" in cells:
        current = inv.cell_text(cells["dok_nr"]).strip()
        info["dok_nr"] = current
        m_old = re.search(r"(\d+\.\d+)\s*$", current)
        info["old_version"] = m_old.group(1) if m_old else ""
        bumped, version = bump_version(current, opts.version_mode)
        if version:
            set_cell_text(cells["dok_nr"], bumped)
            info["new_version"] = version
            info["fields"] += 1
    return info


STRIP_TAGS = {
    f"{{{W}}}bookmarkStart", f"{{{W}}}bookmarkEnd",
    f"{{{W}}}commentRangeStart", f"{{{W}}}commentRangeEnd",
    f"{{{W}}}permStart", f"{{{W}}}permEnd",
}


def strip_transient(el):
    """
    Remove bookmarks, comment anchors and permission ranges from a cloned row.
    Their ids must be unique document-wide, so copying them verbatim produces
    duplicates that fail validation and make Word offer to repair the file.
    """
    pm = parent_map(el)
    for node in list(el.iter()):
        if node.tag in STRIP_TAGS:
            parent = pm.get(node)
            if parent is not None:
                parent.remove(node)

    pm = parent_map(el)
    for run in list(el.iter(f"{{{W}}}r")):
        if any(child.tag == f"{{{W}}}commentReference" for child in run):
            parent = pm.get(run)
            if parent is not None:
                parent.remove(run)


def new_para_id() -> str:
    return f"{random.randint(1, 0x7FFFFFFE):08X}"


def append_changelog_row(root, version: str, date: str, text: str):
    """Clone the last 'Version ...' row, refresh its ids, and append it."""
    for tbl in root.iter(f"{{{W}}}tbl"):
        rows = table_grid(tbl)
        idx = [i for i, (_tr, row) in enumerate(rows)
               if row and inv.cell_text(row[0]).strip().lower().startswith("version")]
        if not idx:
            continue

        last_tr, last_cells = rows[idx[-1]]
        new_tr = copy.deepcopy(last_tr)
        strip_transient(new_tr)

        for el in new_tr.iter():
            if el.get(f"{{{W14}}}paraId") is not None:
                el.set(f"{{{W14}}}paraId", new_para_id())
            if el.get(f"{{{W14}}}textId") is not None:
                el.set(f"{{{W14}}}textId", new_para_id())

        cells = inv._direct(new_tr, f"{{{W}}}tc")
        if len(cells) < 2:
            return False

        # mirror the existing "Version x.y<spaces>date" layout of column 1
        sample = inv.cell_text(last_cells[0])
        m = re.match(r"(Version\s+\S+)(\s+)(\S+)", sample)
        gap = m.group(2) if m else "  "
        set_cell_text(cells[0], f"Version {version}{gap}{date}")
        set_cell_text(cells[1], text)
        for spare in cells[2:]:
            set_cell_text(spare, "")

        tbl.append(new_tr)
        return True
    return False


VERSION_CELL = re.compile(r"^v?\s*\d+\.\d+\s*$", re.I)


def looks_like_changelog(rows) -> bool:
    """
    Loose detector: does this table resemble a changelog even if it does not
    match the strict 'first cell starts with Version' layout? Used only to
    avoid synthesising a second changelog next to one we failed to parse.
    """
    for row in rows:
        for cell in row:
            t = (cell or "").strip()
            if not t:
                continue
            if VERSION_CELL.match(t) or t.lower().startswith("version "):
                return True
            if t.lower() in ("neuerstellung", "änderung", "aenderung"):
                return True
    return False


def body_content_width(droot) -> int:
    """Printable width in DXA, from the body-level section properties."""
    sects = list(droot.iter(f"{{{W}}}sectPr"))
    for sect in reversed(sects):
        pg = sect.find(f"{{{W}}}pgSz")
        mar = sect.find(f"{{{W}}}pgMar")
        if pg is None or mar is None:
            continue
        try:
            width = int(pg.get(f"{{{W}}}w"))
            left = int(mar.get(f"{{{W}}}left"))
            right = int(mar.get(f"{{{W}}}right"))
            if width > 0:
                return max(2000, width - left - right)
        except (TypeError, ValueError):
            continue
    return 9062  # A4 with 1417 dxa margins, the shape the existing tables use


def changelog_line(version: str, date: str) -> str:
    """'Version 1.0' padded so the date starts where the existing tables put it."""
    label = f"Version {version}"
    return label + " " * max(1, 33 - len(label)) + date


def _el(parent, tag, attrs=None):
    e = ET.SubElement(parent, f"{{{W}}}{tag}")
    for k, v in (attrs or {}).items():
        e.set(f"{{{W}}}{k}" if not k.startswith("{") else k, v)
    return e


TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
    "tblCaption", "tblDescription",
]


def insert_ordered(parent, child):
    """Insert into w:tblPr at the position the schema requires."""
    tag = child.tag.split("}")[-1]
    try:
        rank = TBLPR_ORDER.index(tag)
    except ValueError:
        parent.append(child)
        return
    for i, existing in enumerate(parent):
        etag = existing.tag.split("}")[-1]
        erank = TBLPR_ORDER.index(etag) if etag in TBLPR_ORDER else len(TBLPR_ORDER)
        if erank > rank:
            parent.insert(i, child)
            return
    parent.append(child)


def canonical(el) -> str:
    """Stable string for an element, ignoring Word's revision-tracking noise."""
    s = ET.tostring(el, encoding="unicode")
    s = re.sub(r'\s+w:rsid\w*="[^"]*"', "", s)
    return re.sub(r"\s+", " ", s).strip()


def changelog_template(files):
    """
    Scan the corpus for changelog tables that already exist and return the most
    common shape, so synthesised tables match the estate rather than a guess.
    Returns {'pr': Element|None, 'props': [...], 'votes': n, 'seen': n, 'desc': str}.
    """
    counter = Counter()
    samples = {}
    seen = 0

    for path in files:
        try:
            with zipfile.ZipFile(path, "r") as z:
                root = ET.fromstring(z.read("word/document.xml"))
        except Exception:
            continue
        for tbl, rows in inv.parse_tables(root):
            if not any(r and r[0].strip().lower().startswith("version") for r in rows):
                continue
            pr = tbl.find(f"{{{W}}}tblPr")
            grid = tbl.find(f"{{{W}}}tblGrid")
            if pr is None or grid is None:
                continue
            cols = []
            for g in grid.findall(f"{{{W}}}gridCol"):
                try:
                    cols.append(int(g.get(f"{{{W}}}w", "0")))
                except (TypeError, ValueError):
                    cols = []
                    break
            if len(cols) < 2 or sum(cols) <= 0:
                continue
            props = tuple(round(c / sum(cols), 4) for c in cols)
            key = (canonical(pr), props)
            counter[key] += 1
            samples.setdefault(key, copy.deepcopy(pr))
            seen += 1
            break

    if not counter:
        return {"pr": None, "props": [0.5, 0.5], "votes": 0, "seen": 0,
                "desc": "no existing changelog found - using a standard bordered table"}

    (pr_key, props), votes = counter.most_common(1)[0]
    pr = samples[(pr_key, props)]
    style = pr.find(f"{{{W}}}tblStyle")
    desc = (f"{len(props)} cols "
            f"{'/'.join(str(int(round(p * 100))) for p in props)}, "
            f"style={style.get(f'{{{W}}}val') if style is not None else 'none'}, "
            f"borders={'explicit' if pr.find(f'{{{W}}}tblBorders') is not None else 'from style'}")
    return {"pr": pr, "props": list(props), "votes": votes, "seen": seen, "desc": desc}


def ensure_borders(pr, styles_blob: bytes) -> str:
    """
    Guarantee the table draws a grid. If the template carries explicit borders,
    or names a table style that this document actually defines, leave it alone;
    otherwise write standard single-line borders so the table is never invisible.
    """
    if pr.find(f"{{{W}}}tblBorders") is not None:
        return "template"
    style = pr.find(f"{{{W}}}tblStyle")
    if style is not None:
        sid = style.get(f"{{{W}}}val")
        if sid and f'w:styleId="{sid}"'.encode() in styles_blob:
            return "style"
        if sid:
            pr.remove(style)  # style is missing here, so it would draw nothing
    borders = ET.Element(f"{{{W}}}tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = ET.SubElement(borders, f"{{{W}}}{side}")
        el.set(f"{{{W}}}val", "single")
        el.set(f"{{{W}}}sz", "4")
        el.set(f"{{{W}}}space", "0")
        el.set(f"{{{W}}}color", "auto")
    insert_ordered(pr, borders)
    return "explicit"


def build_changelog_table(rows, width_dxa: int, template, styles_blob: bytes):
    """
    Build a changelog table matching the estate's most common changelog shape,
    scaled to this document's own printable width so it spans the full text
    column, and guaranteed to have visible borders.
    """
    props = template.get("props") or [0.5, 0.5]
    cols = [int(round(width_dxa * p)) for p in props]
    cols[-1] = width_dxa - sum(cols[:-1])  # absorb rounding into the last column

    tbl = ET.Element(f"{{{W}}}tbl")
    if template.get("pr") is not None:
        pr = copy.deepcopy(template["pr"])
        tbl.append(pr)
        source = "corpus"
    else:
        pr = _el(tbl, "tblPr")
        _el(pr, "tblW", {"w": "0", "type": "auto"})
        _el(pr, "tblLook", {"val": "04A0", "firstRow": "1", "lastRow": "0",
                            "firstColumn": "1", "lastColumn": "0",
                            "noHBand": "0", "noVBand": "1"})
        source = "standard"
    border_src = ensure_borders(pr, styles_blob)

    grid = _el(tbl, "tblGrid")
    for c in cols:
        _el(grid, "gridCol", {"w": str(c)})

    texts = list(rows)
    for row_texts in texts:
        tr = _el(tbl, "tr")
        tr.set(f"{{{W14}}}paraId", new_para_id())
        tr.set(f"{{{W14}}}textId", new_para_id())
        padded = list(row_texts) + [""] * (len(cols) - len(row_texts))
        for text, cw in zip(padded, cols):
            tc = _el(tr, "tc")
            tcpr = _el(tc, "tcPr")
            _el(tcpr, "tcW", {"w": str(cw), "type": "dxa"})
            p = _el(tc, "p")
            p.set(f"{{{W14}}}paraId", new_para_id())
            p.set(f"{{{W14}}}textId", new_para_id())
            run = _el(p, "r")
            t = _el(run, "t")
            t.text = text
            if text != text.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return tbl, f"{source}/{border_src}"


def insert_changelog_table(droot, tbl) -> bool:
    """Place the table at the end of the body, before the section properties."""
    body = droot.find(f"{{{W}}}body")
    if body is None:
        return False
    children = list(body)
    sect_idx = len(children)
    for i, child in enumerate(children):
        if child.tag == f"{{{W}}}sectPr":
            sect_idx = i
            break

    spacer = ET.Element(f"{{{W}}}p")
    spacer.set(f"{{{W14}}}paraId", new_para_id())
    spacer.set(f"{{{W14}}}textId", new_para_id())
    trailing = ET.Element(f"{{{W}}}p")
    trailing.set(f"{{{W14}}}paraId", new_para_id())
    trailing.set(f"{{{W14}}}textId", new_para_id())

    # a table must be followed by a paragraph when it ends the body
    for offset, node in enumerate((spacer, tbl, trailing)):
        body.insert(sect_idx + offset, node)
    return True


def bump_version(dok_nr: str, mode: str) -> tuple[str, str]:
    """'20103 / 1.0' -> ('20103 / 2.0', '2.0'). Falls back safely."""
    m = re.match(r"^(.*?)(\d+)\.(\d+)\s*$", dok_nr.strip())
    if not m:
        return dok_nr, ""
    prefix, major, minor = m.group(1), int(m.group(2)), int(m.group(3))
    if mode == "minor":
        major, minor = major, minor + 1
    else:
        major, minor = major + 1, 0
    version = f"{major}.{minor}"
    return f"{prefix}{version}", version


# --------------------------------------------------------------------------
# the patch itself
# --------------------------------------------------------------------------

class PatchError(Exception):
    pass


def patch_file(src: Path, dst: Path, logo_bytes: bytes, logo_dims, opts) -> dict:
    """Patch one file. Raises PatchError on anything unexpected."""
    result = {"logo_parts": 0, "meta_fields": 0, "media_mode": "",
              "changelog": "no", "version": "", "sdt_date": "no",
              "old_version": "", "old_date": "", "meta_tables": 0,
              "warning": ""}

    lw, lh = logo_dims
    aspect = lw / lh
    ext = os.path.splitext(opts.logo)[1].lower().lstrip(".") or "jpg"
    if ext == "jpg":
        ext = "jpeg"
    new_media = f"word/media/logo_rebrand.{ext}"

    with zipfile.ZipFile(src, "r") as z:
        names = z.namelist()
        parts = {n: z.read(n) for n in names}
        infos = {i.filename: i for i in z.infolist()}

    headers = sorted(n for n in parts if re.match(r"word/header\d+\.xml$", n))

    # --- which media part is the logo -----------------------------------
    hdr_imgs = []
    with zipfile.ZipFile(src, "r") as z:
        for h in headers:
            for im in inv.images_in(z, h):
                im["part"] = h
                hdr_imgs.append(im)
    targets = {im["target"] for im in hdr_imgs if im["target"]}
    if len(targets) != 1:
        raise PatchError(f"expected exactly one header image, found {len(targets)}")
    old_media = next(iter(targets))

    # --- swap the logo bytes --------------------------------------------
    # Preferred path: overwrite the existing media part in place. That needs no
    # relationship or content-type changes at all, so there is far less to break.
    # Only when the new logo's format differs from the old part's do we add a
    # new part and repoint the headers at it.
    old_ext = os.path.splitext(old_media)[1].lower().lstrip(".")
    same_format = {old_ext, ext} <= {"jpeg", "jpg"} or old_ext == ext
    new_rids = {}

    if same_format:
        parts[old_media] = logo_bytes
        result["media_mode"] = "in-place"
    else:
        result["media_mode"] = "new-part"
        parts[new_media] = logo_bytes
        for h in headers:
            rels_name = f"word/_rels/{os.path.basename(h)}.rels"
            if rels_name not in parts:
                raise PatchError(f"missing relationships for {h}")
            rroot = load_xml(parts[rels_name])
            existing = rroot.findall(f"{{{PKG_REL}}}Relationship")
            used = {int(m.group(1)) for r in existing
                    if (m := re.match(r"rId(\d+)$", r.get("Id") or ""))}
            rid = f"rId{max(used) + 1 if used else 1}"
            ET.SubElement(rroot, f"{{{PKG_REL}}}Relationship", {
                "Id": rid, "Type": IMG_REL,
                "Target": new_media.replace("word/", "", 1),
            })
            # drop the relationship the old logo used, if nothing else needs it
            for rel in existing:
                tgt = "word/" + re.sub(r"^\.\./", "", rel.get("Target", "")).lstrip("/")
                if tgt == old_media:
                    rroot.remove(rel)
            parts[rels_name] = dump_xml(rroot, parts[rels_name])
            new_rids[h] = rid

        ct = load_xml(parts["[Content_Types].xml"])
        have = {d.get("Extension", "").lower() for d in ct.findall(f"{{{CT_NS}}}Default")}
        if ext not in have:
            mime = {"jpeg": "image/jpeg", "png": "image/png",
                    "gif": "image/gif"}.get(ext, "image/jpeg")
            ET.SubElement(ct, f"{{{CT_NS}}}Default",
                          {"Extension": ext, "ContentType": mime})
            parts["[Content_Types].xml"] = dump_xml(ct, parts["[Content_Types].xml"])

    # --- per-header: descr, relationship id, geometry --------------------
    for h in headers:
        hroot = load_xml(parts[h])
        old_rids = {im["rid"] for im in hdr_imgs
                    if im["part"] == h and im["target"] == old_media}

        if not same_format:
            touched = 0
            for blip in hroot.iter(f"{{{A}}}blip"):
                if blip.get(f"{{{R}}}embed") in old_rids:
                    blip.set(f"{{{R}}}embed", new_rids[h])
                    touched += 1
            if touched == 0 and old_rids:
                raise PatchError(f"could not repoint logo relationship in {h}")

        for cnv in hroot.iter(f"{{{PIC}}}cNvPr"):
            if cnv.get("descr"):
                cnv.set("descr", os.path.basename(opts.logo))

        # geometry: true aspect, same right edge, centred vertically
        for anchor_tag in (f"{{{WP}}}anchor", f"{{{WP}}}inline"):
            for node in hroot.iter(anchor_tag):
                extent = node.find(f"{{{WP}}}extent")
                if extent is None:
                    continue
                try:
                    old_cx = int(extent.get("cx"))
                    old_cy = int(extent.get("cy"))
                except (TypeError, ValueError):
                    continue

                new_cx = opts.logo_width_emu
                new_cy = round(new_cx / aspect)
                extent.set("cx", str(new_cx))
                extent.set("cy", str(new_cy))
                for a_ext in node.iter(f"{{{A}}}ext"):
                    # a:extLst also contains <a:ext uri="..."> elements, which
                    # take no cx/cy - only touch the sizing one inside a:xfrm
                    if a_ext.get("cx") is None:
                        continue
                    a_ext.set("cx", str(new_cx))
                    a_ext.set("cy", str(new_cy))

                if anchor_tag.endswith("anchor"):
                    ph = node.find(f"{{{WP}}}positionH")
                    if ph is not None:
                        off = ph.find(f"{{{WP}}}posOffset")
                        if off is not None and (off.text or "").strip().lstrip("-").isdigit():
                            right_edge = int(off.text) + old_cx
                            off.text = str(right_edge - new_cx)
                    pv = node.find(f"{{{WP}}}positionV")
                    if pv is not None:
                        off = pv.find(f"{{{WP}}}posOffset")
                        if off is not None and (off.text or "").strip().lstrip("-").isdigit():
                            centre = int(off.text) + old_cy / 2
                            off.text = str(round(centre - new_cy / 2))
                result["logo_parts"] += 1

        parts[h] = dump_xml(hroot, parts[h])

    if result["logo_parts"] == 0:
        raise PatchError("no logo placement was resized")

    if not same_format:
        still_used = any(
            re.search(rf'Target="[^"]*{re.escape(os.path.basename(old_media))}"',
                      v.decode("utf-8", "replace"))
            for k, v in parts.items() if k.endswith(".rels")
        )
        if not still_used:
            parts.pop(old_media, None)

    # --- metadata tables (one per header is normal) -----------------------
    iso = opts.iso_date
    patched = []
    for cand in meta_part_order(parts, headers):
        if cand not in parts:
            continue
        try:
            root = load_xml(parts[cand])
        except ET.ParseError:
            continue
        tables = find_meta_tables(root)
        if not tables:
            continue
        for _tbl, cells in tables:
            patched.append((cand, apply_meta(cells, opts, iso)))
        parts[cand] = dump_xml(root, parts[cand])

    if not patched:
        raise PatchError("metadata table not found")

    meta_part = patched[0][0]
    lead = patched[0][1]
    result["meta_fields"] = sum(info["fields"] for _p, info in patched)
    result["meta_tables"] = len(patched)
    result["sdt_date"] = "yes" if any(info["sdt"] for _p, info in patched) else "no"
    result["version"] = lead["new_version"]
    result["old_version"] = lead["old_version"]
    result["old_date"] = lead["old_date"]

    # A file carrying two metadata tables with different document numbers is
    # really two documents in one - patch both, but say so loudly.
    doks = {info["dok_nr"].split("/")[0].strip()
            for _p, info in patched if info["dok_nr"]}
    if len(doks) > 1:
        result["warning"] = ("conflicting Dok-Nr across " + str(len(patched)) +
                             " metadata tables: " + ", ".join(sorted(doks)))

    if not result["version"]:
        raise PatchError("could not parse a version number from the Dok.-Nr. cell")

    # --- changelog ------------------------------------------------------
    droot = load_xml(parts["word/document.xml"])
    if append_changelog_row(droot, result["version"], opts.date, opts.changelog):
        result["changelog"] = "extended"
        parts["word/document.xml"] = dump_xml(droot, parts["word/document.xml"])
    elif opts.no_synth_changelog:
        raise PatchError("no changelog table to extend")
    else:
        # Guard: only synthesise when nothing changelog-like exists anywhere.
        # If something is there that we could not parse, stop rather than
        # risk leaving the document with two changelogs.
        for part_name, blob in parts.items():
            if not re.match(r"word/(document|header\d+|footer\d+)\.xml$", part_name):
                continue
            try:
                probe = ET.fromstring(blob)
            except ET.ParseError:
                continue
            for _tbl, rws in inv.parse_tables(probe):
                if looks_like_changelog(rws):
                    raise PatchError(
                        f"a changelog-like table exists in {part_name} but its layout "
                        "was not recognised - needs manual review")

        old_version = result.get("old_version") or "1.0"
        old_date = result.get("old_date") or ""
        if not old_date:
            raise PatchError("cannot synthesise a changelog: no previous Freigabedatum found")

        tbl, source = build_changelog_table(
            [
                (changelog_line(old_version, old_date), opts.initial_changelog),
                (changelog_line(result["version"], opts.date), opts.changelog),
            ],
            body_content_width(droot),
            opts.changelog_template,
            parts.get("word/styles.xml", b""),
        )
        if not insert_changelog_table(droot, tbl):
            raise PatchError("could not insert a changelog table (no w:body)")
        result["changelog"] = f"created({source})"
        parts["word/document.xml"] = dump_xml(droot, parts["word/document.xml"])

    # --- write atomically ------------------------------------------------
    dst.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".docx", dir=str(dst.parent))
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for name in names:
                if name not in parts:
                    continue
                info = infos.get(name)
                ct_ = info.compress_type if info else zipfile.ZIP_DEFLATED
                out.writestr(name, parts[name], compress_type=ct_)
            if new_media in parts and new_media not in names:
                out.writestr(new_media, parts[new_media], zipfile.ZIP_DEFLATED)

        verify(tmp, opts, result)
        shutil.move(tmp, dst)
    except Exception:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise

    return result


def verify(path: str, opts, result):
    """Re-open the freshly written file and confirm the edits actually took."""
    with zipfile.ZipFile(path, "r") as z:
        bad = z.testzip()
        if bad:
            raise PatchError(f"corrupt entry in output: {bad}")
        for name in z.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                try:
                    ET.fromstring(z.read(name))
                except ET.ParseError as e:
                    raise PatchError(f"malformed XML in {name}: {e}")

        checked = 0
        for part in sorted(n for n in z.namelist()
                           if re.match(r"word/(header\d+|footer\d+|document)\.xml$", n)):
            try:
                root = ET.fromstring(z.read(part))
            except ET.ParseError:
                continue
            for _tbl, cells in find_meta_tables(root):
                checks = {
                    "bearbeiter": opts.bearbeiter,
                    "freigeber": opts.freigeber,
                    "pruefer": opts.pruefer,
                    "freigabedatum": opts.date,
                }
                for key, want in checks.items():
                    if key in cells:
                        got = inv.cell_text(cells[key]).strip()
                        if got != want:
                            raise PatchError(
                                f"verify failed in {part}: {key} is {got!r}, "
                                f"expected {want!r}")
                checked += 1
        if checked == 0:
            raise PatchError("verify failed: no metadata table in the output")
        if checked != result.get("meta_tables", checked):
            raise PatchError(
                f"verify failed: patched {result['meta_tables']} metadata "
                f"table(s) but found {checked} in the output")

        droot = ET.fromstring(z.read("word/document.xml"))
        joined = " ".join(t.text or "" for t in droot.iter(f"{{{W}}}t"))
        if opts.changelog not in joined:
            raise PatchError("verify failed: changelog row missing from output")


# --------------------------------------------------------------------------
# driver
# --------------------------------------------------------------------------

COLUMNS = ["file", "status", "reason", "warning", "version", "meta_fields",
           "meta_tables", "logo_parts", "media_mode", "sdt_date", "changelog"]


def main(argv=None):
    ap = argparse.ArgumentParser(description="Batch-apply a document rebrand to .docx files.")
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--logo", required=True, help="new logo image (jpg/png)")
    ap.add_argument("--bearbeiter", required=True)
    ap.add_argument("--freigeber", required=True)
    ap.add_argument("--pruefer", required=True)
    ap.add_argument("--date", required=True, help="Freigabedatum as shown, e.g. 30.07.2026")
    ap.add_argument("--changelog", default="Anpassung Rebranding")
    ap.add_argument("--initial-changelog", default="Neuerstellung",
                    help="text for the seeded first row when a changelog has to be created")
    ap.add_argument("--no-synth-changelog", action="store_true",
                    help="skip files that have no changelog instead of creating one")
    ap.add_argument("--version-mode", choices=["major", "minor"], default="major")
    ap.add_argument("--logo-width-mm", type=float, default=47.0,
                    help="rendered logo width in mm (default 47)")
    ap.add_argument("--report", default="rebrand_report.csv")
    ap.add_argument("--errors", default="rebrand_errors.log")
    ap.add_argument("--dry-run", action="store_true", help="preflight only, write nothing")
    ap.add_argument("--limit", type=int, help="process at most N files (for a pilot run)")
    opts = ap.parse_args(argv)

    src_root = Path(opts.input).resolve()
    out_root = Path(opts.output).resolve()
    if not src_root.is_dir():
        print(f"error: {src_root} is not a directory", file=sys.stderr)
        return 2
    if out_root == src_root:
        print("error: --output must differ from --input", file=sys.stderr)
        return 2

    logo_path = Path(opts.logo)
    if not logo_path.is_file():
        print(f"error: logo not found: {logo_path}", file=sys.stderr)
        return 2
    logo_bytes = logo_path.read_bytes()
    dims = image_size(logo_bytes)
    if not dims:
        print("error: could not read the logo's dimensions", file=sys.stderr)
        return 2

    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", opts.date.strip())
    if not m:
        print("error: --date must look like DD.MM.YYYY", file=sys.stderr)
        return 2
    opts.iso_date = f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}T00:00:00Z"
    opts.logo_width_emu = int(round(opts.logo_width_mm * 36000))

    files = sorted(p for p in src_root.rglob("*")
                   if p.suffix.lower() == ".docx" and p.is_file()
                   and not p.name.startswith("~$"))
    if opts.limit:
        files = files[:opts.limit]
    if not files:
        print(f"No .docx files found under {src_root}", file=sys.stderr)
        return 1

    opts.changelog_template = {"pr": None, "props": [0.5, 0.5], "votes": 0,
                               "seen": 0, "desc": "changelog synthesis disabled"}
    if not opts.no_synth_changelog:
        opts.changelog_template = changelog_template(files)
        tpl = opts.changelog_template
        if tpl["seen"]:
            print(f"changelog template: {tpl['votes']}/{tpl['seen']} existing "
                  f"changelogs agree -> {tpl['desc']}", file=sys.stderr)
        else:
            print(f"changelog template: {tpl['desc']}", file=sys.stderr)

    rows, failures = [], []
    for i, src in enumerate(files, 1):
        rel = src.relative_to(src_root)
        row = {"file": str(rel), "status": "", "reason": "", "warning": "",
               "version": "", "meta_fields": "", "meta_tables": "",
               "logo_parts": "", "media_mode": "", "sdt_date": "", "changelog": ""}
        try:
            rep = inv.inspect(src, src_root)
            if rep["status"] != "OK":
                row["status"] = "SKIPPED"
                row["reason"] = rep["error"] or "; ".join(rep["blockers"])
            elif opts.dry_run:
                row["status"] = "WOULD_PATCH"
            else:
                res = patch_file(src, out_root / rel, logo_bytes, dims, opts)
                row.update(status="PATCHED", **{k: res[k] for k in
                           ("version", "meta_fields", "meta_tables", "logo_parts",
                            "media_mode", "sdt_date", "changelog", "warning")})
        except PatchError as e:
            row["status"] = "FAILED"
            row["reason"] = str(e)
        except Exception as e:
            row["status"] = "FAILED"
            row["reason"] = f"{e.__class__.__name__}: {e}"
            row["_tb"] = traceback.format_exc()
        rows.append(row)
        if row["status"] in ("FAILED", "SKIPPED"):
            failures.append(row)
        print(f"\r  {i}/{len(files)}  {row['status']:<12}", end="", file=sys.stderr, flush=True)
    print("", file=sys.stderr)

    with open(opts.report, "w", newline="", encoding="utf-8-sig") as fh:
        wtr = csv.DictWriter(fh, fieldnames=COLUMNS, extrasaction="ignore")
        wtr.writeheader()
        wtr.writerows(rows)

    with open(opts.errors, "w", encoding="utf-8") as fh:
        for r in failures:
            fh.write(f"{r['file']}\n    status: {r['status']}\n    reason: {r['reason']}\n")
            if r.get("_tb"):
                fh.write("    " + r["_tb"].replace("\n", "\n    ") + "\n")
            fh.write("\n")

    done = sum(1 for r in rows if r["status"] in ("PATCHED", "WOULD_PATCH"))
    print(f"\n{len(rows)} file(s)")
    print(f"  {'would patch' if opts.dry_run else 'patched':<12}: {done}")
    print(f"  skipped     : {sum(1 for r in rows if r['status'] == 'SKIPPED')}")
    print(f"  failed      : {sum(1 for r in rows if r['status'] == 'FAILED')}")
    print(f"\nreport -> {opts.report}")
    if failures:
        print(f"needs attention -> {opts.errors}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
