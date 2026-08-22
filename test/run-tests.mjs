// Real-JS regression harness: runs the exact shipped rebrand-core.mjs logic
// (via setup.mjs, which supplies Node equivalents of the browser globals) over
// synthetic fixtures + real sample documents. Run: node --test test/
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, existsSync } from "node:fs";
import * as core from "./setup.mjs";
import { unzipSync } from "./zip-shim.mjs";
import { makeParts, para, run, BOLD, companyOpts } from "./make-fixtures.mjs";

const NS = core.NS;
const dec = (u8) => new TextDecoder().decode(u8);
const docStr = (parts) => dec(parts["word/document.xml"]);

// Concatenate the text of the given element tags in document order.
function orderedText(root, tags) {
  let s = "";
  (function walk(n) {
    for (const c of n.childNodes || []) {
      if (c.nodeType !== 1) continue;
      if (c.namespaceURI === NS.w && tags.includes(c.localName)) s += c.textContent || "";
      else walk(c);
    }
  })(root);
  return s;
}
const parseDoc = (str) => core.parse(str).documentElement;
const visible = (str) => orderedText(parseDoc(str), ["t"]);            // what a reader sees
const accepted = (str) => {                                            // accept all revisions
  const d = core.parse(str);
  for (const del of core.els(d, NS.w, "del")) del.parentNode.removeChild(del);
  return orderedText(d.documentElement, ["t"]);
};
const rejected = (str) => {                                            // reject all revisions
  const d = core.parse(str);
  for (const ins of core.els(d, NS.w, "ins")) ins.parentNode.removeChild(ins);
  return orderedText(d.documentElement, ["t", "delText"]);
};
const assertAllPartsParse = (parts) => {
  for (const [n, b] of Object.entries(parts))
    if (n.endsWith(".xml") || n.endsWith(".rels")) core.parse(b);
};

// ---------------------------------------------------------------------------
// Company replacement — plain mode
// ---------------------------------------------------------------------------
test("plain: single-run, case-insensitive, exact target casing", () => {
  const parts = makeParts({ doc: para(run("Willkommen bei Altmarke und ALTMARKE heute")) });
  const res = core.replaceCompany(parts, companyOpts());
  assert.equal(res.count, 2);
  assert.equal(res.fallbacks.length, 0);
  assert.equal(visible(docStr(parts)), "Willkommen bei Neumarke und Neumarke heute");
});

test("plain: cross-run match collapses into one run", () => {
  const parts = makeParts({ doc: para(run("Va"), run("med rocks", BOLD)) });
  const res = core.replaceCompany(parts, companyOpts());
  assert.equal(res.count, 1);
  assert.equal(visible(docStr(parts)), "Neumarke rocks");
});

test("plain: regex-special 'from' is escaped, no accidental matches", () => {
  const parts = makeParts({ doc: para(run("A.B and AxB")) });
  const res = core.replaceCompany(parts, companyOpts({ companyFrom: "A.B", companyTo: "Z" }));
  assert.equal(res.count, 1);
  assert.equal(visible(docStr(parts)), "Z and AxB");
});

// ---------------------------------------------------------------------------
// Company replacement — tracked-changes mode
// ---------------------------------------------------------------------------
test("tracked: single-run emits well-formed ins/del, formatting preserved", () => {
  const parts = makeParts({ doc: para(run("Hallo Altmarke Welt", BOLD)) });
  const res = core.replaceCompany(parts, companyOpts({ companyTracked: true }));
  assert.equal(res.count, 1);
  assert.equal(res.fallbacks.length, 0);
  const s = docStr(parts);
  const d = core.parse(s);
  const del = core.els(d, NS.w, "del")[0], ins = core.els(d, NS.w, "ins")[0];
  assert.ok(del && ins, "has both del and ins");
  assert.equal(core.els(del, NS.w, "delText")[0].textContent, "Altmarke");
  assert.equal(core.els(ins, NS.w, "t")[0].textContent, "Neumarke");
  for (const el of [del, ins]) {
    assert.equal(core.attr(el, NS.w, "author"), "Chef");
    assert.equal(core.attr(el, NS.w, "date"), "2026-08-22T00:00:00Z");
    assert.ok(/^\d+$/.test(core.attr(el, NS.w, "id")), "numeric w:id");
  }
  // bold rPr carried onto the inserted run
  assert.ok(core.els(ins, NS.w, "b").length === 1, "inserted run keeps bold");
  assert.equal(accepted(s), "Hallo Neumarke Welt");
  assert.equal(rejected(s), "Hallo Altmarke Welt");
});

test("tracked: multiple single-run matches get unique revision ids", () => {
  const parts = makeParts({ doc: para(run("Altmarke and Altmarke done")) });
  const res = core.replaceCompany(parts, companyOpts({ companyTracked: true }));
  assert.equal(res.count, 2);
  const d = core.parse(docStr(parts));
  const ids = [...core.els(d, NS.w, "del"), ...core.els(d, NS.w, "ins")].map((e) => core.attr(e, NS.w, "id"));
  assert.equal(new Set(ids).size, ids.length, "all revision ids unique");
  assert.equal(accepted(docStr(parts)), "Neumarke and Neumarke done");
});

test("tracked: cross-run match falls back to plain replace and is logged", () => {
  const parts = makeParts({ doc: para(run("Contact Va"), run("med now")) });
  const res = core.replaceCompany(parts, companyOpts({ companyTracked: true }));
  assert.equal(res.count, 1);
  assert.equal(res.fallbacks.length, 1);
  assert.equal(res.fallbacks[0].part, "document.xml");
  assert.match(res.fallbacks[0].snippet, /Altmarke/i);
  const d = core.parse(docStr(parts));
  assert.equal(core.els(d, NS.w, "ins").length, 0, "no tracked ins for cross-run");
  assert.equal(core.els(d, NS.w, "del").length, 0, "no tracked del for cross-run");
  assert.equal(visible(docStr(parts)), "Contact Neumarke now");
});

// ---------------------------------------------------------------------------
// Scope + no-op
// ---------------------------------------------------------------------------
test("scope: replaces across document, header and footer", () => {
  const parts = makeParts({ doc: para(run("Altmarke A")), header: para(run("Altmarke H")), footer: para(run("Altmarke F")) });
  const res = core.replaceCompany(parts, companyOpts());
  assert.equal(res.count, 3);
  assert.equal(visible(docStr(parts)), "Neumarke A");
  assert.equal(visible(dec(parts["word/header1.xml"])), "Neumarke H");
  assert.equal(visible(dec(parts["word/footer1.xml"])), "Neumarke F");
  assertAllPartsParse(parts);
});

test("no-op: absent company name leaves the document byte-identical", () => {
  const parts = makeParts({ doc: para(run("Nichts zu tun hier")) });
  const before = docStr(parts);
  const res = core.replaceCompany(parts, companyOpts());
  assert.equal(res.count, 0);
  assert.equal(docStr(parts), before);
});

// ---------------------------------------------------------------------------
// Step-aware patch() on a real sample document (skipped if input/ is absent)
// ---------------------------------------------------------------------------
const REAL = "input/QUAM/Prozesse/3.08.01 Partnerbewertung.docx";
const haveReal = existsSync(REAL);
const loadReal = () => unzipSync(new Uint8Array(readFileSync(REAL)));
const LOGO = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const DIMS = [100, 40];
const fullOpts = (steps, extra = {}) => Object.assign({
  steps, bearbeiter: "ZZTESTEDITOR", freigeber: "Max Chef", pruefer: "Eva Prüf",
  date: "22.08.2026", iso: "2026-08-22T00:00:00Z", changelog: "Anpassung Rebranding",
  width: 47 * 36000, logoName: "logo.png",
  companyFrom: "Altmarke", companyTo: "Neumarke", companyTracked: true, author: "ZZTESTEDITOR",
}, extra);

test("real: all steps together — verify passes, everything applied", { skip: !haveReal }, () => {
  const parts = loadReal();
  const steps = { logo: true, meta: true, changelog: true, company: true };
  const insp = core.inspect(parts, steps);
  assert.ok(insp.ok, "inspect ok: " + insp.reason);
  const res = core.patch(parts, insp, LOGO, DIMS, fullOpts(steps)); // verify() throws on any failure
  const out = unzipSync(res.bytes);
  assert.ok(res.company.count >= 1, "company replaced at least once");
  assert.match(docStr(out), /Neumarke/);
  assert.match(docStr(out), /Anpassung Rebranding/);
  assert.ok(docStr(out).includes("ZZTESTEDITOR") || dec(out["word/header1.xml"]).includes("ZZTESTEDITOR"), "editor written");
  assertAllPartsParse(out);
});

test("real: meta step off — responsibility fields not written", { skip: !haveReal }, () => {
  const parts = loadReal();
  const steps = { logo: false, meta: false, changelog: false, company: true };
  const insp = core.inspect(parts, steps);
  assert.ok(insp.ok);
  const res = core.patch(parts, insp, null, null, fullOpts(steps, { companyTracked: false }));
  const out = unzipSync(res.bytes);
  const joined = Object.entries(out).filter(([n]) => n.endsWith(".xml")).map(([, b]) => dec(b)).join("");
  assert.ok(!joined.includes("ZZTESTEDITOR"), "editor name must not appear when meta off");
  assert.ok(res.company.count >= 1, "company still replaced");
});

test("real: changelog toggle controls the changelog row", { skip: !haveReal }, () => {
  const marker = "REBRAND-MARKER-ZQX";
  const on = core.patch(loadReal(), core.inspect(loadReal(), { logo: false, meta: true, changelog: true, company: false }),
    null, null, fullOpts({ logo: false, meta: true, changelog: true, company: false }, { changelog: marker }));
  const off = core.patch(loadReal(), core.inspect(loadReal(), { logo: false, meta: true, changelog: false, company: false }),
    null, null, fullOpts({ logo: false, meta: true, changelog: false, company: false }, { changelog: marker }));
  assert.match(docStr(unzipSync(on.bytes)), new RegExp(marker), "changelog on: marker present");
  assert.ok(!docStr(unzipSync(off.bytes)).includes(marker), "changelog off: marker absent");
});

test("real: logo step off — doc still processes for company only", { skip: !haveReal }, () => {
  const parts = loadReal();
  const steps = { logo: false, meta: false, changelog: false, company: true };
  const insp = core.inspect(parts, steps);
  assert.ok(insp.ok, "ready without logo checks");
  const res = core.patch(parts, insp, null, null, fullOpts(steps, { companyTracked: false }));
  assert.ok(res.company.count >= 1);
  assertAllPartsParse(unzipSync(res.bytes));
});
