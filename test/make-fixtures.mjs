// Synthetic DOCX fixtures for the company-name replacement tests. Each builder
// returns an in-memory { partName: Uint8Array } package (the same shape the core
// operates on), so no files touch disk. Real-document fixtures are read straight
// from input/ by the test runner when present.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const enc = (s) => new TextEncoder().encode(s);

export const run = (text, rpr = "") =>
  `<w:r>${rpr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
export const para = (...runs) => `<w:p>${runs.join("")}</w:p>`;
export const BOLD = "<w:rPr><w:b/></w:rPr>";

const docXml = (inner) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document xmlns:w="${W}"><w:body>${inner}<w:sectPr/></w:body></w:document>`;
const hdrXml = (inner) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr xmlns:w="${W}">${inner}</w:hdr>`;
const ftrXml = (inner) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:ftr xmlns:w="${W}">${inner}</w:ftr>`;

// Build a minimal, well-formed-enough package with an optional header/footer.
export function makeParts({ doc = "", header = null, footer = null } = {}) {
  const overrides = [`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`];
  if (header != null) overrides.push(`<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`);
  if (footer != null) overrides.push(`<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`);
  const parts = {
    "[Content_Types].xml": enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.join("")}</Types>`),
    "_rels/.rels": enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": enc(docXml(doc)),
  };
  if (header != null) parts["word/header1.xml"] = enc(hdrXml(header));
  if (footer != null) parts["word/footer1.xml"] = enc(ftrXml(footer));
  return parts;
}

// Standard options object for company-only runs (no logo/meta/changelog).
export function companyOpts(overrides = {}) {
  return Object.assign({
    steps: { logo: false, meta: false, changelog: false, company: true },
    companyFrom: "Altmarke", companyTo: "Neumarke", companyTracked: false,
    author: "Chef", iso: "2026-08-22T00:00:00Z", date: "22.08.2026",
    bearbeiter: "", freigeber: "", pruefer: "", changelog: "", width: 0, logoName: "",
  }, overrides);
}
