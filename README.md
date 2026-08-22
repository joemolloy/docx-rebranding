# DOCX Rebranding Tools

Client-side and command-line tools for applying a controlled rebrand to Word
`.docx` files. The browser tool performs processing locally; document contents
are not uploaded.

## Browser tool

Open `index.html` directly, or serve this folder with any static file server.
Choose which steps run (logo swap, responsibility/date/version
metadata, changelog row, company-name replacement), fill the relevant fields,
select `.docx` files (and a logo if the logo step is on), review the preflight
results, then download a patched document or batch ZIP.

Company-name replacement (e.g. `Vamed` → `VITREA`) is case-insensitive and
applies to the document body, headers and footers. Enable **Als Änderung
markieren** to record each swap as a tracked Word revision; matches that span a
run boundary fall back to a plain replacement and are listed in the results and
the `rebrand_report.csv` fallback log.

The page currently loads the small `fflate` ZIP library from jsDelivr. Vendor
that dependency locally if the published GitHub Pages site must work offline.

### Editing the browser tool

The document logic lives in **`rebrand-core.mjs`** (the single source of truth).
`index.html` embeds it between `/*CORE_START*/` … `/*CORE_END*/` markers; the UI
glue lives directly in `index.html` outside those markers. After editing
`rebrand-core.mjs`, regenerate the page:

```sh
node build.mjs        # inlines the core into index.html
```

### Tests

The same `rebrand-core.mjs` runs headlessly under Node (via `test/setup.mjs`,
which supplies `@xmldom/xmldom` for the DOM and a small `node:zlib` ZIP shim,
`test/zip-shim.mjs`, in place of `fflate`):

```sh
pnpm install
pnpm test
```

`@xmldom/xmldom` is the only dependency; ZIP handling in Node reuses the
built-in `node:zlib` (the browser keeps using `fflate` from the CDN). The
real-document tests read from `input/` and are skipped automatically when that
tree is absent.

## Python tools

```sh
python3 inventory.py --input ./documents --out inventory.csv
python3 rebrand.py --input ./documents --output ./rebranded \
  --logo new-logo.jpg \
  --bearbeiter "Editor Name" \
  --freigeber "Approver Name" \
  --pruefer "Reviewer Name" \
  --date 30.07.2026
```

The Python implementation keeps originals untouched and writes results to a
separate output tree. Review the generated report before distributing files.

## GitHub Pages

Publish the repository root as a GitHub Pages site. The HTML file is static and
does not require a backend. A repository workflow can be added later if the
site needs automated deployment.
