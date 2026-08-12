# DOCX Rebranding Tools

Client-side and command-line tools for applying a controlled rebrand to Word
`.docx` files. The browser tool performs processing locally; document contents
are not uploaded.

## Browser tool

Open `rebrand-prototype.html` directly, or serve this folder with any static
file server. Select a logo and one or more `.docx` files, review the preflight
results, then download a patched document or batch ZIP.

The page currently loads the small `fflate` ZIP library from jsDelivr. Vendor
that dependency locally if the published GitHub Pages site must work offline.

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
