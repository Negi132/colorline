# PDF → Excel Converter (browser version)

A single-page website that converts Color Line Cargo invoice PDFs into Excel
files **entirely in the visitor's browser**. There is no server and no upload —
the PDF never leaves the user's computer. It can be hosted for free on GitHub
Pages.

## Files

| File                | Purpose                                                        |
|---------------------|----------------------------------------------------------------|
| `index.html`        | The web page and all of its logic (loads Python via Pyodide).  |
| `converter_core.py` | The conversion engine (pure Python: pdfminer.six + openpyxl).  |

Both files must sit in the **same folder**. `index.html` fetches
`converter_core.py` at runtime, so editing the Python is all you need to tweak
the conversion rules — no rebuild step.

## How it works

When the page opens it downloads the Pyodide runtime (Python compiled to
WebAssembly) from a CDN, installs `pdfminer.six` and `openpyxl` with `micropip`,
and loads `converter_core.py`. After that, converting is instant and offline.
The first load takes a few seconds because of the one-time runtime download;
the browser caches it afterwards.

## Editing the page text

Open `index.html` and look for the block marked
`EDITABLE TEXT` near the top. Change the heading, the instructions, and the
"only tested with…" notice freely — it's plain HTML.

## Deploy to GitHub Pages

1. Create a new repository on GitHub (e.g. `pdf-to-excel`).
2. Upload `index.html` and `converter_core.py` to the repository root
   (drag them onto the GitHub page, or use `git`).
3. In the repository: **Settings → Pages**.
4. Under **Build and deployment → Source**, choose **Deploy from a branch**.
5. Pick branch **main** and folder **/ (root)**, then **Save**.
6. Wait ~1 minute. GitHub shows the live URL, e.g.
   `https://<your-username>.github.io/pdf-to-excel/`
7. Send that link to the customer. Nothing to install on their side — any
   modern browser works.

To update later, just commit new versions of the two files; the site refreshes
automatically.

## Testing locally first (optional)

You can't just double-click `index.html` (the browser blocks the Python
download from `file://`). Serve it instead:

```bash
cd this-folder
python -m http.server 8000
```

Then open <http://localhost:8000>.

## Notes

- Pyodide version is pinned in `index.html` (the CDN URL). If a newer version
  is released, bump the number; check <https://pyodide.org>.
- The converter has only been validated against Color Line Cargo invoices.
- The two conversion options on the desktop version are exposed in the page as
  a checkbox (configuration block) and a dropdown (buyer/recipient mapping).
