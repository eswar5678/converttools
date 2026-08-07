# ⚡ ConvertTools

Free, private, in-browser file tools — a fully **static website** ready for **GitHub Pages**.
Every conversion happens client-side with JavaScript. Files never leave the user's device.

The published site is **one single file — `index.html`** (all CSS and all libraries are
inlined into it). It works from any static host, from `file://`, even offline once opened.

## Tools

| Tool | What it does | Default download name |
|---|---|---|
| **Merge PDF** | Combine multiple PDFs into one (reorderable) | `merged.pdf` |
| **Merge Word Docs** | Join several `.docx` files with page breaks | `merged.docx` |
| **Photos → PDF** | JPG/PNG/WebP/GIF → one PDF, page per photo | `images.pdf` |
| **Word → PDF** | Convert `.docx` into a clean PDF | `converted.pdf` |
| **PDF → Word** | Extract PDF text into an editable `.docx` | `converted.docx` |
| **PDF → Photos** | Render PDF pages as PNG/JPG (zipped if many) | `page-1.png` / `pages.zip` |
| **TXT → PDF** | Plain text → paginated A4 PDF | `text.pdf` |
| **PDF → TXT** | Pull all text out of a PDF | `extracted.txt` |

After every operation the result screen shows the **default filename in an editable
field** — the user can rename it (base name is pre-selected) before clicking **Download**.
The correct extension is appended automatically if removed, and a fallback
"open file" link appears in case an embedded preview blocks downloads.

## Project layout

```
index.html          ← THE SITE (generated single-file build — publish this)
.nojekyll           ← tells GitHub Pages to serve files as-is
build-single.py     ← build script: python3 build-single.py
src/index.src.html  ← editable HTML template
src/css/style.css   ← editable styles
src/js/app.js       ← editable application logic
vendor/             ← JS libraries used by the build and by tests
tests/              ← sample files + node smoke test
```

**To change anything:** edit files in `src/`, then run `python3 build-single.py`
to regenerate `index.html`. (Never edit `index.html` directly — it is overwritten.)

## How it works (libraries, all inlined — no CDN, no third-party requests)

- [pdf-lib](https://pdf-lib.js.org/) — PDF merging, image embedding, text layout
- [PDF.js](https://mozilla.github.io/pdf.js/) — reading & rendering PDFs (worker runs
  from an inline-Blob, so the whole site stays one file)
- [mammoth.js](https://github.com/mwilliamson/mammoth.js) — `.docx` → HTML
- [docx](https://docx.js.org/) — building `.docx` files
- [html2pdf.js](https://ekoopmans.github.io/html2pdf.js/) — HTML → PDF rendering
- [JSZip](https://stuk.github.io/jszip/) — zipping multi-page image output

### Honest limitations (shown in the UI where relevant)

- Old binary **`.doc`** (Word 97–2003) **cannot be parsed in a browser** — only modern `.docx`.
- **Word → PDF** and **PDF → Word** are *approximate*: text, headings, lists, basic tables and
  images carry over; exact fonts, columns, headers/footers and complex layouts may shift.
  Perfect fidelity requires a server (no server exists on GitHub Pages).
- Scanned/image-only PDFs have no text layer, so **PDF → Word/TXT** can't extract from them
  (use **PDF → Photos** instead).
- **TXT → PDF**'s built-in font supports Latin characters only.

## Security (why this site is hard to attack)

There is deliberately **nothing to hack**:

- **No server, no database, no accounts, no cookies** — the site is 100% static files.
  There is no backend to breach and no user data is ever collected or stored anywhere.
- **Files never leave the device** — all processing happens in browser memory, so there is
  no upload channel to intercept.
- **Content-Security-Policy** (see `index.html`) — no plugin/object embeds, no forms, no
  `base-uri` tricks, no third-party connect/img/font origins. The single-file build inlines
  its scripts, so `script-src` is `'unsafe-inline'` by necessity — but there is **no user
  input rendered as HTML anywhere** (file names are HTML-escaped before display), which
  removes the injection paths that 'unsafe-inline' would otherwise expose. PDF.js also runs
  with `isEvalSupported: false` so `eval`-style execution stays blocked.
- **All libraries are inlined from `/vendor`** — no CDN that could be hijacked or swapped
  for malicious code (supply-chain safe).
- **No referrer leakage** (`referrer: no-referrer`) and HTTPS is enforced automatically
  by GitHub Pages.
- **Download filenames are sanitized** (path characters stripped) before saving.

## Run locally

Any static server works. From the project root:

```bash
python3 -m http.server 8000
# then open http://localhost:8000
```

(Opening `index.html` directly with `file://` also works — it's one self-contained file.)

## Publish on GitHub Pages

1. Create a new repository on GitHub (e.g. `converttools`).
2. The only file the site strictly needs is **`index.html`** (everything is inlined).
   Push the whole folder, or just `index.html` + `.nojekyll`:
   ```bash
   git init
   git add .
   git commit -m "ConvertTools: static in-browser file tools"
   git branch -M main
   git remote add origin https://github.com/<you>/converttools.git
   git push -u origin main
   ```
3. In the repo → **Settings → Pages** → Source: **Deploy from a branch** →
   Branch: `main` / `(root)` → **Save**.
4. Wait ~1 minute — the site is live at `https://<you>.github.io/converttools/`.

No build step on Pages, no secrets, no backend.

## Test files

The `tests/` folder contains ready-made samples (`chapter-1.pdf`, `chapter-2.pdf`,
`report.docx`, `letter.docx`, `photo-1.png`, `photo-2.jpg`, `story.txt`) so every tool
can be tried instantly. `tests/smoke-test.js` runs a Node-based sanity check of the
core document operations (`node tests/smoke-test.js`).
