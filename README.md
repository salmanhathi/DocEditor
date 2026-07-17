# DocForge

A local, no-AI file toolkit: XML→Excel, PDF→Excel, PDF unlocking, and a full
PDF edit suite (merge/split/reorder/delete pages, rotate, compress, add/remove
watermarks, PDF→PNG/JPG, add/edit page text).

Every conversion happens in server memory using open-source libraries
(PyMuPDF, pikepdf, pdfplumber, pandas, xmltodict). Nothing is written to disk
permanently, nothing is sent to any external API or AI service, and uploads
are discarded as soon as the response is sent.

## Run locally

```bash
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000

## Deploy to Render

1. Push this folder to a GitHub repo.
2. In Render: **New → Web Service**, connect the repo.
3. Render will pick up `render.yaml` automatically (or set manually):
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app --bind 0.0.0.0:$PORT --timeout 120 --workers 2`
4. Free tier is fine to start; bump the instance size if you're processing
   large PDFs (image-heavy compress/convert jobs are the most memory-hungry).

## What each tool does

| Tool | Notes |
|---|---|
| XML → Excel | Auto-detects the repeating record element and flattens it into rows (handles attributes and nesting). |
| PDF → Excel | Detects tables per page via pdfplumber; one sheet per table. Falls back to raw text per page if no tables are found. |
| Unlock PDF | Removes password/permission restrictions. Enter the password if it's user-password protected; leave blank for owner-only restrictions. |
| Merge / Split / Reorder / Delete | Merge multiple PDFs, split into single pages or custom ranges (downloads a zip), or pick pages visually to keep/reorder. |
| Rotate | Rotate all or selected pages by 90/180/270°. |
| Compress | Recompresses embedded images (JPEG re-encode + optional downscale) — most effective on scan-heavy PDFs. |
| Add Watermark/Stamp | Diagonal or horizontal text stamp with adjustable color/size/opacity. |
| Remove Watermark | Give the exact watermark text for a precise redaction, or leave blank to auto-detect text/images repeating across most pages. Won't reliably strip watermarks baked into a flattened background image — spot-check results. |
| PDF → PNG/JPG | Renders pages to images at a chosen DPI. |
| Add/Edit Text | Loads a page's text into an editable box and rewrites the page plainly (left-aligned, single font) with your edits. It does not preserve exact original layout/fonts — no local tool can guarantee that on an arbitrary PDF. |

## Known limitations (by design, not bugs)

- **Text editing** rebuilds the page's text layer rather than doing true
  in-place WYSIWYG editing — faithful layout-preserving edits on an arbitrary
  PDF aren't reliably possible without proprietary format reconstruction.
- **Watermark removal** is heuristic (exact-text match or cross-page
  repetition detection). It works well for standard stamped watermarks but
  can miss ones rendered directly into a background image.
- 200MB upload cap (`MAX_CONTENT_LENGTH` in `app.py`) — raise it if needed.
