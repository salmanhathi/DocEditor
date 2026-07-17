"""
DocForge — local file toolkit
XML→Excel, PDF→Excel, Unlock PDF, and a full PDF edit suite.
Everything runs in-process. No file ever leaves this server, no AI calls, no
external services. Uploads live only in memory / a per-request temp dir and
are discarded after the response is sent.
"""
import io
import os
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict

import fitz  # PyMuPDF
import pikepdf
import pdfplumber
import pandas as pd
import xmltodict
from flask import Flask, request, send_file, jsonify, render_template, after_this_request

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB safety cap

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def error(msg, code=400):
    return jsonify({"error": msg}), code


def cleanup_after(*paths):
    @after_this_request
    def _cleanup(response):
        for p in paths:
            try:
                if p and os.path.exists(p):
                    if os.path.isdir(p):
                        shutil.rmtree(p, ignore_errors=True)
                    else:
                        os.remove(p)
            except Exception:
                pass
        return response
    return _cleanup


def parse_page_ranges(spec, max_pages):
    """'1-3,5,8-9' -> sorted list of 0-indexed page numbers (deduped)."""
    if not spec or not spec.strip():
        return list(range(max_pages))
    pages = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a, b = int(a), int(b)
            for p in range(min(a, b), max(a, b) + 1):
                if 1 <= p <= max_pages:
                    pages.add(p - 1)
        else:
            p = int(part)
            if 1 <= p <= max_pages:
                pages.add(p - 1)
    return sorted(pages)


def package_pdf_output(doc, fmt, base_name, dpi=150):
    """Return a Flask response for a finished fitz.Document, in the format
    the user asked for: 'pdf' returns the document as-is; 'png'/'jpg' render
    every page to an image (a single image file if there's one page, a .zip
    of images otherwise). Used by every Edit-group tool so each one can
    export as PDF or as images."""
    fmt = (fmt or "pdf").lower()
    if fmt not in ("png", "jpg", "jpeg"):
        buf = io.BytesIO(doc.tobytes(garbage=3, deflate=True))
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name=f"{base_name}.pdf",
                          mimetype="application/pdf")

    ext = "jpg" if fmt in ("jpg", "jpeg") else "png"
    mime = "image/jpeg" if ext == "jpg" else "image/png"
    mat = fitz.Matrix(dpi / 72, dpi / 72)

    if len(doc) == 1:
        pix = doc[0].get_pixmap(matrix=mat)
        data = pix.tobytes("jpeg" if ext == "jpg" else "png")
        buf = io.BytesIO(data)
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name=f"{base_name}.{ext}", mimetype=mime)

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=mat)
            data = pix.tobytes("jpeg" if ext == "jpg" else "png")
            zf.writestr(f"{base_name}_page{i+1}.{ext}", data)
    zip_buf.seek(0)
    return send_file(zip_buf, as_attachment=True, download_name=f"{base_name}_images.zip",
                      mimetype="application/zip")


# ---------------------------------------------------------------------------
# XML -> Excel
# ---------------------------------------------------------------------------

def xml_to_rows(xml_bytes):
    """Flatten XML into tabular rows. Repeats the first repeating element
    found as one row per record; falls back to a single flattened row."""
    data = xmltodict.parse(xml_bytes)
    root_key = list(data.keys())[0]
    root = data[root_key]

    def find_records(node):
        # look for the first dict value that is a list of dicts -> treat as records
        if isinstance(node, dict):
            for k, v in node.items():
                if isinstance(v, list) and v and isinstance(v[0], dict):
                    return v
            for k, v in node.items():
                found = find_records(v) if isinstance(v, (dict, list)) else None
                if found:
                    return found
        return None

    records = find_records(root)
    if records is None:
        records = [root] if isinstance(root, dict) else [{"value": root}]

    df = pd.json_normalize(records)
    return df


@app.route("/api/xml-to-excel", methods=["POST"])
def xml_to_excel():
    f = request.files.get("file")
    if not f:
        return error("No file uploaded")
    try:
        df = xml_to_rows(f.read())
    except Exception as e:
        return error(f"Could not parse XML: {e}")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
    buf.seek(0)
    name = os.path.splitext(f.filename)[0] + ".xlsx"
    return send_file(buf, as_attachment=True, download_name=name,
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# PDF -> Excel (table extraction)
# ---------------------------------------------------------------------------

TABLE_STRATEGIES = [
    # default: works when the table has visible ruling lines
    {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
    # stricter line detection - fewer false positives on noisy scans
    {"vertical_strategy": "lines_strict", "horizontal_strategy": "lines_strict"},
    # no ruling lines at all (common in bank statements / invoices) -
    # infer columns/rows from text alignment instead. min_words_vertical=1
    # is important: pdfplumber's default (3) needs 3+ aligned words to call
    # a column boundary, which merges adjacent narrow numeric columns
    # (e.g. separate Debit/Credit columns) into one.
    {"vertical_strategy": "text", "horizontal_strategy": "text",
     "min_words_vertical": 1, "min_words_horizontal": 1,
     "intersection_tolerance": 5, "snap_tolerance": 4},
    # mixed: lines where present, text-based gap detection otherwise
    {"vertical_strategy": "text", "horizontal_strategy": "lines", "min_words_vertical": 1},
]


def backfill_missing_cells(page, tbl):
    """pdfplumber returns None for a cell's text when it can't resolve a
    bounding box for it - this happens when a table's ruling lines have a
    gap along just one cell's edge (seen in the uploaded invoice: the
    header row's rightmost border isn't drawn, so the last header cell's
    text - 'TOTAL AMOUNT' - comes back None even though every data row
    below it is fine). Recover it by inferring that cell's x-range from a
    sibling row whose cells are all fully resolved, paired with this row's
    own y-range."""
    grid = tbl.extract()
    rows = tbl.rows
    if not rows:
        return grid
    ncols = len(rows[0].cells)
    ref = next((r for r in rows if len(r.cells) == ncols and all(c is not None for c in r.cells)), None)
    if not ref:
        return grid
    for ri, row in enumerate(rows):
        if ri >= len(grid) or len(row.cells) != ncols:
            continue
        for ci, cellbbox in enumerate(row.cells):
            if ci >= len(grid[ri]) or grid[ri][ci] or cellbbox is not None:
                continue  # already has text, or has a bbox and is genuinely empty
            try:
                x0, x1 = ref.cells[ci][0], ref.cells[ci][2]
                y0, y1 = row.bbox[1], row.bbox[3]
                crop = page.within_bbox((x0, y0, x1, y1 + 2))
                text = (crop.extract_text() or "").strip()
                if text:
                    grid[ri][ci] = text
            except Exception:
                continue
    return grid


def is_plausible_table(table, min_cols=2, max_cols=8):
    """Reject obviously-wrong 'tables' - e.g. plain prose that the
    aggressive text-based strategy split one word per column. The max_cols
    cap only matters for that text-based fallback: a table backed by real
    ruling lines is trustworthy at any column count (invoices routinely
    have 9-10+ columns), so line-based callers should pass max_cols=None."""
    if not table or len(table) < 1:
        return False
    ncols = max(len(r) for r in table)
    nrows = len(table)
    if ncols < min_cols:
        return False
    if max_cols is not None and ncols > max_cols:
        return False
    if nrows < 2:
        return False
    return True


def extract_tables_best_effort(page):
    """Layered strategy: a table with real ruling lines is authoritative and
    returned immediately (min_cols=1, no column cap - a single-column ruled
    table is still real structure, and ruled invoices commonly run 9-10+
    columns). Cells pdfplumber couldn't resolve get backfilled (see
    backfill_missing_cells). Only when no ruled table exists do we fall back
    to text-alignment-based detection - trying the more precise (but more
    false-positive-prone) settings and keeping whichever plausible result
    resolves the most columns (e.g. correctly split adjacent numeric columns
    like Debit/Credit rather than merging them). The text fallback keeps the
    8-column cap since that's what guards against word-by-word false
    positives on plain prose."""
    line_strategies = TABLE_STRATEGIES[:2]
    text_strategies = TABLE_STRATEGIES[2:]

    for settings in line_strategies:
        try:
            found = page.find_tables(table_settings=settings)
        except Exception:
            continue
        plausible = []
        for tbl in found:
            grid = backfill_missing_cells(page, tbl)
            if is_plausible_table(grid, min_cols=1, max_cols=None):
                plausible.append(grid)
        if plausible:
            return plausible

    best = []
    for settings in text_strategies:
        try:
            tables = page.extract_tables(table_settings=settings)
        except Exception:
            continue
        plausible = [t for t in tables if is_plausible_table(t, min_cols=2, max_cols=8)]
        if plausible:
            best_cols = max(len(r) for r in best[0]) if best else 0
            new_cols = max(len(r) for r in plausible[0])
            if new_cols >= best_cols:
                best = plausible
    return best


_DIGIT_SPACE_ARTIFACT = re.compile(r"(?<=\d)\s+(?=\d)")


def table_to_df(table):
    """Turn a raw pdfplumber table (list of rows, possibly jagged, possibly
    with None cells or duplicate/blank headers) into a clean DataFrame
    without ever raising on malformed input."""
    if not table or len(table) == 0:
        return None
    width = max(len(r) for r in table)

    def clean_cell(c):
        if not isinstance(c, str):
            return c or ""
        c = c.strip()
        # some PDF generators emit numbers as two text runs with a
        # near-zero-width space between digits (e.g. "171.05" comes out as
        # "1 71.05") - collapse whitespace strictly between two digits.
        # Never touches real digit+word spacing like "2 PCS".
        return _DIGIT_SPACE_ARTIFACT.sub("", c)

    cleaned = [[clean_cell(c) for c in (row + [None] * (width - len(row)))] for row in table]

    header_row = cleaned[0]
    has_usable_header = any(h for h in header_row) and len(set(header_row)) == len(header_row)
    if has_usable_header:
        header = []
        for i, h in enumerate(header_row):
            header.append(h if h else f"col_{i+1}")
        rows = cleaned[1:]
    else:
        header = [f"col_{i+1}" for i in range(width)]
        rows = cleaned

    # drop fully-empty rows - these are usually phantom row boundaries from
    # the text-based detection strategy, not real blank rows in the source
    rows = [r for r in rows if any(str(c).strip() for c in r)]

    if not rows:
        rows = [[""] * width]
    df = pd.DataFrame(rows, columns=header)
    return df


@app.route("/api/pdf-to-excel", methods=["POST"])
def pdf_to_excel():
    f = request.files.get("file")
    combine = request.form.get("combine", "true").lower() != "false"
    if not f:
        return error("No file uploaded")
    pdf_bytes = f.read()
    try:
        buf = io.BytesIO()
        sheets_written = 0
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                found_tables = []  # (page_idx, table_idx, df)
                for i, page in enumerate(pdf.pages):
                    tables = extract_tables_best_effort(page)
                    for j, table in enumerate(tables):
                        df = table_to_df(table)
                        if df is None or df.empty:
                            continue
                        found_tables.append((i, j, df))

                if combine:
                    # group tables that share the exact same header (same
                    # column names, same order) - these are almost always
                    # the same table continuing across pages
                    groups = OrderedDict()
                    for i, j, df in found_tables:
                        key = tuple(df.columns)
                        groups.setdefault(key, []).append((i, j, df))

                    group_num = 0
                    for key, items in groups.items():
                        if len(items) > 1:
                            group_num += 1
                            combined_df = pd.concat([d for (_, _, d) in items], ignore_index=True)
                            sheet_name = f"combined_{group_num}"[:31]
                            combined_df.to_excel(writer, index=False, sheet_name=sheet_name)
                            sheets_written += 1
                        else:
                            i, j, df = items[0]
                            sheet_name = f"p{i+1}_t{j+1}"[:31]
                            df.to_excel(writer, index=False, sheet_name=sheet_name)
                            sheets_written += 1
                else:
                    for i, j, df in found_tables:
                        sheet_name = f"p{i+1}_t{j+1}"[:31]
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                        sheets_written += 1

                if sheets_written == 0:
                    # no tables detected anywhere - fall back to raw text per page
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text() or ""
                        df = pd.DataFrame({"text": text.split("\n")})
                        df.to_excel(writer, index=False, sheet_name=f"page{i+1}"[:31])
        buf.seek(0)
    except Exception as e:
        return error(f"Could not process PDF: {e}")

    name = os.path.splitext(f.filename)[0] + ".xlsx"
    return send_file(buf, as_attachment=True, download_name=name,
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# Unlock PDF
# ---------------------------------------------------------------------------

@app.route("/api/unlock-pdf", methods=["POST"])
def unlock_pdf():
    f = request.files.get("file")
    password = request.form.get("password", "")
    if not f:
        return error("No file uploaded")
    pdf_bytes = f.read()
    try:
        try:
            pdf = pikepdf.open(io.BytesIO(pdf_bytes), password=password)
        except pikepdf.PasswordError:
            return error("Incorrect password (or a password is required).")
        buf = io.BytesIO()
        pdf.save(buf)
        pdf.close()
        buf.seek(0)
    except Exception as e:
        return error(f"Could not unlock PDF: {e}")

    name = os.path.splitext(f.filename)[0] + "_unlocked.pdf"
    return send_file(buf, as_attachment=True, download_name=name, mimetype="application/pdf")


# ---------------------------------------------------------------------------
# PDF page ops: merge / split / delete / reorder / rotate / compress
# ---------------------------------------------------------------------------

@app.route("/api/pdf-merge", methods=["POST"])
def pdf_merge():
    files = request.files.getlist("files")
    out_format = request.form.get("format", "pdf")
    dpi = int(request.form.get("dpi", 150))
    if len(files) < 2:
        return error("Upload at least two PDFs to merge.")
    out = fitz.open()
    try:
        for f in files:
            with fitz.open(stream=f.read(), filetype="pdf") as doc:
                out.insert_pdf(doc)
        response = package_pdf_output(out, out_format, "merged", dpi)
    except Exception as e:
        return error(f"Could not merge PDFs: {e}")
    finally:
        out.close()
    return response


@app.route("/api/pdf-split", methods=["POST"])
def pdf_split():
    f = request.files.get("file")
    ranges = request.form.get("ranges", "").strip()
    mode = request.form.get("mode", "single")  # 'single' pages each, or 'ranges' as groups
    out_format = request.form.get("format", "pdf").lower()
    dpi = int(request.form.get("dpi", 150))
    if not f:
        return error("No file uploaded")
    if out_format not in ("png", "jpg", "jpeg"):
        out_format = "pdf"
    ext = "jpg" if out_format in ("jpg", "jpeg") else ("png" if out_format == "png" else "pdf")
    mat = fitz.Matrix(dpi / 72, dpi / 72)

    def add_part_to_zip(zf, part_doc, name_prefix):
        if out_format == "pdf":
            zf.writestr(f"{name_prefix}.pdf", part_doc.tobytes(garbage=3, deflate=True))
        else:
            for pi, page in enumerate(part_doc):
                pix = page.get_pixmap(matrix=mat)
                data = pix.tobytes("jpeg" if ext == "jpg" else "png")
                suffix = f"_p{pi+1}" if len(part_doc) > 1 else ""
                zf.writestr(f"{name_prefix}{suffix}.{ext}", data)

    try:
        doc = fitz.open(stream=f.read(), filetype="pdf")
        n = len(doc)
        zip_buf = io.BytesIO()
        base = os.path.splitext(f.filename)[0]
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            if mode == "ranges" and ranges:
                groups = [g.strip() for g in ranges.split(";") if g.strip()]
                for idx, g in enumerate(groups, 1):
                    pages = parse_page_ranges(g, n)
                    if not pages:
                        continue
                    part = fitz.open()
                    for p in pages:
                        part.insert_pdf(doc, from_page=p, to_page=p)
                    add_part_to_zip(zf, part, f"{base}_part{idx}")
                    part.close()
            else:
                pages = parse_page_ranges(ranges, n) if ranges else list(range(n))
                for p in pages:
                    part = fitz.open()
                    part.insert_pdf(doc, from_page=p, to_page=p)
                    add_part_to_zip(zf, part, f"{base}_page{p+1}")
                    part.close()
        doc.close()
        zip_buf.seek(0)
    except Exception as e:
        return error(f"Could not split PDF: {e}")
    return send_file(zip_buf, as_attachment=True, download_name=f"{base}_split.zip", mimetype="application/zip")


@app.route("/api/pdf-organize", methods=["POST"])
def pdf_organize():
    """Delete / reorder pages in one pass. 'order' is a comma-separated
    1-indexed page list in the desired final order (omit pages to delete them)."""
    f = request.files.get("file")
    order = request.form.get("order", "").strip()
    out_format = request.form.get("format", "pdf")
    dpi = int(request.form.get("dpi", 150))
    if not f:
        return error("No file uploaded")
    try:
        doc = fitz.open(stream=f.read(), filetype="pdf")
        n = len(doc)
        if order:
            new_order = [int(x.strip()) - 1 for x in order.split(",") if x.strip()]
            new_order = [p for p in new_order if 0 <= p < n]
        else:
            new_order = list(range(n))
        if not new_order:
            return error("No valid pages selected.")
        doc.select(new_order)
        response = package_pdf_output(doc, out_format, "organized", dpi)
        doc.close()
    except Exception as e:
        return error(f"Could not organize PDF: {e}")
    return response


@app.route("/api/pdf-rotate", methods=["POST"])
def pdf_rotate():
    f = request.files.get("file")
    angle = int(request.form.get("angle", 90))
    pages_spec = request.form.get("pages", "").strip()
    out_format = request.form.get("format", "pdf")
    dpi = int(request.form.get("dpi", 150))
    if not f:
        return error("No file uploaded")
    try:
        doc = fitz.open(stream=f.read(), filetype="pdf")
        n = len(doc)
        pages = parse_page_ranges(pages_spec, n) if pages_spec else list(range(n))
        for p in pages:
            page = doc[p]
            page.set_rotation((page.rotation + angle) % 360)
        response = package_pdf_output(doc, out_format, "rotated", dpi)
        doc.close()
    except Exception as e:
        return error(f"Could not rotate PDF: {e}")
    return response


@app.route("/api/pdf-compress", methods=["POST"])
def pdf_compress():
    f = request.files.get("file")
    quality = request.form.get("quality", "medium")  # low/medium/high image dpi
    out_format = request.form.get("format", "pdf")
    dpi = int(request.form.get("dpi", 150))
    if not f:
        return error("No file uploaded")
    quality_map = {"low": 40, "medium": 60, "high": 80}
    scale_map = {"low": 0.6, "medium": 0.8, "high": 1.0}
    jpg_quality = quality_map.get(quality, 60)
    scale = scale_map.get(quality, 0.8)
    try:
        original_bytes = f.read()
        doc = fitz.open(stream=original_bytes, filetype="pdf")
        seen_xrefs = set()
        for page in doc:
            for img in page.get_images(full=True):
                xref = img[0]
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)
                try:
                    pix = fitz.Pixmap(doc, xref)
                    if pix.n - pix.alpha >= 4:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    if pix.alpha:
                        pix = fitz.Pixmap(pix, 0)
                    if scale < 1.0:
                        new_w = max(1, int(pix.width * scale))
                        new_h = max(1, int(pix.height * scale))
                        pix = fitz.Pixmap(pix, new_w, new_h)
                    doc.update_stream(xref, pix.tobytes("jpeg", jpg_quality=jpg_quality))
                except Exception:
                    continue
        compressed_bytes = doc.tobytes(garbage=4, deflate=True, clean=True)
        doc.close()
        # only use compressed version if it's actually smaller
        if len(compressed_bytes) >= len(original_bytes):
            compressed_bytes = fitz.open(stream=original_bytes, filetype="pdf").tobytes(garbage=4, deflate=True, clean=True)
        out_doc = fitz.open(stream=compressed_bytes, filetype="pdf")
        response = package_pdf_output(out_doc, out_format, "compressed", dpi)
        out_doc.close()
    except Exception as e:
        return error(f"Could not compress PDF: {e}")
    return response


# ---------------------------------------------------------------------------
# Watermark: add / remove
# ---------------------------------------------------------------------------

@app.route("/api/pdf-watermark-add", methods=["POST"])
def pdf_watermark_add():
    f = request.files.get("file")
    text = request.form.get("text", "").strip()
    opacity = float(request.form.get("opacity", 0.3))
    size = int(request.form.get("size", 48))
    color = request.form.get("color", "#808080").lstrip("#")
    position = request.form.get("position", "center")  # center/diagonal
    if not f:
        return error("No file uploaded")
    if not text:
        return error("Watermark text is required.")
    try:
        r, g, b = tuple(int(color[i:i+2], 16) / 255.0 for i in (0, 2, 4))
    except Exception:
        r, g, b = 0.5, 0.5, 0.5
    try:
        doc = fitz.open(stream=f.read(), filetype="pdf")
        for page in doc:
            rect = page.rect
            point = fitz.Point(rect.width * 0.2, rect.height * 0.55)
            morph = (point, fitz.Matrix(45)) if position == "diagonal" else None
            page.insert_text(
                point,
                text,
                fontsize=size,
                color=(r, g, b),
                fill_opacity=opacity,
                morph=morph,
                overlay=True,
            )
        buf = io.BytesIO(doc.tobytes(garbage=3, deflate=True))
        doc.close()
        buf.seek(0)
    except Exception as e:
        return error(f"Could not add watermark: {e}")
    return send_file(buf, as_attachment=True, download_name="watermarked.pdf", mimetype="application/pdf")


@app.route("/api/pdf-watermark-remove", methods=["POST"])
def pdf_watermark_remove():
    """Best-effort watermark removal:
    - if watermark_text given, redacts every instance of that text on every page
    - otherwise, auto-detects text spans that repeat identically across most
      pages (the signature of a stamped watermark) and redacts those
    Image watermarks that repeat on every page (same xref) are also stripped.
    This will not perfectly handle every watermark (e.g. watermarks baked into
    a flattened background image) - results should be spot-checked.
    """
    f = request.files.get("file")
    watermark_text = request.form.get("watermark_text", "").strip()
    if not f:
        return error("No file uploaded")
    try:
        doc = fitz.open(stream=f.read(), filetype="pdf")
        n = len(doc)

        if watermark_text:
            for page in doc:
                rects = page.search_for(watermark_text)
                for rect in rects:
                    page.add_redact_annot(rect, fill=(1, 1, 1))
                page.apply_redactions()
        else:
            # auto-detect: text spans whose exact string repeats on >=60% of pages
            from collections import Counter
            span_counter = Counter()
            span_examples = {}
            for page in doc:
                seen_this_page = set()
                d = page.get_text("dict")
                for block in d.get("blocks", []):
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            t = span["text"].strip()
                            if len(t) >= 3 and t not in seen_this_page:
                                span_counter[t] += 1
                                seen_this_page.add(t)
                                span_examples[t] = span
            threshold = max(2, int(n * 0.6))
            watermark_strings = {t for t, c in span_counter.items() if c >= threshold}
            if watermark_strings:
                for page in doc:
                    for t in watermark_strings:
                        rects = page.search_for(t)
                        for rect in rects:
                            page.add_redact_annot(rect, fill=(1, 1, 1))
                    page.apply_redactions()

            # repeated image xrefs (same image on every/most page) -> remove
            xref_pages = {}
            for i, page in enumerate(doc):
                for img in page.get_images(full=True):
                    xref = img[0]
                    xref_pages.setdefault(xref, set()).add(i)
            for xref, pages_seen in xref_pages.items():
                if len(pages_seen) >= threshold:
                    for i in pages_seen:
                        doc[i].delete_image(xref)

        buf = io.BytesIO(doc.tobytes(garbage=4, deflate=True))
        doc.close()
        buf.seek(0)
    except Exception as e:
        return error(f"Could not remove watermark: {e}")
    return send_file(buf, as_attachment=True, download_name="watermark_removed.pdf", mimetype="application/pdf")


# ---------------------------------------------------------------------------
# PDF -> images (PNG/JPG)
# ---------------------------------------------------------------------------

@app.route("/api/pdf-to-image", methods=["POST"])
def pdf_to_image():
    f = request.files.get("file")
    fmt = request.form.get("format", "png").lower()
    dpi = int(request.form.get("dpi", 150))
    pages_spec = request.form.get("pages", "").strip()
    if not f:
        return error("No file uploaded")
    if fmt not in ("png", "jpg", "jpeg"):
        fmt = "png"
    try:
        doc = fitz.open(stream=f.read(), filetype="pdf")
        n = len(doc)
        pages = parse_page_ranges(pages_spec, n) if pages_spec else list(range(n))
        mat = fitz.Matrix(dpi / 72, dpi / 72)

        if len(pages) == 1:
            page = doc[pages[0]]
            pix = page.get_pixmap(matrix=mat)
            ext = "jpg" if fmt in ("jpg", "jpeg") else "png"
            data = pix.tobytes(ext if ext != "jpg" else "jpeg")
            doc.close()
            buf = io.BytesIO(data)
            buf.seek(0)
            mime = "image/jpeg" if ext == "jpg" else "image/png"
            return send_file(buf, as_attachment=True, download_name=f"page.{ext}", mimetype=mime)

        zip_buf = io.BytesIO()
        base = os.path.splitext(f.filename)[0]
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in pages:
                page = doc[p]
                pix = page.get_pixmap(matrix=mat)
                ext = "jpg" if fmt in ("jpg", "jpeg") else "png"
                data = pix.tobytes(ext if ext != "jpg" else "jpeg")
                zf.writestr(f"{base}_page{p+1}.{ext}", data)
        doc.close()
        zip_buf.seek(0)
        return send_file(zip_buf, as_attachment=True, download_name=f"{base}_images.zip", mimetype="application/zip")
    except Exception as e:
        return error(f"Could not convert PDF to image: {e}")


# ---------------------------------------------------------------------------
# PDF info (page count etc, used by the frontend to build page pickers)
# ---------------------------------------------------------------------------

@app.route("/api/pdf-info", methods=["POST"])
def pdf_info():
    f = request.files.get("file")
    if not f:
        return error("No file uploaded")
    try:
        doc = fitz.open(stream=f.read(), filetype="pdf")
        info = {"pages": len(doc), "encrypted": doc.is_encrypted}
        doc.close()
    except Exception as e:
        return error(f"Could not read PDF: {e}")
    return jsonify(info)


@app.route("/")
def index():
    return render_template("index.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
