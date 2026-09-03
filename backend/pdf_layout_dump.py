"""
"Preserve layout" PDF -> Word conversion.

Unlike docx_fill.py (which maps extracted data into an existing company
template), this module builds a *brand new* .docx from scratch that
reconstructs the source PDF's own reading order: paragraphs (with heading /
centering detection from font size and position), real Word tables wherever
the PDF has a ruled/bordered table, and images (logos, stamps, figures)
cropped straight out of the rendered page and dropped in at roughly the right
point in the flow.

This is a best-effort reconstruction, not a pixel-perfect clone -- PDF is a
fixed page-description format and Word is a flowing document format, so exact
fonts/kerning/margins can't be guaranteed to match 1:1 (the same is true of
Word's own built-in "Open PDF" feature). What IS preserved: reading order,
paragraph/heading structure, table structure and cell contents, and images.

Pages with a real text layer are read directly (fast, exact). Pages that are
actually a scan or photo (no text layer -- the "page" is just one big picture)
are run through OCR instead: the page is rendered to a high-resolution image,
text is recognized word-by-word, a ruled-table grid is detected from the
page's own line structure (if present) and OCR'd cell-by-cell, and any
leftover graphical regions (logos, stamps, signatures) that aren't part of
the recognized text are cropped out as images -- so a scanned page ends up
just as editable as a normal one, instead of being dropped in as one flat
picture.

No network access; pdfplumber + python-docx + Pillow, with OpenCV + Tesseract
for the OCR path.
"""
from __future__ import annotations

import dataclasses
import io
from typing import List, Optional, Tuple

import numpy as np
import pdfplumber
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from PIL import Image

try:
    import cv2
    import pytesseract
    _OCR_AVAILABLE = True
except ImportError:
    _OCR_AVAILABLE = False

OCR_DPI = 300


# --------------------------------------------------------------------------- #
# Block extraction (per page): paragraphs, tables, images, each tagged with a
# vertical position so they can be interleaved in the correct reading order.
# --------------------------------------------------------------------------- #

@dataclasses.dataclass
class TextBlock:
    kind: str  # "paragraph"
    top: float
    lines: List[str]
    avg_size: float
    centered: bool
    bold: bool


@dataclasses.dataclass
class TableBlock:
    kind: str  # "table"
    top: float
    rows: List[List[str]]


@dataclasses.dataclass
class ImageBlock:
    kind: str  # "image"
    top: float
    png_bytes: bytes
    width_pt: float
    height_pt: float


def _bbox_overlaps(word, table_bboxes) -> bool:
    wx0, wtop, wx1, wbottom = word["x0"], word["top"], word["x1"], word["bottom"]
    for (x0, top, x1, bottom) in table_bboxes:
        if wx0 >= x0 - 1 and wx1 <= x1 + 1 and wtop >= top - 1 and wbottom <= bottom + 1:
            return True
    return False


def _group_words_into_paragraphs(words: List[dict], page_width: float) -> List[TextBlock]:
    if not words:
        return []
    words = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))

    # 1) group into lines by close 'top' values
    lines: List[List[dict]] = []
    current: List[dict] = [words[0]]
    for w in words[1:]:
        if abs(w["top"] - current[-1]["top"]) <= 3:
            current.append(w)
        else:
            lines.append(current)
            current = [w]
    lines.append(current)

    line_records = []
    for line_words in lines:
        line_words.sort(key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in line_words)
        top = min(w["top"] for w in line_words)
        bottom = max(w["bottom"] for w in line_words)
        x0 = min(w["x0"] for w in line_words)
        x1 = max(w["x1"] for w in line_words)
        avg_size = sum(w["size"] for w in line_words) / len(line_words)
        is_bold = any("bold" in (w.get("fontname") or "").lower() for w in line_words)
        line_records.append(dict(text=text, top=top, bottom=bottom, x0=x0, x1=x1, size=avg_size, bold=is_bold))

    # 2) group lines into paragraphs: new paragraph when the vertical gap to the
    # previous line is large relative to the line height, or font size changes a lot
    paragraphs: List[TextBlock] = []
    cur_lines: List[dict] = [line_records[0]]
    for rec in line_records[1:]:
        prev = cur_lines[-1]
        gap = rec["top"] - prev["bottom"]
        line_height = max(prev["bottom"] - prev["top"], 1)
        size_changed = abs(rec["size"] - prev["size"]) > 2
        if gap > line_height * 0.9 or size_changed:
            paragraphs.append(_finish_paragraph(cur_lines, page_width))
            cur_lines = [rec]
        else:
            cur_lines.append(rec)
    paragraphs.append(_finish_paragraph(cur_lines, page_width))
    return paragraphs


def _finish_paragraph(lines: List[dict], page_width: float) -> TextBlock:
    top = min(l["top"] for l in lines)
    avg_size = sum(l["size"] for l in lines) / len(lines)
    bold = sum(1 for l in lines if l["bold"]) >= len(lines) / 2
    x0 = min(l["x0"] for l in lines)
    x1 = max(l["x1"] for l in lines)
    block_center = (x0 + x1) / 2
    page_center = page_width / 2
    centered = abs(block_center - page_center) < page_width * 0.08 and (x0 > page_width * 0.15)
    return TextBlock(kind="paragraph", top=top, lines=[l["text"] for l in lines], avg_size=avg_size, centered=centered, bold=bold)


# --------------------------------------------------------------------------- #
# OCR path: for pages that are actually a scan/photo (no real text layer).
# --------------------------------------------------------------------------- #

def _page_is_scanned(page) -> bool:
    """A page counts as 'scanned' if pdfplumber can't find a usable text
    layer on it (a handful of stray characters from a stamp/watermark don't
    count) -- in that case the page content is really just a picture and
    needs OCR."""
    try:
        text = page.extract_text() or ""
    except Exception:
        text = ""
    return len(text.strip()) < 20


def _render_page_image(page, dpi: int = OCR_DPI) -> "Image.Image":
    return page.to_image(resolution=dpi).original.convert("RGB")


def _ocr_words(pil_image: "Image.Image", dpi: int) -> List[dict]:
    """Run Tesseract on a rendered page image and return word dicts in the
    same shape pdfplumber's extract_words() produces (x0/x1/top/bottom in PDF
    points, plus a synthetic 'size' from the box height so the existing
    paragraph/heading logic works unchanged)."""
    if not _OCR_AVAILABLE:
        return []
    scale = 72.0 / dpi
    data = pytesseract.image_to_data(pil_image, output_type=pytesseract.Output.DICT)
    words = []
    for i in range(len(data["text"])):
        text = (data["text"][i] or "").strip()
        try:
            conf = float(data["conf"][i])
        except (ValueError, TypeError):
            conf = -1
        if not text or conf < 40:
            continue
        x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
        words.append({
            "text": text,
            "x0": x * scale, "x1": (x + w) * scale,
            "top": y * scale, "bottom": (y + h) * scale,
            "size": max(6.0, h * scale * 0.8),
            "fontname": "",
        })
    return words


def _peak_positions(profile: "np.ndarray", min_gap: int) -> List[int]:
    """Given a 1D 'how much line ink is here' profile, return the center of
    each distinct peak (i.e. each detected grid line's position)."""
    if profile.max() <= 0:
        return []
    threshold = profile.max() * 0.3
    positions = np.where(profile > threshold)[0]
    if len(positions) == 0:
        return []
    clusters = [[int(positions[0])]]
    for p in positions[1:]:
        if p - clusters[-1][-1] <= min_gap:
            clusters[-1].append(int(p))
        else:
            clusters.append([int(p)])
    return [int(np.mean(c)) for c in clusters]


def _join_words_in_reading_order(words: List[dict]) -> str:
    """Join a bag of OCR words (e.g. everything inside one detected table
    cell) into reading-order text. Sorting directly by each word's raw
    OCR `top` is NOT safe: two words on the same visual line routinely get
    slightly different `top` values (a few points, from OCR bbox noise,
    differing glyph heights, or tiny page skew), so a plain (top, x0) sort
    ends up ordering primarily by that noise and scrambles left-to-right
    order within a line. Instead, cluster words into visual lines first
    (grouping by top within a tolerance derived from the words' own font
    size), sort each line left-to-right by x0, then stack lines top to
    bottom."""
    if not words:
        return ""
    ordered = sorted(words, key=lambda w_: w_["top"])
    lines: List[List[dict]] = []
    for w_ in ordered:
        if lines:
            line_top = min(x["top"] for x in lines[-1])
            tol = 0.6 * max((x.get("size") or 8.0) for x in lines[-1])
            if w_["top"] - line_top <= tol:
                lines[-1].append(w_)
                continue
        lines.append([w_])
    parts = []
    for line in lines:
        line.sort(key=lambda w_: w_["x0"])
        parts.append(" ".join(w_["text"] for w_ in line))
    return "\n".join(parts)


def _ocr_cell_text(pil_image: "Image.Image", px0: float, py0: float, px1: float, py1: float) -> str:
    """Run a fresh, isolated OCR pass on just one table cell's pixel region.

    This exists because a single OCR pass over the WHOLE page (what
    `_ocr_words` does) can silently drop words that sit inside a dense grid
    of ruled lines -- verified on a real scanned CoA where Tesseract's
    automatic page-segmentation, confused by all the surrounding table
    rules, dropped entire cell labels ('Appearance', 'Solubility', 'Bulk
    density') even though those same words OCR perfectly (conf ~96) once
    cropped down to just their own cell. Re-OCRing each cell in isolation
    (small, uniform, single-block image -- ideal Tesseract conditions) is
    slower but far more reliable, and it's naturally already in correct
    reading order, sidestepping the word-scrambling problem entirely."""
    w_px, h_px = pil_image.size
    pad = 2
    x0, y0 = max(0, int(px0) - pad), max(0, int(py0) - pad)
    x1, y1 = min(w_px, int(px1) + pad), min(h_px, int(py1) + pad)
    if x1 - x0 < 4 or y1 - y0 < 4:
        return ""
    crop = pil_image.crop((x0, y0, x1, y1))
    try:
        # psm 6: "assume a single uniform block of text" -- right fit for one
        # table cell, unlike the full-page auto layout analysis.
        raw = pytesseract.image_to_string(crop, config="--psm 6")
    except Exception:
        return ""
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    return "\n".join(lines)


def _detect_scanned_tables(gray: "np.ndarray", ocr_words: List[dict], dpi: int,
                            pil_image: "Image.Image" = None) -> List[dict]:
    """Detect ruled table grids in a scanned page via line morphology (the
    standard 'erode/dilate with a long thin kernel' trick to isolate
    horizontal and vertical rules), then read each detected cell. Returns
    [{'bbox_pt': (x0,top,x1,bottom), 'rows': [[str,...]]}, ...]. Silently
    returns [] if the page has no ruled grid (e.g. a borderless scan) --
    callers should fall back to plain paragraphs in that case."""
    scale = 72.0 / dpi
    h_px, w_px = gray.shape
    bw = cv2.adaptiveThreshold(255 - gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -2)

    horiz_size = max(20, w_px // 40)
    horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (horiz_size, 1))
    horiz = cv2.dilate(cv2.erode(bw, horiz_kernel), horiz_kernel)

    vert_size = max(20, h_px // 40)
    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, vert_size))
    vert = cv2.dilate(cv2.erode(bw, vert_kernel), vert_kernel)

    grid = cv2.bitwise_or(horiz, vert)
    contours, _ = cv2.findContours(grid, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    tables = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w < w_px * 0.25 or h < h_px * 0.03:
            continue

        # min_gap must scale with DPI, not be a fixed pixel count: at 300dpi a
        # single ruled line's ink can spread across several adjacent profile
        # peaks (anti-aliasing, slight page skew), which a too-small min_gap
        # (previously a flat 8px) reports as multiple distinct grid lines
        # instead of merging them into the one real line -- this was
        # corrupting column/row counts on real scanned tables (verified: a
        # true 4-column table was coming out as 6 columns at min_gap=8; a
        # margin of dpi*0.12 was still occasionally 1 pixel-cluster short,
        # dpi*0.14 merges reliably and stays stable well past that point).
        line_gap = max(8, int(dpi * 0.14))
        row_lines = _peak_positions(horiz[y:y + h, x:x + w].sum(axis=1), min_gap=line_gap)
        col_lines = _peak_positions(vert[y:y + h, x:x + w].sum(axis=0), min_gap=line_gap)
        if len(row_lines) < 2 or len(col_lines) < 2:
            continue

        row_bounds = [y + r for r in row_lines]
        col_bounds = [x + c for c in col_lines]
        rows_data = []
        for ri in range(len(row_bounds) - 1):
            row_cells = []
            cy0, cy1 = row_bounds[ri] * scale, row_bounds[ri + 1] * scale
            for ci in range(len(col_bounds) - 1):
                cx0, cx1 = col_bounds[ci] * scale, col_bounds[ci + 1] * scale
                cell_text = ""
                if pil_image is not None:
                    cell_text = _ocr_cell_text(pil_image, col_bounds[ci], row_bounds[ri],
                                                col_bounds[ci + 1], row_bounds[ri + 1])
                if not cell_text:
                    # Fallback: assign from the whole-page OCR word list (still
                    # useful if the isolated per-cell OCR pass came back empty,
                    # e.g. a cell that's really blank vs one Tesseract choked on).
                    cell_words = [
                        w_ for w_ in ocr_words
                        if cx0 <= (w_["x0"] + w_["x1"]) / 2 <= cx1
                        and cy0 <= (w_["top"] + w_["bottom"]) / 2 <= cy1
                    ]
                    cell_text = _join_words_in_reading_order(cell_words)
                row_cells.append(cell_text)
            rows_data.append(row_cells)

        if not any(any(c.strip() for c in r) for r in rows_data):
            continue

        bbox_pt = (x * scale, y * scale, (x + w) * scale, (y + h) * scale)
        tables.append({"bbox_pt": bbox_pt, "rows": rows_data})
    return tables


def _word_in_bbox(word: dict, bbox_pt) -> bool:
    x0, top, x1, bottom = bbox_pt
    cx, cy = (word["x0"] + word["x1"]) / 2, (word["top"] + word["bottom"]) / 2
    return x0 <= cx <= x1 and top <= cy <= bottom


def _detect_logo_regions(pil_image: "Image.Image", gray: "np.ndarray", ocr_words: List[dict],
                          table_bboxes_pt: List[tuple], dpi: int) -> List[dict]:
    """Find graphical regions (logos, stamps, signatures) on a scanned page:
    areas with real ink that AREN'T explained by recognized text or a
    detected table. Returns [{'bbox_pt': (...), 'png_bytes': ...}, ...]."""
    scale = 72.0 / dpi
    h_px, w_px = gray.shape
    ink = (255 - gray) > 40  # True where there's non-background ink

    mask = np.zeros_like(ink, dtype=bool)
    pad = int(0.15 * dpi)  # a bit of headroom around each word/table box
    for w_ in ocr_words:
        x0 = max(0, int(w_["x0"] / scale) - pad); x1 = min(w_px, int(w_["x1"] / scale) + pad)
        y0 = max(0, int(w_["top"] / scale) - pad); y1 = min(h_px, int(w_["bottom"] / scale) + pad)
        mask[y0:y1, x0:x1] = True
    for bbox in table_bboxes_pt:
        x0, top, x1, bottom = bbox
        mask[max(0, int(top / scale)):min(h_px, int(bottom / scale)),
             max(0, int(x0 / scale)):min(w_px, int(x1 / scale))] = True

    leftover = (ink & ~mask).astype(np.uint8) * 255
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 25))
    leftover = cv2.dilate(leftover, kernel, iterations=1)
    contours, _ = cv2.findContours(leftover, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    page_area = h_px * w_px
    regions = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        if area < page_area * 0.001 or area > page_area * 0.5:
            continue
        if w > w_px * 0.97 and h > h_px * 0.97:
            continue  # a border/frame around the whole page, not a real graphic
        if w * scale < 15 or h * scale < 15:
            continue  # too thin in one dimension to be a real graphic -- a stray rule/underline
        if max(w, h) / max(1, min(w, h)) > 12:
            continue  # extreme aspect ratio -- almost certainly a line, not a logo/stamp
        x0, y0 = max(0, x - 5), max(0, y - 5)
        x1, y1 = min(w_px, x + w + 5), min(h_px, y + h + 5)
        crop = pil_image.crop((x0, y0, x1, y1))
        buf = io.BytesIO()
        crop.save(buf, format="PNG")
        regions.append({
            "bbox_pt": (x0 * scale, y0 * scale, x1 * scale, y1 * scale),
            "png_bytes": buf.getvalue(),
        })
    return regions


def _extract_scanned_page_blocks(page, page_width: float, page_height: float) -> List:
    """OCR-based replacement for the normal pdfplumber path, used when a page
    turns out to be a scan/photo rather than real digital text. Returns a
    list of (top, block) tuples for this page only."""
    blocks = []
    if not _OCR_AVAILABLE:
        return blocks

    pil_image = _render_page_image(page, OCR_DPI)
    gray = np.array(pil_image.convert("L"))
    ocr_words = _ocr_words(pil_image, OCR_DPI)

    tables = _detect_scanned_tables(gray, ocr_words, OCR_DPI, pil_image=pil_image)
    table_bboxes_pt = [t["bbox_pt"] for t in tables]
    for t in tables:
        tb = TableBlock(kind="table", top=t["bbox_pt"][1], rows=t["rows"])
        blocks.append((tb.top, tb))

    remaining_words = [w for w in ocr_words if not any(_word_in_bbox(w, b) for b in table_bboxes_pt)]
    for tblock in _group_words_into_paragraphs(remaining_words, page_width):
        blocks.append((tblock.top, tblock))

    for region in _detect_logo_regions(pil_image, gray, ocr_words, table_bboxes_pt, OCR_DPI):
        x0, top, x1, bottom = region["bbox_pt"]
        ib = ImageBlock(kind="image", top=top, png_bytes=region["png_bytes"],
                         width_pt=(x1 - x0), height_pt=(bottom - top))
        blocks.append((ib.top, ib))

    return blocks


def extract_blocks(pdf_path: str, max_pages: int = 15, image_resolution: int = 200) -> Tuple[List, float, int]:
    """Returns (ordered_blocks, median_body_font_size, ocr_pages_used) across
    all pages, each block additionally tagged with its 1-based page number
    via block.page. ocr_pages_used counts how many pages had no real text
    layer and were read via OCR instead."""
    all_blocks = []
    sizes_seen = []
    ocr_pages_used = 0

    with pdfplumber.open(pdf_path) as pdf:
        n_pages = min(len(pdf.pages), max_pages)
        for page_idx in range(n_pages):
            page = pdf.pages[page_idx]
            page_width = page.width

            if _page_is_scanned(page):
                ocr_pages_used += 1
                for top, block in _extract_scanned_page_blocks(page, page.width, page.height):
                    if isinstance(block, TextBlock):
                        sizes_seen.append(block.avg_size)
                    all_blocks.append((top, page_idx + 1, block))
                continue

            tables = page.find_tables()
            table_bboxes = [t.bbox for t in tables]

            words = page.extract_words(extra_attrs=["size", "fontname"])
            words = [w for w in words if not _bbox_overlaps(w, table_bboxes)]
            sizes_seen.extend(w["size"] for w in words)

            text_blocks = _group_words_into_paragraphs(words, page_width)
            for b in text_blocks:
                b_page = page_idx + 1
                all_blocks.append((b.top, b_page, b))

            for t in tables:
                try:
                    rows = t.extract()
                except Exception:
                    continue
                if not rows:
                    continue
                rows = [[(c or "").strip() for c in row] for row in rows]
                tb = TableBlock(kind="table", top=t.bbox[1], rows=rows)
                all_blocks.append((tb.top, page_idx + 1, tb))

            for im in page.images:
                try:
                    x0, top, x1, bottom = im["x0"], im["top"], im["x1"], im["bottom"]
                    if x1 - x0 < 15 or bottom - top < 15:
                        continue
                    x0 = max(0, x0)
                    top = max(0, top)
                    x1 = min(page_width, x1)
                    bottom = min(page.height, bottom)
                    if x1 <= x0 or bottom <= top:
                        continue
                    cropped = page.crop((x0, top, x1, bottom))
                    pil_img = cropped.to_image(resolution=image_resolution).original
                    import io as _io
                    buf = _io.BytesIO()
                    pil_img.save(buf, format="PNG")
                    ib = ImageBlock(kind="image", top=top, png_bytes=buf.getvalue(), width_pt=(x1 - x0), height_pt=(bottom - top))
                    all_blocks.append((top, page_idx + 1, ib))
                except Exception:
                    continue

    all_blocks.sort(key=lambda tup: (tup[1], tup[0]))
    median_size = sorted(sizes_seen)[len(sizes_seen) // 2] if sizes_seen else 11.0
    return [b for _, _, b in all_blocks], median_size, ocr_pages_used


# --------------------------------------------------------------------------- #
# Rendering the extracted blocks into a new .docx
# --------------------------------------------------------------------------- #

def _style_run(run, size_pt: float, bold: bool):
    run.font.size = Pt(max(8, min(size_pt, 36)))
    run.bold = bold


def build_docx_from_blocks(blocks: List, median_size: float, output_path: str) -> None:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    for block in blocks:
        if block.kind == "paragraph":
            text = " ".join(l for l in block.lines if l.strip())
            if not text.strip():
                continue
            p = doc.add_paragraph()
            is_heading = block.avg_size > median_size + 2 or (block.bold and block.avg_size >= median_size)
            if block.centered:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            _style_run(run, block.avg_size, block.bold or is_heading)
            if is_heading:
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(6)
            else:
                p.paragraph_format.space_after = Pt(4)

        elif block.kind == "table":
            rows = block.rows
            n_rows = len(rows)
            n_cols = max(len(r) for r in rows)
            table = doc.add_table(rows=n_rows, cols=n_cols)
            try:
                table.style = "Table Grid"
            except KeyError:
                pass
            for ri, row in enumerate(rows):
                for ci in range(n_cols):
                    text = row[ci] if ci < len(row) else ""
                    cell = table.cell(ri, ci)
                    cell.text = text or ""
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.size = Pt(10)
                            if ri == 0:
                                run.bold = True
            doc.add_paragraph().paragraph_format.space_after = Pt(2)

        elif block.kind == "image":
            import io as _io
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            # Convert PDF points -> inches (72pt = 1in), cap width to page.
            width_in = min(block.width_pt / 72.0, 6.0)
            run.add_picture(_io.BytesIO(block.png_bytes), width=Inches(width_in))

    doc.save(output_path)


def convert_pdf_preserve_layout(pdf_path: str, output_path: str) -> dict:
    """Top-level entry point used by main.py for the 'preserve layout' mode."""
    blocks, median_size, ocr_pages_used = extract_blocks(pdf_path)
    build_docx_from_blocks(blocks, median_size, output_path)
    n_tables = sum(1 for b in blocks if b.kind == "table")
    n_images = sum(1 for b in blocks if b.kind == "image")
    n_paragraphs = sum(1 for b in blocks if b.kind == "paragraph")
    return {
        "paragraphs": n_paragraphs,
        "tables": n_tables,
        "images": n_images,
        "ocr_pages_used": ocr_pages_used,
        "ocr_available": _OCR_AVAILABLE,
    }
