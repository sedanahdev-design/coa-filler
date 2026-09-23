"""
PDF extraction utilities.

Given an uploaded PDF (a Certificate of Analysis, typically 1-2 pages), this module:
  1. Extracts raw text (for text-based PDFs).
  2. Renders each page to a PNG image (base64) so it can be sent to a vision-capable
     OpenAI model -- this also covers scanned/image-only PDFs where text extraction
     yields nothing.
  3. Extracts embedded raster images (candidates for logos / signatures / stamps)
     together with metadata (page number, width, height) so the AI step and the
     docx-fill step can reason about which one is a signature/stamp.

No network access is required for anything in this file.
"""
from __future__ import annotations

import base64
import io
import dataclasses
import unicodedata
from typing import List, Optional, Tuple

import pdfplumber
import pypdfium2 as pdfium
from pypdf import PdfReader
from PIL import Image


def looks_garbled(text: str) -> bool:
    """Heuristic: does this extracted text look like it came from a PDF with a
    broken/custom embedded font, where pdfplumber pulls out the right glyph
    SHAPES but the wrong underlying Unicode codepoints (common with some
    Chinese lab-report generators)? The symptom is individual words that mix
    normal ASCII Latin letters with stray non-ASCII "letter-like" characters
    mid-word -- e.g. a batch number rendered as 'YsCˉ1519ˉ2508002' instead of
    'YSC-1519-2508002', or 'Ⅱame' instead of 'Name'. Confirmed against a real
    such document, where downstream (an AI transcription step fed this raw
    text alongside the correctly-rendered page image) reproduced the garbled
    batch number instead of reading the correct one off the image -- so
    callers should treat text flagged here as unreliable for exact
    transcription and prefer the rendered page images instead.

    Deliberately requires *mixed* ASCII+non-ASCII within tokens (not just "a
    lot of non-ASCII text"), so a genuinely Chinese/Arabic/etc-language
    document isn't mistaken for a corrupted one."""
    if not text or len(text) < 50:
        return False
    tokens = text.split()
    mixed = 0
    checked = 0
    for tok in tokens:
        has_ascii_letter = any(c.isalpha() and ord(c) < 128 for c in tok)
        has_odd_letter = any(
            c.isalpha() and ord(c) >= 128 and unicodedata.category(c).startswith("L") for c in tok
        )
        if has_ascii_letter and has_odd_letter:
            mixed += 1
        if has_ascii_letter or has_odd_letter:
            checked += 1
    if checked == 0:
        return False
    return mixed >= 5 or (checked > 20 and mixed / checked > 0.08)


@dataclasses.dataclass
class EmbeddedImage:
    index: int
    page_number: int  # 1-based
    width: int
    height: int
    data: bytes  # raw image bytes (as stored in the PDF, decoded to a standard format)
    format: str  # "png" or "jpeg"
    # Where the image is drawn on its page, as fractions of page width/height
    # with a TOP-left origin ([x0, y0, x1, y1], same convention as the AI's
    # signature bbox). None when the placement couldn't be determined.
    bbox: Optional[Tuple[float, float, float, float]] = None

    @property
    def aspect_ratio(self) -> float:
        return self.width / self.height if self.height else 0.0

    def as_base64(self) -> str:
        return base64.b64encode(self.data).decode("ascii")


@dataclasses.dataclass
class PageRender:
    page_number: int  # 1-based
    width: int
    height: int
    png_bytes: bytes

    def as_base64(self) -> str:
        return base64.b64encode(self.png_bytes).decode("ascii")

    def as_data_url(self) -> str:
        return f"data:image/png;base64,{self.as_base64()}"


@dataclasses.dataclass
class PdfExtraction:
    text: str
    pages: List[PageRender]
    images: List[EmbeddedImage]
    text_reliable: bool = True


def extract_text(pdf_path: str) -> str:
    """Extract raw text from every page, best-effort. Returns '' for scanned PDFs."""
    chunks = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                if t.strip():
                    chunks.append(t)
    except Exception:
        pass
    return "\n\n".join(chunks).strip()


def render_pages(pdf_path: str, dpi: int = 200, max_pages: int = 5) -> List[PageRender]:
    """Render each page (up to max_pages) to a PNG for vision-model consumption."""
    renders: List[PageRender] = []
    pdf = pdfium.PdfDocument(pdf_path)
    try:
        n_pages = min(len(pdf), max_pages)
        scale = dpi / 72.0
        for i in range(n_pages):
            page = pdf[i]
            bitmap = page.render(scale=scale)
            pil_image = bitmap.to_pil()
            buf = io.BytesIO()
            pil_image.save(buf, format="PNG")
            renders.append(
                PageRender(
                    page_number=i + 1,
                    width=pil_image.width,
                    height=pil_image.height,
                    png_bytes=buf.getvalue(),
                )
            )
    finally:
        pdf.close()
    return renders


def _extract_embedded_images_pdfium(pdf_path: str, min_side: int, max_count: int) -> List[EmbeddedImage]:
    """Embedded images WITH their on-page placement (needed to match a
    stamp/signature the AI located to the clean image object behind it)."""
    import pypdfium2.raw as pdfium_c

    results: List[EmbeddedImage] = []
    pdf = pdfium.PdfDocument(pdf_path)
    try:
        for page_index in range(len(pdf)):
            page = pdf[page_index]
            page_w, page_h = page.get_size()
            for obj in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_IMAGE]):
                try:
                    get_bounds = getattr(obj, "get_bounds", None) or getattr(obj, "get_pos")
                    left, bottom, right, top = get_bounds()
                    try:
                        # render=True applies soft masks / colour space / flips
                        pil_image = obj.get_bitmap(render=True).to_pil()
                    except Exception:
                        pil_image = obj.get_bitmap(render=False).to_pil()
                except Exception:
                    continue
                w, h = pil_image.size
                if w < min_side or h < min_side:
                    continue
                if pil_image.mode not in ("RGB", "RGBA", "L"):
                    pil_image = pil_image.convert("RGBA")
                buf = io.BytesIO()
                pil_image.save(buf, format="PNG")
                bbox = None
                if page_w and page_h:
                    bbox = (
                        max(0.0, min(1.0, left / page_w)),
                        max(0.0, min(1.0, 1.0 - top / page_h)),
                        max(0.0, min(1.0, right / page_w)),
                        max(0.0, min(1.0, 1.0 - bottom / page_h)),
                    )
                results.append(
                    EmbeddedImage(
                        index=len(results),
                        page_number=page_index + 1,
                        width=w,
                        height=h,
                        data=buf.getvalue(),
                        format="png",
                        bbox=bbox,
                    )
                )
                if len(results) >= max_count:
                    return results
    finally:
        pdf.close()
    return results


def extract_embedded_images(pdf_path: str, min_side: int = 25, max_count: int = 40) -> List[EmbeddedImage]:
    try:
        return _extract_embedded_images_pdfium(pdf_path, min_side, max_count)
    except Exception:
        pass  # fall back to pypdf below (no placement info)
    return _extract_embedded_images_pypdf(pdf_path, min_side, max_count)


def _extract_embedded_images_pypdf(pdf_path: str, min_side: int = 25, max_count: int = 40) -> List[EmbeddedImage]:
    """
    Pull embedded raster images out of the PDF using pypdf. These are candidates for
    logos / signatures / stamps. Very small images (icons/bullets) are skipped.
    Images are normalized to PNG bytes via Pillow so downstream code doesn't need to
    care about the original PDF filter (DCTDecode/FlateDecode/etc).
    """
    results: List[EmbeddedImage] = []
    reader = PdfReader(pdf_path)
    idx = 0
    for page_number, page in enumerate(reader.pages, start=1):
        try:
            images = page.images
        except Exception:
            images = []
        for img in images:
            try:
                pil_image = Image.open(io.BytesIO(img.data))
                pil_image.load()
            except Exception:
                continue
            w, h = pil_image.size
            if w < min_side or h < min_side:
                continue
            if pil_image.mode not in ("RGB", "RGBA", "L"):
                pil_image = pil_image.convert("RGBA")
            buf = io.BytesIO()
            pil_image.save(buf, format="PNG")
            results.append(
                EmbeddedImage(
                    index=idx,
                    page_number=page_number,
                    width=w,
                    height=h,
                    data=buf.getvalue(),
                    format="png",
                )
            )
            idx += 1
            if len(results) >= max_count:
                return results
    return results


def extract_all(pdf_path: str, dpi: int = 200) -> PdfExtraction:
    text = extract_text(pdf_path)
    return PdfExtraction(
        text=text,
        pages=render_pages(pdf_path, dpi=dpi),
        images=extract_embedded_images(pdf_path),
        text_reliable=not looks_garbled(text),
    )
