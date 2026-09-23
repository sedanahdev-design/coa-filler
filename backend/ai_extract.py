"""
OpenAI-powered extraction of structured data out of a Certificate of Analysis PDF.

Sends the rendered page images (vision) plus any extracted raw text to an OpenAI
chat model with a strict JSON schema (Structured Outputs), so the result always
parses. Also locates a signature/stamp (if any) as a normalized bounding box on a
given page, which pdf_extract's rendered page image can then be cropped against.

Requires: OPENAI_API_KEY (env var or passed explicitly), `openai` python package.
"""
from __future__ import annotations

import io
import json
import dataclasses
from typing import List, Optional

from PIL import Image

from pdf_extract import PdfExtraction

DEFAULT_MODEL = "gpt-4o"

JSON_SCHEMA = {
    "name": "coa_extraction",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "header_fields": {
                "type": "array",
                "description": (
                    "Every labeled key/value pair that appears in the document's header, "
                    "footer, or metadata block -- e.g. Product Name, Batch No, Batch Size, "
                    "COA No, AR No, Mfg Date, Exp Date, Date of Sampling, Test Completion "
                    "Date, Quantity, Manufacturer, Remarks, Storage Condition, etc. Use the "
                    "label text exactly as it appears on the document (without the trailing "
                    "colon). This also applies to simpler 'reference standard' style CoAs "
                    "laid out as one label/value pair per row (rather than a multi-column "
                    "field grid) -- extract every one of those rows too, e.g. Lot#, Original "
                    "Lot#, Qualification Date, Re-qualification Date, Formula Weight, CAS, "
                    "COA#, Version#, Storage condition, Usage, use method. Don't skip a row "
                    "just because it isn't one of the more common field names above -- if it's "
                    "a labeled value in the header/metadata area, report it."
                ),
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "label": {"type": "string"},
                        "value": {"type": "string"},
                    },
                    "required": ["label", "value"],
                },
            },
            "test_results": {
                "type": "array",
                "description": (
                    "Every row of the test-results / specifications table. When a category "
                    "groups several named sub-items under one heading (e.g. 'Related "
                    "Substances' listing several named impurities, 'Residual Solvents' "
                    "listing several named solvents, 'Identification' tested by more than "
                    "one method such as HPLC and IR), list EVERY sub-item as its own separate "
                    "entry -- never summarize, collapse, or skip any of them, even if there "
                    "are many. If such a group's sub-items continue onto a later page image, "
                    "include all of them as one continuous set of entries in the order they "
                    "appear -- do not stop just because the page changed. Some CoAs (especially "
                    "simpler reference-standard certificates) have no multi-row specifications "
                    "table at all -- just one or two individual quality attributes stated "
                    "inline in the header area (e.g. a single row reading 'Assay (By HPLC, "
                    "%w/w): 99.3% (on dried basis)'). Still report each of those as its own "
                    "test_results entry (parameter='Assay', result='99.3% (on dried basis)') "
                    "rather than leaving test_results empty just because there's no table -- "
                    "any explicit test/assay/purity-type result belongs here even outside a "
                    "table."
                ),
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "parameter": {
                            "type": "string",
                            "description": (
                                "Test / parameter name, e.g. 'Appearance', 'Assay'. For a "
                                "named sub-item within a grouped category, include BOTH the "
                                "group's own heading and the sub-item's specific name, e.g. "
                                "'Related Substances ZP061-S010 (Isomer)', 'Residual Solvents "
                                "Ethanol', 'Identification HPLC'. Note that an Identification "
                                "test performed 'by HPLC' (or by IR, UV, etc.) is a DIFFERENT "
                                "test from 'Assay' even when Assay is also performed by HPLC -- "
                                "they test different things and must be reported as separate, "
                                "unrelated entries, never merged or cross-substituted just "
                                "because they share an analytical method name."
                            ),
                        },
                        "specification": {"type": "string", "description": "The specification/limit text, if shown. Empty string if not present in the source."},
                        "result": {"type": "string", "description": "The observed/reported result value for this parameter, transcribed exactly as shown (e.g. keep a 'RRT2.23: 0.05%' style result exactly as printed, don't simplify it to just the percentage)."},
                    },
                    "required": ["parameter", "specification", "result"],
                },
            },
            "signature": {
                "type": "object",
                "description": "Location of a handwritten signature or company stamp/chop, if any is visible.",
                "additionalProperties": False,
                "properties": {
                    "present": {"type": "boolean"},
                    "page_number": {"type": "integer", "description": "1-based page number it appears on. 0 if not present."},
                    "description": {"type": "string", "description": "Brief description, e.g. 'blue handwritten signature' or 'round red company chop'."},
                    "bbox": {
                        "type": "array",
                        "description": "Bounding box [x0, y0, x1, y1] as fractions (0-1) of the page width/height.",
                        "items": {"type": "number"},
                        "minItems": 4,
                        "maxItems": 4,
                    },
                },
                "required": ["present", "page_number", "description", "bbox"],
            },
        },
        "required": ["header_fields", "test_results", "signature"],
    },
}

SYSTEM_PROMPT = """You are a meticulous pharmaceutical QA document analyst. You will be shown
one or more page images (and possibly raw extracted text) of a Certificate of Analysis (CoA)
PDF. Extract ALL data from it precisely into the given JSON schema:

- header_fields: every labeled field in the document header/footer (product name, batch
  number, dates, quantities, report numbers, remarks, storage conditions, etc). Do not
  invent fields that aren't present. Do not include the test-results table here. Some CoAs
  -- especially simpler "reference standard" certificates -- lay out their header as a
  single column of one label/value pair per row rather than a multi-column field grid (e.g.
  Lot#, Original Lot#, Qualification Date, Re-qualification Date, Formula Weight, CAS,
  COA#, Version#, Storage condition, Usage, use method). Extract every one of those rows
  too, the same as any other header field -- don't skip a row just because its label isn't
  one of the more common ones.
- test_results: every row of the specifications/test-results table, in the same order as
  the source document. Preserve numbers/units exactly as written. A category that groups
  several named sub-items under one heading (Related Substances, Residual Solvents, an
  Identification tested by several methods, an Amino acids ratio broken down per amino
  acid, Microbial tests, etc.) must have EVERY sub-item listed as its own entry -- go
  through the whole list slowly and check you haven't skipped or merged any, especially
  when the list is long or continues onto a following page image (residual-solvent lists in
  particular are often split across a page break -- keep listing every solvent from both
  pages as one continuous set, in order). Two DIFFERENT tests that happen to share an
  analytical method name (e.g. an Identification done "by HPLC" vs. the separate "Assay"
  test, even though Assay is also run by HPLC) are NOT the same thing -- report each under
  its own real test name, never combined or substituted for the other. If the document has
  no multi-row specifications table at all -- just one or two quality attributes stated
  inline (e.g. a reference-standard CoA's single "Assay (By HPLC, %w/w): 99.3%" line) --
  still report each as its own test_results entry rather than leaving test_results empty.
- signature: if a handwritten signature, stamp, or company chop/seal is visible anywhere in
  the document, report which page and a normalized bounding box tightly around just that
  mark (not the whole signature block/table cell -- just the ink/stamp itself). If none is
  visible, set present to false and bbox to [0,0,0,0].

Some CoAs show every label and value in two languages (e.g. English then Chinese) stacked in
the same cell -- that is still ONE field/row, not two; read both lines to confirm they agree,
then report it once. Do not let the extra lines cause you to lose track of, skip, or
duplicate a row.

Be exhaustive and accurate. This data will be used to fill a regulatory document, so
transcribe values exactly as printed (do not normalize units or reformat dates)."""


@dataclasses.dataclass
class ExtractionResult:
    header_fields: List[dict]
    test_results: List[dict]
    signature: dict
    raw: dict


def _build_messages(extraction: PdfExtraction) -> list:
    if not extraction.text:
        text_note = (
            "Here is the document. Raw extracted text (may be empty if the PDF is a "
            "scan -- rely on the images in that case):\n\n"
            "(no extractable text -- this looks like a scanned document, read it from the images)"
        )
    elif not extraction.text_reliable:
        # Some PDFs (seen in practice with certain Chinese lab-report
        # generators) embed a custom/subset font where the visible glyph
        # shapes are correct but the underlying character codes are not --
        # pdfplumber then extracts text that LOOKS plausible but has individual
        # characters silently wrong (e.g. a batch number 'YSC-1519-2508002'
        # comes out as 'YsCˉ1519ˉ2508002'). Presenting that text as normally
        # trustworthy risks the model transcribing the corrupted version
        # instead of the correct one visible in the page image, which is
        # exactly what happened before this warning was added. Flag it
        # explicitly rather than silently dropping it (it can still be useful
        # for context/wording), so the model knows to verify anything
        # character-exact (batch numbers, codes, dates) against the image.
        text_note = (
            "Here is the document. WARNING: the raw extracted text below appears to come "
            "from a PDF with a broken/subset font -- individual characters in it may be "
            "WRONG even though the words look plausible (a real example: a batch number "
            "printed as 'YSC-1519-2508002' was extracted as 'YsCˉ1519ˉ2508002'). Do NOT "
            "transcribe exact values (batch numbers, codes, dates, any alphanumeric "
            "string) from this text -- read those directly off the page images instead. "
            "Only use this text for general context/wording, if at all:\n\n" + extraction.text
        )
    else:
        text_note = "Here is the document. Raw extracted text:\n\n" + extraction.text
    content = [{"type": "text", "text": text_note}]
    for page in extraction.pages:
        content.append({"type": "text", "text": f"--- Page {page.page_number} image below ---"})
        content.append({"type": "image_url", "image_url": {"url": page.as_data_url(), "detail": "high"}})
    return [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": content},
    ]


def extract_coa_data(extraction: PdfExtraction, api_key: str, model: str = DEFAULT_MODEL) -> ExtractionResult:
    """Call OpenAI to extract structured CoA data. Raises on API/parse failure."""
    from openai import OpenAI

    # Explicit timeout so a slow/stuck vision call fails with a clear,
    # catchable error within a bounded time instead of hanging until some
    # upstream proxy or load balancer gives up first (which tends to produce
    # an opaque "Bad Gateway" with no useful detail).
    client = OpenAI(api_key=api_key, timeout=150.0, max_retries=1)
    messages = _build_messages(extraction)

    response = client.chat.completions.create(
        model=model,
        messages=messages,
        response_format={"type": "json_schema", "json_schema": JSON_SCHEMA},
        temperature=0,
    )
    raw_text = response.choices[0].message.content
    data = json.loads(raw_text)
    return ExtractionResult(
        header_fields=data.get("header_fields", []),
        test_results=data.get("test_results", []),
        signature=data.get("signature", {"present": False, "page_number": 0, "description": "", "bbox": [0, 0, 0, 0]}),
        raw=data,
    )


def _whiten_to_transparent(png_bytes: bytes, threshold: int = 235) -> bytes:
    """Make the near-white background of a stamp/signature transparent so it
    doesn't blank out template text it's placed over."""
    img = Image.open(io.BytesIO(png_bytes)).convert("RGBA")
    px = img.getdata()
    img.putdata([
        (r, g, b, 0) if (r >= threshold and g >= threshold and b >= threshold) else (r, g, b, a)
        for r, g, b, a in px
    ])
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _embedded_image_for_bbox(extraction: PdfExtraction, page_number: int, bbox) -> Optional[bytes]:
    """The embedded image on `page_number` that best matches the AI's
    signature bbox, as clean PNG bytes -- or None.

    Match = the image's placement overlaps the AI box (a good share of the
    smaller of the two), or failing that its centre is very close. Skips
    full-page scans (the page itself is one big image) and images whose
    placement is unknown."""
    ax0, ay0, ax1, ay1 = bbox
    a_area = (ax1 - ax0) * (ay1 - ay0)
    acx, acy = (ax0 + ax1) / 2, (ay0 + ay1) / 2
    best, best_score = None, 0.0
    for im in getattr(extraction, "images", None) or []:
        if im.page_number != page_number or not getattr(im, "bbox", None):
            continue
        bx0, by0, bx1, by1 = im.bbox
        b_area = (bx1 - bx0) * (by1 - by0)
        if b_area <= 0 or b_area > 0.5:
            continue  # scanned page / background, not a stamp
        ix = max(0.0, min(ax1, bx1) - max(ax0, bx0))
        iy = max(0.0, min(ay1, by1) - max(ay0, by0))
        overlap = ix * iy / min(a_area, b_area)
        if overlap < 0.3:
            bcx, bcy = (bx0 + bx1) / 2, (by0 + by1) / 2
            dist = ((acx - bcx) ** 2 + (acy - bcy) ** 2) ** 0.5
            if dist > 0.08:
                continue
            overlap = 0.3 - dist  # weak match, ranked below any real overlap
        if overlap > best_score:
            best, best_score = im, overlap
    if best is None:
        return None
    try:
        return _whiten_to_transparent(best.data)
    except Exception:
        return best.data


def crop_signature(extraction: PdfExtraction, signature: dict) -> Optional[bytes]:
    """
    Given the signature dict returned by extract_coa_data, crop the corresponding
    region out of the matching rendered page and return PNG bytes. Returns None if
    no signature was detected or the page/bbox is invalid.
    """
    if not signature or not signature.get("present"):
        return None
    page_number = signature.get("page_number", 0)
    bbox = signature.get("bbox") or [0, 0, 0, 0]
    if len(bbox) != 4 or page_number <= 0:
        return None

    x0, y0, x1, y1 = bbox
    # Clamp + sanity check
    x0, x1 = sorted((max(0.0, min(1.0, x0)), max(0.0, min(1.0, x1))))
    y0, y1 = sorted((max(0.0, min(1.0, y0)), max(0.0, min(1.0, y1))))
    if (x1 - x0) < 0.005 or (y1 - y0) < 0.005:
        return None

    # Prefer the PDF's own embedded stamp/signature image when there is one
    # where the AI pointed. Cropping the rendered page instead picks up any
    # text printed over the stamp and depends on the AI's rough box -- real
    # case (Shandong Fangxing COA): the red chop sits under "Analyst",
    # "FINAL BATCH DISPOSITION" and "Approved", and the page crop came out as
    # a clipped corner of the stamp with that text across it.
    embedded = _embedded_image_for_bbox(extraction, page_number, (x0, y0, x1, y1))
    if embedded is not None:
        return embedded

    page = next((p for p in extraction.pages if p.page_number == page_number), None)
    if page is None:
        return None

    # Add a small margin so we don't clip the ink, then re-clamp.
    margin_x = (x1 - x0) * 0.08
    margin_y = (y1 - y0) * 0.15
    x0 = max(0.0, x0 - margin_x)
    x1 = min(1.0, x1 + margin_x)
    y0 = max(0.0, y0 - margin_y)
    y1 = min(1.0, y1 + margin_y)

    img = Image.open(io.BytesIO(page.png_bytes))
    w, h = img.size
    box = (int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h))
    if box[2] <= box[0] or box[3] <= box[1]:
        return None
    cropped = img.crop(box)

    buf = io.BytesIO()
    cropped.save(buf, format="PNG")
    return buf.getvalue()
