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
                    "colon)."
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
                "description": "Every row of the test-results / specifications table.",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "parameter": {"type": "string", "description": "Test / parameter name, e.g. 'Appearance', 'Assay', 'Identification by IR'."},
                        "specification": {"type": "string", "description": "The specification/limit text, if shown. Empty string if not present in the source."},
                        "result": {"type": "string", "description": "The observed/reported result value for this parameter."},
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
  invent fields that aren't present. Do not include the test-results table here.
- test_results: every row of the specifications/test-results table, in the same order as
  the source document. Preserve numbers/units exactly as written.
- signature: if a handwritten signature, stamp, or company chop/seal is visible anywhere in
  the document, report which page and a normalized bounding box tightly around just that
  mark (not the whole signature block/table cell -- just the ink/stamp itself). If none is
  visible, set present to false and bbox to [0,0,0,0].

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

    client = OpenAI(api_key=api_key)
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

    page = next((p for p in extraction.pages if p.page_number == page_number), None)
    if page is None:
        return None

    x0, y0, x1, y1 = bbox
    # Clamp + sanity check
    x0, x1 = sorted((max(0.0, min(1.0, x0)), max(0.0, min(1.0, x1))))
    y0, y1 = sorted((max(0.0, min(1.0, y0)), max(0.0, min(1.0, y1))))
    if (x1 - x0) < 0.005 or (y1 - y0) < 0.005:
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
