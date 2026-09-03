"""
COA vs. supporting-document comparison.

Given two documents (a Certificate of Analysis and any other document -- an
image, a Word file, or another PDF -- e.g. an invoice, packing list, or a
supplier's own COA), asks the AI to extract a fixed set of key fields from
each and cross-check them:

    - material / product name
    - recipe / batch number
    - manufacturing date
    - expiry date
    - country of origin (inferred from origin-indicating keywords/phrasing --
      "Made in", "Product of", "Origin:", a stamped country name, etc.)

Everything else either document contains is returned as unstructured
"additional info" (not compared, just surfaced for context). The result
states, per field, whether the two documents agree, and if not, what each one
actually says and which document holds which value.
"""
from __future__ import annotations

import dataclasses
import json
from typing import List

from any_doc_extract import DocContent

DEFAULT_MODEL = "gpt-4o"

COMPARE_FIELDS = [
    "material_name",
    "recipe_or_batch_number",
    "manufacturing_date",
    "expiry_date",
    "origin",
]

FIELD_LABELS = {
    "material_name": "Material / Product Name",
    "recipe_or_batch_number": "Recipe / Batch Number",
    "manufacturing_date": "Manufacturing Date",
    "expiry_date": "Expiry Date",
    "origin": "Country of Origin",
}

_DOC_FIELDS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "material_name": {"type": "string"},
        "recipe_or_batch_number": {"type": "string"},
        "manufacturing_date": {"type": "string"},
        "expiry_date": {"type": "string"},
        "origin": {"type": "string", "description": "Country of origin, inferred from any origin-indicating wording (e.g. 'Made in', 'Product of', 'Origin:', a stamp/seal naming a country). Empty string if genuinely not determinable."},
        "origin_evidence": {"type": "string", "description": "The exact phrase/keyword in the document that the origin was inferred from. Empty string if origin is empty."},
        "additional_info": {
            "type": "array",
            "description": "Every other notable labeled field in the document not covered above (report number, quantity, supplier, test results summary, etc).",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {"label": {"type": "string"}, "value": {"type": "string"}},
                "required": ["label", "value"],
            },
        },
    },
    "required": ["material_name", "recipe_or_batch_number", "manufacturing_date", "expiry_date", "origin", "origin_evidence", "additional_info"],
}

JSON_SCHEMA = {
    "name": "coa_comparison",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "coa": _DOC_FIELDS_SCHEMA,
            "other": _DOC_FIELDS_SCHEMA,
            "comparison": {
                "type": "array",
                "description": "One entry for each of: material_name, recipe_or_batch_number, manufacturing_date, expiry_date, origin -- in that order.",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "field": {"type": "string", "enum": COMPARE_FIELDS},
                        "status": {"type": "string", "enum": ["match", "mismatch", "missing_in_coa", "missing_in_other"]},
                        "explanation": {"type": "string", "description": "Brief note, especially for mismatches: what each document says and why they disagree."},
                    },
                    "required": ["field", "status", "explanation"],
                },
            },
            "overall_status": {"type": "string", "enum": ["match", "mismatch"]},
            "summary": {"type": "string", "description": "One or two sentence plain-language summary of the result."},
        },
        "required": ["coa", "other", "comparison", "overall_status", "summary"],
    },
}

SYSTEM_PROMPT = """You are a meticulous pharmaceutical QA auditor. You will be shown two
documents: one labeled COA (a Certificate of Analysis) and one labeled OTHER (any
supporting document -- could be an image, a Word document, or another PDF, e.g. an
invoice, packing list, or a different COA).

From EACH document independently, extract:
- material_name: the product/material name
- recipe_or_batch_number: batch/lot/recipe number
- manufacturing_date, expiry_date
- origin: the country of origin. This is often not in a field literally labeled
  "Origin" -- look for phrasing like "Made in X", "Product of X", "Country of
  Origin: X", a company address whose country implies origin, or a stamp/seal
  naming a country. Record the exact phrase you used as evidence in origin_evidence.
- additional_info: every other labeled field present (report/invoice number, quantity,
  supplier/manufacturer name, storage conditions, test results, etc) -- these are NOT
  compared, just listed for context.

Then compare the two documents field-by-field for material_name, recipe_or_batch_number,
manufacturing_date, expiry_date, and origin. Two values that are the same fact written
differently (e.g. "16.07.2026" vs "July 16, 2026", or "India" vs "Made in India") should
be marked "match". Only mark "mismatch" when the underlying facts actually differ. Use
"missing_in_coa" / "missing_in_other" when one document simply doesn't state that field.
For every mismatch, explain clearly what each document says so the reader knows exactly
where the discrepancy is."""


def _content_block(doc: DocContent, role_label: str) -> list:
    content = [{"type": "text", "text": f"--- Document: {role_label} (filename: {doc.filename}) ---"}]
    if doc.text.strip():
        content.append({"type": "text", "text": "Extracted text:\n" + doc.text})
    elif not doc.image_data_urls:
        content.append({"type": "text", "text": "(no extractable text or images found in this file)"})
    for url in doc.image_data_urls:
        content.append({"type": "image_url", "image_url": {"url": url, "detail": "high"}})
    return content


@dataclasses.dataclass
class ComparisonResult:
    coa: dict
    other: dict
    comparison: List[dict]
    overall_status: str
    summary: str
    raw: dict


def compare_documents(coa_doc: DocContent, other_doc: DocContent, api_key: str, model: str = DEFAULT_MODEL) -> ComparisonResult:
    from openai import OpenAI

    client = OpenAI(api_key=api_key)

    content = []
    content.extend(_content_block(coa_doc, "COA"))
    content.extend(_content_block(other_doc, "OTHER"))

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": content},
    ]

    response = client.chat.completions.create(
        model=model,
        messages=messages,
        response_format={"type": "json_schema", "json_schema": JSON_SCHEMA},
        temperature=0,
    )
    data = json.loads(response.choices[0].message.content)
    return ComparisonResult(
        coa=data["coa"],
        other=data["other"],
        comparison=data["comparison"],
        overall_status=data["overall_status"],
        summary=data["summary"],
        raw=data,
    )
