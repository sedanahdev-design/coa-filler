"""
Shipment-data extraction for the document-forms pipeline (Feature 3).

Takes one or more COA files (one per material/recipe -- a shipment can cover
several) plus one or more supplier documents (the source invoice/packing
paperwork to be reformatted into the chosen customer form), and asks the AI
to produce one structured "shipment" with a line-item per material, cross-
referencing both sets of documents (e.g. manufacturer comes from the COA,
batch/dates/packaging typically come from the supplier docs, per the forms'
own field-mapping rules -- but the AI is given both so it can fill gaps
either way and cross-check).
"""
from __future__ import annotations

import dataclasses
import json
from typing import List

from any_doc_extract import DocContent

DEFAULT_MODEL = "gpt-4o"

BATCH_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "batch_no": {"type": "string"},
        "manufacturing_date": {"type": "string"},
        "expiry_date": {
            "type": "string",
            "description": (
                "The batch's EXPIRY date only -- look for a label like 'Exp Date', "
                "'Expiry Date', 'EXP', or 'Use By'. This is a DIFFERENT field from a "
                "'Retest Date' / 'Re-test Date' / 'Re-qualification Date' (sometimes shown "
                "on a COA instead of, or alongside, an expiry date) -- do not substitute "
                "one for the other; report that under retest_date below instead. Prefer an "
                "explicit expiry date on the supplier/invoice/packing documents if one is "
                "shown there. If no genuine expiry date is stated anywhere for this batch "
                "(only a retest date, or nothing at all), leave this as an empty string."
            ),
        },
        "retest_date": {
            "type": "string",
            "description": (
                "The batch's Retest Date / Re-test Date / Re-qualification Date, if the "
                "source states one -- common on reference-standard or bulk-API CoAs that "
                "don't state a hard expiry at all, just a future date by which the "
                "material must be re-tested/re-qualified to confirm it's still usable. "
                "Leave as an empty string if the source states a genuine expiry_date "
                "instead (don't fill both), or if neither is present."
            ),
        },
    },
    "required": ["batch_no", "manufacturing_date", "expiry_date", "retest_date"],
}

ITEM_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "material_name": {"type": "string"},
        "manufacturer": {
            "type": "string",
            "description": (
                "Manufacturer name, taken from the COA for this material. If no COA was "
                "provided for this material, use the manufacturer stated for it on the "
                "supplier documents; if none is stated anywhere, use an empty string -- "
                "never drop the item just because its manufacturer is unknown."
            ),
        },
        "quantity_text": {"type": "string", "description": "TOTAL quantity across all this item's batches, exactly as written, e.g. '100 G' or '2,500 KG'."},
        "quantity_value": {"type": "number", "description": "Best-effort numeric TOTAL quantity (same unit as unit_price is quoted per), 0 if unknown."},
        "unit": {"type": "string", "description": "Unit the quantity/price are in, e.g. 'G', 'KG'."},
        "unit_price": {"type": "number", "description": "Price per unit in the shipment currency, 0 if not present in the source docs."},
        "currency": {"type": "string", "description": "e.g. USD. Empty string if not stated."},
        "packaging_description": {"type": "string", "description": "How it's packed, e.g. 'PACKED IN 01 BOX OF 100 G NET EACH' or '100 DRUMS OF 25 KG NET'."},
        "package_count_text": {"type": "string", "description": "Just the package count + kind, e.g. '01 BOX' or '100 DRUMS'."},
        "gross_weight_text": {"type": "string"},
        "net_weight_text": {"type": "string"},
        "batches": {
            "type": "array",
            "items": BATCH_SCHEMA,
            "description": "One or more batches/lots making up this item's total quantity. A single shipment of one material is very often split across several batch numbers -- list every distinct batch/lot found in the source documents as its own entry here, each with its own batch number, manufacturing date, and expiry date. If only one batch is mentioned, this array still has exactly one entry.",
        },
    },
    "required": [
        "material_name", "manufacturer",
        "quantity_text", "quantity_value", "unit", "unit_price", "currency",
        "packaging_description", "package_count_text", "gross_weight_text", "net_weight_text", "batches",
    ],
}

JSON_SCHEMA = {
    "name": "shipment_extraction",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "invoice_no": {
                "type": "string",
                "description": (
                    "The invoice number, from a field explicitly labeled 'Invoice No.' (or "
                    "clear equivalent) on the supplier documents. Do not confuse with a "
                    "'Buyer's Order No.', PO number, container number, or any other "
                    "reference number nearby. If no field is explicitly labeled as the "
                    "invoice number, leave this as an empty string rather than guessing."
                ),
            },
            "invoice_date": {
                "type": "string",
                "description": (
                    "The invoice's own date, from a field explicitly labeled 'Date' "
                    "immediately next to/under the invoice number (or clear equivalent). "
                    "Do not confuse with a buyer's order date, shipment date, or any other "
                    "nearby date -- some invoices show more than one 'DATE:' field for "
                    "different things. If you can't tell which date is the invoice date "
                    "with confidence, leave this as an empty string rather than guessing."
                ),
            },
            "origin": {"type": "string", "description": "Country of origin for the shipment, inferred from origin-indicating wording anywhere in the COAs or docs."},
            "origin_evidence": {"type": "string"},
            "total_package_description": {"type": "string", "description": "Overall package count/kind across all items if explicitly stated (e.g. '100 DRUMS ON 7 PALLETS'), else empty string."},
            "port_of_loading": {"type": "string", "description": "Port/place of loading, empty string if not stated."},
            "port_of_discharge": {"type": "string", "description": "Port of discharge, empty string if not stated."},
            "port_of_final_destination": {"type": "string", "description": "Port of final destination, empty string if not stated (may be the same as port_of_discharge)."},
            "country_of_final_destination": {"type": "string", "description": "Country the goods are ultimately headed to, empty string if not stated."},
            "carriage_by": {"type": "string", "description": "Mode/route of carriage if explicitly stated, e.g. 'SEA - ALEX', else empty string."},
            "items": {
                "type": "array",
                "items": ITEM_SCHEMA,
                "description": (
                    "One entry for EVERY distinct material/product line on the supplier "
                    "invoice/packing list, in the same order they appear there. The number "
                    "of entries must equal the number of material lines on the supplier "
                    "documents, regardless of how many COA files were provided -- a "
                    "material with no matching COA is still listed (fill what the supplier "
                    "documents state, leave the rest empty/0). Only fall back to one item "
                    "per COA if no supplier document lists any materials at all."
                ),
            },
        },
        "required": [
            "invoice_no", "invoice_date", "origin", "origin_evidence", "total_package_description",
            "port_of_loading", "port_of_discharge", "port_of_final_destination",
            "country_of_final_destination", "carriage_by", "items",
        ],
    },
}

SYSTEM_PROMPT = """You are a meticulous export-documentation assistant for a pharmaceutical
manufacturer. You will be shown one or more COA (Certificate of Analysis) documents -- one
per material/recipe in this shipment -- and one or more supplier/source documents (an
invoice, packing list, or similar paperwork) describing the actual shipment.

Produce ONE shipment record with one line item per material. For each material:
- material_name: prefer the supplier documents for shipment-specific values, but use the
  COA to confirm the material name and as a fallback if the supplier docs don't state it.
- manufacturer: always take this from the COA (the company that tested/certifies it).
- quantity/unit/unit_price/currency and packaging/weights: from the supplier documents,
  representing the TOTAL for this material across all its batches.
- batches: a material is very often shipped across several batch/lot numbers (e.g. one
  product, six different batch numbers each with its own manufacturing/expiry date) --
  list every distinct batch found in the source documents as its own entry in this
  item's "batches" array, in the order they appear. If the source only mentions one
  batch for a material, "batches" still has exactly one entry.
- Match each COA to the correct line item by material name.

COMPLETENESS IS CRITICAL: the supplier invoice/packing list is the authoritative list of
what is in this shipment. Before answering, count the material/product lines on the
supplier documents and make sure "items" has exactly that many entries, in the same order.
Do NOT drop, merge, or skip a material because no COA was uploaded for it, because its
manufacturer/batch details are missing, or because it looks similar to another line. A
material without a matching COA is still listed: take its name, quantity, price, packaging,
weights, and any batch/dates from the supplier documents, and leave only the fields that
truly aren't stated anywhere empty (or 0). Two lines are the same item only if they are the
same material; different materials are always separate items.

Also extract shipment-level invoice_no and invoice_date, origin (country of origin --
look for phrasing like "Made in X", "Product of X", "Origin: X", or a manufacturer's/COA's
address implying the country; report the exact phrase as origin_evidence), and any shipping
route details mentioned in the source documents: port_of_loading, port_of_discharge,
port_of_final_destination, country_of_final_destination, and carriage_by (mode/route, e.g.
"SEA - ALEXANDRIA").

If a field truly isn't present anywhere, use an empty string (or 0 for numbers) rather than
guessing."""


def _content_block(doc: DocContent, label: str) -> list:
    content = [{"type": "text", "text": f"--- Document: {label} (filename: {doc.filename}) ---"}]
    if doc.text.strip():
        if doc.text_reliable:
            content.append({"type": "text", "text": "Extracted text:\n" + doc.text})
        else:
            # See pdf_extract.looks_garbled -- some PDFs' embedded fonts make
            # pdfplumber pull out the wrong character codes for the right
            # glyph shapes, so text like a batch number can come out
            # corrupted even though it reads fine on the page image itself.
            content.append({
                "type": "text",
                "text": (
                    "Extracted text (WARNING: this PDF's text layer looks corrupted -- "
                    "individual characters may be wrong even though words look plausible, "
                    "e.g. a batch number 'YSC-1519-2508002' extracted as "
                    "'YsCˉ1519ˉ2508002'. Do NOT transcribe exact values -- batch numbers, "
                    "codes, dates -- from this text; read those off the page image(s) "
                    "below instead. Only use this for general context/wording:\n" + doc.text
                ),
            })
    for url in doc.image_data_urls:
        content.append({"type": "image_url", "image_url": {"url": url, "detail": "high"}})
    if not doc.text.strip() and not doc.image_data_urls:
        content.append({"type": "text", "text": "(no extractable content found)"})
    return content


@dataclasses.dataclass
class ShipmentData:
    invoice_no: str
    invoice_date: str
    origin: str
    origin_evidence: str
    total_package_description: str
    port_of_loading: str
    port_of_discharge: str
    port_of_final_destination: str
    country_of_final_destination: str
    carriage_by: str
    items: List[dict]
    raw: dict


def extract_shipment_data(coa_docs: List[DocContent], supplier_docs: List[DocContent], api_key: str, model: str = DEFAULT_MODEL) -> ShipmentData:
    from openai import OpenAI

    client = OpenAI(api_key=api_key, timeout=150.0, max_retries=1)

    content = []
    for i, d in enumerate(coa_docs, 1):
        content.extend(_content_block(d, f"COA #{i}"))
    for i, d in enumerate(supplier_docs, 1):
        content.extend(_content_block(d, f"Supplier document #{i}"))

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
    return ShipmentData(
        invoice_no=data["invoice_no"],
        invoice_date=data["invoice_date"],
        origin=data["origin"],
        origin_evidence=data["origin_evidence"],
        total_package_description=data["total_package_description"],
        port_of_loading=data["port_of_loading"],
        port_of_discharge=data["port_of_discharge"],
        port_of_final_destination=data["port_of_final_destination"],
        country_of_final_destination=data["country_of_final_destination"],
        carriage_by=data["carriage_by"],
        items=data["items"],
        raw=data,
    )
