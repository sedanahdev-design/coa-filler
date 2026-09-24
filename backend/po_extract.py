"""
Purchase-Order extraction for the Shipping Instructions page.

Takes a PO (Purchase Order) PDF and pulls out the handful of fields the
Shipping Instructions sheet needs:

  * product name    -- the material itself, from the Description column of the
                       item table at the bottom of the PO, WITHOUT the trailing
                       spec/qualifier wording ("Adapalene EP with Bacterial test
                       as per typical" -> "Adapalene EP")
  * quantity + unit -- same table
  * unit price      -- same table ("2340/kg" -> 2340)
  * total price     -- same table (Total USD)
  * shipping terms  -- the PO's "Shipping terms" box, e.g. "CIF Amman by Air"
  * shipment mode   -- Air or Sea, from those same shipping terms
  * origin          -- the country in the Vendor block (India / China / ...)
  * consignee       -- the whole Consignee block, as its own lines

Same shape as ai_extract: a strict JSON schema handed to an OpenAI vision model
along with the page text AND the page images, so it works on scanned POs too.
"""
from __future__ import annotations

import dataclasses
import json
from typing import List, Optional

from pdf_extract import PdfExtraction

DEFAULT_MODEL = "gpt-4o"

JSON_SCHEMA = {
    "name": "purchase_order_extraction",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "product_name": {
                "type": "string",
                "description": (
                    "The product/material NAME ONLY, taken from the Description column of "
                    "the item table at the bottom of the PO. Keep the material name and its "
                    "pharmacopoeia/grade suffix if one is shown (EP, USP, BP, IP, JP, Ph.Eur.) "
                    "-- typically the first two words -- and drop everything after it: testing "
                    "requirements, packing notes, 'as per ...' wording, micronised/grade "
                    "qualifiers, etc. Example: 'Adapalene EP with Bacterial test as per typical' "
                    "-> 'Adapalene EP'. Example: 'Paracetamol USP micronised, packed in drums' "
                    "-> 'Paracetamol USP'."
                ),
            },
            "description_full": {
                "type": "string",
                "description": "The Description cell exactly as printed, in full. Empty string if there is none.",
            },
            "quantity": {
                "type": "number",
                "description": "The Qty from the item table as a plain number (e.g. 1, 25, 100). 0 if not stated.",
            },
            "unit": {
                "type": "string",
                "description": "The Unit from the item table exactly as printed, e.g. 'kg', 'KG', 'g', 'drums'. Empty string if not stated.",
            },
            "unit_price": {
                "type": "number",
                "description": (
                    "The Unit Price from the item table as a plain number -- strip any currency "
                    "symbol and any '/kg' style suffix ('2340/kg' -> 2340). 0 if not stated."
                ),
            },
            "total_price": {
                "type": "number",
                "description": "The line/grand Total from the item table as a plain number ('2,340' -> 2340). 0 if not stated.",
            },
            "currency": {"type": "string", "description": "Currency of the prices, e.g. 'USD'. Empty string if not stated."},
            "shipping_terms": {
                "type": "string",
                "description": (
                    "The PO's Shipping terms, assembled into one line exactly as the PO words "
                    "it, including the incoterm, the place, and the mode when the PO shows them "
                    "in that box -- e.g. 'CIF Amman by Air', 'FOB Shanghai by Sea', 'CIP Amman'. "
                    "Note the mode is often printed to the RIGHT of, or under, the incoterm and "
                    "place. Empty string if the PO states no shipping terms."
                ),
            },
            "shipment_mode": {
                "type": "string",
                "description": (
                    "How the goods travel, from the shipping terms / transport wording anywhere "
                    "on the PO. Exactly one of: 'Air', 'Sea', or '' (empty) if the PO doesn't say."
                ),
            },
            "origin_country": {
                "type": "string",
                "description": (
                    "Country of origin of the goods = the country of the VENDOR/supplier in the "
                    "Vendor block (usually the last line of its address), e.g. 'India', 'China'. "
                    "Report the country name alone. Not the buyer's or the consignee's country. "
                    "Empty string if it can't be determined."
                ),
            },
            "consignee": {
                "type": "string",
                "description": (
                    "The full Consignee block exactly as printed, INCLUDING the company name and "
                    "every address line, newline-separated, one PO line per line. Do not include "
                    "the 'Consignee:' heading itself, and do not include anything from the Buyer "
                    "or Vendor blocks. Empty string if there is no consignee block."
                ),
            },
            "po_number": {"type": "string", "description": "The PO number ('Nr'), empty string if not shown."},
            "po_date": {"type": "string", "description": "The PO date exactly as printed, empty string if not shown."},
        },
        "required": [
            "product_name", "description_full", "quantity", "unit", "unit_price", "total_price",
            "currency", "shipping_terms", "shipment_mode", "origin_country", "consignee",
            "po_number", "po_date",
        ],
    },
}

SYSTEM_PROMPT = """You are a meticulous export-documentation assistant. You are shown a
Purchase Order (PO) for pharmaceutical raw materials: its extracted text and an image of
every page. Pull out exactly the fields described in the schema and nothing else.

Read values off the page images when the extracted text looks jumbled -- PO layouts put
several boxes side by side, so the text layer often interleaves the Vendor, Buyer and
Consignee blocks. Keep those three apart: Vendor = the supplier the goods come from (its
country is the origin), Buyer = who is purchasing, Consignee = who the goods are shipped
to. If a PO lists several items, report the FIRST item's row.

Transcribe text values exactly as printed (don't translate, reword or reformat), and
report numbers as plain numbers. If a field genuinely isn't on the PO, use an empty
string (or 0 for numbers) rather than guessing."""


@dataclasses.dataclass
class PurchaseOrderData:
    product_name: str
    description_full: str
    quantity: float
    unit: str
    unit_price: float
    total_price: float
    currency: str
    shipping_terms: str
    shipment_mode: str
    origin_country: str
    consignee: str
    po_number: str
    po_date: str
    raw: dict

    def as_dict(self) -> dict:
        d = dataclasses.asdict(self)
        d.pop("raw", None)
        return d


def _normalize_mode(value: str, shipping_terms: str) -> str:
    """'Air'/'Sea' -- falls back to reading the shipping terms wording, so a PO
    that only says 'CIF Amman by Air' still drives the sheet's Air/Sea logic."""
    text = f"{value} {shipping_terms}".lower()
    if "air" in text:
        return "Air"
    if any(word in text for word in ("sea", "ocean", "vessel", "fcl", "lcl")):
        return "Sea"
    return ""


def extract_po_data(
    extraction: PdfExtraction, api_key: str, model: str = DEFAULT_MODEL
) -> PurchaseOrderData:
    from openai import OpenAI

    client = OpenAI(api_key=api_key, timeout=150.0, max_retries=1)

    content = []
    if extraction.text.strip():
        if extraction.text_reliable:
            content.append({"type": "text", "text": "Extracted text:\n" + extraction.text})
        else:
            content.append({
                "type": "text",
                "text": (
                    "Extracted text (WARNING: this PDF's text layer looks corrupted -- read "
                    "exact values off the page images below instead, and use this only for "
                    "general context):\n" + extraction.text
                ),
            })
    for page in extraction.pages:
        content.append({"type": "image_url", "image_url": {"url": page.as_data_url(), "detail": "high"}})
    if not content:
        content.append({"type": "text", "text": "(no extractable content found)"})

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": content},
        ],
        response_format={"type": "json_schema", "json_schema": JSON_SCHEMA},
        temperature=0,
    )
    data = json.loads(response.choices[0].message.content)
    return PurchaseOrderData(
        product_name=(data.get("product_name") or "").strip(),
        description_full=(data.get("description_full") or "").strip(),
        quantity=data.get("quantity") or 0,
        unit=(data.get("unit") or "").strip(),
        unit_price=data.get("unit_price") or 0,
        total_price=data.get("total_price") or 0,
        currency=(data.get("currency") or "").strip(),
        shipping_terms=(data.get("shipping_terms") or "").strip(),
        shipment_mode=_normalize_mode(data.get("shipment_mode") or "", data.get("shipping_terms") or ""),
        origin_country=(data.get("origin_country") or "").strip(),
        consignee=(data.get("consignee") or "").strip(),
        po_number=(data.get("po_number") or "").strip(),
        po_date=(data.get("po_date") or "").strip(),
        raw=data,
    )
