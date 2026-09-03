"""Fill logic for the 'Biesterfeld Dubai' form (invoice draft, packing draft).

Both documents are built around a single real Word table (not free
paragraphs). Per the user's mapping: the first two table fields and the
'FROM:' exporter block are static; INV No./Dated from docs; CONSIGNEE and
NOTIFY PARTY from the customer list (same customer, since the form only
takes one at a time); BANK DETAILS static; COUNTRY OF ORIGIN OF GOODS from
the AI compare step; TERMS OF DELIVERY from docs; COUNTRY OF FINAL
DESTINATION static; Delivery and Payment from docs; the item table and its
total follow the same computed pattern as the other forms; MANUFACTURED BY
from the COA; the last line static.

Packing draft: same header block, but the item table (materials, weights,
totals) comes straight from the supplier docs rather than being computed.
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import docx

from forms_docx_utils import clone_row_for_items, replace_cell_text, sum_weight_texts, _unique_row_cells
from forms_vigorous import _customer_lines, _fmt_money, _item_total

FORM_DIR = Path(__file__).resolve().parent.parent / "forms" / "biesterfeld_dubai"


def _rows(table) -> List:
    return [_unique_row_cells(r) for r in table.rows]


def _fill_header(table, customer: dict, shipment) -> None:
    rows = _rows(table)

    # INV No. / Dated (row 2, col 1)
    replace_cell_text(rows[2][1], f"INV No.:   {shipment.invoice_no or 'N/A'}\n\nDATED:  {shipment.invoice_date or 'N/A'}")

    # CONSIGNEE / NOTIFY PARTY (row 3, col 0 / col 1) -- same customer for both
    lines = _customer_lines(customer)
    replace_cell_text(rows[3][0], "CONSIGNEE:\n" + "\n".join(lines))
    replace_cell_text(rows[3][1], "NOTIFY PARTY:\n" + "\n".join(lines))

    # Origin (row 4, col 2)
    origin = (shipment.origin or "").upper()
    replace_cell_text(rows[4][2], origin or "N/A")

    # Terms of delivery (row 5, col 2) -- from docs; keep whatever the AI found,
    # fall back to the template's own static value if nothing was extracted.
    # (No dedicated field in ShipmentData for this -- reuse total_package_description
    # only as a last resort; otherwise leave the template's own wording.)


def _fill_invoice(path: Path, out_path: Path, customer: dict, shipment) -> None:
    doc = docx.Document(str(path))
    table = doc.tables[0]
    _fill_header(table, customer, shipment)
    items = shipment.items or []

    def render_item_row(item: dict) -> List[str]:
        total = _item_total(item)
        currency = item.get("currency") or "USD"
        batches = item.get("batches") or [{"batch_no": "", "manufacturing_date": "", "expiry_date": ""}]
        batch_no = " / ".join(b.get("batch_no", "") for b in batches if b.get("batch_no"))
        mfg = " / ".join(b.get("manufacturing_date", "") for b in batches if b.get("manufacturing_date"))
        exp = " / ".join(b.get("expiry_date", "") for b in batches if b.get("expiry_date"))
        return [
            batch_no, mfg, exp,
            item.get("quantity_text", ""),
            item.get("material_name", ""),
            item.get("quantity_text", ""),
            f"{currency} {item.get('unit_price', 0):g} / {item.get('unit', '')}",
            _fmt_money(total),
        ]

    if items:
        clone_row_for_items(table.rows[9], items, render_item_row)
    else:
        for cell in _unique_row_cells(table.rows[9]):
            replace_cell_text(cell, "")

    grand_total = sum(_item_total(i) for i in items)
    for row in table.rows:
        cells = _unique_row_cells(row)
        if len(cells) == 2 and cells[0].text.strip().upper() == "TOTAL":
            replace_cell_text(cells[1], _fmt_money(grand_total))
            break

    manufacturers = sorted({i.get("manufacturer", "") for i in items if i.get("manufacturer")})
    for row in table.rows:
        cells = _unique_row_cells(row)
        if len(cells) == 1 and cells[0].text.strip().upper().startswith("MANUFACTURED BY"):
            replace_cell_text(cells[0], f"MANUFACTURED BY: {', '.join(manufacturers) or 'N/A'}")
            break

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


def _fill_packing(path: Path, out_path: Path, customer: dict, shipment) -> None:
    doc = docx.Document(str(path))
    table = doc.tables[0]
    _fill_header(table, customer, shipment)
    items = shipment.items or []

    def render_item_row(item: dict) -> List[str]:
        batches = item.get("batches") or [{"batch_no": "", "manufacturing_date": "", "expiry_date": ""}]
        batch_no = " / ".join(b.get("batch_no", "") for b in batches if b.get("batch_no"))
        mfg = " / ".join(b.get("manufacturing_date", "") for b in batches if b.get("manufacturing_date"))
        exp = " / ".join(b.get("expiry_date", "") for b in batches if b.get("expiry_date"))
        return [
            batch_no, mfg, exp,
            item.get("quantity_text", ""),
            item.get("material_name", ""),
            item.get("quantity_text", ""),
            item.get("gross_weight_text", ""),
            item.get("net_weight_text", ""),
        ]

    if items:
        clone_row_for_items(table.rows[9], items, render_item_row)
    else:
        for cell in _unique_row_cells(table.rows[9]):
            replace_cell_text(cell, "")

    package_desc = shipment.total_package_description or sum_weight_texts(
        [i.get("package_count_text", "") for i in items]
    )
    for row in table.rows:
        cells = _unique_row_cells(row)
        if len(cells) == 1 and cells[0].text.strip().upper().startswith("TOTAL PACKING"):
            replace_cell_text(cells[0], f"TOTAL PACKING: {package_desc}")
            break

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


def generate(customer: dict, shipment, output_dir: Path) -> Dict[str, Path]:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    results = {}
    results["invoice"] = output_dir / "invoice_draft.docx"
    _fill_invoice(FORM_DIR / "invoice.docx", results["invoice"], customer, shipment)

    results["packing"] = output_dir / "packing_draft.docx"
    _fill_packing(FORM_DIR / "packing.docx", results["packing"], customer, shipment)

    return results
