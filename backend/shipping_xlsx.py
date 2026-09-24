"""
Fill the "Shipping Instructions" Excel sheet from a Purchase Order.

The template (forms/shipping/shipping_instructions.xlsx) is a real, working
sheet: most of its intelligence lives in its OWN formulas, which must keep
working after we fill it --

    B16  =$B$8                              product name on the label block
    B22  =B12                               "Made in" follows Origin
    A26  =B14                               label word under Samples & WS
    C14  =IF($B$14="Neutral", ... , "")     neutral-label note
    A27  =IF(A6="Air Shipment","AWB:","BL Telex Release:")
    A29  / B29                              Air vs Sea shipping notes
    C33/C35/C38/C39, B38, B39
         =IF($B$12="China","... CCPIT","... Chamber of Commerce")

So this module only writes the handful of *input* cells (A6, B8, B9, B10, C9,
B11, B12, B14, B28) and leaves every formula alone.

It edits the sheet XML inside the .xlsx zip in place rather than going through
a library that rewrites the whole workbook, so styling, merged cells, column
widths, conditional formatting, print setup and the custom XML parts all come
out byte-identical.

Because Excel caches each formula's last result inside the file, the cached
values of the formulas above would still show the PREVIOUS product's wording
until Excel recalculated. We therefore recompute those cached values ourselves
(see _eval_formula) and also set fullCalcOnLoad so Excel/LibreOffice recalc the
whole sheet the moment the file is opened.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from typing import Dict, Optional, Union

from lxml import etree

FORM_DIR = Path(__file__).resolve().parent.parent / "forms" / "shipping"
TEMPLATE_PATH = FORM_DIR / "shipping_instructions.xlsx"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"m": NS_MAIN}


def _q(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


# --------------------------------------------------------------------------- #
# Tiny cell-reference helpers
# --------------------------------------------------------------------------- #

_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def _split_ref(ref: str):
    m = _REF_RE.match(ref.strip())
    if not m:
        return None, None
    return m.group(1).upper(), int(m.group(2))


def _col_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _norm_ref(ref: str) -> str:
    col, row = _split_ref(ref)
    return f"{col}{row}" if col else ref.upper()


# --------------------------------------------------------------------------- #
# Minimal formula evaluator: enough for this sheet's IF / = / * / refs
# --------------------------------------------------------------------------- #

_TOKEN_RE = re.compile(
    r"""\s*(?:
        (?P<string>"(?:[^"]|"")*")
      | (?P<number>\d+(?:\.\d+)?)
      | (?P<func>[A-Za-z][A-Za-z0-9_.]*)\s*\(
      | (?P<ref>\$?[A-Za-z]{1,3}\$?\d+)
      | (?P<op><>|<=|>=|[=<>+\-*/&(),])
    )""",
    re.VERBOSE,
)


def _tokenize(formula: str):
    pos, tokens = 0, []
    while pos < len(formula):
        m = _TOKEN_RE.match(formula, pos)
        if not m:
            if formula[pos].isspace():
                pos += 1
                continue
            raise ValueError(f"cannot tokenize at {formula[pos:][:20]!r}")
        pos = m.end()
        if m.lastgroup == "string":
            tokens.append(("string", m.group("string")[1:-1].replace('""', '"')))
        elif m.lastgroup == "number":
            tokens.append(("number", float(m.group("number"))))
        elif m.lastgroup == "func":
            tokens.append(("func", m.group("func").upper()))
        elif m.lastgroup == "ref":
            tokens.append(("ref", _norm_ref(m.group("ref"))))
        else:
            tokens.append(("op", m.group("op")))
    return tokens


class _Parser:
    """Supports: IF(), cell refs, string/number literals, = <> < > <= >=,
    + - * / and & (concatenation). Anything else raises, and the caller then
    just drops the cached value and lets Excel recalculate on open."""

    def __init__(self, tokens, values: Dict[str, object]):
        self.t, self.i, self.values = tokens, 0, values

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expect_op(self, op: str):
        kind, val = self.take()
        if kind != "op" or val != op:
            raise ValueError(f"expected {op!r}")

    def parse(self):
        value = self.comparison()
        if self.i != len(self.t):
            raise ValueError("trailing tokens")
        return value

    def comparison(self):
        left = self.concat()
        kind, val = self.peek()
        if kind == "op" and val in ("=", "<>", "<", ">", "<=", ">="):
            self.take()
            right = self.concat()
            return _compare(left, val, right)
        return left

    def concat(self):
        left = self.additive()
        while self.peek() == ("op", "&"):
            self.take()
            left = _as_text(left) + _as_text(self.additive())
        return left

    def additive(self):
        left = self.term()
        while self.peek()[0] == "op" and self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            right = self.term()
            left = _as_number(left) + _as_number(right) if op == "+" else _as_number(left) - _as_number(right)
        return left

    def term(self):
        left = self.unary()
        while self.peek()[0] == "op" and self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            right = self.unary()
            if op == "*":
                left = _as_number(left) * _as_number(right)
            else:
                divisor = _as_number(right)
                if divisor == 0:
                    raise ValueError("division by zero")
                left = _as_number(left) / divisor
        return left

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return -_as_number(self.unary())
        return self.primary()

    def primary(self):
        kind, val = self.take()
        if kind in ("string", "number"):
            return val
        if kind == "ref":
            return self.values.get(val, "")
        if kind == "op" and val == "(":
            inner = self.comparison()
            self.expect_op(")")
            return inner
        if kind == "func":
            args = []
            if self.peek() != ("op", ")"):
                args.append(self.comparison())
                while self.peek() == ("op", ","):
                    self.take()
                    args.append(self.comparison())
            self.expect_op(")")
            if val == "IF" and len(args) >= 2:
                cond = args[0]
                truthy = cond is True or (isinstance(cond, (int, float)) and cond != 0)
                if truthy:
                    return args[1]
                return args[2] if len(args) > 2 else False
            raise ValueError(f"unsupported function {val}")
        raise ValueError(f"unexpected token {kind}:{val}")


def _as_text(value) -> str:
    if value is None or value is False and not isinstance(value, bool):
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _as_number(value) -> float:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    return float(text) if text else 0.0


def _compare(left, op: str, right) -> bool:
    if isinstance(left, str) or isinstance(right, str):
        # Excel's text comparison is case-insensitive ("china" = "China").
        a, b = _as_text(left).strip().lower(), _as_text(right).strip().lower()
    else:
        a, b = _as_number(left), _as_number(right)
    return {
        "=": a == b,
        "<>": a != b,
        "<": a < b,
        ">": a > b,
        "<=": a <= b,
        ">=": a >= b,
    }[op]


def _eval_formula(formula: str, values: Dict[str, object]):
    """Evaluate one of this sheet's formulas, or return None if it uses
    anything this little evaluator doesn't know."""
    try:
        return _Parser(_tokenize(formula.lstrip("=")), values).parse()
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# Reading / writing the sheet XML
# --------------------------------------------------------------------------- #


def _load_shared_strings(zf: zipfile.ZipFile):
    try:
        root = etree.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall(_q("si")):
        out.append("".join(t.text or "" for t in si.iter(_q("t"))))
    return out


def _cell_value(cell, shared):
    ctype = cell.get("t")
    if ctype == "inlineStr":
        is_el = cell.find(_q("is"))
        return "".join(t.text or "" for t in is_el.iter(_q("t"))) if is_el is not None else ""
    v = cell.find(_q("v"))
    if v is None or v.text is None:
        return ""
    if ctype == "s":
        try:
            return shared[int(v.text)]
        except (ValueError, IndexError):
            return ""
    if ctype in ("str", "e"):
        return v.text
    if ctype == "b":
        return v.text == "1"
    try:
        return float(v.text)
    except ValueError:
        return v.text


def _get_or_create_cell(sheet_data, ref: str):
    col, row_num = _split_ref(ref)
    if col is None:
        raise ValueError(f"bad cell reference {ref!r}")
    row = None
    for r in sheet_data.findall(_q("row")):
        r_num = int(r.get("r") or 0)
        if r_num == row_num:
            row = r
            break
        if r_num > row_num:
            row = etree.Element(_q("row"), r=str(row_num))
            r.addprevious(row)
            break
    if row is None:
        row = etree.SubElement(sheet_data, _q("row"), r=str(row_num))

    target_idx = _col_index(col)
    for c in row.findall(_q("c")):
        c_col, _ = _split_ref(c.get("r") or "")
        if c_col == col:
            return c
        if c_col and _col_index(c_col) > target_idx:
            cell = etree.Element(_q("c"), r=ref)
            c.addprevious(cell)
            return cell
    return etree.SubElement(row, _q("c"), r=ref)


def _set_cell(sheet_data, ref: str, value: Union[str, int, float, None]) -> None:
    """Write a literal value, keeping the cell's style (s= attribute)."""
    cell = _get_or_create_cell(sheet_data, ref)
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    if value is None or value == "":
        return  # empty cell, style preserved
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        v = etree.SubElement(cell, _q("v"))
        v.text = repr(float(value)) if isinstance(value, float) and not float(value).is_integer() else str(int(value))
        return
    cell.set("t", "inlineStr")
    is_el = etree.SubElement(cell, _q("is"))
    t = etree.SubElement(is_el, _q("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = str(value)


def _set_formula_cached(cell, result) -> None:
    """Refresh a formula cell's cached result (the <v> Excel shows before it
    recalculates), leaving its <f> formula untouched."""
    for child in list(cell):
        if child.tag != _q("f"):
            cell.remove(child)
    cell.attrib.pop("t", None)
    if result is None or result == "":
        return
    if isinstance(result, bool):
        cell.set("t", "b")
        etree.SubElement(cell, _q("v")).text = "1" if result else "0"
    elif isinstance(result, (int, float)):
        v = etree.SubElement(cell, _q("v"))
        v.text = repr(float(result)) if not float(result).is_integer() else str(int(result))
    else:
        cell.set("t", "str")
        etree.SubElement(cell, _q("v")).text = str(result)


def _recalculate(sheet_data, shared) -> None:
    """Recompute every formula cell's cached value from the sheet's current
    literal values, repeatedly until nothing changes (so chains like
    B12 -> B22 and A6 -> A27 -> B29 settle)."""
    cells = {}
    formulas = []
    for row in sheet_data.findall(_q("row")):
        for c in row.findall(_q("c")):
            ref = _norm_ref(c.get("r") or "")
            f = c.find(_q("f"))
            if f is not None and (f.text or "").strip():
                formulas.append((ref, c, f.text.strip()))
                cells[ref] = _cell_value(c, shared)
            else:
                cells[ref] = _cell_value(c, shared)

    for _pass in range(6):
        changed = False
        for ref, _c, formula in formulas:
            result = _eval_formula(formula, cells)
            if result is None:
                result = ""
            if cells.get(ref) != result:
                cells[ref] = result
                changed = True
        if not changed:
            break

    for ref, cell, formula in formulas:
        _set_formula_cached(cell, cells.get(ref))


def _force_full_recalc(zf_bytes: Dict[str, bytes]) -> None:
    """Tell Excel/LibreOffice to recalculate everything on open, so the sheet
    is right even if our own cached values were left blank."""
    data = zf_bytes.get("xl/workbook.xml")
    if not data:
        return
    root = etree.fromstring(data)
    calc = root.find(_q("calcPr"))
    if calc is None:
        calc = etree.SubElement(root, _q("calcPr"))
    calc.set("fullCalcOnLoad", "1")
    zf_bytes["xl/workbook.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# --------------------------------------------------------------------------- #
# The Shipping Instructions sheet itself
# --------------------------------------------------------------------------- #

CELLS = {
    "shipment_mode": "A6",     # "Air Shipment" / "Sea Shipment"
    "product": "B8",
    "quantity": "B9",
    "price_label": "A10",      # "Price per KG:" -- follows the PO's unit
    "price_per_unit": "B10",
    "incoterms": "C9",         # "Inco terms: CIF Amman by Air"
    "total": "C10",            # formula =B9*B10 (kept unless the PO disagrees)
    "specs": "B11",
    "origin": "B12",
    "label_type": "B14",
    "consignee": "B28",
}

DEFAULTS = {
    "label_type": "full",
    "shipment_mode": "Air",
}


def _clean_number(value) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"[^\d.\-]", "", str(value))
    try:
        return float(text)
    except ValueError:
        return 0.0


def build_values(po, overrides: Optional[dict] = None) -> dict:
    """The sheet's input values from a PurchaseOrderData (po_extract), with any
    user edits from the review form applied on top."""
    values = {
        "shipment_mode": po.shipment_mode or DEFAULTS["shipment_mode"],
        "product": po.product_name,
        "quantity": po.quantity,
        "unit": po.unit or "KG",
        "price_per_unit": po.unit_price,
        "total": po.total_price,
        "incoterms": po.shipping_terms,
        "origin": po.origin_country,
        "consignee": po.consignee,
        "label_type": DEFAULTS["label_type"],
        "specs": "",  # keeps the template's own wording unless the user types something
    }
    for key, value in (overrides or {}).items():
        if key in values and value not in (None, ""):
            values[key] = value
    return values


def fill_shipping_instructions(values: dict, out_path, template_path=None) -> dict:
    """Write the Shipping Instructions workbook. Returns a small report:
    {"cells": {ref: written value}, "notes": [str]}."""
    template_path = Path(template_path or TEMPLATE_PATH)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    notes = []

    parts: Dict[str, bytes] = {}
    order = []
    with zipfile.ZipFile(template_path) as zf:
        shared = _load_shared_strings(zf)
        for info in zf.infolist():
            order.append(info)
            parts[info.filename] = zf.read(info.filename)

    sheet_name = "xl/worksheets/sheet1.xml"
    if sheet_name not in parts:
        sheet_name = next(n for n in parts if n.startswith("xl/worksheets/sheet"))
    root = etree.fromstring(parts[sheet_name])
    sheet_data = root.find(_q("sheetData"))

    mode = (values.get("shipment_mode") or DEFAULTS["shipment_mode"]).strip()
    mode = "Sea" if mode.lower().startswith("sea") else "Air"
    # NB single space: the sheet's own A27/A29 formulas test A6="Air Shipment",
    # and the template shipped with "Air  Shipment" (two spaces), which made
    # those tests fail and show the Sea wording on an air shipment.
    written = {CELLS["shipment_mode"]: f"{mode} Shipment"}

    product = (values.get("product") or "").strip()
    written[CELLS["product"]] = product

    quantity = _clean_number(values.get("quantity"))
    written[CELLS["quantity"]] = quantity

    unit = (values.get("unit") or "KG").strip()
    written[CELLS["price_label"]] = f"Price per {unit.upper()}:"

    price = _clean_number(values.get("price_per_unit"))
    written[CELLS["price_per_unit"]] = price

    incoterms = (values.get("incoterms") or "").strip()
    written[CELLS["incoterms"]] = f"Inco terms: {incoterms}" if incoterms else "Inco terms:"

    origin = (values.get("origin") or "").strip()
    written[CELLS["origin"]] = origin

    label_type = (values.get("label_type") or DEFAULTS["label_type"]).strip()
    written[CELLS["label_type"]] = "Neutral" if label_type.lower().startswith("neutral") else "full"

    consignee = (values.get("consignee") or "").strip()
    written[CELLS["consignee"]] = consignee

    specs = (values.get("specs") or "").strip()
    if specs:
        written[CELLS["specs"]] = specs

    for ref, value in written.items():
        _set_cell(sheet_data, ref, value)

    # Total: keep the sheet's own =B9*B10 formula (so editing qty or price still
    # updates it) unless the PO's stated total disagrees with that product --
    # then the PO wins and we say so.
    po_total = _clean_number(values.get("total"))
    computed = quantity * price
    total_cell = _get_or_create_cell(sheet_data, CELLS["total"])
    has_formula = total_cell.find(_q("f")) is not None
    if po_total and abs(po_total - computed) > 0.01:
        _set_cell(sheet_data, CELLS["total"], po_total)
        notes.append(
            f"The PO's total ({po_total:,.2f}) doesn't match quantity x price ({computed:,.2f}), "
            f"so the total cell holds the PO's number instead of the formula."
        )
    elif not has_formula:
        _set_cell(sheet_data, CELLS["total"], computed)

    _recalculate(sheet_data, shared)
    parts[sheet_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _force_full_recalc(parts)

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for info in order:
            zf.writestr(info.filename, parts[info.filename])

    return {"cells": written, "notes": notes}


def suggested_filename(values: dict) -> str:
    product = re.sub(r"[^A-Za-z0-9 _-]", "", (values.get("product") or "").strip()) or "Product"
    return f"Shipping Instructions - {product}.xlsx"
