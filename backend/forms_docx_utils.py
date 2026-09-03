"""Shared low-level python-docx helpers for the bespoke per-form fill modules
(forms_vigorous.py etc). Builds on the same span-replacement technique used in
docx_fill.py so existing character formatting is preserved wherever possible."""
from __future__ import annotations

import copy
import re
from typing import Callable, List, Optional

from docx.table import _Cell, _Row, Table
from docx.text.paragraph import Paragraph

from docx_fill import replace_paragraph_span, replace_cell_text, _unique_row_cells  # noqa: F401


def find_paragraph(paragraphs: List[Paragraph], contains: str) -> Optional[Paragraph]:
    """First paragraph whose text contains `contains` (plain substring, case-sensitive)."""
    for p in paragraphs:
        if contains in p.text:
            return p
    return None


def replace_regex_span(paragraph: Paragraph, pattern: str, new_text: str, group: int = 0, flags=0) -> bool:
    """Find `pattern` in paragraph.text and replace the given capture group's
    span with new_text, preserving surrounding text/formatting. Returns True
    if a replacement was made."""
    m = re.search(pattern, paragraph.text, flags)
    if not m:
        return False
    replace_paragraph_span(paragraph, m.start(group), m.end(group), new_text)
    return True


def set_paragraph_full_text(paragraph: Paragraph, text: str) -> None:
    """Overwrite a paragraph's entire visible text, keeping the first run's
    formatting and clearing any others (same technique as replace_cell_text)."""
    runs = paragraph.runs
    if not runs:
        paragraph.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def clone_paragraph_for_items(paragraph: Paragraph, items: list, render_fn: Callable) -> List[Paragraph]:
    """Given a template paragraph representing ONE item/line, fill it with
    items[0] and insert a deep-copied clone (same formatting) right after for
    each remaining item. Returns the list of paragraphs used (including the
    original), in order. No-ops (clears the paragraph) if items is empty."""
    if not items:
        set_paragraph_full_text(paragraph, "")
        return [paragraph]

    set_paragraph_full_text(paragraph, render_fn(items[0]))
    result = [paragraph]
    insert_after_xml = paragraph._p
    for item in items[1:]:
        new_xml = copy.deepcopy(paragraph._p)
        insert_after_xml.addnext(new_xml)
        insert_after_xml = new_xml
        new_para = Paragraph(new_xml, paragraph._parent)
        set_paragraph_full_text(new_para, render_fn(item))
        result.append(new_para)
    return result


# Mass units normalized to a common base (grams) so mixed-unit item lists
# (e.g. one recipe in G, another in KG) sum into a single clean total instead
# of an unhelpful "100.00 G + 25.00 KG" string.
_MASS_TO_GRAMS = {"MG": 0.001, "G": 1.0, "GM": 1.0, "GRAM": 1.0, "KG": 1000.0, "KGS": 1000.0}


def sum_weight_texts(texts: List[str]) -> str:
    """Best-effort sum of weight strings like '12.00 KG' / '2,500.00 KGS' /
    '100 G'. Recognized mass units are normalized to grams and summed into
    one clean total (displayed in KG if >=1000 G, else G); anything with a
    non-mass or unrecognized unit is grouped and summed separately by unit;
    unparsed strings are appended as-is."""
    grams_total = 0.0
    has_mass = False
    other_totals: dict[str, float] = {}
    unparsed: List[str] = []
    for t in texts:
        t = (t or "").strip()
        if not t:
            continue
        m = re.match(r"^([\d,]+(?:\.\d+)?)\s*([A-Za-z\.]+)", t)
        if not m:
            unparsed.append(t)
            continue
        num = float(m.group(1).replace(",", ""))
        unit = m.group(2).upper().rstrip("S").rstrip(".")
        mass_factor = _MASS_TO_GRAMS.get(unit) or _MASS_TO_GRAMS.get(unit + "S")
        if mass_factor:
            grams_total += num * mass_factor
            has_mass = True
        else:
            other_totals[unit] = other_totals.get(unit, 0.0) + num

    parts = []
    if has_mass:
        if grams_total >= 1000:
            parts.append(f"{grams_total / 1000:,.2f} KG")
        else:
            parts.append(f"{grams_total:,.2f} G")
    parts.extend(f"{v:,.2f} {u}" for u, v in other_totals.items())
    parts.extend(unparsed)
    return " + ".join(parts) if parts else ""


def sum_numeric(values: List[float]) -> float:
    return sum(v or 0 for v in values)


def fill_multiline_block(paragraphs: List[Paragraph], lines: List[str]) -> None:
    """Fill a fixed run of consecutive template paragraphs (e.g. a customer's
    address block) with `lines`. If there are fewer lines than paragraph
    slots, extra slots are cleared; if there are more lines than slots, the
    overflow is joined (comma-separated) into the last slot so nothing is
    lost."""
    lines = [l.strip() for l in lines if l and l.strip()]
    n = len(paragraphs)
    if not lines:
        for p in paragraphs:
            set_paragraph_full_text(p, "")
        return
    if len(lines) <= n:
        for i, p in enumerate(paragraphs):
            set_paragraph_full_text(p, lines[i] if i < len(lines) else "")
    else:
        for i in range(n - 1):
            set_paragraph_full_text(paragraphs[i], lines[i])
        set_paragraph_full_text(paragraphs[n - 1], ", ".join(lines[n - 1:]))


def clone_paragraph_group_for_items(paragraphs: List[Paragraph], items: list, render_fns: List[Callable]) -> List[List[Paragraph]]:
    """Like clone_paragraph_for_items but for a BLOCK of consecutive template
    paragraphs that together describe one item (e.g. a 'PRODUCT:' line plus a
    'PACKED IN:' line). The whole block is cloned as a unit per item so
    multi-paragraph item descriptions stay interleaved correctly (item1's
    lines, then item2's lines) instead of all-of-paragraph-1 then
    all-of-paragraph-2. `render_fns` must be the same length as `paragraphs`,
    each render_fns[i](item) -> text for paragraphs[i]. Returns a list of
    groups (each a list of paragraphs), one group per item."""
    if not items:
        for p in paragraphs:
            set_paragraph_full_text(p, "")
        return [list(paragraphs)]

    for p, fn in zip(paragraphs, render_fns):
        set_paragraph_full_text(p, fn(items[0]))
    groups = [list(paragraphs)]
    insert_after_xml = paragraphs[-1]._p
    for item in items[1:]:
        new_group = []
        for p, fn in zip(paragraphs, render_fns):
            new_xml = copy.deepcopy(p._p)
            insert_after_xml.addnext(new_xml)
            insert_after_xml = new_xml
            new_p = Paragraph(new_xml, p._parent)
            set_paragraph_full_text(new_p, fn(item))
            new_group.append(new_p)
        groups.append(new_group)
    return groups


def clone_row_for_items(row: _Row, items: list, render_fn: Callable) -> List[_Row]:
    """Given a template table row representing ONE item, fill it with items[0]
    and insert a deep-copied clone (same formatting/cell count) right after
    for each remaining item. render_fn(item) must return a list of cell
    strings (one per unique cell in the row, left-to-right). Returns the list
    of Row objects used, in order."""
    cells0 = _unique_row_cells(row)
    if not items:
        for cell in cells0:
            replace_cell_text(cell, "")
        return [row]

    def apply(r: _Row, item) -> None:
        cells = _unique_row_cells(r)
        values = render_fn(item)
        for cell, value in zip(cells, values):
            replace_cell_text(cell, value)

    apply(row, items[0])
    result = [row]
    insert_after_xml = row._tr
    for item in items[1:]:
        new_tr = copy.deepcopy(row._tr)
        insert_after_xml.addnext(new_tr)
        insert_after_xml = new_tr
        new_row = _Row(new_tr, row._parent)
        apply(new_row, item)
        result.append(new_row)
    return result
