"""
Generic Certificate-of-Analysis docx fill engine.

Templates in templates/ are real exemplar COA documents (not Jinja-style blank
forms) with company-specific layout, logo, and table structure already baked in.
This module treats each template as a *layout* to reuse: it locates "Label: value"
fields anywhere in the document (body paragraphs or table cells), fuzzy-matches
them against AI-extracted header fields and overwrites just the value while
preserving the run's formatting; locates the test-results table and updates the
Result column row-by-row by matching on parameter name (appending new rows for
tests present in the source PDF but not in the template); and drops the
cropped signature/stamp image from the source PDF into a detected signature
anchor, without touching the template's own logo.

Nothing here calls the network -- it only depends on python-docx + Pillow.
"""
from __future__ import annotations

import copy
import io
import re
import difflib
import dataclasses
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.table import Table, _Cell, _Row
from docx.text.paragraph import Paragraph
from PIL import Image


# --------------------------------------------------------------------------- #
# Text normalization helpers
# --------------------------------------------------------------------------- #

_WS_RE = re.compile(r"\s+")
_PUNCT_RE = re.compile(r"[^a-z0-9 ]")


def normalize(text: str) -> str:
    text = (text or "").lower().strip()
    text = text.replace("\n", " ").replace("\t", " ")
    text = _PUNCT_RE.sub(" ", text)
    text = _WS_RE.sub(" ", text).strip()
    return text


# Common CoA header-field synonyms across manufacturers. Applied on top of
# normalize() before comparing labels, so e.g. a source PDF's "Date of Expiry"
# lines up with a template's "Exp Date", and "Net weight" lines up with
# "Quantity". Word-order/synonym swaps like this defeat plain character-level
# fuzzy matching, so we canonicalize known synonym tokens first.
_TOKEN_SYNONYMS = {
    "expiry": "exp", "expire": "exp", "expiration": "exp",
    "manufacture": "mfg", "manufacturing": "mfg", "manufactured": "mfg",
    "quantity": "qty", "weight": "qty", "netweight": "qty",
    "number": "no", "numbers": "no",
    "lot": "batch",
    "date": "dt",
    "analytical": "ar", "report": "ar",
    "certificate": "coa", "analysis": "coa",
}
_STOPWORDS = {"of", "the", "a", "an"}


def _canonicalize(norm_text: str) -> str:
    tokens = [t for t in norm_text.split() if t not in _STOPWORDS]
    tokens = [_TOKEN_SYNONYMS.get(t, t) for t in tokens]
    return " ".join(sorted(tokens))


def best_match(query: str, candidates: Dict[str, str], cutoff: float = 0.72) -> Optional[str]:
    """Return the candidates-key that best matches query, or None.

    Order matters here: canonicalized token matching (synonym + word-order aware,
    e.g. 'manufacture date' <-> 'date of manufacture') is a much more precise
    signal than raw character similarity, so it's tried *before* difflib -- doing
    it after would let a merely-character-similar wrong candidate (e.g.
    "Manufacturer's Name" sharing a "manufactur..." prefix with "Manufacture
    Date") win before the semantically-correct one is even considered."""
    if not query:
        return None
    if query in candidates:
        return query
    keys = list(candidates.keys())

    # 1) canonicalized token match (synonyms + reordering + stopword-insensitive)
    q_canon = _canonicalize(query)
    if q_canon:
        for k in keys:
            if _canonicalize(k) == q_canon:
                return k

    # 2) character-level fuzzy match
    matches = difflib.get_close_matches(query, keys, n=1, cutoff=cutoff)
    if matches:
        return matches[0]

    # 3) token containment either direction (handles "batch no" vs "batch number")
    q_tokens = set(query.split())
    best_key, best_score = None, 0.0
    for k in keys:
        k_tokens = set(k.split())
        if not q_tokens or not k_tokens:
            continue
        overlap = len(q_tokens & k_tokens) / max(1, min(len(q_tokens), len(k_tokens)))
        if overlap > best_score and overlap >= 0.8:
            best_score = overlap
            best_key = k
    if best_key is not None:
        return best_key

    # 4) "quantity family" fallback: a template's "Net weight" field and a
    # source document's "Batch Quantity" (or similar) are, in practice, very
    # often the same number under a different name for a single-batch CoA --
    # per an explicit product rule, not a general synonym. Only kick in when
    # the query itself is fundamentally a quantity/weight field (canonicalizes
    # to contain the "qty" token) AND there's exactly one unmatched candidate
    # that's also in that family, so this never guesses between several
    # different quantity-like fields that genuinely differ.
    if "qty" in q_canon.split():
        qty_candidates = [k for k in keys if "qty" in _canonicalize(k).split()]
        if len(qty_candidates) == 1:
            return qty_candidates[0]
    return None


# --------------------------------------------------------------------------- #
# Run-level span replacement (preserves formatting as much as possible)
# --------------------------------------------------------------------------- #

def _paragraph_run_spans(paragraph: Paragraph) -> List[Tuple[object, int, int]]:
    """Returns [(run, start, end), ...] offsets into paragraph.text (run.text may
    include '\\n'/'\\t' for breaks/tabs, consistent with paragraph.text)."""
    spans = []
    pos = 0
    for run in paragraph.runs:
        t = run.text or ""
        spans.append((run, pos, pos + len(t)))
        pos += len(t)
    return spans


def replace_paragraph_span(paragraph: Paragraph, start: int, end: int, new_text: str) -> None:
    """Replace paragraph.text[start:end] with new_text, keeping runs outside the
    span untouched and assigning new_text to the first run overlapping the span
    (so its character formatting is preserved); other overlapped runs are cleared."""
    if start >= end:
        return
    spans = _paragraph_run_spans(paragraph)
    assigned = False
    for run, r_start, r_end in spans:
        if r_end <= start or r_start >= end:
            continue  # no overlap
        overlap_start = max(start, r_start)
        overlap_end = min(end, r_end)
        local_start = overlap_start - r_start
        local_end = overlap_end - r_start
        text = run.text or ""
        before = text[:local_start]
        after = text[local_end:]
        if not assigned:
            run.text = before + new_text + after
            assigned = True
        else:
            run.text = before + after


def _run_has_drawing(run) -> bool:
    return bool(run._r.findall(".//" + qn("w:drawing")))


def replace_cell_text(cell: _Cell, new_text: str) -> None:
    """Replace all text in a table cell with new_text, preserving the formatting
    of the first run of the first non-empty paragraph found (falls back to plain
    text if the cell has no runs at all).

    Any run carrying a picture (a `w:drawing` -- e.g. a stamp/signature that
    happens to be anchored inside this cell) is left completely untouched:
    setting `run.text` on a run clears ALL of that run's XML children, not
    just its visible text, so a naive "clear every run" pass silently deletes
    embedded images too. Confirmed on the CoA DERUN template, where a
    template's own red stamp is a floating image anchored inside the
    'Single impurity' row's specification cell -- refilling that cell's
    specification text was wiping the stamp out along with it."""
    paragraphs = cell.paragraphs
    donor_run = None
    for p in paragraphs:
        for run in p.runs:
            if not _run_has_drawing(run):
                donor_run = run
                break
        if donor_run is not None:
            break
    if donor_run is None:
        # No plain-text run to reuse (cell may be image-only) -- append a new
        # paragraph for the text rather than overwriting paragraphs[0], which
        # could itself be carrying a drawing.
        target = None
        for p in paragraphs:
            if not any(_run_has_drawing(r) for r in p.runs):
                target = p
                break
        if target is not None:
            target.text = new_text
        else:
            cell.add_paragraph(new_text)
        return
    # Clear every paragraph's non-drawing runs
    target_paragraph = None
    for p in paragraphs:
        for run in list(p.runs):
            if _run_has_drawing(run):
                continue
            if target_paragraph is None and run is donor_run:
                target_paragraph = p
            run.text = ""
    if target_paragraph is None:
        target_paragraph = paragraphs[0]
    donor_run.text = new_text
    # Remove any now-empty extra paragraphs after the first (keeps cell tidy),
    # but never one that still carries a drawing.
    for p in paragraphs[1:]:
        if not p.text and not any(_run_has_drawing(r) for r in p.runs) and p._p.getparent() is not None:
            p._p.getparent().remove(p._p)


# --------------------------------------------------------------------------- #
# Header "Label: value" field scanning + replacement
# --------------------------------------------------------------------------- #

_FIELD_RE = re.compile(r"^(.*?)\s*[:：]\s*(.*)$", re.DOTALL)
_COLON_CHARS = ":："


# A field is "label:value", where value runs until either (a) a run of 2+
# whitespace followed by what looks like the START of another field on the
# same line (another label ending in a colon -- the common "Label: value
# Label2: value2" layout), or (b) end of line. Deliberately does NOT split on
# a bare 2+ space run right after the colon itself (e.g. "Product Name:
# Vonoprazan Fumarate", a very common template style padding the value away
# from its label) -- an earlier version split chunks on ANY 2+ whitespace
# run, which tore that kind of "Label:  value" apart into a colon-only
# "Product Name:" chunk (empty value, silently skipped) and a colon-less
# "Vonoprazan Fumarate" chunk (dropped for having no colon), so the value was
# lost entirely and the field could never be filled -- confirmed on a real
# HANSOH COA template where every header field uses this double-space style
# and none of them were being filled at all.
_LINE_FIELD_RE = re.compile(
    r"(?P<label>[^\n:：]+?)[:：]\s*(?P<value>.*?)(?=\s{2,}\S.*?[:：]|\s*$)"
)


def _iter_field_candidates(paragraph: Paragraph):
    """Yield (label_raw, value_start, value_end) for each 'Label: value' looking
    segment inside a paragraph, one line at a time. Accepts both ASCII ':' and
    fullwidth '：' separators (common in Chinese manufacturer templates)."""
    text = paragraph.text
    if not any(c in text for c in _COLON_CHARS):
        return
    for line_match in re.finditer(r"[^\n]+", text):
        line = line_match.group(0)
        line_offset = line_match.start()
        for m in _LINE_FIELD_RE.finditer(line):
            label_raw = m.group("label").strip()
            if not label_raw or len(label_raw) > 60:
                continue
            value_start = line_offset + m.start("value")
            value_end = line_offset + m.end("value")
            yield label_raw, value_start, value_end


def _iter_nocolon_candidates(paragraph: Paragraph, remaining_labels: List[str]):
    """Fallback for templates that render a field as 'Label Value' with no colon
    (e.g. 'Quantity 185.00 KG'). Only tried for labels not already matched via the
    colon-based pass. Matches a candidate label as a leading, whole-word prefix of
    a line and treats the rest of that line as the value span."""
    text = paragraph.text
    for line_match in re.finditer(r"[^\n]+", text):
        line = line_match.group(0)
        line_offset = line_match.start()
        norm_line = normalize(line)
        for label in remaining_labels:
            norm_label = normalize(label)
            if not norm_label:
                continue
            if norm_line == norm_label:
                continue  # nothing left to treat as a value
            if norm_line.startswith(norm_label + " "):
                # map back from normalized offset to raw offset is unreliable char-for-char
                # (normalization strips punctuation), so instead re-locate using a regex
                # built from the raw label, allowing flexible whitespace/punctuation.
                pattern = re.escape(label)
                pattern = re.sub(r"\\?\s+", r"[\\s.]*", pattern)
                m = re.match(r"^\s*" + pattern + r"\s*[:\-]?\s*", line, re.IGNORECASE)
                if m:
                    value_start = line_offset + m.end()
                    value_end = line_offset + len(line)
                    remainder = text[value_start:value_end]
                    if any(c in remainder for c in _COLON_CHARS):
                        # The "value" we'd be grabbing still contains a colon --
                        # that means `label` was just a prefix of a *longer*,
                        # different label in this line (e.g. candidate "Product"
                        # matching the start of template text "Product Name:  "),
                        # not a real no-colon value. Let the adjacent-cell pass
                        # handle that case instead of mangling it here.
                        continue
                    if value_end > value_start:
                        yield label, value_start, value_end
                        break


def _cell_is_bare_label(cell_text: str) -> Optional[str]:
    """Returns the label text if this cell looks like a *label with no inline
    value* -- e.g. 'Product Name:  ' or 'Batch No.：' with nothing meaningful
    after the colon. Returns None if the cell has real content after the colon
    (handled by the inline colon-based pass already) or no colon at all."""
    text = (cell_text or "").strip()
    if not text or not any(c in text for c in _COLON_CHARS):
        return None
    m = _FIELD_RE.match(text)
    if not m:
        return None
    label = m.group(1).strip()
    value = m.group(2).strip()
    if value:
        return None  # already has an inline value; not "bare"
    if not label or len(label) > 60:
        return None
    return label


def apply_header_fields_adjacent_cells(doc: DocumentObject, header_fields: List[dict], used: set, skip_table: Optional[Table] = None) -> List[str]:
    """Handles the very common table layout where a 'Label:' cell has its value
    in the *next* cell of the same row (e.g. ['Product Name:', 'Empagliflozin',
    'Test Date:', '25.06.2026']) rather than 'Label: value' combined in one
    cell/paragraph. Mutates `used` in place and returns the list of labels
    matched here."""
    candidates = {normalize(f["label"]): f["value"] for f in header_fields if f.get("label")}
    matched_in_doc: List[str] = []

    def process_table(table: Table):
        if table is skip_table:
            return
        for row in table.rows:
            cells = _unique_row_cells(row)
            i = 0
            while i < len(cells) - 1:
                label = _cell_is_bare_label(cells[i].text)
                if label:
                    next_is_label = _cell_is_bare_label(cells[i + 1].text) is not None
                    if not next_is_label:
                        key = best_match(normalize(label), candidates)
                        if key is not None and key not in used:
                            replace_cell_text(cells[i + 1], candidates[key])
                            used.add(key)
                            matched_in_doc.append(label)
                        i += 2
                        continue
                i += 1
            for cell in cells:
                for nested in cell.tables:
                    process_table(nested)

    for table in doc.tables:
        process_table(table)

    return matched_in_doc


def apply_header_fields(doc: DocumentObject, header_fields: List[dict]) -> dict:
    """Scans every paragraph in the body and every table cell for 'Label: value'
    patterns and overwrites the value when the label fuzzy-matches an extracted
    header field. Returns a report of matched/unmatched labels."""
    candidates = {normalize(f["label"]): f["value"] for f in header_fields if f.get("label")}
    used = set()
    matched_in_doc = []
    all_paragraphs: List[Paragraph] = []

    def process_paragraph(paragraph: Paragraph):
        all_paragraphs.append(paragraph)
        # Collect all field spans first (so replacing one doesn't shift offsets
        # for ones we haven't processed yet) -- process back-to-front.
        fields = list(_iter_field_candidates(paragraph))
        for label_raw, value_start, value_end in reversed(fields):
            if value_end <= value_start:
                # No inline value to overwrite (e.g. a bare "Product Name:" cell
                # whose value lives in a *different* cell) -- leave it for the
                # adjacent-cell pass below rather than falsely marking it "used".
                continue
            key = best_match(normalize(label_raw), candidates)
            if key is None:
                continue
            new_value = candidates[key]
            replace_paragraph_span(paragraph, value_start, value_end, new_value)
            used.add(key)
            matched_in_doc.append(label_raw)

    for paragraph in doc.paragraphs:
        process_paragraph(paragraph)

    def process_table(table: Table):
        for row in table.rows:
            seen_tcs = set()
            for cell in row.cells:
                if id(cell._tc) in seen_tcs:
                    continue
                seen_tcs.add(id(cell._tc))
                for paragraph in cell.paragraphs:
                    process_paragraph(paragraph)
                for nested in cell.tables:
                    process_table(nested)

    for table in doc.tables:
        process_table(table)

    # Second pass: no-colon "Label Value" fallback for header fields the first pass missed.
    remaining = [f["label"] for f in header_fields if normalize(f["label"]) not in used]
    if remaining:
        for paragraph in all_paragraphs:
            if not remaining:
                break
            hits = list(_iter_nocolon_candidates(paragraph, remaining))
            for label_raw, value_start, value_end in reversed(hits):
                key = normalize(label_raw)
                if key not in candidates or key in used:
                    continue
                replace_paragraph_span(paragraph, value_start, value_end, candidates[key])
                used.add(key)
                matched_in_doc.append(label_raw)
                remaining = [l for l in remaining if normalize(l) != key]

    # Third pass: 'Label:' cell + value in the *next* cell of the same table row --
    # very common layout (e.g. CoA DERUN's header table) that the two passes above
    # can't see because the label and value never share one paragraph/cell.
    results = find_results_table(doc)
    skip_table = results[0] if results else None
    matched_in_doc.extend(apply_header_fields_adjacent_cells(doc, header_fields, used, skip_table))

    unmatched = [f["label"] for f in header_fields if normalize(f["label"]) not in used]
    return {"matched": matched_in_doc, "unmatched_extracted_fields": unmatched}


# --------------------------------------------------------------------------- #
# Results table detection + fill
# --------------------------------------------------------------------------- #

# Order matters: more specific/discriminating keywords are checked first so a
# header cell like "Test Result" (which contains the generic word "test") is
# classified as "result" rather than the weaker, catch-all "parameter" role.
HEADER_KEYWORDS = {
    "no": ["s.no", "sr.no", "sl.no", "no.", "sr no", "sl no"],
    "result": ["result", "observation", "finding"],
    "specification": ["specification", "limit", "standard", "acceptance"],
    "parameter": ["test", "parameter", "particular", "characteristic", "item"],
}


def _unique_row_cells(row: _Row) -> List[_Cell]:
    seen = set()
    cells = []
    for cell in row.cells:
        if id(cell._tc) in seen:
            continue
        seen.add(id(cell._tc))
        cells.append(cell)
    return cells


def _classify_header_cell(text: str) -> Optional[str]:
    norm = normalize(text)
    for role, keywords in HEADER_KEYWORDS.items():
        for kw in keywords:
            if normalize(kw) in norm:
                return role
    return None


def find_results_table(doc: DocumentObject) -> Optional[Tuple[Table, int, Dict[int, str]]]:
    """Returns (table, header_row_index, {unique_cell_position: role}) for the
    table + row that looks like the test-results header, or None."""
    best = None
    best_score = 0
    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            cells = _unique_row_cells(row)
            roles = {}
            for pos, cell in enumerate(cells):
                role = _classify_header_cell(cell.text)
                if role:
                    roles[pos] = role
            score = len(set(roles.values()))
            has_result = "result" in roles.values()
            has_param = "parameter" in roles.values()
            if has_result and has_param and score > best_score:
                best_score = score
                best = (table, row_idx, roles)
    return best


def apply_test_results(doc: DocumentObject, test_results: List[dict]) -> dict:
    found = find_results_table(doc)
    if not found:
        return {"status": "no_results_table_found", "matched": 0, "appended": 0, "deleted": 0, "unmatched_template_rows": []}

    table, header_row_idx, roles = found
    role_to_pos = {}
    for pos, role in roles.items():
        role_to_pos.setdefault(role, pos)
    param_pos = role_to_pos.get("parameter")
    result_pos = role_to_pos.get("result")
    spec_pos = role_to_pos.get("specification")
    no_pos = role_to_pos.get("no")

    if param_pos is None or result_pos is None:
        return {"status": "incomplete_header", "matched": 0, "appended": 0, "deleted": 0, "unmatched_template_rows": []}

    # Build match keys for extracted results: parameter + first part of spec (disambiguates
    # generically-named rows like "By IR" / "By HPLC" that repeat under a shared heading).
    def extracted_key(item: dict) -> str:
        return normalize(item.get("parameter", "") + " " + (item.get("specification", "") or "")[:80])

    extracted = [
        {**item, "_key": extracted_key(item), "_param_key": normalize(item.get("parameter", ""))}
        for item in test_results
    ]
    used_extracted_idx = set()

    matched_count = 0
    unmatched_template_rows = []
    rows_to_delete: List[_Row] = []

    data_rows = table.rows[header_row_idx + 1:]
    expected_ncols = len(_unique_row_cells(table.rows[header_row_idx]))
    last_row_for_clone = None  # only ever set to a row that looks like a *real* data row

    # Collect every real (non-blank) data row's info first, WITHOUT matching yet.
    # Matching used to happen row-by-row, top to bottom, greedily grabbing the
    # first extracted item that cleared the 0.55 threshold -- but when a table
    # has several similarly-worded rows (very common: a "Related substances"
    # group with one row per named impurity, all sharing near-identical
    # spec/result formatting like "<0.10%"), an EARLIER row can score just
    # high enough to steal an item that a LATER row actually matches far
    # better. Confirmed on a real CoA DERUN template: rows for "Impurity B/C/
    # D/E/G" all scored ~0.56 against the extracted "Single impurity" result
    # (same generic wording + same "<0.10%" spec), so the first one of them
    # claimed it, leaving the genuine "Single impurity" row unmatched -- which
    # then got deleted as a "template row not in the source PDF", silently
    # destroying it (and, in that specific template, a stamp image anchored
    # inside that row) even though the data WAS present in the source PDF, in
    # the right place, just assigned to the wrong row.
    #
    # Fix: score every (row, item) pair up front, then assign matches
    # greedily in order of DESCENDING score across the whole table, so the
    # best-fitting pair always wins a contested item, regardless of row order.
    candidate_rows = []  # [{"row":, "cells":, "param_text":, "key":, "param_key":}]
    for row in data_rows:
        cells = _unique_row_cells(row)
        if param_pos >= len(cells) or result_pos >= len(cells):
            continue
        param_text = cells[param_pos].text.strip()
        if not param_text:
            # Blank-parameter row -- usually the second half of a vertically
            # merged specification cell (Word shows the merged content only
            # once, spanning both rows) left over as an empty shell. It carries
            # no independent data once its "real" row above has been updated,
            # so remove it rather than leave a stray half-empty row behind.
            rows_to_delete.append(row)
            continue
        # A trailing "CONCLUSION: ..." / "REMARK: ..." style row is often a single
        # fully-merged cell still inside this same table -- don't treat it as a
        # template for new rows, or appended rows will lose their column layout.
        if len(cells) == expected_ncols:
            last_row_for_clone = row
        spec_text = cells[spec_pos].text.strip() if spec_pos is not None and spec_pos < len(cells) else ""
        candidate_rows.append({
            "row": row,
            "cells": cells,
            "param_text": param_text,
            "key": normalize(param_text + " " + spec_text[:80]),
            "param_key": normalize(param_text),
        })

    # Score every (row, item) pair -- best of the full key (param+spec) and the
    # bare-parameter-name comparison, mirroring the two-tier fallback the old
    # per-row logic used, but computed for all pairs before any assignment.
    scored_pairs = []  # (score, row_idx_in_candidate_rows, item_idx)
    for ri, cand in enumerate(candidate_rows):
        for ii, item in enumerate(extracted):
            full_score = difflib.SequenceMatcher(None, cand["key"], item["_key"]).ratio()
            param_score = difflib.SequenceMatcher(None, cand["param_key"], item["_param_key"]).ratio()
            score = max(full_score, param_score)
            if score >= 0.55:
                scored_pairs.append((score, ri, ii))
    scored_pairs.sort(key=lambda t: t[0], reverse=True)

    row_match = {}  # row_idx_in_candidate_rows -> item_idx
    for score, ri, ii in scored_pairs:
        if ri in row_match or ii in used_extracted_idx:
            continue
        row_match[ri] = ii
        used_extracted_idx.add(ii)

    for ri, cand in enumerate(candidate_rows):
        row, cells, param_text = cand["row"], cand["cells"], cand["param_text"]
        if ri in row_match:
            matched_item = extracted[row_match[ri]]
            new_result = matched_item.get("result", "")
            new_spec = matched_item.get("specification", "")
            if new_result:
                replace_cell_text(cells[result_pos], new_result)
                matched_count += 1
            # Overwrite the specification too, not just the result: the template
            # is only reused for its *layout*, so its original specification text
            # belongs to whatever product it was last filled out for -- keeping it
            # would silently pair the source PDF's result with a wrong, unrelated
            # spec whenever the template and the uploaded PDF are for different
            # products (the common case for a generic multi-template tool).
            if new_spec and spec_pos is not None and spec_pos < len(cells):
                replace_cell_text(cells[spec_pos], new_spec)
        else:
            unmatched_template_rows.append(param_text)
            rows_to_delete.append(row)

    # Append rows for extracted tests that had no matching template row
    appended = 0
    next_no = None
    if no_pos is not None:
        seen_numbers = []
        for row in data_rows:
            cells = _unique_row_cells(row)
            if no_pos < len(cells):
                m = re.search(r"\d+", cells[no_pos].text)
                if m:
                    seen_numbers.append(int(m.group(0)))
        next_no = (max(seen_numbers) + 1) if seen_numbers else None

    if last_row_for_clone is not None:
        insert_after_tr = last_row_for_clone._tr
        for i, item in enumerate(extracted):
            if i in used_extracted_idx:
                continue
            new_tr = copy.deepcopy(insert_after_tr)
            # Insert right after the last real data row (not at the very end of
            # the table, which could be past a trailing merged "CONCLUSION" /
            # "REMARK" row) so appended tests land in a sensible place and stay
            # grouped together in source order.
            insert_after_tr.addnext(new_tr)
            insert_after_tr = new_tr
            new_row = next(r for r in table.rows if r._tr is new_tr)
            new_cells = _unique_row_cells(new_row)
            for pos, cell in enumerate(new_cells):
                if pos == param_pos:
                    replace_cell_text(cell, item.get("parameter", ""))
                elif pos == spec_pos:
                    replace_cell_text(cell, item.get("specification", ""))
                elif pos == result_pos:
                    replace_cell_text(cell, item.get("result", ""))
                elif pos == no_pos and next_no is not None:
                    replace_cell_text(cell, str(next_no))
                    next_no += 1
                else:
                    replace_cell_text(cell, "")
            appended += 1

    # Remove template rows for tests that don't exist in the source PDF at all --
    # done last, after cloning/appending is finished (appended rows are already
    # independent copies by this point, so deleting their donor row is safe).
    deleted = 0
    for row in rows_to_delete:
        tr = row._tr
        parent = tr.getparent()
        if parent is not None:
            parent.remove(tr)
            deleted += 1

    return {
        "status": "ok",
        "matched": matched_count,
        "appended": appended,
        "deleted": deleted,
        "unmatched_template_rows": unmatched_template_rows,
    }


# --------------------------------------------------------------------------- #
# Signature / stamp insertion
# --------------------------------------------------------------------------- #

SIGNATURE_KEYWORDS = [
    "signature", "sign here", "authorized", "authorised", "checked by",
    "approved by", "prepared by", "verified by", "qa manager", "qc manager",
    "analyst", "chemist", "authorised signatory", "authorized signatory",
]


def _paragraph_has_signature_keyword(text: str) -> bool:
    norm = normalize(text)
    return any(normalize(kw) in norm for kw in SIGNATURE_KEYWORDS)


def _iter_all_paragraphs_with_container(doc: DocumentObject):
    """Yield (paragraph, container_cell_or_None) for every paragraph in the body
    and in every (nested) table cell, in document order-ish (body then tables)."""
    for p in doc.paragraphs:
        yield p, None

    def walk_table(table: Table):
        for row in table.rows:
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                for p in cell.paragraphs:
                    yield p, cell
                for nested in cell.tables:
                    yield from walk_table(nested)

    for table in doc.tables:
        yield from walk_table(table)


def _find_drawing_run_in_paragraph(paragraph: Paragraph):
    for run in paragraph.runs:
        drawings = run._r.findall(".//" + qn("w:drawing"))
        if drawings:
            return run, drawings[0]
    return None, None


def _replace_image_in_run_drawing(paragraph: Paragraph, run, image_png_bytes: bytes) -> bool:
    """Replace the picture referenced by this run's drawing with new image bytes,
    keeping the existing size/position (swaps the underlying image part blob)."""
    part = paragraph.part
    blips = run._r.findall(".//" + qn("a:blip"))
    if not blips:
        return False
    blip = blips[0]
    rId = blip.get(qn("r:embed"))
    if not rId:
        return False
    try:
        image_part = part.related_parts[rId]
    except KeyError:
        return False
    image_part._blob = image_png_bytes
    return True


def _append_image_to_paragraph(paragraph: Paragraph, image_png_bytes: bytes, width_inches: float = 1.5):
    from docx.shared import Inches
    run = paragraph.add_run()
    run.add_picture(io.BytesIO(image_png_bytes), width=Inches(width_inches))


def _all_drawing_runs(doc: DocumentObject):
    """Document-order list of (paragraph, run) for every run containing a drawing."""
    results = []
    for paragraph, _container in _iter_all_paragraphs_with_container(doc):
        for run in paragraph.runs:
            if run._r.findall(".//" + qn("w:drawing")):
                results.append((paragraph, run))
    return results


def apply_signature(doc: DocumentObject, signature_png_bytes: Optional[bytes]) -> dict:
    if not signature_png_bytes:
        return {"status": "no_signature_detected"}

    # 1) Look for a paragraph containing a signature-related keyword.
    anchor_paragraph = None
    anchor_container = None
    for paragraph, container in _iter_all_paragraphs_with_container(doc):
        if _paragraph_has_signature_keyword(paragraph.text):
            anchor_paragraph = paragraph
            anchor_container = container
            break

    if anchor_paragraph is not None:
        # An existing signature/stamp image is often in a *sibling* paragraph within
        # the same table cell (or a nearby body paragraph) rather than the exact
        # paragraph carrying the "Checked by:" text -- widen the search before
        # falling back to just appending a new image.
        if anchor_container is not None:
            search_paragraphs = list(anchor_container.paragraphs)
        else:
            body_paragraphs = doc.paragraphs
            try:
                idx = body_paragraphs.index(anchor_paragraph)
                lo, hi = max(0, idx - 3), min(len(body_paragraphs), idx + 4)
                search_paragraphs = body_paragraphs[lo:hi]
            except ValueError:
                search_paragraphs = [anchor_paragraph]

        for p in search_paragraphs:
            run, _drawing = _find_drawing_run_in_paragraph(p)
            if run is not None and _replace_image_in_run_drawing(p, run, signature_png_bytes):
                return {"status": "replaced_existing_image_at_anchor"}

        _append_image_to_paragraph(anchor_paragraph, signature_png_bytes)
        return {"status": "appended_at_keyword_anchor"}

    # 2) No keyword anchor found -- fall back to replacing the LAST image in the
    #    document, but only if there's more than one image total (so we don't
    #    clobber a document's sole logo).
    drawing_runs = _all_drawing_runs(doc)
    if len(drawing_runs) >= 2:
        paragraph, run = drawing_runs[-1]
        if _replace_image_in_run_drawing(paragraph, run, signature_png_bytes):
            return {"status": "replaced_last_image_fallback"}

    return {"status": "no_anchor_found_signature_not_inserted"}


# --------------------------------------------------------------------------- #
# Top-level entry point
# --------------------------------------------------------------------------- #

@dataclasses.dataclass
class FillReport:
    header_fields: dict
    test_results: dict
    signature: dict


def fill_template(
    template_path: str,
    output_path: str,
    header_fields: List[dict],
    test_results: List[dict],
    signature_png_bytes: Optional[bytes] = None,
) -> FillReport:
    doc = Document(template_path)

    header_report = apply_header_fields(doc, header_fields)
    results_report = apply_test_results(doc, test_results)
    signature_report = apply_signature(doc, signature_png_bytes)

    doc.save(output_path)

    return FillReport(
        header_fields=header_report,
        test_results=results_report,
        signature=signature_report,
    )
