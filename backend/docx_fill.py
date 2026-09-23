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
    "quantity": "qty", "weight": "qty", "wt": "qty", "netweight": "qty",
    "kg": "qty", "kgs": "qty", "kilogram": "qty", "kilograms": "qty",
    "number": "no", "numbers": "no",
    "lot": "batch",
    "date": "dt",
    "analytical": "ar", "report": "ar",
    "certificate": "coa", "analysis": "coa",
    # A reference-standard CoA's "Qualification Date" (when the standard was
    # tested/certified) is the same underlying concept as a bulk-API
    # template's "Test Date" -- per explicit product rule, map it onto that
    # field rather than leaving it unmatched. Note this does NOT catch
    # "Re-qualification Date" (a future re-check date, handled separately by
    # the Exp/Retest relabeling logic) since that canonicalizes with an
    # extra leftover "re" token that keeps it from equaling "Test Date".
    "qualification": "test",
}
_STOPWORDS = {"of", "the", "a", "an"}

# Tokens (after canonicalization) that flag a "qty family" candidate as a
# per-molecule/formula constant rather than a genuine batch/shipment
# quantity -- used to prefer the latter when both are present on one CoA
# (see best_match's quantity-family fallback).
_MOLECULAR_QTY_TOKENS = {"formula", "molecular", "mol", "mw"}


def _canonicalize(norm_text: str) -> str:
    tokens = [t for t in norm_text.split() if t not in _STOPWORDS]
    tokens = [_TOKEN_SYNONYMS.get(t, t) for t in tokens]
    return " ".join(sorted(tokens))


def _blocks_fuzzy_match(a: str, b: str) -> bool:
    """True if a and b must NEVER be accepted as a character-fuzzy match,
    despite scoring high, because they differ only by a leading "re" on one
    otherwise-identical token -- confirmed a real false positive: "Retest
    date" (Apr.21.2028, a future compliance re-check date) scored 0.9
    against a template's "Test Date" (meant to be when THIS analysis was
    performed) purely on string shape, and would have overwritten it with a
    completely different, 2-years-out date. A "re-" prefixed QC/regulatory
    term (retest, recheck, reprocess, resample, ...) is essentially always a
    materially different concept from its bare counterpart, not a spelling
    variant of it, so this rules the pairing out regardless of how similar
    the strings look character-by-character."""
    a_tokens, b_tokens = a.split(), b.split()
    if len(a_tokens) != len(b_tokens):
        return False
    saw_re_diff = False
    for ta, tb in zip(a_tokens, b_tokens):
        if ta == tb:
            continue
        shorter, longer = (ta, tb) if len(ta) < len(tb) else (tb, ta)
        if longer == "re" + shorter:
            saw_re_diff = True
            continue
        return False
    return saw_re_diff


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

    # 2) canonical token-set CONTAINMENT (handles a missing/extra trailing
    # word that exact canonical equality -- method 1 -- can't bridge, e.g. a
    # source document's bare "Lot#" canonicalizes to just {"batch"} while a
    # template's "Batch No" canonicalizes to {"batch", "no"}; neither is a
    # character-fuzzy match to the other (different lengths/shapes) and
    # they're not equal as canonical sets, but one is a clean subset of the
    # other. Only fires when exactly one candidate qualifies, so it never
    # guesses between several genuinely different fields that happen to
    # share a token (e.g. both "Batch No" and "Batch Size" containing
    # "batch").
    if q_canon:
        q_set = set(q_canon.split())
        subset_keys = [
            k for k in keys
            if (k_set := set(_canonicalize(k).split())) and (q_set <= k_set or k_set <= q_set)
        ]
        if len(subset_keys) == 1:
            return subset_keys[0]

    # 3) character-level fuzzy match
    matches = difflib.get_close_matches(query, keys, n=3, cutoff=cutoff)
    for m in matches:
        if not _blocks_fuzzy_match(query, m):
            return m

    # 4) token containment either direction (handles "batch no" vs "batch number")
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

    # 5) "quantity family" fallback: a template's "Net weight" field and a
    # source document's "Batch Quantity" (or similar) are, in practice, very
    # often the same number under a different name for a single-batch CoA --
    # per an explicit product rule, not a general synonym. Only kick in when
    # the query itself is fundamentally a quantity/weight field (canonicalizes
    # to contain the "qty" token) AND there's exactly one unmatched candidate
    # that's also in that family, so this never guesses between several
    # different quantity-like fields that genuinely differ.
    #
    # Includes "Mol. wt." / "Formula Weight" (molecular weight) in that
    # family on explicit instruction -- it's a different physical quantity
    # from a batch's net weight (a per-molecule mass constant vs. a shipment
    # weight), but when nothing else on the CoA is labeled as a weight/
    # quantity at all, use it anyway rather than leave the field blank.
    # Genuine shipment-quantity candidates are preferred over molecular ones
    # whenever both are present, though (confirmed real case: a reference-
    # standard CoA listing BOTH "Quantity: 1g" and "Formula Weight: 475.6" --
    # the two used to tie and cancel each other out via the plain
    # len(qty_candidates) == 1 check below, leaving Net weight blank even
    # though "Quantity" was right there and unambiguous).
    if "qty" in q_canon.split():
        qty_candidates = [k for k in keys if "qty" in _canonicalize(k).split()]
        if qty_candidates:
            non_molecular = [
                k for k in qty_candidates
                if not (set(_canonicalize(k).split()) & _MOLECULAR_QTY_TOKENS)
            ]
            pool = non_molecular or qty_candidates
            if len(pool) == 1:
                return pool[0]
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
        if r_end == r_start and _run_has_drawing(run):
            continue  # zero-width picture run inside the span: setting .text would delete the image
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
    """True if the run carries a picture: a modern DrawingML `w:drawing` or a
    legacy VML `w:pict` (common in templates converted from old .doc files)."""
    r = run._r
    return bool(r.findall(".//" + qn("w:drawing")) or r.findall(".//" + qn("w:pict")))


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


def apply_header_fields_adjacent_cells(
    doc: DocumentObject,
    header_fields: List[dict],
    used: set,
    skip_table: Optional[Table] = None,
    handled_cells: Optional[set] = None,
    touched_cells: Optional[set] = None,
) -> List[str]:
    """Handles the very common table layout where a 'Label:' cell has its value
    in the *next* cell of the same row (e.g. ['Product Name:', 'Empagliflozin',
    'Test Date:', '25.06.2026']) rather than 'Label: value' combined in one
    cell/paragraph. Mutates `used` in place and returns the list of labels
    matched here.

    `handled_cells` (from apply_header_fields's inline pass) marks cells that
    already had their OWN genuine inline "Label: value" -- skip treating those
    as the "value" slot for a *different*, unrelated bare-label cell next to
    them. Without this, an inline field that legitimately had no match and
    got cleared to "Label: " would look exactly like a real bare label, and
    this pass would then blank whatever independent (possibly
    already-correctly-filled) cell happens to sit next to it."""
    candidates = {normalize(f["label"]): f.get("value", "") for f in header_fields if f.get("label")}
    matched_in_doc: List[str] = []
    handled_cells = handled_cells or set()

    def process_table(table: Table):
        if table is skip_table:
            return
        for row in table.rows:
            cells = _unique_row_cells(row)
            i = 0
            while i < len(cells) - 1:
                label = _cell_is_bare_label(cells[i].text)
                if (
                    label
                    and id(cells[i]._tc) not in handled_cells
                    and id(cells[i + 1]._tc) not in handled_cells
                ):
                    next_is_label = _cell_is_bare_label(cells[i + 1].text) is not None
                    if not next_is_label:
                        key = best_match(normalize(label), candidates)
                        if touched_cells is not None:
                            touched_cells.add(id(cells[i + 1]._tc))
                            touched_cells.add(id(cells[i]._tc))
                        vm = _VALUE_COLON_PREFIX_RE.match(cells[i + 1].text)
                        vprefix = vm.group(0).strip() + " " if vm else ""
                        if key is not None and key not in used:
                            replace_cell_text(cells[i + 1], vprefix + candidates[key])
                            used.add(key)
                            matched_in_doc.append(label)
                        elif key is None:
                            # This label genuinely has no corresponding value
                            # anywhere in the source document -- clear the
                            # value cell rather than leave whatever this
                            # exemplar template's own original product had
                            # there (see the matching comment in
                            # apply_header_fields's process_paragraph for the
                            # confirmed real-world case this fixes: a
                            # template's leftover "Net weight: 50 kg" showing
                            # up in output for a source CoA that never states
                            # a net weight/quantity at all).
                            replace_cell_text(cells[i + 1], vprefix)
                        i += 2
                        continue
                i += 1
            for cell in cells:
                for nested in cell.tables:
                    process_table(nested)

    for table in doc.tables:
        process_table(table)

    return matched_in_doc


# --------------------------------------------------------------------------- #
# Conditional Exp Date <-> Retest Date relabeling
# --------------------------------------------------------------------------- #
#
# Some CoAs state a genuine expiry date; others (common for APIs/intermediates
# with a shelf-life-extension program) state only a *retest* date -- a future
# date by which the material must be re-tested to confirm it's still usable,
# which is a materially different thing from an expiry date. Per explicit
# product rule: if the source document has a real Exp Date, fill the
# template's "Exp Date" field exactly as always. If it has NO Exp Date but
# DOES have a Retest Date, the template's own "Exp Date:" label is rewritten
# to "Retest date:" (in place, keeping the colon/formatting) so the field's
# name accurately reflects what's actually being reported, and the retest
# value is what ends up filled there.

_EXP_LABEL_ALT = r"Exp\.?\s*Date|Expiry\s*Date|Expiration\s*Date|Date\s+of\s+Exp(?:iry)?\.?"
_RETEST_LABEL_ALT = r"Re-?\s*test\s*Date|Date\s+of\s+Re-?\s*test"
_EXP_DATE_LABEL_RE = re.compile(r"(" + _EXP_LABEL_ALT + r")(\s*[:：])", re.IGNORECASE)
# Same labels written with no colon at all, at the start of a line and
# followed by whitespace (the value) or nothing (a label-only cell) -- e.g.
# Elder's "Exp date Apr 2031" or a bare "Exp. Date" table cell.
_EXP_DATE_LABEL_NOCOLON_RE = re.compile(
    r"^(?:\s*)(" + _EXP_LABEL_ALT + r")(?=\s|$)", re.IGNORECASE | re.MULTILINE
)
_RETEST_DATE_LABEL_RE = re.compile(r"(" + _RETEST_LABEL_ALT + r")(\s*[:：])", re.IGNORECASE)
_RETEST_DATE_LABEL_NOCOLON_RE = re.compile(
    r"^(?:\s*)(" + _RETEST_LABEL_ALT + r")(?=\s|$)", re.IGNORECASE | re.MULTILINE
)


def _is_exp_date_label(norm_label: str) -> bool:
    tokens = norm_label.split()
    return "date" in tokens and any(t in ("exp", "expiry", "expiration") for t in tokens)


_RETEST_LIKE_KEYWORDS = (
    "retest",
    "requalification",
    "recheck",
    "reevaluation",
    "reassessment",
    "reverification",
)


def _is_retest_date_label(norm_label: str) -> bool:
    """True for 'Retest Date' and the same underlying concept under other
    QC/regulatory wording -- e.g. a reference-standard CoA's 'Re-qualification
    Date' (confirmed real example: a Junye reference-standard CoA states
    '再确认期 / Re-qualification Date' instead of an expiry date, meaning
    exactly the same thing a retest date does -- a future date to re-confirm
    the material is still usable, not a hard expiry)."""
    tokens = norm_label.split()
    if "date" not in tokens:
        return False
    joined = norm_label.replace(" ", "")
    return any(kw in joined for kw in _RETEST_LIKE_KEYWORDS)


def _find_header_field(header_fields: List[dict], matcher) -> Optional[dict]:
    """First header field (with a non-blank value) whose normalized label
    satisfies `matcher`, or None."""
    for f in header_fields:
        value = (f.get("value") or "").strip()
        if not value:
            continue
        label = f.get("label") or ""
        if matcher(normalize(label)):
            return f
    return None


def _relabel_date_labels(doc: DocumentObject, colon_re, nocolon_re, new_label: str, dry_run: bool = False) -> int:
    """Rewrites every 'Exp Date:' / 'Expiry Date:' / 'Expiration Date:' label
    found anywhere in the document (body paragraphs or table cells, including
    nested tables) to 'Retest date:', preserving the colon and the label
    run's own formatting. Only the label text is touched -- any inline value
    on the same line is left alone (and gets filled normally afterwards by
    the regular header-field matching, which will now line up with the
    source's own 'Retest date' field instead). Returns how many labels were
    relabeled (for logging/debugging)."""
    count = 0

    def label_for(old: str) -> str:
        # Keep the template's casing style (e.g. "DATE OF EXP:" -> "RETEST DATE:").
        return new_label.upper() if old.isupper() else new_label

    def process_paragraph(paragraph: Paragraph):
        nonlocal count
        text = paragraph.text
        matches = list(colon_re.finditer(text))
        if not matches and not any(c in text for c in _COLON_CHARS):
            matches = list(nocolon_re.finditer(text))
        for m in reversed(matches):
            count += 1
            if not dry_run:
                replace_paragraph_span(paragraph, m.start(1), m.end(1), label_for(m.group(1)))

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

    return count


def _relabel_exp_date_to_retest(doc: DocumentObject, dry_run: bool = False) -> int:
    return _relabel_date_labels(doc, _EXP_DATE_LABEL_RE, _EXP_DATE_LABEL_NOCOLON_RE, "Retest date", dry_run)


def _relabel_retest_date_to_exp(doc: DocumentObject, dry_run: bool = False) -> int:
    """Mirror image of _relabel_exp_date_to_retest: a template whose only
    date slot is labeled 'Retest date' (e.g. HENAN LIHUA) filled from a
    source that states a genuine expiry date and no retest date."""
    return _relabel_date_labels(doc, _RETEST_DATE_LABEL_RE, _RETEST_DATE_LABEL_NOCOLON_RE, "Exp date", dry_run)


# Header labels real CoA templates use WITHOUT a colon. The three passes in
# apply_header_fields all key off a colon (or off the *source's* label being
# a literal prefix of the template text), so templates laid out like these
# were never filled -- and worse, kept the previous product's name/batch/
# dates. Confirmed on real templates:
#   Elder     : one cell "Name of the Product Fluticasone ..." / "Exp date Apr 2031"
#   Anuh      : cells ["Product Name", ":", "Betamethasone ..."]
#   HENAN     : cells ["PRODUCT", "PREDNISOLONE", "DATE OF SAMPLING", "MAR.20.2026"]
#   Vaikunth  : paragraph "BATCH/LOT NO. CTZ00323"
# Longest phrases are tried first. Single-word labels are only accepted when
# they make up a whole cell (never as a paragraph prefix).
_KNOWN_HEADER_LABELS = [
    "quantity dispatched/sample quantity",
    "name of the product", "name of product", "product name", "material name",
    "batch/lot no", "batch/lot number", "batch no", "batch number", "lot no", "lot number",
    "batch size", "batch qty", "batch quantity",
    "mfg date", "date of mfg", "manufacturing date", "manufacture date",
    "date of manufacture", "date of manufacturing", "production date",
    "exp date", "date of exp", "expiry date", "date of expiry", "expiration date",
    "retest date", "re-test date", "date of retest",
    "date of sampling", "sampling date", "date of report", "report date",
    "date of analysis", "analysis date", "date of release", "release date",
    "analytical report no", "analysis report no", "a.r. no", "report no",
    "reporting date", "article no", "order no", "executive standard",
    "product", "quantity", "package", "packing",
]
_SEP = r"[\s./\-]*"
_VALUE_COLON_PREFIX_RE = re.compile(r"^\s*[:：]\s*")


def _label_regex(phrase: str) -> "re.Pattern":
    tokens = re.findall(r"[a-z0-9]+", phrase.lower())
    return re.compile(r"(?i)" + _SEP.join(re.escape(t) for t in tokens) + r"(?![a-z0-9])")


_KNOWN_LABEL_PATTERNS = sorted(
    ((p, _label_regex(p), len(re.findall(r"[a-z0-9]+", p))) for p in _KNOWN_HEADER_LABELS),
    key=lambda x: -len(x[0]),
)


def _cell_is_known_label(text: str) -> Optional[str]:
    t = (text or "").strip().rstrip(".-").strip()
    if not t or any(c in t for c in _COLON_CHARS) or len(t) > 45:
        return None
    for phrase, rx, _n in _KNOWN_LABEL_PATTERNS:
        m = rx.match(t)
        if m and m.end() == len(t):
            return t
    return None


def _apply_known_label_fields(
    doc: DocumentObject,
    candidates: Dict[str, str],
    used: set,
    results,
    handled_cells: set,
    touched_cells: set,
    touched_paragraphs: set,
) -> List[str]:
    matched: List[str] = []
    res_table, res_header_idx = (results[0], results[1]) if results else (None, None)

    def fill(label_text: str):
        key = best_match(normalize(label_text), candidates)
        if key is None:
            return None, ""  # known label, but the source doesn't state it: clear
        used.add(key)
        matched.append(label_text)
        return key, candidates[key]

    def do_paragraph(paragraph: Paragraph) -> bool:
        """Returns True if a known label was found and its value replaced."""
        if id(paragraph._p) in touched_paragraphs:
            return False
        text = paragraph.text
        if any(c in text for c in _COLON_CHARS):
            return False
        for line_match in re.finditer(r"[^\n]+", text):
            line = line_match.group(0)
            if len(line) > 150:
                continue
            lead = len(line) - len(line.lstrip())
            for phrase, rx, ntok in _KNOWN_LABEL_PATTERNS:
                if ntok < 2:
                    continue
                m = rx.match(line, lead)
                if not m:
                    continue
                sep = re.compile(r"[\s.\-]*").match(line, m.end())
                value_start = sep.end()
                value = line[value_start:]
                # Needs real whitespace between label and value, and a value.
                if not value.strip() or not re.search(r"\s", line[m.end():value_start] or ""):
                    break
                label_text = line[lead:m.end()]
                _key, new_value = fill(label_text)
                off = line_match.start()
                end = off + len(line)
                # A long value often wraps onto following line(s) of the same
                # paragraph (Elder: "... Fluticasone Propionate BP" +
                # "(Micronized)"); those belong to the old value too, up to
                # the next line that starts another known label.
                for nxt in re.finditer(r"[^\n]+", text[end:]):
                    nl = nxt.group(0)
                    if any(rx.match(nl, len(nl) - len(nl.lstrip())) for _p, rx, _n in _KNOWN_LABEL_PATTERNS):
                        break
                    end = end + nxt.end()
                replace_paragraph_span(paragraph, off + value_start, end, new_value)
                touched_paragraphs.add(id(paragraph._p))
                return True  # one field per paragraph; offsets changed
        return False

    def starts_with_known_label(text: str) -> bool:
        t = text.lstrip()
        return any(rx.match(t) for _p, rx, _n in _KNOWN_LABEL_PATTERNS)

    def do_table(table: Table):
        for row_idx, row in enumerate(table.rows):
            in_results = table is res_table and row_idx >= res_header_idx
            cells = _unique_row_cells(row)
            if not in_results:
                i = 0
                while i < len(cells):
                    c = cells[i]
                    label = None
                    if id(c._tc) not in handled_cells and id(c._tc) not in touched_cells:
                        label = _cell_is_known_label(c.text)
                    if label is None:
                        i += 1
                        continue
                    j = i + 1
                    if j < len(cells) and cells[j].text.strip() in (":", "：", "-"):
                        j += 1
                    # CTX style: the colon starts the VALUE cell
                    # (["Batch No.", ": 24HT0001"]) -- keep that prefix.
                    vtext = cells[j].text if j < len(cells) else ""
                    prefix_m = _VALUE_COLON_PREFIX_RE.match(vtext)
                    prefix = prefix_m.group(0).strip() + " " if prefix_m else ""
                    rest = vtext[prefix_m.end():] if prefix_m else vtext
                    if (
                        j < len(cells)
                        and _cell_is_known_label(cells[j].text) is None
                        and id(cells[j]._tc) not in handled_cells
                        and id(cells[j]._tc) not in touched_cells
                        and not any(ch in rest for ch in _COLON_CHARS)
                    ):
                        _key, new_value = fill(label)
                        replace_cell_text(cells[j], prefix + new_value)
                        touched_cells.update({id(c._tc), id(cells[j]._tc)})
                        i = j + 1
                        continue
                    i += 1
                for c in cells:
                    if id(c._tc) in handled_cells or id(c._tc) in touched_cells:
                        continue
                    continuing = False
                    for paragraph in c.paragraphs:
                        ptext = paragraph.text
                        if (
                            continuing
                            and ptext.strip()
                            and not starts_with_known_label(ptext)
                            and not any(ch in ptext for ch in _COLON_CHARS)
                            and not any(_run_has_drawing(r) for r in paragraph.runs)
                        ):
                            # Old value wrapped into its own paragraph inside
                            # the same cell (Elder: "(Micronized)") -- clear it.
                            replace_paragraph_span(paragraph, 0, len(ptext), "")
                            continue
                        continuing = do_paragraph(paragraph)
            for c in cells:
                for nested in c.tables:
                    do_table(nested)

    for paragraph in doc.paragraphs:
        do_paragraph(paragraph)
    for table in doc.tables:
        do_table(table)
    return matched


def apply_header_fields(doc: DocumentObject, header_fields: List[dict]) -> dict:
    """Scans every paragraph in the body and every table cell for 'Label: value'
    patterns and overwrites the value when the label fuzzy-matches an extracted
    header field. Returns a report of matched/unmatched labels."""
    candidates = {normalize(f["label"]): f.get("value", "") for f in header_fields if f.get("label")}
    used = set()
    matched_in_doc = []
    all_paragraphs: List[Paragraph] = []
    # Cells where a genuine inline "Label: value" (non-empty value) was found,
    # whether or not it ended up matching -- tracked so the adjacent-cell pass
    # below never reinterprets one of these as a "bare label" wanting its
    # value from the NEXT cell. Without this, clearing an unmatched inline
    # field's value (right below) turns e.g. "Batch No: KALM26001A" into
    # "Batch No: " -- which now LOOKS bare -- and the adjacent-cell pass would
    # then treat it as needing a value from whatever cell happens to sit next
    # to it, blanking that unrelated (and possibly already correctly filled)
    # cell too. Confirmed on a real template (DURGA-COA) where this cascade
    # wiped out an already-correct "Batch No" value because the preceding
    # "Invoice No:" cell had no match and got cleared.
    handled_cells = set()
    # Paragraphs / cells any pass below already filled or cleared -- the
    # final known-label pass never second-guesses those.
    touched_paragraphs = set()
    touched_cells = set()

    def process_paragraph(paragraph: Paragraph, cell=None):
        all_paragraphs.append(paragraph)
        # Collect all field spans first (so replacing one doesn't shift offsets
        # for ones we haven't processed yet) -- process back-to-front.
        fields = list(_iter_field_candidates(paragraph))
        if fields:
            touched_paragraphs.add(id(paragraph._p))
        for label_raw, value_start, value_end in reversed(fields):
            if value_end <= value_start:
                # No inline value to overwrite (e.g. a bare "Product Name:" cell
                # whose value lives in a *different* cell) -- leave it for the
                # adjacent-cell pass below rather than falsely marking it "used".
                continue
            if cell is not None:
                handled_cells.add(id(cell._tc))
            key = best_match(normalize(label_raw), candidates)
            if key is None:
                # Templates here are real, previously-filled-out CoAs reused
                # purely for their layout -- every "Label: value" field in the
                # header still carries whatever product it was last filled
                # out for. If the new source PDF genuinely doesn't state this
                # field at all, leaving the template's old value in place
                # means the output silently shows real-looking data (a batch
                # quantity, a date, a spec) that has nothing to do with the
                # document just processed -- confirmed on a real example: a
                # CoA with no net-weight/quantity field anywhere came out
                # with "Net weight: 50 kg", the unrelated leftover value from
                # whatever product the template was originally created for.
                # Clear it instead, so an unfilled field is visibly blank
                # rather than silently wrong.
                replace_paragraph_span(paragraph, value_start, value_end, "")
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
                    process_paragraph(paragraph, cell=cell)
                for nested in cell.tables:
                    process_table(nested)

    for table in doc.tables:
        process_table(table)

    # Second pass: no-colon "Label Value" fallback for header fields the first pass missed.
    # Guard against a malformed extracted field missing "label" entirely (should never
    # happen given the AI schema requires it, but a defensive .get() here is cheap
    # insurance against a KeyError crashing the whole request over one bad field).
    remaining = [f["label"] for f in header_fields if f.get("label") and normalize(f["label"]) not in used]
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
                touched_paragraphs.add(id(paragraph._p))
                used.add(key)
                matched_in_doc.append(label_raw)
                remaining = [l for l in remaining if normalize(l) != key]

    # Third pass: 'Label:' cell + value in the *next* cell of the same table row --
    # very common layout (e.g. CoA DERUN's header table) that the two passes above
    # can't see because the label and value never share one paragraph/cell.
    results = find_results_table(doc)
    skip_table = results[0] if results else None
    matched_in_doc.extend(
        apply_header_fields_adjacent_cells(
            doc, header_fields, used, skip_table, handled_cells, touched_cells
        )
    )

    # Fourth pass: known header labels written WITHOUT a colon, either as
    # "Label Value" in one cell/line or as separate label / [":"] / value
    # cells. See _apply_known_label_fields.
    matched_in_doc.extend(
        _apply_known_label_fields(
            doc, candidates, used, results, handled_cells, touched_cells, touched_paragraphs
        )
    )

    unmatched = [f["label"] for f in header_fields if f.get("label") and normalize(f["label"]) not in used]
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


def _find_header_table(doc: DocumentObject, skip_table: Optional[Table] = None) -> Optional[Table]:
    """Best-effort locate the template's own 'header info' table -- the one
    apply_header_fields_adjacent_cells fills (Product Name / Batch No /
    dates / etc), as opposed to the test-results table or some unrelated
    nested table. Heuristic: the first table (other than the results table)
    that has at least one bare 'Label:' cell recognized by
    _cell_is_bare_label -- exactly the same signal the adjacent-cell filler
    itself keys off of."""
    for table in doc.tables:
        if table is skip_table:
            continue
        for row in table.rows:
            for cell in _unique_row_cells(row):
                if _cell_is_bare_label(cell.text):
                    return table
    return None


def apply_unmatched_header_fields_as_new_rows(
    doc: DocumentObject,
    header_fields: List[dict],
    unmatched_labels: List[str],
    skip_table: Optional[Table] = None,
) -> List[str]:
    """Appends new row(s) to the template's own header table for extracted
    header fields that have a real value but matched no existing template
    field at all -- confirmed real gap: a reference-standard CoA's own
    Storage Condition / Usage / Use Method fields have no corresponding row
    anywhere in a bulk-API template's fixed header table, so they were
    silently dropped from the output entirely instead of showing up
    anywhere. Mirrors the test-results table's own "append rows for tests
    the template didn't anticipate" behavior, applied to the header table.

    New rows reuse the header table's own last row as a formatting donor
    (cloned via deepcopy, same as the results-table append logic), packing
    two label/value pairs per row when the table has 4 columns (matching the
    template's existing density) or one pair per row otherwise. Returns the
    list of labels that got appended, so the caller can fold them into the
    header report as "matched" instead of "unmatched"."""
    remaining = [
        f for f in header_fields
        if f.get("label") in unmatched_labels and (f.get("value") or "").strip()
    ]
    if not remaining:
        return []

    header_table = _find_header_table(doc, skip_table)
    if header_table is None:
        return []

    donor_row = header_table.rows[-1]
    donor_cells = _unique_row_cells(donor_row)
    n_cells = len(donor_cells)
    if n_cells < 2:
        return []

    added_labels: List[str] = []
    insert_after_tr = donor_row._tr
    idx = 0
    while idx < len(remaining):
        new_tr = copy.deepcopy(donor_row._tr)
        insert_after_tr.addnext(new_tr)
        insert_after_tr = new_tr
        new_row = next(r for r in header_table.rows if r._tr is new_tr)
        new_cells = _unique_row_cells(new_row)
        pos = 0
        while pos + 1 < n_cells and idx < len(remaining):
            field = remaining[idx]
            label_text = (field.get("label") or "").rstrip(":： ").strip() + ":"
            replace_cell_text(new_cells[pos], label_text)
            replace_cell_text(new_cells[pos + 1], field.get("value", ""))
            added_labels.append(field["label"])
            idx += 1
            pos += 2
        while pos < n_cells:
            replace_cell_text(new_cells[pos], "")
            pos += 1

    return added_labels


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

    # A handful of pharma CoA test-parameter phrasings that mean the same
    # regulatory concept but share almost no characters, so no amount of
    # string-similarity tuning would ever line them up: a template's row for
    # the generic "any one impurity, whichever is largest" catch-all test is
    # as likely to be labeled "Single impurity" as a source document's own
    # version is to say "Any individual impurity" -- confirmed both exist,
    # for the exact same concept, across two real documents. Canonicalize
    # known variants to one phrase before building match keys, on both the
    # template-row side and the extracted-item side, so whichever wording
    # either happens to use, they still line up.
    _TEST_PHRASE_SYNONYMS = [
        (re.compile(r"\bany\s+(?:single\s+|one\s+)?individual\s+impurity\b", re.I), "single impurity"),
        (re.compile(r"\bany\s+single\s+impurity\b", re.I), "single impurity"),
        (re.compile(r"\bmax(?:imum)?\s+individual\s+impurity\b", re.I), "single impurity"),
        (re.compile(r"\blargest\s+(?:single\s+|individual\s+)?impurity\b", re.I), "single impurity"),
    ]

    def _canon_test_phrases(text: str) -> str:
        for pat, repl in _TEST_PHRASE_SYNONYMS:
            text = pat.sub(repl, text)
        return text

    # Build match keys for extracted results: parameter + first part of spec (disambiguates
    # generically-named rows like "By IR" / "By HPLC" that repeat under a shared heading).
    def extracted_key(item: dict) -> str:
        return normalize(_canon_test_phrases(
            item.get("parameter", "") + " " + (item.get("specification", "") or "")[:80]
        ))

    extracted = [
        {
            **item,
            "_key": extracted_key(item),
            "_param_key": normalize(_canon_test_phrases(item.get("parameter", ""))),
        }
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
            "key": normalize(_canon_test_phrases(param_text + " " + spec_text[:80])),
            "param_key": normalize(_canon_test_phrases(param_text)),
        })

    # --------------------------------------------------------------------- #
    # Grouped/repeated-label sections (e.g. a template's "Related substances"
    # or "Residual solvents" block, where every row in the group repeats the
    # exact same label and each row's real identity -- which specific
    # impurity/solvent -- lives only in its spec text, e.g. "Impurity B<0.10%")
    # get matched POSITIONALLY against the source's own same-category items,
    # in the source's own order, instead of via fuzzy text similarity.
    #
    # Fuzzy matching is actively unreliable for this shape of data: every row
    # in such a group shares near-identical wording ("Impurity X NMT 0.5%"),
    # differing only in the specific letter/number, so almost ANY pairing
    # within the group can clear the similarity cutoff -- confirmed for real:
    # a template's "Impurity E" row ended up paired with an unrelated source
    # impurity purely because the surrounding text happened to align better
    # than the semantically-correct pairing would. The template is reused
    # across unrelated products, so its own specific impurity/solvent names
    # are themselves just leftover placeholder data from whatever product it
    # was last filled out for -- there's no reason to expect them to
    # correspond by name to whatever the newly-uploaded PDF contains. So:
    # treat the whole run of same-labeled rows as interchangeable slots, and
    # place the source's own list of same-category items into them in order,
    # growing the group (cloning the last row) or shrinking it (deleting
    # extra rows) as needed -- and relabel EVERY row in the group with its
    # own item's specific parameter text, so the output actually distinguishes
    # them (rather than every row reading the same generic group name).
    def _group_key(text: str) -> str:
        return normalize(_canon_test_phrases(text))

    def _singularize_tokens(text: str) -> List[str]:
        return [t[:-1] if t.endswith("s") and len(t) > 3 else t for t in text.split()]

    def _contains_group_label(group_key: str, item_key: str) -> bool:
        """True if item_key contains group_key as a contiguous run of tokens,
        tolerant of trivial singular/plural differences (a template's
        "Residual solvents" section vs a source document's own "Residual
        solvent" column heading, transcribed singular by the AI)."""
        g_tokens = _singularize_tokens(group_key)
        i_tokens = _singularize_tokens(item_key)
        n, m = len(g_tokens), len(i_tokens)
        if n == 0:
            return False
        return any(i_tokens[start:start + n] == g_tokens for start in range(m - n + 1))

    def _specific_item_name(item: dict, group_label: str) -> str:
        """Strip a leading occurrence of the group's own label off an
        extracted item's full parameter text (e.g. "Related substances(HPLC)
        Impurity A" with group label "Related substances" -> "Impurity A"),
        tolerant of a trailing plural and a parenthetical like "(HPLC)".
        Falls back to the original text unchanged if no such prefix is
        found."""
        raw_param = item.get("parameter", "") or ""
        base = group_label.strip()
        if base.endswith("s"):
            base = base[:-1]
        pattern = r"(?i)^\s*" + re.escape(base) + r"s?(?:\s*\([^)]*\))?\s*[-:]?\s*"
        m = re.match(pattern, raw_param)
        specific_name = raw_param[m.end():].strip() if m else raw_param
        return specific_name or raw_param

    def _row_display_text(item: dict, group_label: str) -> str:
        """Build the visible text for one row of a grouped section: the
        item's own specific name plus its specification, combined into one
        string (e.g. "Ethanol NMT 0.5%(5000ppm)"), matching how the source
        document itself displays a named sub-item under a shared group
        title. The template's Items/parameter column for this group is a
        single cell vertically merged across every row in the group (one
        shared "Residual solvent" / "Related substances" title, confirmed
        directly in the template's XML) -- Word only ever displays that one
        title for the whole span, so a per-row identity has to live in the
        row's own (independent, non-merged) specification text instead, not
        in the Items column."""
        specific_name = _specific_item_name(item, group_label)
        spec = (item.get("specification", "") or "").strip()
        if specific_name and spec:
            return f"{specific_name} {spec}"
        return specific_name or spec

    row_groups = []
    gi = 0
    while gi < len(candidate_rows):
        gj = gi + 1
        gkey = candidate_rows[gi]["param_key"]
        while gj < len(candidate_rows) and candidate_rows[gj]["param_key"] == gkey and gkey:
            gj += 1
        if gj - gi >= 2:
            row_groups.append({"key": gkey, "rows": candidate_rows[gi:gj]})
        gi = gj

    # Labels shared by 2+ sibling rows carry no per-row identity of their
    # own -- used below (even for rows the group-matching above couldn't
    # confidently reassign, e.g. no extracted item's parameter text happened
    # to contain the group's label) so the generic matching pass still knows
    # to relabel them with whatever specific item ends up matched there,
    # rather than leaving every row reading the same generic group name.
    repeated_param_keys = {g["key"] for g in row_groups}

    grouped_row_trs = set()
    group_appended = 0

    for group in row_groups:
        gkey = group["key"]
        if not gkey:
            continue
        # Items whose own parameter text contains this group's shared label
        # as a prefix/substring (e.g. group label "related substances"
        # inside an extracted parameter "Related substances(HPLC) Impurity
        # 16"), kept in their original (source-document) order.
        member_idxs = [
            idx for idx, item in enumerate(extracted)
            if idx not in used_extracted_idx and _contains_group_label(gkey, item["_param_key"])
        ]
        if not member_idxs:
            continue

        group_rows = group["rows"]
        n_rows = len(group_rows)
        n_items = len(member_idxs)
        group_label = group_rows[0]["param_text"]

        for pos in range(min(n_rows, n_items)):
            cand = group_rows[pos]
            item = extracted[member_idxs[pos]]
            cells = cand["cells"]
            new_spec = _row_display_text(item, group_label)
            new_result = item.get("result", "")
            # The Items/parameter column is left exactly as the template
            # wrote it (the single merged group title, e.g. "Residual
            # solvent") -- this row's own specific identity goes into the
            # specification column instead, combined with its limit text.
            if new_spec and spec_pos is not None and spec_pos < len(cells):
                replace_cell_text(cells[spec_pos], new_spec)
            if new_result:
                replace_cell_text(cells[result_pos], new_result)
                matched_count += 1
            grouped_row_trs.add(id(cand["row"]._tr))
            used_extracted_idx.add(member_idxs[pos])

        if n_rows > n_items:
            # Extra template slots this source doc has no data for -- delete.
            for cand in group_rows[n_items:]:
                rows_to_delete.append(cand["row"])
                grouped_row_trs.add(id(cand["row"]._tr))
        elif n_items > n_rows:
            # Extra source items this template has no slot for -- clone the
            # group's last row for each one, inserted right after the
            # group's own last row so it stays with its siblings instead of
            # landing wherever the table's single "last real row" happens to
            # be (which is often far away, e.g. right before CONCLUSION).
            insert_after_tr = group_rows[-1]["row"]._tr
            for extra_idx in member_idxs[n_rows:]:
                item = extracted[extra_idx]
                new_tr = copy.deepcopy(insert_after_tr)
                insert_after_tr.addnext(new_tr)
                insert_after_tr = new_tr
                new_row = next(r for r in table.rows if r._tr is new_tr)
                new_cells = _unique_row_cells(new_row)
                for pos_, cell in enumerate(new_cells):
                    if pos_ == param_pos:
                        # Leave as-is: cloned from the group's own last row,
                        # it inherits that row's vertical-merge continuation
                        # marker, so it naturally extends the same merged
                        # title cell (e.g. "Residual solvent") the rest of
                        # the group shares -- no separate write needed.
                        continue
                    elif pos_ == spec_pos:
                        replace_cell_text(cell, _row_display_text(item, group_label))
                    elif pos_ == result_pos:
                        replace_cell_text(cell, item.get("result", ""))
                    else:
                        replace_cell_text(cell, "")
                used_extracted_idx.add(extra_idx)
                group_appended += 1

    # Every row handled by the group logic above is done -- remove it from
    # candidate_rows so the general fuzzy-matching pass below never
    # reconsiders it (its extracted-item counterpart is already marked used,
    # so it's excluded from that side too).
    if grouped_row_trs:
        candidate_rows = [c for c in candidate_rows if id(c["row"]._tr) not in grouped_row_trs]

    # A template row that's the ONLY one for its category (e.g. a single
    # "Identification" row) still needs to grow into multiple rows when the
    # source splits the category into several named sub-items (e.g.
    # Identification via both an HPLC and an IR method) -- otherwise one
    # sub-item wins the row via ordinary fuzzy matching (arbitrarily,
    # whichever scores higher) and the other is left to land as a
    # disconnected standalone row wherever the table's single "append new
    # rows here" point happens to be, often far from its sibling. Detected
    # the same way a brand-new group is (2+ extracted items sharing a common
    # multi-word leading phrase that also matches this row's own label), but
    # reusing the template's existing row as the first of the new rows
    # instead of cloning a fresh one.
    #
    # Unlike the template's own pre-existing multi-row groups (Related
    # substances, Residual solvents), which keep ONE merged title cell with
    # the specific identity folded into the specification text, a category
    # that only had a single row to begin with gets a plain, distinct label
    # per row instead -- e.g. "Identification HPLC" / "Identification IR" --
    # rather than a vertical merge repeating just "Identification" on both.
    # A short method name like "HPLC"/"IR" sitting at the very front of a
    # long spec paragraph reads far less clearly than it does as its own
    # row label, which is exactly why this case gets different treatment
    # from the "Ethanol NMT 0.5%(5000ppm)"-style rows above.
    label_counts: Dict[str, int] = {}
    for cand in candidate_rows:
        label_counts[cand["param_key"]] = label_counts.get(cand["param_key"], 0) + 1

    growable_row_trs = set()
    for cand in candidate_rows:
        if label_counts.get(cand["param_key"], 0) != 1:
            continue  # not a lone row for its label -- leave to the generic pass
        label = cand["param_text"]
        member_idxs = [
            idx for idx, item in enumerate(extracted)
            if idx not in used_extracted_idx and _contains_group_label(cand["param_key"], item["_param_key"])
        ]
        if len(member_idxs) < 2:
            continue
        cells = cand["cells"]
        first_item = extracted[member_idxs[0]]
        if param_pos < len(cells):
            replace_cell_text(cells[param_pos], f"{label} {_specific_item_name(first_item, label)}".strip())
        if spec_pos is not None and spec_pos < len(cells):
            replace_cell_text(cells[spec_pos], first_item.get("specification", ""))
        if result_pos < len(cells):
            replace_cell_text(cells[result_pos], first_item.get("result", ""))
        used_extracted_idx.add(member_idxs[0])
        matched_count += 1

        insert_after_tr = cand["row"]._tr
        for extra_idx in member_idxs[1:]:
            item = extracted[extra_idx]
            new_tr = copy.deepcopy(insert_after_tr)
            insert_after_tr.addnext(new_tr)
            insert_after_tr = new_tr
            new_row = next(r for r in table.rows if r._tr is new_tr)
            new_cells = _unique_row_cells(new_row)
            for pos_, cell in enumerate(new_cells):
                if pos_ == param_pos:
                    replace_cell_text(cell, f"{label} {_specific_item_name(item, label)}".strip())
                elif pos_ == spec_pos:
                    replace_cell_text(cell, item.get("specification", ""))
                elif pos_ == result_pos:
                    replace_cell_text(cell, item.get("result", ""))
                else:
                    replace_cell_text(cell, "")
            used_extracted_idx.add(extra_idx)
            matched_count += 1
        growable_row_trs.add(id(cand["row"]._tr))

    if growable_row_trs:
        candidate_rows = [c for c in candidate_rows if id(c["row"]._tr) not in growable_row_trs]

    # Score every (row, item) pair -- best of the full key (param+spec), the
    # bare-parameter-name comparison, and a substring-containment check,
    # mirroring the two-tier fallback the old per-row logic used, but computed
    # for all pairs before any assignment.
    #
    # The containment check exists for a specific, confirmed real-world gap:
    # templates often carry a generic, repeated parameter label across a
    # whole group of rows (e.g. "Related substances" for every named
    # impurity), with the row's REAL identity only living inside its spec
    # text alongside the limit -- e.g. "Impurity B<0.10%". A freshly
    # extracted item just as reasonably reports that identity as a clean
    # parameter name ("Impurity B") with the limit separately ("NMT 0.5%").
    # Those two representations share the identity substring but differ
    # enough elsewhere (different limit wording/format) that character-level
    # SequenceMatcher on the combined strings can land just under the 0.55
    # cutoff (confirmed: 0.50 on a real CoA DERUN case), so the correct row
    # never matched, got deleted as "not in the source", and in that specific
    # template took a stamp image anchored inside it down with it. If the
    # extracted item's parameter name appears verbatim inside the row's
    # combined param+spec text, that's strong, near-certain evidence they're
    # the same row -- treat it as a high-confidence match regardless of how
    # different the rest of the text (the limit wording/format) looks.
    def _containment_score(a: str, b: str) -> float:
        shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
        if len(shorter) < 3:
            return 0.0
        # Only trust containment when the longer side is itself short/
        # structured (a repeated-label row's "Name<limit>" style spec, e.g.
        # "Related substances Impurity B 0 10"). A long, narrative spec (a
        # written-out Identification/method description) can easily contain
        # an unrelated short word by pure coincidence -- confirmed for real:
        # an "Identification" row's spec ends with "...as obtained in the
        # Assay.", a cross-reference to a DIFFERENT row, which made an
        # extracted "Assay" item wrongly match onto this "Identification"
        # row before this length guard was added.
        if len(longer) > 60:
            return 0.0
        # Word-boundary match, NOT a raw substring check -- a naive `shorter
        # in longer` would (and, when first tried, did) let "Purity" falsely
        # match a "...Impurity B..." row, since "purity" is a plain substring
        # of "impurity". \b anchors require the match to start/end on a real
        # word boundary, so "purity" no longer matches inside "impurity" but
        # "impurity b" still matches inside "...impurity b 0 10" as intended.
        return 0.92 if re.search(r"\b" + re.escape(shorter) + r"\b", longer) else 0.0

    # A template's bare, unique row label (e.g. "Assay") is very often
    # exactly what an extracted item's own parameter text STARTS WITH, plus
    # some trailing method/condition annotation the source itself adds (e.g.
    # "Assay (By HPLC, %w/w)", "pH (1% solution)") -- confirmed real case: a
    # reference-standard CoA's own "Assay (By HPLC, %w/w)" scored only 0.545
    # character-similarity against the template's bare "Assay" (just under
    # the 0.55 cutoff), so the match was missed and a whole new row got
    # appended instead of reusing the template's own Assay row. The identity
    # here is completely unambiguous -- the template's label is a strict,
    # whole-word PREFIX of the item's parameter text -- so this is scored as
    # a high-confidence match regardless of how much annotation follows.
    def _leading_label_score(cand_param_key: str, item_param_key: str) -> float:
        if len(cand_param_key) < 3:
            return 0.0
        if re.match(r"^" + re.escape(cand_param_key) + r"\b", item_param_key):
            return 0.9
        return 0.0

    scored_pairs = []  # (score, row_idx_in_candidate_rows, item_idx)
    for ri, cand in enumerate(candidate_rows):
        for ii, item in enumerate(extracted):
            full_score = difflib.SequenceMatcher(None, cand["key"], item["_key"]).ratio()
            param_score = difflib.SequenceMatcher(None, cand["param_key"], item["_param_key"]).ratio()
            contain_score = _containment_score(cand["key"], item["_param_key"])
            leading_score = _leading_label_score(cand["param_key"], item["_param_key"])
            score = max(full_score, param_score, contain_score, leading_score)
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
            # A row whose own label is one of several *identical* repeated
            # labels (e.g. "Residual solvents" shared by every solvent row)
            # is part of a merged-title group whose Items column must stay
            # exactly as the template wrote it (see _row_display_text) --
            # its per-row identity goes into the specification text instead,
            # combined with its limit. Uniquely-labeled rows (e.g. "Assay
            # (HPLC)", "Appearance") just get their specification replaced
            # outright, unchanged from before.
            if cand["param_key"] in repeated_param_keys:
                new_spec = _row_display_text(matched_item, param_text)
            else:
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
            # Written unconditionally, even when new_spec is empty -- a source
            # that genuinely states no spec/limit for this test (confirmed
            # real case: a reference-standard CoA's Assay result with no
            # separate limit column) must still clear the template's own
            # leftover spec rather than silently leave some OTHER product's
            # numeric range sitting there looking like it applies here.
            if spec_pos is not None and spec_pos < len(cells):
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

    def _shared_leading_words(a: str, b: str, min_words: int = 2) -> str:
        aw, bw = a.split(), b.split()
        n = 0
        while n < len(aw) and n < len(bw) and aw[n].lower() == bw[n].lower():
            n += 1
        return " ".join(aw[:n]) if n >= min_words else ""

    # A whole new category the template has NO existing row for at all (e.g.
    # this template was never built with an "Amino acids ratio" section, but
    # the source document has one, peptide products having amino-acid-ratio
    # tests that a small-molecule template's original product never needed)
    # still deserves the SAME title-plus-rows presentation as a category the
    # template already knew about (like "Related substances" above) -- not a
    # full compound label repeated on every single row. Detect this by
    # grouping consecutive not-yet-matched extracted items that share a
    # common multi-word leading phrase ("Amino acids ratio Asp" / "... Glu" /
    # ...); a real, unrelated pair of items is very unlikely to coincidentally
    # share 2+ leading words, so this is a safe, generic signal, not one
    # tied to any specific category name.
    unmatched_idxs = [i for i in range(len(extracted)) if i not in used_extracted_idx]
    append_runs = []  # (shared_prefix_or_None, [extracted_idx, ...])
    ri = 0
    while ri < len(unmatched_idxs):
        idx = unmatched_idxs[ri]
        run = [idx]
        shared = extracted[idx].get("parameter", "") or ""
        rj = ri + 1
        while rj < len(unmatched_idxs) and unmatched_idxs[rj] == unmatched_idxs[rj - 1] + 1:
            candidate = extracted[unmatched_idxs[rj]].get("parameter", "") or ""
            new_shared = _shared_leading_words(shared, candidate)
            if not new_shared:
                break
            shared = new_shared
            run.append(unmatched_idxs[rj])
            rj += 1
        append_runs.append((shared if len(run) >= 2 else None, run))
        ri = rj

    if last_row_for_clone is not None:
        # Always clone from the pristine, original donor row, never from a
        # just-inserted row -- cloning the previous new row (as an earlier
        # version of this did) silently carries forward whatever w:vMerge
        # marker that row ended up with (e.g. the "continue" marker a titled
        # run's own cleanup step sets on its later rows), so an UNRELATED
        # row created right after a titled run would inherit that marker and
        # get silently absorbed into the previous run's merged title. Only
        # the insertion POSITION needs to advance; the content always comes
        # from the same clean donor.
        donor_tr = last_row_for_clone._tr
        insert_after_tr = donor_tr
        for shared, run in append_runs:
            new_rows = []
            for idx in run:
                item = extracted[idx]
                new_tr = copy.deepcopy(donor_tr)
                # Insert right after the last real data row (not at the very end
                # of the table, which could be past a trailing merged
                # "CONCLUSION"/"REMARK" row) so appended tests land in a
                # sensible place and stay grouped together in source order.
                insert_after_tr.addnext(new_tr)
                insert_after_tr = new_tr
                new_row = next(r for r in table.rows if r._tr is new_tr)
                new_cells = _unique_row_cells(new_row)
                for pos, cell in enumerate(new_cells):
                    if pos == param_pos:
                        # Defensive: this row's own vMerge should already be
                        # unset (cloned from the pristine donor), but clear
                        # it explicitly anyway so an ungrouped row can never
                        # end up silently absorbed into some other titled
                        # section's merge span.
                        tcPr = cell._tc.tcPr
                        if tcPr is not None:
                            tcPr._remove_vMerge()
                        if not shared:
                            replace_cell_text(cell, item.get("parameter", ""))
                        # else: left as-is for now -- merged into one titled
                        # cell across the whole run once all its rows exist.
                    elif pos == spec_pos:
                        new_spec = _row_display_text(item, shared) if shared else item.get("specification", "")
                        replace_cell_text(cell, new_spec)
                    elif pos == result_pos:
                        replace_cell_text(cell, item.get("result", ""))
                    elif pos == no_pos and next_no is not None:
                        replace_cell_text(cell, str(next_no))
                        next_no += 1
                    else:
                        replace_cell_text(cell, "")
                new_rows.append(new_row)
                appended += 1
            if shared and new_rows:
                # Merge the Items/parameter column across this whole run into
                # one titled cell, the same way the confirmed-working
                # template-native groups (e.g. "Related substances") are
                # structured: first row w:vMerge="restart" carrying the
                # title text, every following row w:vMerge="continue" with
                # an empty cell of its own. Done by hand (setting the
                # w:vMerge markers directly) rather than via python-docx's
                # Cell.merge() -- tried first, but it computes the merged
                # span from the cells' current grid position, and on a table
                # whose rows were just inserted via raw XML (bypassing
                # python-docx's own row-management) that computation went
                # wrong, silently merging in several unrelated rows from
                # later runs along with the intended ones.
                first_cells = _unique_row_cells(new_rows[0])
                if param_pos < len(first_cells):
                    tcPr = first_cells[param_pos]._tc.get_or_add_tcPr()
                    tcPr.vMerge_val = "restart"
                    replace_cell_text(first_cells[param_pos], shared)
                for row_ in new_rows[1:]:
                    cells_ = _unique_row_cells(row_)
                    if param_pos < len(cells_):
                        tcPr = cells_[param_pos]._tc.get_or_add_tcPr()
                        tcPr.vMerge_val = "continue"
                        replace_cell_text(cells_[param_pos], "")

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
        "appended": appended + group_appended,
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
    # Shared image part (same picture referenced by several drawings, e.g. a
    # logo repeated elsewhere): don't mutate it in place, or every other use
    # of that picture would turn into the signature too.
    shared = sum(
        1 for b in part.element.iter(qn("a:blip")) if b.get(qn("r:embed")) == rId
    ) > 1
    if shared:
        new_rId, _img = part.get_or_add_image(io.BytesIO(image_png_bytes))
        blip.set(qn("r:embed"), new_rId)
    else:
        image_part._blob = image_png_bytes
    _fit_drawing_extent_to_image(run._r, image_png_bytes)
    return True


def _fit_drawing_extent_to_image(r_element, image_bytes: bytes) -> None:
    """Shrink the drawing's box so the new picture keeps its own aspect ratio,
    fitting inside the original box (centered-ish: only the overflowing
    dimension shrinks). Without this a wide signature crop swapped into a
    square stamp slot came out visibly squashed."""
    try:
        from PIL import Image
        with Image.open(io.BytesIO(image_bytes)) as im:
            w, h = im.size
    except Exception:
        return
    if not w or not h:
        return
    drawing = r_element.find(".//" + qn("w:drawing"))
    if drawing is None or not len(drawing):
        return
    container = drawing[0]
    extent = container.find(qn("wp:extent"))
    if extent is None:
        return
    cx, cy = int(extent.get("cx", 0)), int(extent.get("cy", 0))
    if cx <= 0 or cy <= 0:
        return
    img_ratio = w / h
    if cx / cy > img_ratio:
        new_cx, new_cy = int(cy * img_ratio), cy
    else:
        new_cx, new_cy = cx, int(cx / img_ratio)
    # Keep the picture centered on where the old one was (anchored images).
    for axis, delta in (("wp:positionH", cx - new_cx), ("wp:positionV", cy - new_cy)):
        pos = container.find(qn(axis))
        off = pos.find(qn("wp:posOffset")) if pos is not None else None
        if off is not None and off.text and delta:
            off.text = str(int(off.text) + delta // 2)
    extent.set("cx", str(new_cx))
    extent.set("cy", str(new_cy))
    for ext in container.iter(qn("a:ext")):
        if ext.get("cx") is not None and ext.get("cy") is not None:
            ext.set("cx", str(new_cx))
            ext.set("cy", str(new_cy))


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


_EMU_PER_INCH = 914400


def _is_letterhead_logo(doc: DocumentObject, drawing) -> bool:
    """Heuristic: is this drawing the template's company logo / letterhead
    banner (which must never be swapped for the extracted signature)?

    Across the real templates a letterhead image is always (a) anchored in
    one of the first body elements, at or before the first table, (b) pulled
    up above its anchor paragraph (negative vertical offset) or pinned near
    the top of the page, and (c) wider than it is tall. Stamps and signature
    images sit further down (after the header/results tables) or are
    roughly square. Confirmed real case: on the HANSOH template the logo and
    the red seal share the first paragraph, with the logo second in document
    order -- so the "replace the LAST image" fallback was overwriting the
    logo with the signature crop, which looked like the logo vanishing."""
    body = doc.element.body
    top = drawing
    while top is not None and top.getparent() is not body:
        top = top.getparent()
    if top is None:
        return False
    children = list(body)
    idx = children.index(top)
    first_tbl = next((i for i, el in enumerate(children) if el.tag == qn("w:tbl")), len(children))
    if idx > max(first_tbl, 1):
        return False
    if not len(drawing):
        return False
    container = drawing[0]
    extent = container.find(qn("wp:extent"))
    if extent is None:
        return False
    cx, cy = int(extent.get("cx", 0)), int(extent.get("cy", 0))
    if cy <= 0 or cx / cy < 1.8:
        return False
    if container.tag == qn("wp:inline"):
        # Inline picture in the very first body element = letterhead too.
        return idx == 0
    pos_v = container.find(qn("wp:positionV"))
    if pos_v is None:
        return False
    offset_el = pos_v.find(qn("wp:posOffset"))
    offset = int(offset_el.text) / _EMU_PER_INCH if offset_el is not None and offset_el.text else 0.0
    rel = pos_v.get("relativeFrom", "")
    if rel in ("page", "topMargin", "margin"):
        return offset < 1.5
    return offset < 0


def _run_is_letterhead_logo(doc: DocumentObject, run) -> bool:
    return any(_is_letterhead_logo(doc, dr) for dr in run._r.findall(".//" + qn("w:drawing")))


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
            if run is None or _run_is_letterhead_logo(doc, run):
                continue
            if _replace_image_in_run_drawing(p, run, signature_png_bytes):
                return {"status": "replaced_existing_image_at_anchor"}

        _append_image_to_paragraph(anchor_paragraph, signature_png_bytes)
        return {"status": "appended_at_keyword_anchor"}

    # 2) No keyword anchor found -- fall back to replacing the LAST image in the
    #    document that isn't the letterhead logo, but only if there's more than
    #    one image total (so we don't clobber a document's sole logo).
    drawing_runs = _all_drawing_runs(doc)
    if len(drawing_runs) >= 2:
        candidates = [(p, r) for p, r in drawing_runs if not _run_is_letterhead_logo(doc, r)]
        if candidates:
            paragraph, run = candidates[-1]
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


def _sanitize_str_fields(items: List[dict], keys: List[str]) -> List[dict]:
    """Coerce the given keys to plain strings (None -> ''), and drop any item
    whose value for `keys[0]` (the primary identifying field: label/
    parameter) is missing or empty entirely. The AI extraction schema types
    every one of these fields as a required string, so this should normally
    be a no-op -- but an occasional null slipping through (or any other
    caller passing hand-built data) would otherwise crash deep inside string
    concatenation/regex code with an unhelpful TypeError instead of just
    treating the field as blank."""
    out = []
    for item in items:
        if not isinstance(item, dict):
            continue
        if not item.get(keys[0]):
            continue
        clean = dict(item)
        for k in keys:
            if clean.get(k) is None:
                clean[k] = ""
            elif not isinstance(clean.get(k), str):
                clean[k] = str(clean[k])
        out.append(clean)
    return out


def _drawing_id(drawing) -> Optional[str]:
    doc_pr = drawing.find(".//" + qn("wp:docPr"))
    return doc_pr.get("id") if doc_pr is not None else None


def _snapshot_results_table_images(doc: DocumentObject):
    """Record every picture living inside the results table (with its row /
    cell position) before apply_test_results reshapes that table.

    Some templates anchor a stamp inside an ordinary results row -- CoA
    DERUN's red stamp sits in the 'Single impurity' row. replace_cell_text
    already refuses to wipe it, but when the results engine DELETES that
    row (a test the new product doesn't have) or rebuilds its group, the
    stamp was deleted along with the row; if a row carrying it is cloned,
    the stamp would be duplicated instead."""
    found = find_results_table(doc)
    if not found:
        return None
    table = found[0]
    items = []
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(_unique_row_cells(row)):
            for drawing in cell._tc.iter(qn("w:drawing")):
                did = _drawing_id(drawing)
                if did is None:
                    continue
                run = drawing.getparent()
                while run is not None and run.tag != qn("w:r"):
                    run = run.getparent()
                if run is None:
                    continue
                items.append((did, r_idx, c_idx, copy.deepcopy(run)))
    return (table, items) if items else None


def _restore_results_table_images(doc: DocumentObject, snapshot) -> None:
    if not snapshot:
        return
    table, items = snapshot
    body = doc.element.body
    for did, r_idx, c_idx, run_copy in items:
        hits = [d for d in body.iter(qn("w:drawing")) if _drawing_id(d) == did]
        if len(hits) > 1:
            # Row got cloned: keep the first copy only.
            for extra in hits[1:]:
                run = extra.getparent()
                while run is not None and run.tag != qn("w:r"):
                    run = run.getparent()
                if run is not None and run.getparent() is not None:
                    run.getparent().remove(run)
            continue
        if hits or table._tbl.getparent() is None or not len(table.rows):
            continue
        # Lost with a deleted row: re-attach it to the row now at that
        # position (or the last row), in the same column where possible.
        row = table.rows[min(r_idx, len(table.rows) - 1)]
        cells = _unique_row_cells(row)
        cell = cells[min(c_idx, len(cells) - 1)]
        cell.paragraphs[0]._p.append(run_copy)


# Templates whose own stamp/signature must stay exactly as it is in the
# template -- the signature cropped from the uploaded PDF is NOT inserted.
# Match is on the template file name (case-insensitive, without .docx).
# Add a template's name here to give it the same behaviour.
KEEP_TEMPLATE_STAMP = {
    "hansoh coa",
}


def _keeps_own_stamp(template_path: str) -> bool:
    import os
    stem = os.path.splitext(os.path.basename(str(template_path)))[0]
    return " ".join(stem.lower().split()) in KEEP_TEMPLATE_STAMP


def fill_template(
    template_path: str,
    output_path: str,
    header_fields: List[dict],
    test_results: List[dict],
    signature_png_bytes: Optional[bytes] = None,
) -> FillReport:
    doc = Document(template_path)

    header_fields = _sanitize_str_fields(header_fields or [], ["label", "value"])
    test_results = _sanitize_str_fields(test_results or [], ["parameter", "specification", "result"])

    # Exp Date vs. Retest Date: only relabel the template's "Exp Date" field
    # to "Retest date" when the source genuinely has no expiry date of its
    # own but does state a retest date -- if a real Exp Date is present,
    # leave the template's label alone and fill it exactly as before.
    exp_field = _find_header_field(header_fields, _is_exp_date_label)
    retest_field_src = _find_header_field(header_fields, _is_retest_date_label)
    if (
        exp_field is not None
        and retest_field_src is None
        and _relabel_exp_date_to_retest(doc, dry_run=True) == 0
    ):
        # Template only has a "Retest date" slot but the source states a real
        # expiry date: relabel that slot so the expiry date has somewhere to go.
        _relabel_retest_date_to_exp(doc)
    if exp_field is None:
        retest_field = retest_field_src
        if retest_field is not None:
            if _relabel_retest_date_to_exp(doc, dry_run=True) == 0:
                # (Skipped when the template ALREADY has its own Retest date
                # slot -- relabeling Exp too would give it two of them.)
                _relabel_exp_date_to_retest(doc)
            # The source's own wording for this field (e.g. a reference-
            # standard CoA's "Re-qualification Date") is often nowhere near
            # "Retest date" character- or token-wise, so the generic label
            # matcher below could easily fail to reconnect them even though
            # the template text was just rewritten to say exactly that.
            # Renaming this field's own label to the exact relabeled text
            # (rather than appending a second, duplicate field) guarantees
            # the value lands on the relabeled cell regardless of how
            # differently the source phrased the underlying concept, and
            # also means it doesn't ALSO get picked up as an "unmatched"
            # field and appended a second time as its own new header row.
            header_fields = [
                {"label": "Retest date", "value": retest_field.get("value", "")} if f is retest_field else f
                for f in header_fields
            ]

    header_report = apply_header_fields(doc, header_fields)

    # Any extracted header field that matched no existing template row at
    # all (e.g. a reference-standard CoA's Storage Condition / Usage / Use
    # Method, which a bulk-API template's header table has no slot for)
    # still deserves to show up somewhere rather than being silently
    # dropped -- append new header-table row(s) for it, the same way the
    # results table already appends rows for untemplated tests. Skips the
    # results table itself so a leftover field can never land inside it.
    results = find_results_table(doc)
    skip_table = results[0] if results else None
    appended_header_labels = apply_unmatched_header_fields_as_new_rows(
        doc, header_fields, header_report["unmatched_extracted_fields"], skip_table
    )
    if appended_header_labels:
        header_report["matched"] = header_report["matched"] + appended_header_labels
        header_report["unmatched_extracted_fields"] = [
            l for l in header_report["unmatched_extracted_fields"] if l not in appended_header_labels
        ]

    image_snapshot = _snapshot_results_table_images(doc)
    results_report = apply_test_results(doc, test_results)
    _restore_results_table_images(doc, image_snapshot)
    if _keeps_own_stamp(template_path):
        signature_report = {"status": "kept_template_stamp"}
    else:
        signature_report = apply_signature(doc, signature_png_bytes)

    doc.save(output_path)

    return FillReport(
        header_fields=header_report,
        test_results=results_report,
        signature=signature_report,
    )
