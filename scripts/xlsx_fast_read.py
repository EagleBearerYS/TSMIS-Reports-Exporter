"""Cheaper read-only workbook loads, with openpyxl's own results.

``openpyxl.load_workbook(read_only=True)`` constructs a ``ReadOnlyWorksheet``
for EVERY sheet, and each one immediately looks for the sheet's ``<dimension>``
element by iterating XML end-events until it has seen either that element or
the end of ``<sheetData>``. A workbook written by openpyxl's write-only writer
— which is every consolidated workbook, TSN library workbook and comparison
this app produces — has no ``<dimension>``, so that search parses every sheet
to its last row and then reports "unsized". A comparison opens each of its two
inputs several times (the role gates, the marker check, then the read), and
each open paid one full parse per sheet for an answer that is in the first few
hundred bytes.

`cheap_sizing` makes that one question cheap for the loads made inside it, on
the calling thread only:

  * the ``<dimension>`` element can only precede ``<sheetData>`` (the schema
    orders them), so the sheet's head — up to the ``<sheetData`` tag — decides
    it;
  * no ``<dimension`` before ``<sheetData`` -> the sheet is left unsized, which
    is exactly what the full scan concludes;
  * anything else — a ``<dimension`` in the head, no recognizable
    ``<sheetData`` (a prefixed or re-encoded sheet), an unreadable member — is
    handed to openpyxl's own method unchanged.

So row iteration, padding and every value are openpyxl's, bit for bit. The one
observable difference is WHEN a damaged sheet is noticed: a sheet whose XML is
broken past its head used to fail the load itself; inside `cheap_sizing` it
fails when its rows are read, and a damaged sheet nobody reads is not noticed.
That is why this is opt-in per call site and is NOT applied to the commit
gates, whose read-only load is there to prove a just-written workbook opens.
"""
from __future__ import annotations

import contextlib
import threading

try:
    from openpyxl.worksheet import _reader
    from openpyxl.worksheet._read_only import ReadOnlyWorksheet
    _DEPS_OK = True
except ImportError:                      # pragma: no cover - mirrors the loaders
    _DEPS_OK = False

__all__ = ["cheap_sizing"]

_LOCAL = threading.local()
_HEAD_STEP = 64 * 1024
_HEAD_LIMIT = 4 * 1024 * 1024            # a sheet head larger than this is not ours


def _unsized_by_head(sheet):
    """True when the sheet's head proves it carries no ``<dimension>``."""
    try:
        with sheet._get_source() as src:
            head = b""
            while len(head) < _HEAD_LIMIT:
                block = src.read(_HEAD_STEP)
                if not block:
                    return False
                head += block
                at = head.find(b"<sheetData")
                if at >= 0:
                    return b"<dimension" not in head[:at]
    except Exception:  # noqa: BLE001  # silent-ok: openpyxl's own method then reports it
        return False
    return False


if _DEPS_OK:
    _FULL_GET_SIZE = ReadOnlyWorksheet._get_size

    def _get_size(self):
        if getattr(_LOCAL, "depth", 0) and _unsized_by_head(self):
            return None                  # unsized, as the full scan would conclude
        return _FULL_GET_SIZE(self)

    ReadOnlyWorksheet._get_size = _get_size

    # ----------------------------------------------------------------------- #
    # Plain inline strings.
    #
    # For an inline-string cell openpyxl builds a ``Text`` object from the
    # ``<is>`` element — a descriptor-driven class that also models rich-text
    # runs and phonetic hints — and then asks it for ``content``. The workbooks
    # this app writes hold nothing but ``<is><t>…</t></is>``, where that
    # round-trip returns the ``<t>`` element's text and was more than a third of
    # an input load. Exactly that shape is answered directly; any other (runs,
    # phonetics, attributes, a formula, rich-text mode) is openpyxl's.
    # ----------------------------------------------------------------------- #
    _T_TAG = "{%s}t" % _reader.SHEET_MAIN_NS
    _INLINE_STRING = _reader.INLINE_STRING
    _coordinate_to_tuple = _reader.coordinate_to_tuple
    _FULL_PARSE_CELL = _reader.WorkSheetParser.parse_cell

    def _parse_cell(self, element):
        if (getattr(_LOCAL, "depth", 0) and element.get("t") == "inlineStr"
                and self.data_only and not self.rich_text):
            child = element.find(_INLINE_STRING)
            if child is not None and len(child) == 1 and not child.attrib:
                text = child[0]
                if text.tag == _T_TAG and len(text) == 0:
                    coordinate = element.get("r")
                    style_id = element.get("s", 0)
                    if style_id:
                        style_id = int(style_id)
                    if coordinate:
                        row, column = _coordinate_to_tuple(coordinate)
                        self.col_counter = column
                    else:
                        self.col_counter += 1
                        row, column = self.row_counter, self.col_counter
                    value = text.text
                    return {"row": row, "column": column,
                            "value": "" if value is None else value,
                            "data_type": "s", "style_id": style_id}
        return _FULL_PARSE_CELL(self, element)

    _reader.WorkSheetParser.parse_cell = _parse_cell


@contextlib.contextmanager
def cheap_sizing():
    """Answer read-only sheets' size question from their head, on this thread,
    for the loads made inside the block (re-entrant)."""
    _LOCAL.depth = getattr(_LOCAL, "depth", 0) + 1
    try:
        yield
    finally:
        _LOCAL.depth -= 1
