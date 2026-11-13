"""Direct row serialization for openpyxl write-only worksheets.

openpyxl's write-only path builds a ``WriteOnlyCell`` for every cell, binds its
value through the type-inference machinery, then builds an ElementTree element
(or three) and serializes that through et_xmlfile's incremental writer, one
``file.write`` per fragment. A statewide comparison writes several million cells
that way, and it is most of the time a comparison takes: measured on the shipped
configuration (no lxml), ~8 µs per cell.

The cells those workbooks hold are overwhelmingly plain: an inline string, a
number, a formula, or an empty styled cell. This module writes such a row as one
pre-built string through the SAME text stream openpyxl is writing the sheet to,
so everything around the rows — the sheet header, column widths, conditional
formatting, comments, the styles table, the package — is still openpyxl's, and
the bytes of the row are the ones openpyxl would have produced.

That last claim is the whole contract, and it is held two ways:

  * every fragment below reproduces ``openpyxl.cell._writer.etree_write_cell``
    for exactly one cell shape, using the same escaping function, and anything
    outside those shapes raises `Unsupported`, which makes the caller build that
    row with real openpyxl cells instead;
  * ``build/check_xlsx_fast_rows.py`` writes the same content through both paths
    and requires byte-identical package members, and the comparison canaries do
    the same on real workbooks.

It relies on pinned internals (openpyxl 3.1.5 + et_xmlfile 2.0.0, both
hash-locked in the build). `install` verifies the shape it needs and returns a
stream with ``raw = False`` when it is not there — notably when lxml is
installed, whose incremental writer has no text stream to borrow — and the
caller then writes every row the classic way. ``TSMIS_CLASSIC_XLSX=1`` forces
that for a session.
"""
from __future__ import annotations

import os
import xml.etree.ElementTree as ET

try:
    from openpyxl.cell.cell import ERROR_CODES, ILLEGAL_CHARACTERS_RE
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.exceptions import IllegalCharacterError
    _DEPS_OK = True
except ImportError:                      # pragma: no cover - mirrors compare_core
    _DEPS_OK = False

__all__ = ["Unsupported", "RowStream", "install", "classic_forced", "ERROR_CODES",
           "string_text", "cell_string", "cell_number", "cell_formula",
           "cell_error", "cell_bool", "cell_empty", "column_letter"]

# The escaping et_xmlfile applies to element text. Called, not re-implemented, so
# a Python that changes it changes both paths together.
_escape = ET._escape_cdata
_MAX_STRING = 32767                      # openpyxl's check_string truncation
_illegal = ILLEGAL_CHARACTERS_RE.search if _DEPS_OK else None
_LETTERS = {}


class Unsupported(Exception):
    """This row needs openpyxl's own cell machinery; build it the classic way."""


def classic_forced():
    """True when the session asked for openpyxl's own row path everywhere."""
    return os.environ.get("TSMIS_CLASSIC_XLSX", "").strip().lower() in ("1", "true", "yes")


def column_letter(index):
    """`get_column_letter` for a 1-based column, memoized."""
    letter = _LETTERS.get(index)
    if letter is None:
        letter = _LETTERS[index] = get_column_letter(index)
    return letter


def string_text(value):
    """The text openpyxl would store for a string value (``Cell.check_string``):
    coerced to ``str``, truncated to 32,767 characters, and refused when it
    carries a character XML cannot hold."""
    text = value if type(value) is str else str(value)
    if len(text) > _MAX_STRING:
        text = text[:_MAX_STRING]
    # Printable ASCII cannot hold a control character, which is all the pattern
    # looks for; the two C-level scans are several times cheaper than the search.
    if not (text.isascii() and text.isprintable()) and _illegal(text):
        raise IllegalCharacterError(f"{text} cannot be used in worksheets.")
    return text


# --------------------------------------------------------------------------- #
# One fragment per cell shape. `ref` is the coordinate ("C14"); `attr` is the
# style attribute including its leading space (' s="7"'), or "" for an unstyled
# cell. Each docstring quotes the openpyxl branch it reproduces.
# --------------------------------------------------------------------------- #
def cell_string(ref, attr, text):
    """``data_type == 's'``: an inline string. An empty string is the bare cell;
    text with leading or trailing whitespace carries ``xml:space="preserve"``
    (``openpyxl.xml.functions.whitespace``)."""
    if text == "":
        return f'<c r="{ref}"{attr} t="inlineStr" />'
    stripped = text.strip()
    if stripped and stripped != text:
        return (f'<c r="{ref}"{attr} t="inlineStr"><is><t xml:space="preserve">'
                f'{_escape(text)}</t></is></c>')
    return f'<c r="{ref}"{attr} t="inlineStr"><is><t>{_escape(text)}</t></is></c>'


def cell_number(ref, attr, value):
    """``data_type == 'n'`` with a finite int or float (``safe_string``)."""
    return f'<c r="{ref}"{attr} t="n"><v>{"%.16g" % value}</v></c>'


def cell_formula(ref, attr, text):
    """``data_type == 'f'``: the formula without its leading ``=``, and the empty
    cached value openpyxl writes beside it."""
    return f'<c r="{ref}"{attr}><f>{_escape(text[1:])}</f><v /></c>'


def cell_error(ref, attr, text):
    """``data_type == 'e'``: one of Excel's error literals."""
    return f'<c r="{ref}"{attr} t="e"><v>{_escape(text)}</v></c>'


def cell_bool(ref, attr, value):
    """``data_type == 'b'``."""
    return f'<c r="{ref}"{attr} t="b"><v>{1 if value else 0}</v></c>'


def cell_empty(ref, attr):
    """A styled cell with no value. (An UNSTYLED one is not written at all —
    ``WorksheetWriter.write_row`` skips it — so callers never ask for that.)"""
    return f'<c r="{ref}"{attr} t="n" />'


def _sheet_rows(ws, state):
    """``WriteOnlyWorksheet._write_rows`` with one addition: a row sent as a
    ``str`` is already serialized and is written as-is. Everything else — a list
    of cells or values — takes openpyxl's own path, so the two kinds of row can
    interleave freely and number themselves from one counter."""
    try:
        xf = ws._writer.xf.send(True)
    except StopIteration:
        ws._already_saved()
    write = getattr(xf, "_file", None)
    if callable(write) and type(xf).__module__.split(".")[0] == "et_xmlfile":
        state["write"] = write
    with xf.element("sheetData"):
        row_idx = 1
        try:
            while True:
                row = (yield)
                if type(row) is str:
                    write(row)
                else:
                    ws._writer.write_row(xf, ws._values_to_row(row, row_idx), row_idx)
                row_idx += 1
                state["row"] = row_idx
        except GeneratorExit:  # silent-ok: the sheet is closing; fall through to its tail
            pass
    ws._writer.xf.send(None)


class RowStream:
    """The row sink of one write-only sheet whose stream this module drives."""
    __slots__ = ("ws", "raw", "_state", "_send")

    def __init__(self, ws, generator, state):
        self.ws = ws
        self._state = state
        self._send = generator.send
        self.raw = state["write"] is not None

    @property
    def row(self):
        """The 1-based number the NEXT row will get."""
        return self._state["row"]

    def plain_row(self):
        """True when the next row carries no row-level attributes (height, …),
        which is the only kind `emit` writes."""
        return self.raw and self._state["row"] not in self.ws.row_dimensions

    def emit(self, cells_xml):
        """Write the next row from its already-serialized cells."""
        self._send(f'<row r="{self._state["row"]}">{cells_xml}</row>')


def install(ws):
    """Take over `ws`'s row stream. Call it where the sheet's first
    ``ws.append`` would have been — after every column width and row height is
    set, because creating the stream writes the sheet header.

    Returns a `RowStream`, or None when the sheet already has rows or the
    session forces the classic path. ``ws.append`` keeps working on a sheet
    with a stream installed: it sends rows into the same generator."""
    if not _DEPS_OK or classic_forced():
        return None
    if getattr(ws, "_rows", None) is not None:
        return None
    ws._get_writer()
    state = {"write": None, "row": 1}
    generator = _sheet_rows(ws, state)
    next(generator)
    ws._rows = generator
    return RowStream(ws, generator, state)
