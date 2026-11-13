"""The direct row writer produces openpyxl's bytes.

`compare_core._SheetRows` serializes the big sheets' rows straight into the
sheet stream (`xlsx_fast_rows`) instead of building an openpyxl cell per value.
That is only allowed to be a SPEED change, so this writes the same content
through both paths and requires every stable OOXML package member to be
byte-exact:

  * a sheet of every cell shape the writers produce — strings with every kind
    of whitespace, control characters XML allows, markup, non-ASCII and
    astral characters, 32,767+ characters, numbers of every size, Booleans,
    formulas, error literals, a str subclass, empty cells — guarded and not,
    with and without the exact-numeric flag, in several composite styles;
  * rows the direct path must hand back to openpyxl (a date, a Decimal, a row
    with its own height), interleaved with rows it writes itself, so the row
    numbering and the styles table are shown to stay in step;
  * whole comparison workbooks in both flavors, built with the direct path and
    with ``TSMIS_CLASSIC_XLSX=1``.

It also holds the two refusals: a control character XML cannot carry and a
non-finite guarded number fail exactly as they do through openpyxl.

Run from the repo root:  build\\.venv\\Scripts\\python.exe build\\check_xlsx_fast_rows.py
"""
from __future__ import annotations

import hashlib
import io
import os
import random
import sys
import tempfile
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.exceptions import IllegalCharacterError

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT / "scripts"), str(ROOT / "build"), str(ROOT)]

import compare_core as core  # noqa: E402
import xlsx_fast_rows  # noqa: E402

_failures: list[str] = []


def check(name, cond, detail=""):
    print(f"  [{'OK ' if cond else 'FAIL'}] {name}")
    if not cond:
        if detail:
            print(f"       {detail}")
        _failures.append(name)


def _members(data):
    out = {}
    with ZipFile(io.BytesIO(data) if isinstance(data, bytes) else data) as z:
        for name in sorted(z.namelist()):
            if name != "docProps/core.xml":
                out[name] = z.read(name)
    return out


def _first_difference(a, b):
    for name in sorted(set(a) | set(b)):
        if a.get(name) != b.get(name):
            x, y = a.get(name, b""), b.get(name, b"")
            i = next((k for k in range(min(len(x), len(y))) if x[k] != y[k]),
                     min(len(x), len(y)))
            return f"{name} @ {i}: {x[max(0, i - 80):i + 80]!r} vs {y[max(0, i - 80):i + 80]!r}"
    return ""


def _abandon(wb):
    """Close an unsaved write-only workbook the way run_compare does (sheets
    first), so its row generators end while their stream is still open."""
    for ws in wb.worksheets:
        ws.close()
    wb.close()


class _Classic:
    """Force openpyxl's own row path for a block."""

    def __enter__(self):
        self._saved = os.environ.get("TSMIS_CLASSIC_XLSX")
        os.environ["TSMIS_CLASSIC_XLSX"] = "1"

    def __exit__(self, *exc):
        if self._saved is None:
            os.environ.pop("TSMIS_CLASSIC_XLSX", None)
        else:
            os.environ["TSMIS_CLASSIC_XLSX"] = self._saved


# --------------------------------------------------------------------------- #
# 1. every cell shape
# --------------------------------------------------------------------------- #
class _Key(str):
    """A str subclass, as the comparators' PhysicalKey cells are."""


class _Shouting(str):
    """A str subclass whose text is NOT its own characters — openpyxl stores
    ``str(value)``, so the direct path must too."""

    def __str__(self):
        return str.upper(self)


_KEY = _Key("R012.345")
_SHOUT = _Shouting("quiet = loud")

_STRINGS = [
    "", "plain", " lead", "trail ", "  both  ", "in  ner", "\tlead tab", "tab\tinside",
    "new\nline", "\nlead newline", "trail newline\n", "cr\rinside", "\r\n", " ", " nbsp",
    "a<b&c>d\"e'f", "&amp; already", "<![CDATA[x]]>", "]]>", "é ≠ ü — ½", "🙂 astral 𝔘",
    "=SUM(A1)", "=", "==", "+1", "-x", "@x", "#N/A", "#DIV/0!", "#NULL!", "#REF!", "#NAME?",
    "#NUM!", "#VALUE!", "#N/A ", "TRUE", "FALSE", "1.50", "007", "1e5", "0x1F", "none",
    "x" * 32767, "y" * 32768, "z" * 40000 + " ", "　ideographic space",
    "\x85next line", " line sep", "﻿bom", "a\x7fdel", "é" * 5,
]
_NUMBERS = [
    0, 1, -1, 12, 255, 10 ** 14, 10 ** 15 - 1, 10 ** 15, 10 ** 16, 12345678901234567890,
    -(10 ** 20), 0.0, -0.0, 1.5, -2.25, 1e20, 1e-7, 0.1 + 0.2, 1 / 3, 123456789.123456789,
    1.2345678901234567e-5, 9007199254740993, 4.0, 1e15, 1e16, float(10 ** 22),
    2.5e-320, 1.7976931348623157e308,
]
_OTHER = [None, True, False, _KEY, _SHOUT, _Shouting("=formula-looking")]
_UNSUPPORTED = [date(2026, 1, 2), datetime(2026, 1, 2, 3, 4, 5), Decimal("1.2300"),
                Decimal("12345678901234567890.5")]


def _build_shapes(seed):
    """One workbook of mixed rows through `_SheetRows`; returns the package."""
    rng = random.Random(seed)
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Shapes")
    body = Font(name="Arial", size=10)
    link = core._link_font()
    bold = Font(name="Arial", size=10, bold=True)
    fill = PatternFill("solid", start_color="FFE699")
    wrap = Alignment(wrap_text=True, vertical="top")
    border = Border(bottom=Side(style="thin"))
    plain = Font()                       # every id zero: openpyxl writes no style
    ws.column_dimensions["A"].width = 30
    ws.row_dimensions[1].height = 30     # header carries row attributes
    ws.row_dimensions[9].height = 22.5   # ...and so does one body row
    out = core._SheetRows(ws)
    ws.append(core._header_row(ws, ["A", "B", " C"]))
    values = _STRINGS + _NUMBERS + _OTHER
    styles = [(body, None, None, None), (link, None, None, None), (bold, fill, None, None),
              (body, None, wrap, None), (body, fill, wrap, border), (plain, None, None, None)]

    def one(value_list):
        # The row is fully decided HERE, not inside `fill_row`: a row the direct
        # path declines is filled a second time for openpyxl, so a fill must
        # describe the same cells every time it is called (as the writers' do).
        plan = []
        for value in value_list:
            guard = rng.random() < 0.6
            plan.append((value, rng.choice(styles), guard,
                         guard and rng.random() < 0.5))

        def fill_row(put):
            for value, (font, fl, al, bd), guard, exact in plan:
                put(value, font, fill=fl, align=al, border=bd, guard=guard,
                    exact_source_numeric=exact)
        return fill_row

    # every value alone, in every guard mode, in a fixed style
    for value in values:
        for guard, exact in ((False, False), (True, False), (True, True)):
            out.row(lambda put, v=value, g=guard, e=exact: put(
                v, body, guard=g, exact_source_numeric=e))
    # random wide rows, some carrying a value only openpyxl can type
    for i in range(160):
        row_values = [rng.choice(values) for _ in range(rng.randint(0, 30))]
        if i % 9 == 4:
            row_values.insert(rng.randint(0, len(row_values)), rng.choice(_UNSUPPORTED))
        out.row(one(row_values))
        if i % 25 == 7:                  # a classic row straight through ws.append
            ws.append([core._styled(ws, "classic", bold, fill), None, 7, "tail"])
    ws.append([None, None, core._styled(ws, "footer", body)])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), out


def _shapes():
    print("every cell shape, direct vs openpyxl:")
    direct, sink = _build_shapes(7)
    with _Classic():
        classic, classic_sink = _build_shapes(7)
    check("the direct path was really in use", sink.fast is True,
          "xlsx_fast_rows.install did not hand back a raw stream")
    check("...and the control really went through openpyxl", classic_sink.fast is False)
    a, b = _members(direct), _members(classic)
    check("every package member is byte-exact", a == b, _first_difference(a, b))
    sheet = a.get("xl/worksheets/sheet1.xml", b"")
    check("the fixture reaches every fragment kind",
          all(token in sheet for token in (b't="inlineStr"', b't="n"', b't="b"', b't="e"',
                                           b"<f>", b'xml:space="preserve"', b't="n" />',
                                           b't="inlineStr" />')),
          "a cell shape is no longer exercised")
    for seed in (1, 2, 3):
        direct, _ = _build_shapes(seed)
        with _Classic():
            classic, _ = _build_shapes(seed)
        a, b = _members(direct), _members(classic)
        check(f"random layout {seed} is byte-exact", a == b, _first_difference(a, b))


# --------------------------------------------------------------------------- #
# 2. refusals
# --------------------------------------------------------------------------- #
def _refusals():
    print("values openpyxl refuses are refused the same way:")
    body = Font(name="Arial", size=10)
    for label, value, guard, expected in (
            ("a control character", "bad\x01char", True, IllegalCharacterError),
            ("a control character (unguarded)", "bad\x0bchar", False, IllegalCharacterError),
            ("a guarded NaN", float("nan"), True, ValueError),
            ("a guarded infinity", float("inf"), True, ValueError)):
        results = []
        for classic in (False, True):
            wb = Workbook(write_only=True)
            ws = wb.create_sheet("R")
            ctx = _Classic() if classic else None
            if ctx:
                ctx.__enter__()
            try:
                out = core._SheetRows(ws)
                try:
                    out.row(lambda put: put(value, body, guard=guard,
                                            exact_source_numeric=guard))
                    results.append(None)
                except Exception as e:  # noqa: BLE001
                    results.append((type(e), str(e)))
            finally:
                if ctx:
                    ctx.__exit__()
                _abandon(wb)
        check(f"{label}: both paths raise {expected.__name__} with one message",
              results[0] is not None and results[0] == results[1]
              and issubclass(results[0][0], expected), repr(results))


# --------------------------------------------------------------------------- #
# 3. whole comparison workbooks
# --------------------------------------------------------------------------- #
def _ditto_resolver(_rows, _has_route):
    return {1: {2: "paired value"}}


def _schema(**extra):
    return core.CompareSchema(
        report_name="Direct Rows Lock",
        header=["Loc", "Text", "Ditto", "Med Wid", "Context", "Count"],
        side_a="TSMIS", side_b="TSN", id_noun="row", id_noun_plural="rows",
        medwid_fields=("Med Wid",), context_fields=("Context",), **extra)


def _rows(n, salt):
    rng = random.Random(salt)
    rows = []
    for i in range(n):
        route = f"{(i % 7) + 1:03d}"
        rows.append([route, f"{i:05d}.{rng.randint(0, 9)}",
                     rng.choice(["=1+1", "@literal", " padded ", "plain text", "", None,
                                 "a<b&c", "é ≠ ü"]),
                     rng.choice(["++", "+", "same", ""]),
                     rng.choice(["8V", "08V", "02", "2", "0", None]),
                     rng.choice(["left", "right", "", None]),
                     rng.choice([0, 1, 12, 1.5, 10 ** 16, None, True])])
    return rows


def _workbooks():
    print("whole comparison workbooks, direct vs openpyxl:")
    tmp = Path(tempfile.mkdtemp(prefix="tsmis_fast_rows_"))
    left, right = _rows(260, 1), _rows(250, 2)
    for label, schema, has_route in (
            ("route comparison", _schema(), True),
            ("ditto tint + hover comments",
             _schema(ditto_nonasserting=True, ditto_resolver=_ditto_resolver), True),
            ("flat comparison", _schema(), False)):
        a_rows = left if has_route else [r[1:] for r in left]
        b_rows = right if has_route else [r[1:] for r in right]
        for mode in ("values", "formulas", "both"):
            direct = tmp / f"{label}-{mode}-direct.xlsx"
            classic = tmp / f"{label}-{mode}-classic.xlsx"
            r1 = core.run_compare(schema, a_rows, b_rows, has_route, direct, mode=mode)
            with _Classic():
                r2 = core.run_compare(schema, a_rows, b_rows, has_route, classic, mode=mode)
            ok = r1.status == r2.status == "ok"
            pairs = [(direct, classic)]
            if mode == "both":
                pairs.append((direct.with_name(direct.stem + " (values).xlsx"),
                              classic.with_name(classic.stem + " (values).xlsx")))
            detail = ""
            for d, c in pairs:
                if not (ok and d.is_file() and c.is_file()):
                    ok = False
                    detail = f"{r1.status}/{r2.status}: {r1.message or r2.message}"
                    break
                a, b = _members(str(d)), _members(str(c))
                if a != b:
                    ok = False
                    detail = _first_difference(a, b)
                    break
            check(f"{label}, {mode}: byte-exact", ok, detail)
            if ok and mode == "values":
                digest = hashlib.sha256(direct.read_bytes()).hexdigest()[:12]
                check(f"{label}, {mode}: typed outcome is exact",
                      r1.comparison_outcome.to_dict() == r2.comparison_outcome.to_dict(),
                      digest)


def _guards():
    print("the direct path stands down when it must:")
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("G")
    ws.append(["already", "started"])
    check("a sheet that already has rows is not taken over",
          xlsx_fast_rows.install(ws) is None)
    _abandon(wb)
    saved = core._styled_cell
    core._styled_cell = lambda *a, **k: saved(*a, **k)
    try:
        wb = Workbook(write_only=True)
        out = core._SheetRows(wb.create_sheet("H"))
        check("a swapped cell builder (the output-lock control) builds every cell",
              out.fast is False and out.stream is None)
        _abandon(wb)
    finally:
        core._styled_cell = saved
    check("openpyxl's error literals are exactly the guard's",
          set(xlsx_fast_rows.ERROR_CODES) == set(core._EXCEL_ERROR_CODES))


if __name__ == "__main__":
    print("direct row serialization is byte-exact with openpyxl:")
    _shapes()
    _refusals()
    _workbooks()
    _guards()
    if _failures:
        print(f"\n{len(_failures)} check(s) FAILED")
        raise SystemExit(1)
    print("all good")
