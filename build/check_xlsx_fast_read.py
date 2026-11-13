"""Cheaper read-only loads return openpyxl's own rows.

Inside ``xlsx_fast_read.cheap_sizing()`` two things are answered without
openpyxl's slow path: whether a sheet carries a ``<dimension>`` (from the sheet's
head instead of a parse to its last row) and the text of a plain inline string
(the ``<t>`` element instead of a ``Text`` object). Both must be invisible in the
result, so this reads a range of workbooks both ways and requires the same rows
— values AND types — the same sheet sizes and the same errors:

  * a write-only workbook (no ``<dimension>``): gaps, empty rows, an empty sheet,
    strings with markup / whitespace / line breaks, numbers, numeric text;
  * a normal-mode workbook (has ``<dimension>``, so openpyxl sizes and pads);
  * hand-written sheet XML: rich-text runs, phonetic runs, an empty ``<t/>``,
    a bare ``<is/>``, cells with no ``r`` attribute, a formula with a cached
    inline string.

It also holds that the shortcut is opt-in, per thread, and re-entrant, that the
head scan is what answers (the full scan is not called for an unsized sheet),
and the one documented difference: a sheet damaged past its head fails when its
rows are read instead of when the workbook is opened.

Run from the repo root:  build\\.venv\\Scripts\\python.exe build\\check_xlsx_fast_read.py
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import threading
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT / "scripts"), str(ROOT / "build"), str(ROOT)]

import xlsx_fast_read  # noqa: E402

_failures: list[str] = []


def check(name, cond, detail=""):
    print(f"  [{'OK ' if cond else 'FAIL'}] {name}")
    if not cond:
        if detail:
            print(f"       {detail}")
        _failures.append(name)


def _read(path, data_only=True):
    """Every sheet as (title, max_row, max_column, [(value, type name), ...])."""
    wb = load_workbook(path, read_only=True, data_only=data_only)
    try:
        out = []
        for ws in wb.worksheets:
            rows = [tuple((v, type(v).__name__) for v in row)
                    for row in ws.iter_rows(values_only=True)]
            out.append((ws.title, ws.max_row, ws.max_column, rows))
        return out
    finally:
        wb.close()


def _both(path, data_only=True):
    plain = _read(path, data_only)
    with xlsx_fast_read.cheap_sizing():
        cheap = _read(path, data_only)
    return plain, cheap


def _write_only(path):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Data")
    bold = Font(bold=True)
    head = []
    for label in ("Route", "Text", "Count"):
        cell = WriteOnlyCell(ws, value=label)
        cell.font = bold
        head.append(cell)
    ws.append(head)
    ws.append(["005", "plain", 12])
    ws.append(["005", " padded ", 1.5])
    ws.append([None, "gap before", None, "past the header"])
    ws.append([])
    ws.append(["101", "a<b&c>d \"q\"", 10 ** 16])
    ws.append(["101", "line\nbreak\tand tab", -0.0])
    ws.append(["", "empty string first", True])
    ws.append(["=SUM(A1)", "é ≠ ü 🙂", "007"])
    wb.create_sheet("Empty")
    notes = wb.create_sheet("Notes")
    notes.append(["one cell"])
    wb.save(path)


def _normal(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sized"
    ws.append(["Route", "Text", "Count"])
    ws.append(["005", "plain", 12])
    ws["F9"] = "far corner"
    wb.save(path)


_SHEET_HEAD = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
               '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData>')
_ODD_ROWS = (
    '<row r="1"><c r="A1" t="inlineStr"><is><t>plain</t></is></c>'
    '<c r="B1" t="inlineStr"><is><r><t>rich </t></r><r><rPr><b/></rPr><t>runs</t></r></is></c>'
    '<c r="C1" t="inlineStr"><is><t>base</t><rPh sb="0" eb="1"><t>phon</t></rPh></is></c></row>'
    '<row r="2"><c r="A2" t="inlineStr"><is><t/></is></c>'
    '<c r="B2" t="inlineStr"><is/></c>'
    '<c r="C2" t="inlineStr"><is><t xml:space="preserve">  kept  </t></is></c></row>'
    '<row r="3"><c t="inlineStr"><is><t>no ref 1</t></is></c>'
    '<c t="inlineStr"><is><t>no ref 2</t></is></c><c t="n"><v>42</v></c></row>'
    '<row r="5"><c r="B5" t="inlineStr"><f>A1</f><is><t>cached text</t></is></c>'
    '<c r="C5" t="n"><v>1.25</v></c><c r="D5" t="b"><v>1</v></c>'
    '<c r="E5" t="e"><v>#N/A</v></c><c r="F5" t="str"><v>formula string</v></c></row>')


def _with_sheet_xml(source, target, xml):
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w",
                                                         zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                data = xml.encode("utf-8")
            zout.writestr(info, data)


def main():
    tmp = Path(tempfile.mkdtemp(prefix="tsmis_fast_read_"))
    try:
        print("the rows are openpyxl's, values and types:")
        streamed = tmp / "streamed.xlsx"
        _write_only(streamed)
        plain, cheap = _both(streamed)
        check("a write-only workbook reads identically", plain == cheap,
              f"{plain!r}\n       vs {cheap!r}"[:900])
        check("...and it is unsized either way (no <dimension> to find)",
              all(sheet[1] is None and sheet[2] is None for sheet in plain))
        check("the fixture carries the shapes that matter",
              any(v == (" padded ", "str") for row in plain[0][3] for v in row)
              and any(v == ("line\nbreak\tand tab", "str") for row in plain[0][3] for v in row)
              and any(v[1] == "bool" for row in plain[0][3] for v in row))

        sized = tmp / "sized.xlsx"
        _normal(sized)
        plain, cheap = _both(sized)
        check("a workbook WITH a <dimension> reads identically, padding included",
              plain == cheap and plain[0][1] == 9 and plain[0][2] == 6
              and len(plain[0][3][0]) == 6, repr(plain[0][:3]))

        odd = tmp / "odd.xlsx"
        _with_sheet_xml(streamed, odd, _SHEET_HEAD + _ODD_ROWS + "</sheetData></worksheet>")
        plain, cheap = _both(odd)
        flat = [v for row in plain[0][3] for v in row]
        check("rich runs, phonetics, empty text and ref-less cells read identically",
              plain == cheap, f"{plain[0][3]!r}\n       vs {cheap[0][3]!r}")
        check("...with openpyxl's own answers for each",
              ("rich runs", "str") in flat and ("base", "str") in flat
              and ("", "str") in flat and ("  kept  ", "str") in flat
              and ("no ref 2", "str") in flat and (42, "int") in flat, repr(flat))
        plain, cheap = _both(odd, data_only=False)
        check("formula mode (not data_only) reads identically too", plain == cheap)

        print("the shortcut is opt-in, per thread, re-entrant:")
        full_calls = []
        real_full = xlsx_fast_read._FULL_GET_SIZE

        def counting(sheet):
            full_calls.append(sheet.title)
            return real_full(sheet)

        xlsx_fast_read._FULL_GET_SIZE = counting
        try:
            _read(streamed)
            outside = len(full_calls)
            full_calls.clear()
            with xlsx_fast_read.cheap_sizing():
                _read(streamed)
                inside = len(full_calls)
                with xlsx_fast_read.cheap_sizing():
                    pass
                _read(streamed)
                still_inside = len(full_calls)
                seen = []
                worker = threading.Thread(
                    target=lambda: (_read(streamed), seen.append(len(full_calls))))
                worker.start()
                worker.join()
            full_calls.clear()
            _read(sized)
            with xlsx_fast_read.cheap_sizing():
                _read(sized)
            sized_calls = len(full_calls)
            full_calls.clear()
            _read(streamed)
            after = len(full_calls)
        finally:
            xlsx_fast_read._FULL_GET_SIZE = real_full
        check("outside the block every sheet is sized the full way", outside == 3 == after,
              f"outside {outside}, after {after}")
        check("inside it an unsized sheet is answered from its head",
              inside == 0 and still_inside == 0, f"inside {inside}, nested {still_inside}")
        check("another thread is unaffected by this thread's block",
              seen and seen[0] == 3, repr(seen))
        check("a sheet WITH a <dimension> always goes to openpyxl's own method",
              sized_calls == 2, f"full sizing calls: {sized_calls}")

        print("the one documented difference — a sheet damaged past its head:")
        broken = tmp / "broken.xlsx"
        _with_sheet_xml(streamed, broken, _SHEET_HEAD + _ODD_ROWS[:200])
        try:
            _read(broken)
            opened_plain = "read"
        except Exception as e:  # noqa: BLE001
            opened_plain = type(e).__name__
        wb = None
        try:
            with xlsx_fast_read.cheap_sizing():
                wb = load_workbook(broken, read_only=True, data_only=True)
                opened_cheap = "opened"
                try:
                    list(wb["Data"].iter_rows(values_only=True))
                    read_cheap = "read"
                except Exception as e:  # noqa: BLE001
                    read_cheap = type(e).__name__
        except Exception as e:  # noqa: BLE001
            opened_cheap, read_cheap = type(e).__name__, "-"
        finally:
            if wb is not None:
                wb.close()
        check("openpyxl alone refuses it at open", opened_plain != "read", opened_plain)
        check("inside the block it opens, and fails when its rows are read",
              opened_cheap == "opened" and read_cheap not in ("read", "-"),
              f"open: {opened_cheap}, read: {read_cheap}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    print("cheap read-only loads return openpyxl's rows:")
    main()
    if _failures:
        print(f"\n{len(_failures)} check(s) FAILED")
        raise SystemExit(1)
    print("all good")
