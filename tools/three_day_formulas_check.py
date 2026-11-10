"""Does a three-day comparison's live-formulas copy calculate to its values
workbook? Opens the copy in Excel, recalculates it, and holds every cell of the
six shared sheets against the values workbook the app wrote.

    python tools/three_day_formulas_check.py <values workbook> [<formulas copy>]
                                             [--out report.txt]

The copy defaults to '<values workbook name> (formulas).xlsx' beside the
workbook. Needs Excel on this PC (it is the only thing that can say what the
formulas calculate to) and pywin32, which the build environment has.

What it checks:

  * every cell of Summary, Routes, Changes, Added & removed, Three days (the
    grid's own columns) and Notes equals the values workbook's cell — a number
    for a number, text for text, character for character;
  * every line under the Summary's "Live formulas" heading reads OK;
  * no cell of those sheets holds an Excel error.

Exit code 0 = everything equal, 1 = a difference, 2 = could not run (no Excel,
a file missing or unreadable).

It reads the two workbooks and nothing else: it shares no code with the
comparison or with either writer.
"""
import argparse
import contextlib
import re
import sys
import time
from pathlib import Path

SHARED = ("Summary", "Routes", "Changes", "Added & removed", "Three days", "Notes")
# The copy adds its checks under the Summary and its own paragraphs to the Notes;
# every other sheet has exactly the values workbook's rows.
LONGER_IN_COPY = ("Summary", "Notes")
GRID = "Three days"
GRID_LEAD = 6
CHECK_HEADING = "Live formulas"
CHUNK_ROWS = 4000
_LINK = re.compile(r'^=HYPERLINK\("[^"]*",\s*(\d+)\)$')


class CannotRun(Exception):
    """The check could not be made at all (as opposed to: it found a difference)."""


def formulas_sibling(values_path):
    p = Path(values_path)
    return p.with_name(f"{p.stem} (formulas){p.suffix}")


def _norm(value):
    """A cell as it compares: nothing and the empty string are one thing, a
    whole number is itself whichever way it was stored, a link is its number."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str):
        link = _LINK.match(value)
        if link:
            return int(link.group(1))
    return value


def read_values(path):
    """{sheet: [[cell, ...], ...]} of the values workbook."""
    from openpyxl import load_workbook

    try:
        wb = load_workbook(path, read_only=True)
    except Exception as e:  # noqa: BLE001 - any unreadable workbook is "could not run"
        raise CannotRun(f"cannot read {path} ({type(e).__name__}: {e})") from e
    try:
        missing = [name for name in SHARED if name not in wb.sheetnames]
        if missing:
            raise CannotRun(f"{path} is not a three-day values workbook "
                            f"(no sheet {missing[0]!r})")
        return {name: [list(row) for row in wb[name].iter_rows(values_only=True)]
                for name in SHARED}
    finally:
        wb.close()


@contextlib.contextmanager
def excel():
    """A hidden Excel of this check's own, closed again whatever happens."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as e:
        raise CannotRun(f"pywin32 is not installed ({e})") from e
    pythoncom.CoInitialize()
    try:
        try:
            app = win32.DispatchEx("Excel.Application")
        except Exception as e:  # noqa: BLE001 - no Excel on this PC is "could not run"
            raise CannotRun(
                f"Excel could not be started ({type(e).__name__}: {e})") from e
        try:
            app.Visible = False
            app.DisplayAlerts = False
            app.AskToUpdateLinks = False
            yield app
        finally:
            app.Quit()
    finally:
        pythoncom.CoUninitialize()


def excel_available():
    """True when this PC can run the check (Excel and pywin32 are both here)."""
    try:
        with excel():
            return True
    except CannotRun:
        return False


def _sheet_cells(ws, n_cols=None):
    """Every used cell of one sheet as rows of values, read in bands so a
    statewide grid does not cross the COM boundary in one piece."""
    used = ws.UsedRange
    n_rows = used.Row + used.Rows.Count - 1
    if n_cols is None:
        n_cols = used.Column + used.Columns.Count - 1
    out = []
    for top in range(1, n_rows + 1, CHUNK_ROWS):
        bottom = min(top + CHUNK_ROWS - 1, n_rows)
        block = ws.Range(ws.Cells(top, 1), ws.Cells(bottom, n_cols)).Value
        if not isinstance(block, tuple):
            block = ((block,),)
        out.extend(list(row) for row in block)
    return out


def read_calculated(app, path, grid_cols, log=print, edits=(), before=None):
    """{sheet: rows} of the copy as Excel calculates it.

    `edits` are (sheet, row, column, text) changes made to the open copy before
    it is recalculated — how a check asks what the formulas say once a day's
    value is different — and `before(workbook)` is anything else to do to it
    first (sort a sheet, read a colour). Nothing is saved."""
    started = time.time()
    try:
        wb = app.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0,
                                ReadOnly=True)
    except Exception as e:  # noqa: BLE001 - Excel refusing the file is the finding
        raise CannotRun(f"Excel could not open {Path(path).name} "
                        f"({type(e).__name__}: {e})") from e
    try:
        opened = time.time() - started
        for sheet, row, column, text in edits:
            cell = wb.Worksheets(sheet).Cells(row, column)
            cell.NumberFormat = "@"          # keep the text text ("5" is not 5)
            cell.Value = text
        if before is not None:
            before(wb)
        started = time.time()
        app.CalculateFull()
        log(f"  Excel opened it in {opened:.1f}s and recalculated in "
            f"{time.time() - started:.1f}s")
        names = [wb.Worksheets(i + 1).Name for i in range(wb.Worksheets.Count)]
        missing = [name for name in SHARED if name not in names]
        if missing:
            raise CannotRun(f"{Path(path).name} is not a live-formulas copy "
                            f"(no sheet {missing[0]!r})")
        return {name: _sheet_cells(wb.Worksheets(name),
                                   grid_cols if name == GRID else None)
                for name in SHARED}
    finally:
        wb.Close(SaveChanges=False)


def _is_error(value):
    """Excel hands an error cell over COM as a large negative integer."""
    return (isinstance(value, int) and not isinstance(value, bool)
            and value < -2_000_000_000)


def compare_sheet(name, want, got, width=None):
    """(cells compared, [difference lines]) for one sheet. `width` limits the
    comparison to the first columns (the grid's own, on the grid sheet)."""
    differences, compared = [], 0
    for i, wanted in enumerate(want):
        have = got[i] if i < len(got) else []
        n = width if width is not None else len(wanted)
        for j in range(n):
            a = _norm(wanted[j]) if j < len(wanted) else ""
            b = _norm(have[j]) if j < len(have) else ""
            compared += 1
            if a != b or type(a) is not type(b):
                differences.append(f"{name} row {i + 1} column {j + 1}: the values "
                                   f"workbook has {a!r}, the copy calculates {b!r}")
    for i, have in enumerate(got):
        for j, value in enumerate(have):
            if _is_error(value):
                differences.append(f"{name} row {i + 1} column {j + 1}: the copy "
                                   f"calculates an Excel error ({value})")
    if len(got) < len(want) or (len(got) > len(want) and name not in LONGER_IN_COPY):
        differences.append(f"{name}: the copy has {len(got)} rows, the values "
                           f"workbook {len(want)}")
    return compared, differences


def check_lines(summary):
    """[(label, reading)] of the lines under the Summary's Live formulas heading,
    or None when the sheet has no such heading."""
    for i, row in enumerate(summary):
        if row and isinstance(row[0], str) and row[0].startswith(CHECK_HEADING):
            return [(r[0], r[1] if len(r) > 1 else None) for r in summary[i + 2:]
                    if r and r[0]]
    return None


def grid_width(values):
    """How many columns the values workbook's grid has."""
    grid = values[GRID]
    return len(grid[1]) if len(grid) > 1 else GRID_LEAD


def differences(want, got, log=print, expect_checks_ok=True):
    """Every problem between the values workbook's sheets and the copy's."""
    problems = []
    for name in SHARED:
        width = grid_width(want) if name == GRID else None
        compared, found = compare_sheet(name, want[name], got[name], width)
        extra = len(got[name]) - len(want[name])
        log(f"  {name}: {compared:,} cells compared, {len(found):,} differ"
            + (f" ({extra} more rows in the copy)" if extra > 0 else ""))
        problems.extend(found)
    lines = check_lines(got["Summary"])
    if not lines:
        problems.append("Summary: no check lines under a Live formulas heading")
        return problems
    for label, reading in lines:
        log(f"  check: {reading} — {label}")
        if expect_checks_ok and reading != "OK":
            problems.append(f"Summary check {label!r} reads {reading!r}")
    return problems


def check(values_path, copy_path=None, log=print, app=None):
    """Run the whole check. Returns the list of problems (empty = equal)."""
    values_path = Path(values_path)
    copy_path = Path(copy_path) if copy_path else formulas_sibling(values_path)
    for p in (values_path, copy_path):
        if not p.is_file():
            raise CannotRun(f"no such file: {p}")
    log(f"values workbook: {values_path.name}")
    log(f"live-formulas copy: {copy_path.name} "
        f"({copy_path.stat().st_size / 1e6:.1f} MB)")
    want = read_values(values_path)
    if app is not None:
        got = read_calculated(app, copy_path, grid_width(want), log)
    else:
        with excel() as own:
            got = read_calculated(own, copy_path, grid_width(want), log)
    return differences(want, got, log)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Recalculate a three-day live-formulas copy in Excel and hold "
                    "it against its values workbook.")
    parser.add_argument("values", help="the three-day values workbook")
    parser.add_argument("copy", nargs="?", help="its live-formulas copy "
                        "(default: '<name> (formulas).xlsx' beside it)")
    parser.add_argument("--out", help="also write the report to this file")
    args = parser.parse_args(argv)
    lines = []

    def log(message):
        print(message, flush=True)
        lines.append(message)

    try:
        problems = check(args.values, args.copy, log)
    except CannotRun as e:
        log(f"COULD NOT RUN: {e}")
        code = 2
    else:
        for problem in problems[:60]:
            log(f"  DIFFERENT: {problem}")
        if len(problems) > 60:
            log(f"  … and {len(problems) - 60:,} more")
        log("RESULT: " + ("the copy calculates to the values workbook, cell for cell"
                          if not problems else f"{len(problems):,} differences"))
        code = 1 if problems else 0
    if args.out:
        Path(args.out).write_text("\n".join(lines) + "\n", encoding="utf-8")
    return code


if __name__ == "__main__":
    sys.exit(main())
