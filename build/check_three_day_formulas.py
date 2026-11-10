"""build/check_three_day_formulas.py — the three-day comparison's live-formulas
copy (v0.52.1): the values workbook's sheets with every number an Excel formula.

Drives the SHIPPED path on synthetic data (no corpus files):

  * the copy is written beside the values workbook, from the same computed
    comparison, under the name every matrix gives that copy — and it never
    decides a result: past the size limit, on a writer failure or a cancel the
    values workbook and its counts stand, and a copy left from an earlier build
    is removed rather than left looking current (one Excel still holds open is
    said, not passed over);
  * what the file holds: the seven sheets; each day's values in its block
    exactly as exported; and the rule Excel imposes on a filled-down formula —
    every formula on the grid sheet reads its own sheet only, is stored ONCE per
    column as a shared formula, and reads back from the file as the formula
    that was written;
  * limits: no formula past Excel's 8,192 characters and no text literal past
    255, on a 300-column report and on one whose columns are almost all shown
    only;
  * text Excel would not read back as written (`_xHHHH_`, also where two share
    an underscore; a carriage return; a string that is only line feeds and
    tabs) is stored escaped;
  * IN EXCEL, when this PC has it (`tools/three_day_formulas_check.py`): the
    copy recalculates to the values workbook cell for cell — the hand fixture,
    random data with and without repeated keys, a report wide enough to need
    several state formulas, and values chosen to be awkward (blanks, spaces,
    letter case, numbers, formula-looking text, Excel's error words, text that
    reads like the workbook's own marks, `_xHHHH_`, white space alone); each
    day's block, as EXCEL reads it, is the export's text character for
    character — carriage returns included; every Summary check reads
    OK; after a day's value is CHANGED in Excel the recalculated copy equals a
    fresh comparison of the changed data, while its first check says the
    workbook no longer describes the exports; SORTED from a column header, every
    record keeps its own route, location and values and the copy says it is out
    of order; and the class colours are the values workbook's.

Without Excel the last part is skipped and says so: nothing else can tell what a
formula calculates to.

Run with the build venv:
    build\\.venv\\Scripts\\python.exe build\\check_three_day_formulas.py
"""
import logging
import random
import re
import sys
import zipfile
from pathlib import Path

from _checklib import ROOT, Checker, patch, pinned_site, scripts_path, temp_dir

scripts_path()
sys.path.insert(0, str(ROOT / "tools"))

from openpyxl import load_workbook  # noqa: E402

import check_three_day as base  # noqa: E402  (its fixtures and random data)
import compare_core as core  # noqa: E402
import consolidation_meta  # noqa: E402
import three_day as td  # noqa: E402
import three_day_formulas as live  # noqa: E402
import three_day_formulas_check as oracle  # noqa: E402  (tools/: Excel's own answer)
import three_day_workbook as book  # noqa: E402
from errors import RunCancelled  # noqa: E402
from events import Events  # noqa: E402

c = Checker()
DAYS = base.DAYS
META = {"inputs": ("one", "two", "three")}
# Values chosen to trip a formula: what Excel might read as a formula, a number,
# an error, an escape or one of the workbook's own marks — and plain awkwardness.
AWKWARD = ["", "0", "1", " 2", "2 ", "a", "A", "a  b", "a b", 5, 5.0, 5.5, -5, None,
           "=1+1", "+x", "-5", "@x", "#N/A", "#DIV/0!", td.BLANK_MARK,
           td.ABSENT_MARK, "a → b", "x_x000d_y", "x_x000D_y", "_x005F_", "a\nb",
           "a\tb", "TRUE", True, "true", 0.1 + 0.2, "1E5", "001", " ", "é",
           "a\u00a0b", "'q", '"q"', "a,b", "<&>", "yes", "same", "OK",
           "\n", "\t", " \n", "\n\n", "a\n", "\ta", "_x0041_x0042_", "_x0041\n"]
# Excel reads a bare carriage return as a line feed, and so does the VALUES
# workbook, whose file therefore cannot be the yardstick for one. These go into
# a comparison of their own, held against the exports and the app's counts.
RETURNS = ["a", "a\r", "a\n", "a\r\n", "\r", "_x000D\r", "\r_x0041_", "b\rc", ""]


def _log():
    lines = []
    return lines, Events(on_log=lines.append)


def _run(tmp, name, sc, rows, **kwargs):
    out = tmp / f"{name}.xlsx"
    lines, events = _log()
    res = td.run(sc, rows, DAYS, out, events=events, meta=META, **kwargs)
    return out, td.formulas_sibling(out), res, lines


# --------------------------------------------------------------------------- #
# 1. the copy is written beside the workbook — and never decides the result
# --------------------------------------------------------------------------- #
def test_written(tmp):
    print("the copy is written beside the values workbook:")
    sc, rows = base._schema(), base._fixture_days()
    out, twin, res, lines = _run(tmp, "demo three days", sc, rows, also_formulas=True)
    c.check("asked for, it is written beside the workbook under the matrices' name",
            res.status == "ok" and twin.is_file()
            and twin.name == "demo three days (formulas).xlsx", res.message)
    c.check("the result says so, and the counts are still the values workbook's",
            any(twin.name in line for line in res.summary_lines)
            and td.read_counts(out) is not None, res.summary_lines)
    _out, _twin, res, lines = _run(tmp, "demo three days", sc, rows)
    c.check("a build without the option removes the earlier copy and says so",
            res.status == "ok" and not twin.exists()
            and any("Removed the earlier live-formulas copy" in line for line in lines),
            lines[-3:])
    c.check("...and a build that never had one says nothing about it",
            not any("live-formulas" in line
                    for line in _run(tmp, "plain", sc, rows)[3]))


def test_never_decides(tmp):
    print("the copy never decides the result:")
    sc, rows = base._schema(), base._fixture_days()
    out, twin, res, _lines = _run(tmp, "d", sc, rows, also_formulas=True)
    with patch(td, "FORMULAS_MAX_CELLS", 10):
        _o, _t, res, lines = _run(tmp, "d", sc, rows, also_formulas=True)
    c.check("past the size limit the copy is skipped, said, and the earlier one "
            "removed — the comparison itself is unchanged",
            res.status == "ok" and not twin.exists() and td.read_counts(out) is not None
            and any("Skipping the live-formulas copy" in line for line in lines), lines)

    def boom(*_a, **_k):
        raise RuntimeError("no disk")

    def cancelled(*_a, **_k):
        raise RunCancelled()

    for label, writer, said in (("a writer failure", boom, "was not written"),
                                ("a cancel while it is written", cancelled,
                                 "was cancelled")):
        with patch(live, "write_workbook", writer):
            _o, _t, res, lines = _run(tmp, "d", sc, rows, also_formulas=True)
        c.check(f"{label} leaves the values workbook and its counts standing",
                res.status == "ok" and out.is_file() and not twin.exists()
                and td.read_counts(out) is not None
                and any(said in line for line in lines), lines[-3:])
    _run(tmp, "d", sc, rows, also_formulas=True)
    with patch(consolidation_meta, "write_outcome", lambda *_a, **_k: False):
        _o, _t, res, lines = _run(tmp, "d", sc, rows, also_formulas=True)
    c.check("when the counts cannot be published the build fails — and the copy "
            "from the build before does not stay beside the new workbook",
            res.status == "error" and not twin.exists(), (res.status, lines[-2:]))
    _run(tmp, "d", sc, rows, also_formulas=True)
    with open(twin, "rb"):                    # Excel holding the copy open
        _o, _t, res, lines = _run(tmp, "d", sc, rows)
    c.check("a copy that cannot be removed is said to be out of date",
            res.status == "ok" and twin.exists()
            and any("could not be removed" in line for line in lines), lines[-3:])


# --------------------------------------------------------------------------- #
# 2. what the file holds
# --------------------------------------------------------------------------- #
def _grid_xml(path):
    """The grid sheet's XML (the fifth sheet written)."""
    with zipfile.ZipFile(path) as z:
        return z.read("xl/worksheets/sheet5.xml").decode("utf-8")


def test_file(tmp):
    print("what the copy holds:")
    sc, rows = base._schema(), base._fixture_days()
    _out, twin, _res, _lines = _run(tmp, "f", sc, rows, also_formulas=True)
    compared = td.compute(sc, rows)
    lay = live.Layout(sc, len(compared.records))
    wb = load_workbook(twin)
    c.check("its sheets: the values workbook's six, then the counts as built",
            tuple(wb.sheetnames) == live.SHEETS == book.SHEETS + (live.AS_BUILT,),
            wb.sheetnames)
    grid = wb[book.GRID_SHEET]
    bad = []
    for k, rec in enumerate(compared.records):
        row = lay.first + k
        for day, idx in enumerate(rec):
            flag = grid.cell(row=row, column=lay.flag[day]).value
            if flag != (live.ON_DAY if idx is not None else td.ABSENT_MARK):
                bad.append((k, day, flag))
            for f in range(lay.n):
                held = grid.cell(row=row, column=lay.day0[day] + f).value
                raw = None if idx is None else rows[day][idx][f + 1]
                want = None if raw in (None, "") else raw
                if held != want:
                    bad.append((k, day, f, held, want))
    c.check("each day's block holds that day's values exactly as exported, with "
            "the days a record is on", not bad, bad[:5])
    k = next(i for i, rec in enumerate(compared.records) if rec[0] == 8)
    held = grid.cell(row=lay.first + k, column=lay.day0[0] + 1)
    c.check("formula-looking text in a day block is text, not a formula",
            held.value == "=1+1" and held.data_type == "s",
            (held.value, held.data_type))
    formulas = live.grid_formulas(lay, DAYS)
    read_back = [(col, row) for col in formulas for row in (lay.first, lay.first + 4)
                 if grid.cell(row=row, column=col).value
                 != live.grid_formulas(_at_row(lay, row), DAYS)[col]]
    c.check("every formula column reads back, on its first row and further down, "
            "as the formula written for that row", not read_back, read_back[:5])
    xml = _grid_xml(twin)
    masters = re.findall(r'<f t="shared" ref="([A-Z]+)\d+:\1\d+" si="\d+">([^<]*)</f>',
                         xml)
    plain = re.findall(r"<f>([^<]*)</f>", xml)
    c.check("each formula column is stored once, as one shared formula filled down",
            len(masters) == len(formulas) and not plain
            and xml.count('<f t="shared" si=') == len(formulas)
            * (len(compared.records) - 1), (len(masters), len(formulas), len(plain)))
    c.check("no formula on the grid sheet reads another sheet (Excel cannot share "
            "one that does)", not [text for _col, text in masters if "!" in text])
    summary = wb["Summary"]
    built = wb[live.AS_BUILT]
    numbers = [(cell.coordinate, cell.value) for row in summary.iter_rows(min_col=2)
               for cell in row if isinstance(cell.value, (int, float))]
    c.check("the Summary holds no typed-in number: every count is a formula, and "
            "the counts the app computed are on the As built sheet",
            not numbers and built["B8"].value == compared.counts.records["all"],
            numbers[:5])
    wb.close()


def _at_row(lay, row):
    """`lay` with its first record on `row` — the formulas as written there."""
    moved = live.Layout(lay.sc, lay.last - lay.first + 1)
    moved.first = row
    return moved


def _literals(formula):
    return [m.group(0)[1:-1] for m in re.finditer(r'"(?:[^"]|"")*"', formula)]


def test_limits():
    print("formula limits:")
    wide = ["PM"] + [f"C{i:03d}" for i in range(298)] + ["X_EXTRACT_DATE"]
    shown = [f"S{i:03d}" for i in range(298)]
    cases = (
        ("a 300-column report", base._schema(wide)),
        ("a report whose columns are almost all shown only",
         core.CompareSchema(report_name="Demo", header=["PM", "A"] + shown,
                            id_noun="segment", id_noun_plural="segments",
                            key_field=0, context_fields=tuple(shown))))
    for label, sc in cases:
        lay = live.Layout(sc, core.XL_MAX_ROWS - book.GRID_FIRST_ROW)
        formulas = list(live.grid_formulas(lay, DAYS).values())
        covered = [f for run in lay.chunks for f in run]
        c.check(f"{label}: the state formulas cover every column once, in order",
                covered == list(range(len(sc.header))) and len(lay.chunks) > 1,
                (len(lay.chunks), len(covered)))
        c.check(f"{label}: no formula past 8,192 characters, no text literal past "
                "255", max(map(len, formulas)) <= 8192
                and max(len(t) for f in formulas for t in _literals(f)) <= 255,
                (max(map(len, formulas)),
                 max(len(t) for f in formulas for t in _literals(f))))
        c.check(f"{label}: the sheet fits Excel",
                core.excel_limit_error(lay.last, lay.last_col) is None, lay.last_col)


def test_text():
    print("text Excel would not read back as written:")
    plain = "NB OFF TO A ST"
    c.check("ordinary text is stored as it is", live.excel_text(plain) is plain)
    c.check("_xHHHH_ is escaped so Excel does not decode it, and what only looks "
            "like one is left alone",
            live.excel_text("a_x000d_b") == "a_x005F_x000d_b"
            and live.excel_text("_x005F_") == "_x005F_x005F_"
            and live.excel_text("a_xZZZZ_b_x12_") == "a_xZZZZ_b_x12_")
    c.check("two that share an underscore are both escaped",
            live.excel_text("_x0041_x0042_") == "_x005F_x0041_x005F_x0042_")
    c.check("a carriage return, a line feed and a tab are written as their "
            "escapes — and an underscore that would join one into a token is "
            "escaped too",
            live.excel_text("a\rb\r\n") == "a_x000D_b_x000D__x000A_"
            and live.excel_text("a\nb\tc") == "a_x000A_b_x0009_c"
            and live.excel_text("_x000D\r") == "_x005F_x000D_x000D_")


# --------------------------------------------------------------------------- #
# 3. in Excel: the copy calculates to the values workbook
# --------------------------------------------------------------------------- #
def _pair(tmp, name, sc, rows):
    """(values workbook, its copy) for one comparison."""
    out, twin, res, _lines = _run(tmp, name, sc, rows, also_formulas=True)
    if res.status != "ok" or not twin.is_file():
        raise AssertionError(f"{name}: {res.message}")
    return out, twin


def _quiet(_message):
    pass


def _cases(tmp):
    """(label, schema, rows) — the fixture, random data plain and awkward, with
    and without repeated keys, and a report wide enough for several state
    formulas."""
    yield "the hand fixture", base._schema(), base._fixture_days()
    random_sc = core.CompareSchema(
        report_name="Demo", header=list(base.P_HEADER), id_noun="row",
        id_noun_plural="rows", key_field=0, context_fields=("X_EXTRACT_DATE",))
    for seed in range(6):
        for dup in (False, True):
            rows = base._random_days(random.Random(seed * 2 + dup), dup)
            if all(rows):
                yield f"random data {seed}{' with repeated keys' if dup else ''}", \
                    random_sc, rows
    with patch(base, "_VALUES", AWKWARD):
        awkward = [(seed, dup, base._random_days(random.Random(100 + seed * 2 + dup),
                                                 dup, n_routes=2, n_keys=12))
                   for seed in range(4) for dup in (False, True)]
    for seed, dup, rows in awkward:
        if all(rows):
            yield f"awkward values {seed}{' with repeated keys' if dup else ''}", \
                random_sc, rows
    wide = ["PM"] + [f"C{i:02d}" for i in range(60)] + ["X_EXTRACT_DATE"]
    rng = random.Random(7)

    def wide_row(route, pm, day, like=None):
        cells = [rng.choice(["a", "b", "", 3]) if like is None or rng.random() < 0.1
                 else like[2 + i] for i in range(60)]
        return [route, pm] + cells + [f"2026-0{day}"]

    day1 = [wide_row("001", f"{k}.0", 1) for k in range(12)]
    day2 = [wide_row(r[0], r[1], 2, r) for r in day1[:10]] + [wide_row("001", "20.0", 2)]
    day3 = [wide_row(r[0], r[1], 3, r) for r in day2[1:]] + [wide_row("002", "1.0", 3)]
    yield "a 62-column report", base._schema(wide), (day1, day2, day3)


def _blocks(sc, rows, grid):
    """Every difference between the day blocks AS EXCEL HOLDS THEM and the
    exports: the days each record is on, and each value character for
    character (a number as the text it is compared as; a value that is only
    spaces arrives empty, which is what TRIM makes of it either way)."""
    compared = td.compute(sc, rows)
    lay = live.Layout(sc, len(compared.records))
    bad = []
    for k, rec in enumerate(compared.records):
        held = grid[lay.first - 1 + k]
        for day, idx in enumerate(rec):
            flag = held[lay.flag[day] - 1]
            if flag != (live.ON_DAY if idx is not None else td.ABSENT_MARK):
                bad.append(f"record {k + 1} day {day + 1}: flag {flag!r}")
            for f in range(lay.n):
                raw = None if idx is None else rows[day][idx][f + 1]
                want = (compared.trim[day][idx][f + 1]
                        if raw is not None and not isinstance(raw, str) else raw or "")
                want = want if want.strip(" ") else ""
                got = held[lay.day0[day] - 1 + f]
                if (got or "") != want:
                    bad.append(f"record {k + 1} day {day + 1} column {f + 1}: the "
                               f"export has {want!r}, Excel holds {got!r}")
    return bad


def _in_excel(app, sc, out, twin, rows):
    """(differences from the values workbook, differences in the day blocks,
    the Summary's check lines) for one copy, recalculated in Excel."""
    want = oracle.read_values(out)
    lay = live.Layout(sc, 1)
    got = oracle.read_calculated(app, twin, lay.last_col, _quiet)
    return (oracle.differences(want, got, _quiet), _blocks(sc, rows, got[book.GRID_SHEET]),
            oracle.check_lines(got["Summary"]) or [])


def test_excel(tmp):
    print("in Excel — the copy calculates to the values workbook:")
    if not oracle.excel_available():
        print("  skipped: Excel is not on this PC, so no formula was calculated")
        return
    seen_direct = chunks = 0
    with oracle.excel() as app:
        for n, (label, sc, rows) in enumerate(_cases(tmp)):
            out, twin = _pair(tmp, f"case {n}", sc, rows)
            counts = td.read_counts(out)
            seen_direct += bool(counts and counts.first_last_differs())
            chunks = max(chunks, len(live.Layout(sc, 1).chunks))
            try:
                problems, blocks, _lines = _in_excel(app, sc, out, twin, rows)
            except oracle.CannotRun as e:
                problems, blocks = [str(e)], []
            c.check(f"{label}: every cell equal, every check OK, each day's block "
                    "the export's own text", not problems and not blocks,
                    "\n      ".join((problems + blocks)[:6]))
        _returns(tmp, app)
        c.check("the cases reached a first-vs-last pair that pairs differently "
                "directly (its row is the app's own numbers) and a report needing "
                "several state formulas", seen_direct > 0 and chunks > 1,
                (seen_direct, chunks))
        _edited(tmp, app)
        _sorted(tmp, app)
        _coloured(tmp, app)


def _returns(tmp, app):
    """Carriage returns: the app tells `a\\r` from `a\\n`; a stored string would
    not. Held against the exports and against the app's own counts (the As built
    sheet), since the values workbook's FILE reads a return as a line feed."""
    with patch(base, "_VALUES", RETURNS):
        rows = base._random_days(random.Random(11), False, n_routes=2, n_keys=14)
    sc = core.CompareSchema(
        report_name="Demo", header=list(base.P_HEADER), id_noun="row",
        id_noun_plural="rows", key_field=0, context_fields=("X_EXTRACT_DATE",))
    out, twin = _pair(tmp, "returns", sc, rows)
    _problems, blocks, lines = _in_excel(app, sc, out, twin, rows)
    counts = td.read_counts(out)
    c.check("carriage returns: each day's block is the export's own text in Excel, "
            "and every count equals the app's (which tells a return from a line "
            "feed)", not blocks and lines and all(v == "OK" for _k, v in lines)
            and counts.changed_cells() > 0,
            "\n      ".join(blocks[:6]) or lines)


def _sorted(tmp, app):
    """Sort the grid sheet from a column header, as a reader would: every record
    must keep its own route, location and values, and the copy must say that the
    sheets reading it by position no longer line up."""
    sc, rows = base._schema(), base._fixture_days()
    out, twin = _pair(tmp, "sorted", sc, rows)
    want = oracle.read_values(out)
    width = oracle.grid_width(want)

    def sort(wb):
        ws = wb.Worksheets(book.GRID_SHEET)
        by = ws.AutoFilter.Sort
        by.SortFields.Clear()
        by.SortFields.Add(Key=ws.Range("D2"), SortOn=0, Order=2)   # Changed cells, most first
        by.Header = 1
        by.Apply()

    got = oracle.read_calculated(app, twin, width, _quiet, before=sort)
    was = [tuple(map(oracle._norm, r[:width])) for r in want[book.GRID_SHEET][2:]]
    now = [tuple(map(oracle._norm, r[:width])) for r in got[book.GRID_SHEET][2:]]
    c.check("sorted from a header, the grid's rows are the same records in another "
            "order — each still with its own route, location and values",
            now != was and sorted(map(repr, now)) == sorted(map(repr, was)),
            (now[:2], was[:2]))
    checks = dict(oracle.check_lines(got["Summary"]) or ())
    order = next((v for k, v in checks.items() if "order it was built in" in k), "")
    c.check("...and the copy says the sheet is out of order, while its counts "
            "still match the build", str(order).startswith("SORTED")
            and next(iter(checks.values()), "") == "OK", checks)


def _bgr(rgb):
    """An RRGGBB fill as Excel's colour number."""
    return int(rgb[4:6] + rgb[2:4] + rgb[0:2], 16)


def _coloured(tmp, app):
    """The class colours come from conditional formatting on the state letters."""
    sc, rows = base._schema(), base._fixture_days()
    _out, twin = _pair(tmp, "coloured", sc, rows)
    compared = td.compute(sc, rows)
    lay = live.Layout(sc, len(compared.records))
    seen = {}

    def read(wb):
        ws = wb.Worksheets(book.GRID_SHEET)
        app.CalculateFull()
        for k, state in enumerate(compared.states):
            row = lay.first + k
            for f, letter in enumerate(state):
                if letter in td.DIFF_CODES:
                    seen.setdefault(letter, int(
                        ws.Cells(row, lay.field0 + f).DisplayFormat.Interior.Color))
            tone = book._PRESENCE_TONE.get(td.presence_of(compared.records[k]))
            if tone:
                seen.setdefault(tone, int(
                    ws.Cells(row, 3).DisplayFormat.Interior.Color))
        seen["plain"] = int(ws.Cells(lay.first, lay.field0).DisplayFormat.Interior.Color)

    oracle.read_calculated(app, twin, book.GRID_LEAD + lay.n, _quiet, before=read)
    want = {letter: _bgr(book._FILLS[td.CLASS_OF[letter]])
            for letter in td.CHANGED_CODES}
    want["D"] = _bgr(book._FILLS["two"])
    want.update({tone: _bgr(book._FILLS[tone]) for tone in ("added", "removed", "blip")})
    want["plain"] = 0xFFFFFF
    c.check("each changed cell is filled in its class's colour, each record that "
            "is not on all three days in its own, and an unchanged cell in none",
            seen == want, (seen, want))


def _edited(tmp, app):
    """Change one day's value IN EXCEL: the recalculated copy must equal a fresh
    comparison of the changed data, and must say it no longer matches the build."""
    sc, rows = base._schema(), base._fixture_days()
    out, twin = _pair(tmp, "before the edit", sc, rows)
    compared = td.compute(sc, rows)
    lay = live.Layout(sc, len(compared.records))
    k = next(i for i, rec in enumerate(compared.records)
             if rec == (0, 0, 0))                       # route 001, PM 0.0
    changed = [list(map(list, day)) for day in rows]
    changed[2][0][2] = "5"                              # day 3, column A
    fresh, _twin = _pair(tmp, "after the edit", sc, tuple(changed))
    want = oracle.read_values(fresh)
    got = oracle.read_calculated(
        app, twin, oracle.grid_width(want), _quiet,
        edits=[(book.GRID_SHEET, lay.first + k, lay.day0[2] + 1, "5")])
    live_sheets = ("Routes", book.GRID_SHEET)
    problems = [p for name in live_sheets
                for p in oracle.compare_sheet(
                    name, want[name], got[name],
                    oracle.grid_width(want) if name == book.GRID_SHEET else None)[1]]
    summary = oracle.compare_sheet("Summary", want["Summary"][:len(want["Summary"])],
                                   got["Summary"])[1]
    c.check("a value changed in Excel: the grid and the Routes sheet recalculate "
            "to a fresh comparison of the changed data", not problems,
            "\n      ".join(problems[:6]))
    c.check("...and so does every count on the Summary", not summary,
            "\n      ".join(summary[:6]))
    checks = dict(oracle.check_lines(got["Summary"]) or ())
    first = next(iter(checks.values()), "")
    c.check("...while the copy says it no longer describes the exports, and that "
            "its Changes list is out of date",
            str(first).startswith("DIFFERS")
            and str(checks.get("The Changes sheet lists every changed cell", ""))
            .startswith("OUT OF DATE"), checks)


def main():
    # A writer failure and a copy held open are provoked on purpose; what the
    # user is told about each is asserted, so the log's own lines are noise here.
    logging.disable(logging.CRITICAL)
    with pinned_site("dev"), temp_dir("tsmis_tdf_") as tmp:
        tmp = Path(tmp)
        for name, test in (("w", test_written), ("n", test_never_decides),
                           ("f", test_file), ("x", test_excel)):
            sub = tmp / name
            sub.mkdir()
            test(sub)
        test_limits()
        test_text()
    return c.summary()


if __name__ == "__main__":
    raise SystemExit(main())
