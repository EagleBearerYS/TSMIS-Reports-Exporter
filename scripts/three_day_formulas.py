"""The three-day comparison's live-formulas copy (v0.52.1) — the values
workbook's sheets with every number an Excel formula.

`three_day_workbook` writes what the app computed. This writes the same sheets
so that Excel computes them: the three days' exported values sit on the grid
sheet beside the grid, one block of columns per day, and every grid cell, every
count on the Summary and every figure on the Routes sheet is a formula over
them. A reader can click any number and follow it back to what each day's file
said.

The grid sheet, left to right:

  * the values grid's own columns   Route, Location and four tallies, then every
                                    column of the report — the cells the values
                                    workbook shows, as formulas
  * one block per day               "On <day>" (yes / —), then that day's values
                                    as exported
  * the working columns             the record's number, the days it is on as a
                                    code (4 + 2 + 1), one state letter per
                                    column, and the tallies read off the letters

Every formula on that sheet reads its OWN ROW ONLY. That is what lets a formula
be written once per column and filled down (OOXML's shared formula — the form
Excel itself saves a filled-down column in; Excel has no such form for a
formula that reaches another sheet, which is why the days are not sheets of
their own). A statewide report's millions of formula cells stay a small file,
and the grid can be filtered or sorted without a formula losing its record.

A cell is equal exactly when `EXACT(TRIM(a),TRIM(b))` — compare_core's own rule,
as its live-formulas workbooks write it. WHICH rows are one record is decided
when the workbook is built: the pairing is the engine's, not a formula's. The
foot of the Summary holds the live numbers against the ones the app computed,
kept on the `As built` sheet, so the copy says for itself whether it still
describes the exports.

Nothing here decides anything either: the records, their order and every
expected number arrive in the `Compared`, and `build/check_three_day_formulas.py`
recalculates the copy in Excel and holds every cell against the values workbook.
"""
from __future__ import annotations

import functools
import re

import artifact_store
import compare_core as core
import three_day as td
import three_day_workbook as book

AS_BUILT = "As built"
SHEETS = book.SHEETS + (AS_BUILT,)
ON_DAY = "yes"                    # a day block's flag: the record is on that day
NOT_ON_DAY = td.ABSENT_MARK
NO_LONGER = "same"                # a listed change whose days now agree
OK = "OK"
# The days a record is on as one number: 4 for the first day, 2 for the middle
# day, 1 for the newest.
DAYS_CODE = {"all": 7, "d12": 6, "d13": 5, "d1": 4, "d23": 3, "d2": 2, "d3": 1}
_BY_CODE = sorted(DAYS_CODE, key=DAYS_CODE.get)          # the presence of code 1..7
_ON_ALL = DAYS_CODE["all"]
_TWO_DIFFERS = "D"                # three_day's letter: differs, record on two days
_FORMULA_TARGET = 6_000           # a state formula's length; Excel allows 8,192
_MAX_RUN = 200                    # ...and its columns: a text literal holds 255
_PROGRESS_EVERY = 10_000          # records between progress lines on the grid sheet
# What Excel does not read back from a stored string as it was written (each
# measured on Excel 16): `_xHHHH_` is decoded as the character with that code; a
# carriage return arrives as a line feed; and a string that is nothing but line
# feeds, tabs and spaces arrives empty. Written as escapes, all of it arrives
# whole. An underscore is escaped where it would START a token in what is
# written — which includes the underscore a following escape begins with.
_ESCAPES = {"\r": "_x000D_", "\n": "_x000A_", "\t": "_x0009_"}
_NEEDS_ESCAPE = re.compile(r"_x|[\r\n\t]")
_ESCAPED = re.compile(r"_(?=x[0-9A-Fa-f]{4}[_\r\n\t])|[\r\n\t]")


def _q(text):
    """`text` as a string literal inside a formula."""
    return core._formula_text(text)


def excel_text(text):
    """`text` as a cell value Excel reads back character for character, so a
    formula sees exactly the text the app compared (see `_ESCAPES`)."""
    if not _NEEDS_ESCAPE.search(text):
        return text
    return _ESCAPED.sub(lambda m: _ESCAPES.get(m.group(), "_x005F_"), text)


@functools.lru_cache(maxsize=1)
def _shared_type():
    from openpyxl.worksheet.formula import ArrayFormula

    class SharedFormula(ArrayFormula):
        """One formula filled down a column, stored once. openpyxl writes the
        attributes an `ArrayFormula` yields onto the cell's <f> element, so
        yielding t="shared" makes it the shared formula Excel itself stores:
        the first cell carries the text and the range, every other cell only
        the group number."""
        t = "shared"

        def __init__(self, group, ref=None, text=None):
            super().__init__(ref, text)
            self.group = group

        def __iter__(self):
            yield "t", self.t
            if self.ref:
                yield "ref", self.ref
            yield "si", str(self.group)

    return SharedFormula


class _FillDown:
    """Hands out the two cells of each formula a sheet fills down a column."""

    def __init__(self, first, last):
        self.first, self.last, self._groups = first, last, 0
        self._type = _shared_type()

    def column(self, letter, formula):
        """(the first row's cell, every other row's cell) of `formula`, which is
        written for the first row."""
        core._checked_formula(formula, f"three-day column {letter}")
        if self.last <= self.first:
            return formula, None
        group = self._groups
        self._groups += 1
        ref = f"{letter}{self.first}:{letter}{self.last}"
        return self._type(group, ref, formula), self._type(group)


# --------------------------------------------------------------------------- #
# the grid sheet's geometry and formulas
# --------------------------------------------------------------------------- #
class Layout:
    """Where everything sits on the live-formulas grid sheet (1-based columns)."""

    def __init__(self, sc, n_records):
        from openpyxl.utils import get_column_letter

        self.sc = sc
        self.letter = get_column_letter
        self.n = len(sc.header)
        self.first = book.GRID_FIRST_ROW
        self.last = self.first + n_records - 1
        self.field0 = book.GRID_LEAD + 1                  # the grid's column f
        col = book.GRID_LEAD + self.n
        self.flag, self.day0 = [], []                     # per day: flag, column f
        for _day in range(3):
            self.flag.append(col + 1)
            self.day0.append(col + 2)
            col += self.n + 1
        self.working0 = self.c_number = col + 1
        self.c_days = col + 2
        col += 2
        self.base = td.one_day_state(sc)
        self.chunks = _plan_state(self)
        self.c_chunks = [col + 1 + i for i in range(len(self.chunks))]
        col += len(self.chunks)
        self.c_state = col + 1
        codes = td.CODE[td.SAME] + td.DIFF_CODES
        self.c_count = {code: col + 2 + i for i, code in enumerate(codes)}
        col += 1 + len(codes)
        self.c_pair = {pair: col + 1 + i for i, pair in enumerate(td.PAIRS)}
        self.last_col = col + len(td.PAIRS)

    def at(self, col, row):
        """One cell of the grid sheet, as a formula on that sheet reads it."""
        return f"${self.letter(col)}{row}"

    def cell(self, col, row):
        """One cell of the grid sheet, as another sheet reads it."""
        return f"{core._sref(book.GRID_SHEET)}!${self.letter(col)}${row}"

    def span(self, col, first, last):
        """Rows `first`..`last` of one column, as another sheet reads them."""
        letter = self.letter(col)
        return f"{core._sref(book.GRID_SHEET)}!${letter}${first}:${letter}${last}"

    def whole(self, col):
        """Every record's cell in one column, as another sheet reads it."""
        return self.span(col, self.first, self.last)


def state_formula(lay, row, positions):
    """One record's state letters for the columns `positions` — the letters
    `three_day.compute` gives them, decided by the days the record is on."""
    days = lay.at(lay.c_days, row)

    def value(day, f):
        return f"TRIM({lay.at(lay.day0[day] + f, row)})"

    def three(f):
        a, b, c = (value(day, f) for day in range(3))
        same, newest, middle, back, both = (
            _q(td.CODE[k]) for k in (td.SAME, td.NEWEST, td.MIDDLE, td.BACK, td.BOTH))
        return (f"IF(EXACT({a},{b}),IF(EXACT({b},{c}),{same},{newest}),"
                f"IF(EXACT({b},{c}),{middle},IF(EXACT({a},{c}),{back},{both})))")

    def two(x, y):
        return lambda f: (f'IF(EXACT({value(x, f)},{value(y, f)}),"E",'
                          f"{_q(_TWO_DIFFERS)})")

    def letters(per_column):
        parts, fixed = [], ""
        for f in positions:
            if lay.base[f] == ".":
                if fixed:
                    parts.append(_q(fixed))
                    fixed = ""
                parts.append(per_column(f))
            else:
                fixed += lay.base[f]
        if fixed:
            parts.append(_q(fixed))
        return "&".join(parts)

    one_day = _q("".join(lay.base[f] for f in positions))
    return (f"=IF({days}={_ON_ALL},{letters(three)},"
            f'IF({days}={DAYS_CODE["d12"]},{letters(two(0, 1))},'
            f'IF({days}={DAYS_CODE["d23"]},{letters(two(1, 2))},'
            f'IF({days}={DAYS_CODE["d13"]},{letters(two(0, 2))},{one_day}))))')


def _plan_state(lay):
    """The report's columns in runs, each short enough for one state formula."""
    runs, current = [], []
    for f in range(lay.n):
        longer = current + [f]
        if current and (len(longer) > _MAX_RUN or len(state_formula(
                lay, core.XL_MAX_ROWS, longer)) > _FORMULA_TARGET):
            runs.append(tuple(current))
            current = [f]
        else:
            current = longer
    if current:
        runs.append(tuple(current))
    for run in runs:
        core._checked_formula(state_formula(lay, core.XL_MAX_ROWS, run),
                              "three-day state")
    return runs


def display_formula(lay, row, f):
    """What the grid shows for column `f`: the three days in order where the
    column's state letter is a change, else the newest day the record is on."""
    days, state = lay.at(lay.c_days, row), lay.at(lay.c_state, row)
    a, b, c = (f"TRIM({lay.at(lay.day0[day] + f, row)})" for day in range(3))

    def shown(value, on_that_day):
        return (f'IF({on_that_day},IF({value}="",{_q(td.BLANK_MARK)},{value}),'
                f"{_q(td.ABSENT_MARK)})")

    arrow = _q(td.ARROW)
    history = "&".join((shown(a, f"{days}>=4"), arrow,
                        shown(b, f"OR({days}=2,{days}=3,{days}>=6)"), arrow,
                        shown(c, f"MOD({days},2)=1")))
    newest = f"IF(MOD({days},2)=1,{c},IF({days}=4,{a},{b}))"
    return (f"=IF(ISNUMBER(FIND(MID({state},{f + 1},1),{_q(td.DIFF_CODES)})),"
            f"{history},{newest})")


def _lead_formulas(lay, row, days):
    """The grid's Presence and its three tallies (columns C to F)."""
    code = lay.at(lay.c_days, row)
    labels = ",".join(_q(td.presence_short(p, days)) for p in _BY_CODE)
    changed = "+".join(lay.at(lay.c_count[c], row) for c in td.DIFF_CODES)
    out = [f"=CHOOSE({code},{labels})", f"={changed}"]
    for pair in ("1-2", "2-3"):
        two = DAYS_CODE[td._PAIR_TWO[pair]]
        out.append(f"=IF(OR({code}={_ON_ALL},{code}={two}),"
                   f'{lay.at(lay.c_pair[pair], row)},"")')
    return out


def _working_formulas(lay, row):
    """{column: formula} of the working columns: the days code, the state
    letters, how many of each letter, and the cells that differ per pair."""
    out = {lay.c_days: "=" + "+".join(
        f"IF({lay.at(flag, row)}={_q(ON_DAY)},{weight},0)"
        for flag, weight in zip(lay.flag, (4, 2, 1)))}
    for col, positions in zip(lay.c_chunks, lay.chunks):
        out[col] = state_formula(lay, row, positions)
    out[lay.c_state] = "=" + "&".join(lay.at(col, row) for col in lay.c_chunks)
    state, code = lay.at(lay.c_state, row), lay.at(lay.c_days, row)
    for letter, col in lay.c_count.items():
        out[col] = f'=LEN({state})-LEN(SUBSTITUTE({state},{_q(letter)},""))'
    two_day = lay.at(lay.c_count[_TWO_DIFFERS], row)
    for pair, col in lay.c_pair.items():
        classes = "+".join(lay.at(lay.c_count[td.CODE[c]], row)
                           for c in td._PAIR_DIFFERS[pair])
        out[col] = (f"=IF({code}={_ON_ALL},{classes},"
                    f"IF({code}={DAYS_CODE[td._PAIR_TWO[pair]]},{two_day},0))")
    return out


def grid_formulas(lay, days):
    """{column: the first record's formula} for every formula column of the
    grid sheet."""
    row = lay.first
    out = dict(zip(range(3, book.GRID_LEAD + 1), _lead_formulas(lay, row, days)))
    for f in range(lay.n):
        out[lay.field0 + f] = display_formula(lay, row, f)
    out.update(_working_formulas(lay, row))
    return out


# --------------------------------------------------------------------------- #
# the grid sheet
# --------------------------------------------------------------------------- #
def working_labels(lay, days):
    d1, d2, d3 = days
    return (["#", "Days code"]
            + [f"State {i + 1}" for i in range(len(lay.chunks))]
            + ["State", "Same", f"Changed {d3}", f"Changed {d2}",
               f"Changed {d2}, back {d3}", f"Changed {d2} and {d3}",
               "Differs (two days)", f"Differs {d1} v {d2}", f"Differs {d2} v {d3}",
               f"Differs {d1} v {d3}"])


def _working_note(sc, days):
    d1, d2, d3 = days
    c = td.CODE
    return (f"Working columns. Days code = 4 if on {d1} + 2 if on {d2} + 1 if on "
            f"{d3}. State = one letter per column of the report: {c[td.SAME]} same, "
            f"{c[td.NEWEST]} changed {d3}, {c[td.MIDDLE]} changed {d2}, {c[td.BACK]} "
            f"changed {d2} and back, {c[td.BOTH]} changed both times; E same and "
            f"{_TWO_DIFFERS} differs for a {sc.id_noun} on two days; . nothing to "
            "compare, - shown only, k the key.")


def _grid_columns(ws, sc, lay):
    """Widths, and an outline over everything right of the grid so the day
    blocks and the working columns fold away with one click."""
    fields = book.field_widths(sc)
    book._widths(ws, book.GRID_LEAD_WIDTHS + fields + ([9] + fields) * 3
                 + [8, 10] + [14] * len(lay.chunks) + [18] + [10] * 6 + [13] * 3)
    for col in range(lay.flag[0], lay.last_col + 1):
        ws.column_dimensions[lay.letter(col)].outlineLevel = 1


def _grid_colours(ws, lay):
    """The class colours, read off the state letters by conditional formatting."""
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Font, PatternFill

    def rule(formula, tone):
        return FormulaRule(formula=[formula],
                           fill=PatternFill(bgColor=book._FILLS[tone]),
                           font=Font(color=book._FONTS[tone]))

    cells = (f"{lay.letter(lay.field0)}{lay.first}:"
             f"{lay.letter(lay.field0 + lay.n - 1)}{lay.last}")
    state = lay.at(lay.c_state, lay.first)
    tones = [(c, td.CLASS_OF[c]) for c in td.CHANGED_CODES] + [(_TWO_DIFFERS, "two")]
    for letter, tone in tones:
        ws.conditional_formatting.add(cells, rule(
            f"EXACT(MID({state},COLUMN()-{book.GRID_LEAD},1),{_q(letter)})", tone))
    code = lay.at(lay.c_days, lay.first)
    by_tone = {}
    for presence, tone in book._PRESENCE_TONE.items():
        by_tone.setdefault(tone, []).append(f"{code}={DAYS_CODE[presence]}")
    for tone, tests in by_tone.items():
        ws.conditional_formatting.add(f"C{lay.first}:C{lay.last}",
                                      rule(f"OR({','.join(tests)})", tone))


def _grid_heading_rows(ws, st, sc, days, lay):
    """The legend row and the header row, across all three parts of the sheet."""
    legend = book._grid_legend(ws, st, sc, days)
    legend += [None] * (lay.last_col - len(legend))
    notes = [(col, f"The {day} export, as exported") for day, col in zip(days, lay.flag)]
    notes.append((lay.working0, _working_note(sc, days)))
    for col, text in notes:
        if legend[col - 1] is None:
            legend[col - 1] = book._cell(ws, st, text, st.bold)
    header = book._header(ws, st, book.grid_labels(sc, days), grey=book.shown_only(sc))
    for day in days:
        header += book._header(ws, st, [f"On {day}"]
                               + [f"{day}\n{name}" for name in sc.header])
    working = working_labels(lay, days)
    header += book._header(ws, st, working, grey=working)
    return legend, header


def _stored(ws, raw, text):
    """One exported value as its day block holds it: text as exported, anything
    else (a number) as the text it is compared as."""
    if raw is None or raw == "":
        return None
    return book._plain(ws, excel_text(raw if isinstance(raw, str) else text))


def _day_cells(ws, day_rows, day_trim, idx, n):
    """One record's cells in one day's block."""
    if idx is None:
        return [NOT_ON_DAY] + [None] * n
    raw, text = day_rows[idx], day_trim[idx]
    return [ON_DAY] + [_stored(ws, raw[i], text[i]) for i in range(1, n + 1)]


def _write_grid(wb, st, sc, rows, compared, days, lay, events):
    """One row per record: the grid as formulas, each day's values, and the
    working columns."""
    ws = wb.create_sheet(book.GRID_SHEET)
    ws.sheet_properties.tabColor = book._DARK
    ws.freeze_panes = "C3"
    _grid_columns(ws, sc, lay)
    _grid_colours(ws, lay)
    # The filter spans EVERY column, not the grid's alone: a sort from a header
    # moves the filter's columns, and a record's route and location must move
    # with its three days or each would sit beside another record's values.
    ws.auto_filter.ref = f"A{lay.first - 1}:{lay.letter(lay.last_col)}{lay.last}"
    for heading in _grid_heading_rows(ws, st, sc, days, lay):
        ws.append(heading)
    fill = _FillDown(lay.first, lay.last)
    formulas = grid_formulas(lay, days)
    filled = {col: fill.column(lay.letter(col), text) for col, text in formulas.items()}
    grid_cols = range(3, lay.field0 + lay.n)
    working_cols = range(lay.c_days, lay.last_col + 1)
    first = ([filled[c][0] for c in grid_cols], [filled[c][0] for c in working_cols])
    rest = ([filled[c][1] for c in grid_cols], [filled[c][1] for c in working_cols])
    trim = compared.trim
    total = len(compared.records)
    for k, rec in enumerate(compared.records):
        if k % 2000 == 0:
            book._cancel(events)
            if k and k % _PROGRESS_EVERY == 0 and events is not None:
                events.on_log(f"    … {k:,} of {total:,} {sc.id_noun_plural}")
        grid, working = rest if k else first
        row = [book._plain(ws, book._route_of(rows, rec)),
               book._plain(ws, compared.locations[k])]
        row += grid
        for day, idx in enumerate(rec):
            row += _day_cells(ws, rows[day], trim[day], idx, lay.n)
        row.append(k + 1)
        row += working
        ws.append(row)


# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
class _Live:
    """The Summary's numbers as formulas over the grid sheet."""

    def __init__(self, lay, sc, counts):
        self.lay, self.sc, self.counts = lay, sc, counts
        self.codes = lay.whole(lay.c_days)

    def _records(self, *presences):
        return "+".join(f"COUNTIF({self.codes},{DAYS_CODE[p]})" for p in presences)

    def presence(self, presence):
        return "=" + self._records(presence)

    def cells(self, cls):
        return f"=SUM({self.lay.whole(self.lay.c_count[td.CODE[cls]])})"

    def pair(self, pair):
        """One pair of days as five formulas. The first day against the last is
        the exception when a repeated key pairs differently once the middle day
        is skipped: the grid holds the pairing THROUGH the middle day, so that
        row carries the numbers the app computed by pairing the two directly."""
        if pair == "1-3" and self.counts.first_last_differs():
            direct = self.counts.pairs()[pair]
            return tuple(direct[k] for k in td._PAIR_FIELDS)
        earlier, later = td._PAIR_ONLY[pair]
        differs = self.lay.whole(self.lay.c_pair[pair])
        return ("=" + self._records("all", td._PAIR_TWO[pair]),
                "=" + self._records(*earlier), "=" + self._records(*later),
                f'=COUNTIF({differs},">0")', f"=SUM({differs})")

    def column(self, name):
        """One column's cells by class, then its differing cells on the records
        only two days have — each counted off that column's state letter."""
        place = self.sc.header.index(name) + 1
        state = self.lay.whole(self.lay.c_state)
        letters = [td.CODE[c] for c in td.CLASSES] + [_TWO_DIFFERS]
        return tuple(f"=SUMPRODUCT(--EXACT(MID({state},{place},1),{_q(letter)}))"
                     for letter in letters)


def _formula(ws, st, text, font=None, fill=None):
    """A styled formula cell. The text is this module's own, never data, so it
    is the one kind of cell that must not be guarded into a string."""
    return core._styled_cell(ws, text, font=font or st.body, fill=fill)


def _summary_row(ws, st, kind, cells, live, pairs):
    """One Summary line, as `three_day_workbook` lays it out, with each number
    a formula."""
    if kind == "title":
        return [book._cell(ws, st, cells[0], st.title)]
    if kind == "blank":
        return []
    if kind == "input":
        return [book._cell(ws, st, cells[0], st.bold), book._plain(ws, cells[1])]
    if kind in ("warn", "note"):
        return [book._cell(ws, st, cells[0], st.warn if kind == "warn" else st.muted)]
    if kind == "head":
        return [book._cell(ws, st, c, st.head, st.head_fill) for c in cells if c]
    if kind == "presence":
        return [book._cell(ws, st, cells[0]), live.presence(cells[2])]
    if kind == "class":
        cls = cells[2]
        fill = st.fill.get(cls)
        return [book._cell(ws, st, cells[0], st.font_bold.get(cls, st.bold), fill),
                _formula(ws, st, live.cells(cls), st.bold, fill),
                book._cell(ws, st, td.class_meaning(cls), st.muted)]
    if kind in ("pairhead", "colhead"):
        return book._header(ws, st, cells)
    if kind == "pair":
        return [book._plain(ws, cells[0]), *live.pair(next(pairs))]
    if kind == "column":
        return [book._plain(ws, cells[0]), *live.column(cells[0])]
    raise ValueError(f"unknown Summary line {kind!r}")


def _summary_colours(ws, lines, check_rows):
    """Tints that the values workbook applies to a count that is not zero, and
    green / red on the checks."""
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.styles import Font, PatternFill

    def tint(tone):
        return PatternFill(bgColor=book._FILLS[tone])

    columns = []
    for row, (kind, cells) in enumerate(lines, start=1):
        tone = book._PRESENCE_TONE.get(cells[2]) if kind == "presence" else None
        if tone:
            ws.conditional_formatting.add(f"A{row}:B{row}", FormulaRule(
                formula=[f"$B${row}>0"], fill=tint(tone)))
        if kind == "column":
            columns.append(row)
    if columns:
        for letter, tone in zip("BCDEG", (td.NEWEST, td.MIDDLE, td.BACK, td.BOTH,
                                          "two")):
            ws.conditional_formatting.add(
                f"{letter}{columns[0]}:{letter}{columns[-1]}",
                CellIsRule(operator="greaterThan", formula=["0"], fill=tint(tone)))
    for row in check_rows:
        ws.conditional_formatting.add(f"B{row}", CellIsRule(
            operator="equal", formula=[_q(OK)], fill=PatternFill(bgColor="C6EFCE"),
            font=Font(color="2E7D32", bold=True)))
        ws.conditional_formatting.add(f"B{row}", CellIsRule(
            operator="notEqual", formula=[_q(OK)], fill=PatternFill(bgColor="FFC7CE"),
            font=Font(color="9C0006", bold=True)))


def checks(sc, lay, compared, n_lines):
    """[(what is checked, the condition, what a failure reads)] — the copy's
    own account of whether it still describes the exports."""
    grid, built = book.GRID_SHEET, core._sref(AS_BUILT)
    numbers = f"B2:G{n_lines}"
    order = lay.whole(lay.c_number)
    out = [
        (f"Every number above is the one the app computed (the {AS_BUILT} sheet)",
         f"SUMPRODUCT(--({numbers}<>{built}!{numbers}))=0",
         f"DIFFERS — a value on the {grid} sheet was changed after this workbook "
         "was built"),
        (f"The {grid} sheet is in the order it was built in (Routes, Changes and "
         "Added & removed read its rows by position)",
         f"SUMPRODUCT(--({order}<>ROW({order})-{lay.first - 1}))=0",
         f"SORTED — sort the {grid} sheet by its # column, smallest first"),
    ]
    listed, changed = len(book.changed_cells(compared)), book.total_changed(compared)
    if listed == changed:
        still = (f'COUNTIF(Changes!$G$3:$G${listed + 2},{_q(NO_LONGER)})=0'
                 if listed else "TRUE")
        out.append(("The Changes sheet lists every changed cell",
                    f"AND(SUM({lay.whole(4)})={listed},{still})",
                    "OUT OF DATE — build the comparison again"))
    partial = sum(1 for rec in compared.records if None in rec)
    if len(book.partial_records(compared)) == partial:
        out.append((f"The Added & removed sheet lists every {sc.id_noun} that is "
                    "not on all three days",
                    f'COUNTIF({lay.whole(lay.c_days)},"<{_ON_ALL}")={partial}',
                    "OUT OF DATE — build the comparison again"))
    return out


def _write_summary(wb, st, sc, compared, days, meta, lay):
    """The values Summary, line for line, with each number a formula — then the
    checks."""
    counts = compared.counts
    lines = book.summary_lines(sc, counts, days, meta)
    listed = checks(sc, lay, compared, len(lines))
    first_check = len(lines) + 4
    ws = wb.create_sheet("Summary")
    ws.sheet_properties.tabColor = book._DARK
    book._widths(ws, book.SUMMARY_WIDTHS)
    _summary_colours(ws, lines, range(first_check, first_check + len(listed)))
    live, pairs = _Live(lay, sc, counts), iter(td.PAIRS)
    for kind, cells in lines:
        ws.append(_summary_row(ws, st, kind, cells, live, pairs))
    ws.append([])
    ws.append([book._cell(ws, st, "Live formulas — every line should read OK",
                          st.head, st.head_fill),
               book._cell(ws, st, "check", st.head, st.head_fill)])
    ws.append([book._cell(ws, st, "Excel calculates every number on this sheet from "
                          f"the three days' values on the {book.GRID_SHEET} sheet. "
                          "Notes says how.", st.muted)])
    for label, condition, failure in listed:
        ws.append([book._cell(ws, st, label, st.body, None, st.wrap),
                   _formula(ws, st, f"=IF({condition},{_q(OK)},{_q(failure)})",
                            st.bold)])


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #
def route_blocks(rows, compared):
    """[(route, first record, last record)] in grid order. The grid sorts by
    route first, so each route's records are one run of rows and the Routes
    sheet can total a route by its rows."""
    out = []
    for k, rec in enumerate(compared.records):
        route = book._route_of(rows, rec)
        if out and out[-1][0] == route:
            out[-1][2] = k
        else:
            out.append([route, k, k])
    if len({block[0] for block in out}) != len(out):
        raise ValueError("A route's records are not together in the grid, so the "
                         "live-formulas copy cannot total its routes.")
    return [tuple(block) for block in out]


def _route_colours(ws, n_routes):
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill

    if not n_routes:
        return
    tones = ("added", "removed", "blip", td.NEWEST, td.MIDDLE, td.BACK, td.BOTH, "two")
    for letter, tone in zip("FGHIJKLM", tones):
        ws.conditional_formatting.add(
            f"{letter}3:{letter}{n_routes + 2}",
            CellIsRule(operator="greaterThan", formula=["0"],
                       fill=PatternFill(bgColor=book._FILLS[tone])))


def _write_routes(wb, st, sc, rows, compared, days, lay, events):
    """One row per route, every figure a total over that route's grid rows."""
    blocks = route_blocks(rows, compared)
    ws = book._routes_sheet(wb, st, sc, days, len(blocks),
                            prepare=lambda sheet: _route_colours(sheet, len(blocks)))
    for n, (route, first, last) in enumerate(blocks):
        if n % 200 == 0:
            book._cancel(events)
        a, b = lay.first + first, lay.first + last
        codes = lay.span(lay.c_days, a, b)

        def records(*presences, codes=codes):
            return "=" + "+".join(f"COUNTIF({codes},{DAYS_CODE[p]})"
                                  for p in presences)

        def total(letter, a=a, b=b):
            return f"=SUM({lay.span(lay.c_count[letter], a, b)})"

        ws.append([book._plain(ws, route)]
                  + [f"=COUNTIF({lay.span(flag, a, b)},{_q(ON_DAY)})"
                     for flag in lay.flag]
                  + [records("all"), records("d23", "d3"), records("d12", "d1"),
                     records("d13", "d2")]
                  + [total(letter) for letter in td.DIFF_CODES])


# --------------------------------------------------------------------------- #
# the two lists
# --------------------------------------------------------------------------- #
def _day_value(lay, day, f, row):
    """One day's value of one grid cell, as the Changes list shows it."""
    value = f"TRIM({lay.cell(lay.day0[day] + f, row)})"
    return (f"=IF({lay.cell(lay.flag[day], row)}={_q(ON_DAY)},"
            f'IF({value}="",{_q(td.BLANK_MARK)},{value}),{_q(td.ABSENT_MARK)})')


def _change_formula(lay, f, row, days):
    """What the change is, read off the cell's state letter as it is NOW."""
    letter = f"MID({lay.cell(lay.c_state, row)},{f + 1},1)"
    code = lay.cell(lay.c_days, row)
    classes = ",".join(_q(td.class_short(td.CLASS_OF[c], days))
                       for c in td.CHANGED_CODES)
    two = (f'IF({code}={DAYS_CODE["d12"]},{_q(book.two_day_change("d12", days))},'
           f'IF({code}={DAYS_CODE["d23"]},{_q(book.two_day_change("d23", days))},'
           f'{_q(book.two_day_change("d13", days))}))')
    return (f"=IFERROR(CHOOSE(FIND({letter},{_q(td.DIFF_CODES)}),{classes},{two}),"
            f"{_q(NO_LONGER)})")


def _change_colours(ws, days, n_listed):
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Font, PatternFill

    if not n_listed:
        return
    tones = ([(td.class_short(c, days), c) for c in td.CHANGED]
             + [(book.two_day_change(p, days), "two") for p in td.TWO_DAY])
    for text, tone in tones:
        ws.conditional_formatting.add(
            f"G3:G{n_listed + 2}",
            CellIsRule(operator="equal", formula=[_q(text)],
                       fill=PatternFill(bgColor=book._FILLS[tone]),
                       font=Font(color=book._FONTS[tone])))


def _write_changes(wb, st, sc, rows, compared, days, lay, events):
    """One row per changed cell. WHICH cells are listed is settled when the
    workbook is built; what each row shows is read live from the grid sheet."""
    listed = book.changed_cells(compared)
    ws = book._changes_sheet(
        wb, st, sc, days, len(listed),
        prepare=lambda sheet: _change_colours(sheet, days, len(listed)))
    header = list(sc.header)
    for n, (k, f, _code) in enumerate(listed):
        if n % 2000 == 0:
            book._cancel(events)
        row = lay.first + k
        ws.append([book._plain(ws, book._route_of(rows, compared.records[k])),
                   book._plain(ws, compared.locations[k]), book._plain(ws, header[f]),
                   _day_value(lay, 0, f, row), _day_value(lay, 1, f, row),
                   _day_value(lay, 2, f, row), _change_formula(lay, f, row, days),
                   book._link(ws, st, row)])
    book._changes_foot(ws, st, sc, compared, len(listed))


def _write_records(wb, st, sc, rows, compared, days, lay, events):
    """One row per record that is not on all three days."""
    listed = book.partial_records(compared)
    ws = book._records_sheet(wb, st, sc, days, len(listed))
    labels = ",".join(_q(td.presence_label(p, days)) for p in _BY_CODE)
    two_days = ",".join(f"{{code}}={DAYS_CODE[p]}" for p in td.TWO_DAY)
    for n, k in enumerate(listed):
        if n % 2000 == 0:
            book._cancel(events)
        rec, row = compared.records[k], lay.first + k
        code = lay.cell(lay.c_days, row)
        tone = book._PRESENCE_TONE[td.presence_of(rec)]
        ws.append([book._plain(ws, book._route_of(rows, rec)),
                   book._plain(ws, compared.locations[k]),
                   _formula(ws, st, f"=CHOOSE({code},{labels})", st.font[tone],
                            st.fill[tone])]
                  + [f"=IF({lay.cell(flag, row)}={_q(ON_DAY)},{_q(book.ON_MARK)},"
                     f"{_q(td.ABSENT_MARK)})" for flag in lay.flag]
                  + [f"=IF(OR({two_days.format(code=code)}),"
                     f'{lay.cell(lay.c_count[_TWO_DIFFERS], row)},"")',
                     book._link(ws, st, row)])
    book._records_foot(ws, st, sc, compared, len(listed))


# --------------------------------------------------------------------------- #
# Notes and the entry point
# --------------------------------------------------------------------------- #
def live_notes(sc, days):
    """The Notes the live-formulas copy adds to the standard ones."""
    noun, nouns, grid = sc.id_noun, sc.id_noun_plural, book.GRID_SHEET
    return (
        "LIVE FORMULAS. This is the live-formulas copy of the comparison: the "
        "same sheets as the values workbook, with every number calculated by "
        f"Excel. The three days' exported values are on the {grid} sheet, to the "
        f"right of the grid: one block of columns per day, as exported, then the "
        f"working columns (which days a {noun} is on, and one letter per column "
        "saying how its days compare). Click a cell of the grid, or a number on "
        "the Summary or Routes sheet, to see the formula that reads them. The + "
        f"and - buttons above the {grid} sheet's columns fold those columns away.",
        "EQUAL means the same text once the leading and trailing spaces are "
        "removed and runs of spaces are read as one: EXACT(TRIM(a),TRIM(b)), "
        f"capital and small letters kept apart. Which rows are one {noun} is "
        f"settled when the workbook is built, not by a formula: each row of the "
        f"{grid} sheet holds one {noun}'s three days.",
        "CHECKS. The foot of the Summary holds the live numbers against the ones "
        f"the app computed, which the {AS_BUILT} sheet keeps. Change a day's "
        "value and everything recalculates; that check then reads DIFFERS, "
        f"because the workbook no longer describes the exports. Filter the {grid} "
        "sheet freely. If you sort it, sort it back by its # column: the Routes, "
        f"Changes and Added & removed sheets read its rows by position. The "
        f"Changes sheet lists the cells that had changed when the workbook was "
        f"built, and so does Added & removed for the {nouns}.",
    )


def formula_cells(sc, n_records):
    """How many formula cells the copy's grid sheet holds for `n_records`."""
    lay = Layout(sc, max(n_records, 1))
    return n_records * (lay.last_col - 2 - 3 * (lay.n + 1) - 1)


def write_workbook(sc, rows, compared, days, out_path, meta, events=None,
                   proceed=lambda: True):
    """Write the live-formulas copy to a temp sibling and move it onto
    `out_path` (gated on `proceed()`), so a failed or cancelled write keeps the
    last good one. True iff it landed. Raises ValueError when it cannot fit a
    sheet."""
    from openpyxl import Workbook
    from openpyxl.utils.exceptions import IllegalCharacterError

    lay = Layout(sc, len(compared.records))
    limit = core.excel_limit_error(lay.last, lay.last_col)
    if limit:
        raise ValueError(limit)
    meta = {**meta, "lookalikes": tuple(compared.lookalikes)}
    wb = Workbook(write_only=True)
    st = book._Styles()
    try:
        _write_summary(wb, st, sc, compared, days, meta, lay)
        _write_routes(wb, st, sc, rows, compared, days, lay, events)
        _write_changes(wb, st, sc, rows, compared, days, lay, events)
        _write_records(wb, st, sc, rows, compared, days, lay, events)
        _write_grid(wb, st, sc, rows, compared, days, lay, events)
        book._write_notes(
            wb, st, f"{sc.report_name} — three-day comparison notes",
            (tuple(meta.get("notes", ())) or book.standard_notes(sc, days))
            + live_notes(sc, days))
        book._write_summary(wb, st, sc, compared.counts, days, meta, sheet=AS_BUILT,
                            tab="808080")
        book._cancel(events)
        return artifact_store.atomic_save_if(wb, out_path, proceed)
    except IllegalCharacterError as e:
        book._abandon(wb)
        raise book.UnwritableValue(
            f"One of the {sc.report_name} exports holds a control character Excel "
            f"cannot store ({e!r}), so the live-formulas copy was not written.") from e
    except BaseException:
        book._abandon(wb)
        raise


__all__ = ["AS_BUILT", "DAYS_CODE", "Layout", "NO_LONGER", "OK", "ON_DAY", "SHEETS",
           "checks", "display_formula", "excel_text", "formula_cells",
           "grid_formulas", "live_notes", "route_blocks", "state_formula",
           "working_labels", "write_workbook"]
