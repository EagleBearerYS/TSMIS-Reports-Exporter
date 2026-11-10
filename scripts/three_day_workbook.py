"""The three-day comparison's workbook (v0.52.0) — values only, written from a
computed `three_day.Compared`.

Sheets, in reading order:

  * Summary           the inputs, the records by the days they are on, the
                      compared cells by class, each pair of days as a two-day
                      comparison counts it, and every column's classes
  * Routes            the same counts per route
  * Changes           one row per changed cell: its value on each day
  * Added & removed   one row per record that is not on all three days
  * Three days        the record grid: every record, every column; a changed
                      cell shows its history ("2 → 3 → 3") in its class's colour
  * Notes             how to read it

Every value shown is the TRIMmed text the classes were decided on, so what a
reader sees is what was compared. Nothing here decides anything: the classes,
states and counts all arrive in the `Compared`.
"""
from __future__ import annotations

import artifact_store
import compare_core as core
import outcome
import three_day as td
from errors import RunCancelled

GRID_SHEET = "Three days"
GRID_FIRST_ROW = 3                # the grid's legend row, its header, then records
GRID_LEAD = 6                     # Route, Location, Presence and three tallies
GRID_LEAD_WIDTHS = [8, 30, 22, 9, 11, 11]
SHEETS = ("Summary", "Routes", "Changes", "Added & removed", GRID_SHEET, "Notes")

_BLANK = td.BLANK_MARK
_ABSENT = td.ABSENT_MARK
_ARROW = td.ARROW
ON_MARK = "✓"               # a day the record is on, in the Added & removed list
_MAX_LIST_ROWS = 1_048_000        # one list fits one sheet
_DARK = "1F3864"
_FILLS = {td.NEWEST: "FFC7CE", td.MIDDLE: "FFEB9C", td.BACK: "DDEBF7",
          td.BOTH: "E4DFEC", "two": "FFF2CC", "added": "E2EFDA",
          "removed": "FCE4D6", "blip": "EDEDED"}
_FONTS = {td.NEWEST: "9C0006", td.MIDDLE: "7F6000", td.BACK: "1F4E79",
          td.BOTH: "5B2C6F", "two": "7F6000", "added": "375623",
          "removed": "843C0C", "blip": "595959"}
_PRESENCE_TONE = {"d12": "removed", "d1": "removed", "d23": "added", "d3": "added",
                  "d13": "blip", "d2": "blip"}
SUMMARY_WIDTHS = (62, 22, 24, 24, 24, 20, 26)


class UnwritableValue(Exception):
    """An export value Excel cannot store, so no workbook was written."""


class _Styles:
    """The handful of shared style components (compare_core's `_styled_cell`
    caches the composite per workbook, so sharing instances is what keeps a
    multi-million-cell grid cheap)."""

    def __init__(self):
        from openpyxl.styles import Alignment, Font, PatternFill
        self.body = Font(name="Arial", size=10)
        self.bold = Font(name="Arial", size=10, bold=True)
        self.title = Font(name="Arial", size=14, bold=True, color=_DARK)
        self.head = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        self.head_fill = PatternFill("solid", start_color=_DARK)
        self.grey_head_fill = PatternFill("solid", start_color="808080")
        self.muted = Font(name="Arial", size=10, color="7F7F7F")
        self.warn = Font(name="Arial", size=10, bold=True, color="C00000")
        self.link = Font(name="Arial", size=10, color="0563C1", underline="single")
        self.wrap = Alignment(vertical="top", wrap_text=True)
        self.center = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
        self.fill = {k: PatternFill("solid", start_color=v) for k, v in _FILLS.items()}
        self.font = {k: Font(name="Arial", size=10, color=v) for k, v in _FONTS.items()}
        self.font_bold = {k: Font(name="Arial", size=10, bold=True, color=v)
                          for k, v in _FONTS.items()}


def _cell(ws, st, value, font=None, fill=None, align=None):
    """One streamed styled cell; text Excel would read as a formula stays text."""
    return core._styled_cell(ws, value, font=font or st.body, fill=fill,
                             align=align, guard=True)


def _plain(ws, value):
    """A bare value for the hot path — only formula-looking text needs a cell."""
    if isinstance(value, str) and core.is_formula_injection(value):
        from openpyxl.cell import WriteOnlyCell
        c = WriteOnlyCell(ws, value=value)
        c.data_type = "s"
        return c
    return value


def _link(ws, st, row):
    """A live HYPERLINK to a grid row. Built here from the grid's own sheet name
    and a row number, never from data, so it is the one cell that must NOT be
    guarded."""
    return core._styled_cell(
        ws, f'=HYPERLINK("#\'{GRID_SHEET}\'!{row}:{row}",{row})', font=st.link)


def day_text(text):
    """One day's value as a list shows it: the value, (blank), or — for a day
    the record is not on."""
    if text is None:
        return _ABSENT
    return text if text else _BLANK


def history(texts):
    """A cell's three values in day order: '2 → 3 → 3'."""
    return _ARROW.join(day_text(v) for v in texts)


def _newest(texts):
    """The value on the newest day the record is on."""
    return next(v for v in reversed(texts) if v is not None)


def _route_of(rows, rec):
    day = next(d for d, idx in enumerate(rec) if idx is not None)
    return rows[day][rec[day]][0]


def _widths(ws, widths):
    from openpyxl.utils import get_column_letter
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _header(ws, st, labels, grey=()):
    return [_cell(ws, st, label, st.head,
                  st.grey_head_fill if label in grey else st.head_fill, st.center)
            for label in labels]


def _tinted(ws, st, n, tone):
    """A count, filled in its tone when it is not zero."""
    return _cell(ws, st, n, st.body, st.fill[tone] if n else None)


# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
def summary_lines(sc, counts, days, meta):
    """The Summary sheet's content as (kind, cells) lines — kept apart from the
    writer so a check can read what the sheet says."""
    d1, d2, d3 = days
    out = [("title", (f"{sc.report_name} — three days: {d1} → {d2} → {d3}",)),
           ("blank", ())]
    for day, line in zip(days, meta.get("inputs", ())):
        out.append(("input", (day, line)))
    for line in meta.get("warnings", ()):
        out.append(("warn", (line,)))
    lookalikes = tuple(meta.get("lookalikes", ()))
    if lookalikes:
        shown = ", ".join(repr(v) for v in lookalikes[:5])
        more = f" and {len(lookalikes) - 5} more" if len(lookalikes) > 5 else ""
        out.append(("warn", (
            f"{len(lookalikes)} value{'' if len(lookalikes) == 1 else 's'} in these "
            f"exports read like this workbook's own marks ({_BLANK!r} for an empty "
            f"cell, {_ABSENT!r} for a day a {sc.id_noun} is not on, {_ARROW!r} between "
            f"days): {shown}{more}. The counts are unaffected; where a cell shows "
            "one, read it in the export file.",)))
    if counts.completion != outcome.COMPLETE:
        reasons = list(meta.get("partial_reasons", ()))
        if counts.pairing_quality == "capped":
            reasons.append("a key repeats too many times to pair its rows exactly")
        out.append(("warn", ("PARTIAL — " + ("; ".join(reasons)
                                             or "an input is incomplete"),)))
    out += [("blank", ()), ("head", ("Records", sc.id_noun_plural, ""))]
    for p in td.PRESENCE:
        out.append(("presence", (td.presence_label(p, days), counts.records[p], p)))
    out += [("blank", ()),
            ("head", (f"Compared cells — the {counts.counted_columns} compared "
                      f"columns of the {sc.id_noun_plural} on all three days",
                      "cells", "what it means"))]
    for c in td.CLASSES:
        out.append(("class", (td.class_label(c, days), counts.cells[c], c)))
    out += [("blank", ()),
            ("pairhead", ("Each pair of days, as a two-day comparison counts it",
                          f"{sc.id_noun_plural} on both",
                          "only on the earlier day", "only on the later day",
                          f"{sc.id_noun_plural} that differ", "differing cells"))]
    pairs = counts.pairs()
    for pair, (x, y) in td.PAIRS.items():
        p = pairs[pair]
        out.append(("pair", (f"{days[x]} vs {days[y]}", p["paired"],
                             p["only_earlier"], p["only_later"],
                             p["differing_records"], p["differing_cells"])))
    if counts.first_last_differs():
        t = counts.through_middle()["1-3"]
        out.append(("note", (
            f"Read through {d2} instead, {d1} vs {d3} comes to "
            f"{t['paired']:,} {sc.id_noun_plural} on both and "
            f"{t['differing_cells']:,} differing cells: a key that repeats can "
            "pair one way across each consecutive pair of days and another way "
            "when the middle day is skipped. The classes above read through "
            f"{d2}.",)))
    out += [("blank", ()),
            ("colhead", ("Column", f"Changed {d3}", f"Changed {d2}",
                         f"Changed {d2}, back {d3}", f"Changed {d2} and {d3}",
                         "Same", f"Differs, {sc.id_noun_plural} on two days"))]
    for f in sc.field_indices:
        name = sc.header[f]
        if name not in counts.columns:
            continue
        c = counts.columns[name]
        two = sum(counts.two_day_columns[p].get(name, 0) for p in td.TWO_DAY)
        out.append(("column", (name,) + tuple(c.get(k, 0) for k in td.CLASSES)
                    + (two,)))
    shown = [sc.header[f] for f in sc.field_indices if sc.is_context(f)]
    if shown:
        out += [("blank", ()),
                ("note", ("Shown, never counted: " + ", ".join(shown),))]
    return out


def _write_summary(wb, st, sc, counts, days, meta, sheet="Summary", tab=_DARK):
    ws = wb.create_sheet(sheet)
    ws.sheet_properties.tabColor = tab
    _widths(ws, SUMMARY_WIDTHS)
    for kind, cells in summary_lines(sc, counts, days, meta):
        if kind == "title":
            ws.append([_cell(ws, st, cells[0], st.title)])
        elif kind == "blank":
            ws.append([])
        elif kind == "input":
            ws.append([_cell(ws, st, cells[0], st.bold), _plain(ws, cells[1])])
        elif kind == "warn":
            ws.append([_cell(ws, st, cells[0], st.warn)])
        elif kind == "note":
            ws.append([_cell(ws, st, cells[0], st.muted)])
        elif kind == "head":
            ws.append([_cell(ws, st, c, st.head, st.head_fill) for c in cells if c])
        elif kind == "presence":
            label, n, p = cells
            tone = _PRESENCE_TONE.get(p)
            fill = st.fill.get(tone) if (tone and n) else None
            ws.append([_cell(ws, st, label, st.body, fill),
                       _cell(ws, st, n, st.bold if fill else st.body, fill)])
        elif kind == "class":
            label, n, cls = cells
            fill = st.fill.get(cls)
            ws.append([_cell(ws, st, label, st.font_bold.get(cls, st.bold), fill),
                       _cell(ws, st, n, st.bold, fill),
                       _cell(ws, st, td.class_meaning(cls), st.muted)])
        elif kind in ("pairhead", "colhead"):
            ws.append(_header(ws, st, cells))
        elif kind == "pair":
            ws.append([_plain(ws, cells[0])] + list(cells[1:]))
        elif kind == "column":
            name, n, m, b, x, same, two = cells
            ws.append([_plain(ws, name), _tinted(ws, st, n, td.NEWEST),
                       _tinted(ws, st, m, td.MIDDLE), _tinted(ws, st, b, td.BACK),
                       _tinted(ws, st, x, td.BOTH), same,
                       _tinted(ws, st, two, "two")])


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #
def route_table(rows, compared):
    """[(route, tallies)] in grid order: each route's records per day, by the
    days they are on, and its changed cells by class ('D' = the differing cells
    of its records on two days)."""
    out = {}
    for rec, state in zip(compared.records, compared.states):
        route = _route_of(rows, rec)
        t = out.get(route)
        if t is None:
            t = out[route] = dict.fromkeys(
                ("n1", "n2", "n3", *td.PRESENCE, *td.DIFF_CODES), 0)
        presence = td.presence_of(rec)
        t[presence] += 1
        for day, idx in enumerate(rec):
            if idx is not None:
                t[f"n{day + 1}"] += 1
        if presence == "all":
            for code in td.CHANGED_CODES:
                t[code] += state.count(code)
        elif presence in td.TWO_DAY:
            t["D"] += state.count("D")
    return list(out.items())


def _routes_sheet(wb, st, sc, days, n_routes, prepare=None):
    """The Routes sheet with its two heading rows written, ready for one row
    per route. `prepare(ws)` runs before the first row is streamed — the last
    moment a sheet can be given anything but rows."""
    d1, d2, d3 = days
    ws = wb.create_sheet("Routes")
    ws.sheet_properties.tabColor = "70AD47"
    ws.freeze_panes = "B3"
    _widths(ws, [9, 13, 13, 13, 13, 11, 11, 15, 14, 14, 16, 16, 18])
    if n_routes:
        ws.auto_filter.ref = f"A2:M{n_routes + 2}"
    if prepare is not None:
        prepare(ws)
    ws.append([_cell(ws, st, f"Every route's {sc.id_noun_plural} and changed "
                     "cells. Filter or sort a column to find the routes that "
                     "moved.", st.bold)])
    ws.append(_header(ws, st, (
        "Route", f"{d1} {sc.id_noun_plural}", f"{d2} {sc.id_noun_plural}",
        f"{d3} {sc.id_noun_plural}", "On all three days", "Added", "Removed",
        f"On or off {d2} only", f"Changed {d3}", f"Changed {d2}",
        f"Changed {d2}, back {d3}", f"Changed {d2} and {d3}",
        f"Differs, {sc.id_noun_plural} on two days")))
    return ws


def _write_routes(wb, st, sc, rows, compared, days, events):
    table = route_table(rows, compared)
    ws = _routes_sheet(wb, st, sc, days, len(table))
    for k, (route, t) in enumerate(table):
        if k % 200 == 0:
            _cancel(events)
        ws.append([_plain(ws, route), t["n1"], t["n2"], t["n3"], t["all"],
                   _tinted(ws, st, t["d23"] + t["d3"], "added"),
                   _tinted(ws, st, t["d12"] + t["d1"], "removed"),
                   _tinted(ws, st, t["d13"] + t["d2"], "blip"),
                   _tinted(ws, st, t["N"], td.NEWEST),
                   _tinted(ws, st, t["M"], td.MIDDLE),
                   _tinted(ws, st, t["B"], td.BACK),
                   _tinted(ws, st, t["X"], td.BOTH),
                   _tinted(ws, st, t["D"], "two")])


# --------------------------------------------------------------------------- #
# the two lists
# --------------------------------------------------------------------------- #
def _cancel(events):
    if events is not None and events.is_cancelled():
        raise RunCancelled()


def two_day_change(presence, days):
    """What a differing cell of a record on two days says."""
    d1, d2, d3 = days
    return {"d12": f"changed {d2} (then removed {d3})",
            "d23": f"changed {d3} (added {d2})",
            "d13": f"differs {d1} vs {d3} (missing {d2})"}[presence]


def changed_cells(compared):
    """Every listed change as (record index, column index, state code), in grid
    order — at most `_MAX_LIST_ROWS` of them, the rows one list sheet holds."""
    out = []
    for k, state in enumerate(compared.states):
        if not any(code in state for code in td.DIFF_CODES):
            continue
        for f, code in enumerate(state):
            if code in td.DIFF_CODES:
                if len(out) >= _MAX_LIST_ROWS:
                    return out
                out.append((k, f, code))
    return out


def total_changed(compared):
    """How many compared cells differ between the days their record is on."""
    return sum(sum(state.count(code) for code in td.DIFF_CODES)
               for state in compared.states)


def _changes_sheet(wb, st, sc, days, n_listed, prepare=None):
    """The Changes sheet with its two heading rows written, ready for one row
    per changed cell. `prepare(ws)` runs before the first row is streamed."""
    ws = wb.create_sheet("Changes")
    ws.sheet_properties.tabColor = _FILLS[td.NEWEST]
    ws.freeze_panes = "A3"
    _widths(ws, [8, 30, 32, 26, 26, 26, 38, 9])
    if n_listed:
        ws.auto_filter.ref = f"A2:H{n_listed + 2}"
    if prepare is not None:
        prepare(ws)
    ws.append([_cell(ws, st, "Every compared cell that is not the same on all the "
                     f"days its {sc.id_noun} is on. Filter the Change or the "
                     "Column header.", st.bold)])
    ws.append(_header(ws, st, ("Route", "Location", "Column", days[0], days[1],
                               days[2], "Change", "Grid row")))
    return ws


def _changes_foot(ws, st, sc, compared, n_listed):
    """The line under the list: none to list, or where a full sheet stops."""
    if not n_listed:
        ws.append([_cell(ws, st, "None — every compared cell is the same on every "
                         f"day its {sc.id_noun} is on.", st.muted)])
    elif n_listed < total_changed(compared):
        ws.append([_cell(ws, st, f"… the list stops at {_MAX_LIST_ROWS:,} "
                         f"rows — filter the {GRID_SHEET} sheet for the rest",
                         st.warn)])


def _write_changes(wb, st, sc, rows, compared, days, events):
    """One row per changed cell: the record, the column, its value on each day
    and what the change is."""
    listed = changed_cells(compared)
    ws = _changes_sheet(wb, st, sc, days, len(listed))
    header = list(sc.header)
    for n, (k, f, code) in enumerate(listed):
        if n % 2000 == 0:
            _cancel(events)
        rec = compared.records[k]
        texts = td.texts_of(compared.trim, rec, f)
        if code == "D":
            tone, change = "two", two_day_change(td.presence_of(rec), days)
        else:
            tone = td.CLASS_OF[code]
            change = td.class_short(tone, days)
        # Every cell is built for THIS row: a guarded (formula-looking) text is
        # a cell object, and one object cannot sit in two rows.
        ws.append([_plain(ws, _route_of(rows, rec)), _plain(ws, compared.locations[k]),
                   _plain(ws, header[f]),
                   _plain(ws, day_text(texts[0])), _plain(ws, day_text(texts[1])),
                   _plain(ws, day_text(texts[2])),
                   _cell(ws, st, change, st.font[tone], st.fill[tone]),
                   _link(ws, st, GRID_FIRST_ROW + k)])
    _changes_foot(ws, st, sc, compared, len(listed))
    return len(listed)


def partial_records(compared):
    """The index of every record that is not on all three days, in grid order —
    at most `_MAX_LIST_ROWS` of them, the rows one list sheet holds."""
    out = []
    for k, rec in enumerate(compared.records):
        if None in rec:
            if len(out) >= _MAX_LIST_ROWS:
                break
            out.append(k)
    return out


def _records_sheet(wb, st, sc, days, n_listed):
    """The Added & removed sheet with its two heading rows written, ready for
    one row per record that is not on all three days."""
    ws = wb.create_sheet("Added & removed")
    ws.sheet_properties.tabColor = _FILLS["added"]
    ws.freeze_panes = "A3"
    _widths(ws, [8, 30, 46, 13, 13, 13, 16, 9])
    if n_listed:
        ws.auto_filter.ref = f"A2:H{n_listed + 2}"
    ws.append([_cell(ws, st, f"Every {sc.id_noun} that is not on all three days.",
                     st.bold)])
    ws.append(_header(ws, st, ("Route", "Location", "What happened", days[0],
                               days[1], days[2], "Differing cells", "Grid row")))
    return ws


def _records_foot(ws, st, sc, compared, n_listed):
    """The line under the list: none to list, or where a full sheet stops."""
    partial = sum(1 for rec in compared.records if None in rec)
    if not n_listed:
        ws.append([_cell(ws, st, f"None — every {sc.id_noun} is on all three days.",
                         st.muted)])
    elif n_listed < partial:
        ws.append([_cell(ws, st, f"… the list stops at {_MAX_LIST_ROWS:,} rows — "
                         f"filter the {GRID_SHEET} sheet's Presence column",
                         st.warn)])


def _write_records(wb, st, sc, rows, compared, days, events):
    """One row per record that is not on all three days."""
    listed = partial_records(compared)
    ws = _records_sheet(wb, st, sc, days, len(listed))
    for n, k in enumerate(listed):
        if n % 2000 == 0:
            _cancel(events)
        rec, state = compared.records[k], compared.states[k]
        presence = td.presence_of(rec)
        tone = _PRESENCE_TONE[presence]
        ws.append([_plain(ws, _route_of(rows, rec)),
                   _plain(ws, compared.locations[k]),
                   _cell(ws, st, td.presence_label(presence, days), st.font[tone],
                         st.fill[tone])]
                  + [ON_MARK if idx is not None else _ABSENT for idx in rec]
                  + [state.count("D") if presence in td.TWO_DAY else None,
                     _link(ws, st, GRID_FIRST_ROW + k)])
    _records_foot(ws, st, sc, compared, len(listed))
    return len(listed)


# --------------------------------------------------------------------------- #
# the grid
# --------------------------------------------------------------------------- #
def grid_text(code, texts):
    """The grid text of one cell: its history where the days differ, else its
    value. A column that is shown only (a context column, the key column) shows
    the newest day the record is on."""
    if code in td.DIFF_CODES:
        return history(texts)
    return _newest(texts)


def grid_labels(sc, days):
    """The grid's column headings: its six lead columns, then every column of
    the report."""
    return (["Route", "Location", "Presence", "Changed cells",
             f"Changed {days[1]}", f"Changed {days[2]}"] + list(sc.header))


def field_widths(sc):
    """One width per column of the report, as the grid shows it."""
    wide = sc.data_widths or {}
    return [max(12, min(40, int(wide.get(c) or 16))) for c in sc.header]


def shown_only(sc):
    """The columns that are shown and never counted."""
    return {sc.header[f] for f in sc.field_indices if sc.is_context(f)}


def _grid_legend(ws, st, sc, days):
    """The grid's first row: how to read a changed cell, and the colours."""
    legend = [_cell(ws, st, "A changed cell shows its three days in order "
                    f"({days[0]} → {days[1]} → {days[2]}; — = not on that day):",
                    st.bold), None, None]
    for c in td.CHANGED:
        legend.append(_cell(ws, st, td.class_short(c, days), st.font_bold[c],
                            st.fill[c]))
    legend.append(_cell(ws, st, f"{sc.id_noun} on two days only", st.font["two"],
                        st.fill["two"]))
    legend.append(_cell(ws, st, "grey header = shown, never counted", st.muted))
    return legend


def _write_grid(wb, st, sc, rows, compared, days, events):
    """One row per record, every column, each changed cell showing its history
    and filled by its class."""
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet(GRID_SHEET)
    ws.sheet_properties.tabColor = _DARK
    ws.freeze_panes = "C3"
    labels = grid_labels(sc, days)
    _widths(ws, GRID_LEAD_WIDTHS + field_widths(sc))
    ws.append(_grid_legend(ws, st, sc, days))
    ws.append(_header(ws, st, labels, grey=shown_only(sc)))
    ws.auto_filter.ref = (f"A{GRID_FIRST_ROW - 1}:{get_column_letter(len(labels))}"
                          f"{len(compared.records) + GRID_FIRST_ROW - 1}")
    trim = compared.trim
    for k, (rec, state) in enumerate(zip(compared.records, compared.states)):
        if k % 2000 == 0:
            _cancel(events)
        presence = td.presence_of(rec)
        tone = _PRESENCE_TONE.get(presence)
        on_2 = on_3 = None
        if presence == "all":
            n, m, b, x = (state.count(code) for code in td.CHANGED_CODES)
            changed, on_2, on_3 = n + m + b + x, m + b + x, n + b + x
        else:
            changed = state.count("D")
            if presence == "d12":
                on_2 = changed
            elif presence == "d23":
                on_3 = changed
        out = [_plain(ws, _route_of(rows, rec)), _plain(ws, compared.locations[k]),
               (_cell(ws, st, td.presence_short(presence, days), st.font[tone],
                      st.fill[tone]) if tone else td.presence_short(presence, days)),
               changed, on_2, on_3]
        if not changed:
            # Nothing differs: every cell is the newest day's own text.
            newest = next(trim[d][rec[d]] for d in (2, 1, 0) if rec[d] is not None)
            out.extend(_plain(ws, v) for v in newest[1:])
            ws.append(out)
            continue
        for f, code in enumerate(state):
            texts = td.texts_of(trim, rec, f)
            text = grid_text(code, texts)
            if code == "D":
                out.append(_cell(ws, st, text, st.font["two"], st.fill["two"]))
            elif code in td.CHANGED_CODES:
                cls = td.CLASS_OF[code]
                out.append(_cell(ws, st, text, st.font[cls], st.fill[cls]))
            else:
                out.append(_plain(ws, text))
        ws.append(out)


# --------------------------------------------------------------------------- #
# Notes
# --------------------------------------------------------------------------- #
def standard_notes(sc, days):
    """The Notes every three-day workbook carries."""
    d1, d2, d3 = days
    noun, nouns = sc.id_noun, sc.id_noun_plural
    shown = [sc.header[f] for f in sc.field_indices if sc.is_context(f)]
    return (
        f"Three exports of the {sc.report_name} report from the same source: "
        f"{d1}, {d2} and {d3}. Every {noun} is matched across the three days and "
        "every compared cell is sorted by its history.",
        f"READING A CELL. 'Changed {d3}' = the value was the same on {d1} and {d2} "
        f"and is different on {d3}. 'Changed {d2}' = it changed between {d1} and "
        f"{d2}, and {d3} still has the new value. 'Changed {d2}, back {d3}' = {d1} "
        f"and {d3} agree and {d2} stands alone. 'Changed {d2} and {d3}' = three "
        "different values. A changed cell shows its three values in day order: "
        "'2 → 3 → 3'. (blank) is an empty cell; — means the "
        f"{noun} is not on that day.",
        f"RECORDS. A {noun} on {d2} and {d3} but not {d1} was added on {d2}; on "
        f"{d3} only, added on {d3}; on {d1} and {d2} but not {d3}, removed on {d3}; "
        f"on {d1} only, removed on {d2}. One on {d1} and {d3} but not {d2} is "
        f"missing on {d2} only, and one on {d2} only appeared and was gone again. "
        f"The Added & removed sheet lists them. A {noun} on two days still has "
        "those two days' cells compared (yellow where they differ).",
        f"MATCHING. A {noun} is matched the way the two-day comparison (Compare ▸ "
        "vs Baseline Matrix) matches it: the same route and the same physical "
        f"key, and where a key repeats within a route, the {nouns} that are most "
        f"alike pair. {d2} anchors: it is matched to {d1} and to {d3}, and what "
        f"is on neither side of it is matched between {d1} and {d3}.",
        "THE SAME NUMBERS AS THE TWO-DAY COMPARISONS. For each pair of days the "
        f"Summary shows the {nouns} on both days, the {nouns} only one day has, "
        f"the {nouns} that differ and the differing cells — the numbers a two-day "
        "comparison of the later day against the earlier one reports.",
        "VALUES are compared as the two-day comparison compares them: as text, "
        "with leading and trailing spaces removed and runs of spaces read as one. "
        "Nothing else is normalized, so a change of format between two exports "
        "shows as a change."
        + (" SHOWN, NEVER COUNTED: " + ", ".join(shown) + " — the day the site "
           "made the file, so it differs on every row between two days. The grid "
           f"shows the newest day's value there." if shown else ""),
        f"THE GRID's Changed cells column counts a row's changed cells; the two "
        f"columns after it count the cells that changed on {d2} and on {d3}. "
        "Filter on them, or on Presence, to read one kind of change at a time. "
        "The Grid row link on the two lists jumps to the record's row.",
    )


def _write_notes(wb, st, title, lines):
    ws = wb.create_sheet("Notes")
    ws.sheet_properties.tabColor = "ED7D31"
    ws.column_dimensions["A"].width = 120
    ws.append([_cell(ws, st, title, st.head, st.head_fill)])
    for line in lines:
        ws.append([_cell(ws, st, line, st.body, None, st.wrap)])


def _abandon(wb):
    """Finish and drop a write-only workbook that will not be saved (a cancel, a
    value Excel refuses, a locked destination), so its row generators and temp
    files do not outlive it. Best effort: the workbook is being discarded."""
    for ws in wb.worksheets:
        try:
            ws.close()
        except Exception:  # noqa: BLE001  # silent-ok: discarding an unsaved workbook; nothing to report
            pass
        writer = getattr(ws, "_writer", None)
        try:
            if writer is not None:
                writer.cleanup()
        except Exception:  # noqa: BLE001  # silent-ok: a temp file already gone is the goal
            pass
    try:
        wb.close()
    except Exception:  # noqa: BLE001  # silent-ok: discarding an unsaved workbook; nothing to report
        pass


def write_workbook(sc, rows, compared, days, out_path, meta, events=None,
                   proceed=lambda: True):
    """Write the values workbook to a temp sibling and move it onto `out_path`
    (gated on `proceed()`), so a failed or cancelled write keeps the last good
    one. True iff it landed. Raises ValueError when the grid cannot fit a sheet."""
    from openpyxl import Workbook

    from openpyxl.utils.exceptions import IllegalCharacterError

    limit = core.excel_limit_error(len(compared.records) + GRID_FIRST_ROW - 1,
                                   len(sc.header) + GRID_LEAD)
    if limit:
        raise ValueError(limit)
    meta = {**meta, "lookalikes": tuple(compared.lookalikes)}
    wb = Workbook(write_only=True)
    st = _Styles()
    try:
        _write_summary(wb, st, sc, compared.counts, days, meta)
        _write_routes(wb, st, sc, rows, compared, days, events)
        _write_changes(wb, st, sc, rows, compared, days, events)
        _write_records(wb, st, sc, rows, compared, days, events)
        _write_grid(wb, st, sc, rows, compared, days, events)
        _write_notes(wb, st, f"{sc.report_name} — three-day comparison notes",
                     tuple(meta.get("notes", ())) or standard_notes(sc, days))
        _cancel(events)
        return artifact_store.atomic_save_if(wb, out_path, proceed)
    except IllegalCharacterError as e:
        _abandon(wb)
        raise UnwritableValue(
            f"One of the {sc.report_name} exports holds a control character Excel "
            f"cannot store ({e!r}), so the three-day workbook was not written. Find "
            "the cell in the export and re-export that route.") from e
    except BaseException:
        _abandon(wb)
        raise


__all__ = ["GRID_FIRST_ROW", "GRID_SHEET", "SHEETS", "UnwritableValue",
           "changed_cells", "day_text", "field_widths", "grid_labels", "grid_text",
           "history", "partial_records", "route_table", "shown_only",
           "standard_notes", "summary_lines", "total_changed", "two_day_change",
           "write_workbook"]
