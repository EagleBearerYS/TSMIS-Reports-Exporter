"""Three-way comparisons (v0.51.0) — one report's records from THREE sources,
classified cell by cell: the TSMIS site's export, the TSN extract, and OUR build
of the report from the ArcGIS layers.

Why three. TSMIS replaced TSN, so the TSN side of a TSMIS-vs-TSN comparison is
frozen at the September-2025 cutover and most of its differences are the network
changing since then (roadmap D5) — expected, and noise when the question is
whether the REPORT is wrong. The ArcGIS layers are the data TSMIS itself reads,
so bringing them in as a referee sorts every difference:

  * REPORT DIFFERS  (TSN = ArcGIS ≠ TSMIS) — two independent sources agree and the
    report does not. The one worth chasing. The lane exists because of exactly
    this observation: most TSMIS-vs-TSN mismatches were expected because the
    layers differed from TSN too, but in some the layers matched TSN and the
    report was the one that was wrong;
  * TSN DIFFERS     (TSMIS = ArcGIS ≠ TSN) — the report follows the layers and TSN
    is out of date. Expected;
  * ARCGIS DIFFERS  (TSMIS = TSN ≠ ArcGIS) — our layer build disagrees with both:
    a gap in the build, or the report reads something other than the layer the
    column is traced to;
  * ALL DIFFER, and UNDECIDED where the build has no placeable value (the HF-01
    unavailable anchor).

Records pair the way the two comparisons this combines pair them, anchored on
the report: compare_core's own keying and duplicate assignment, under the
TSMIS-vs-TSN comparison's schema for the TSN side and the Reports-vs-ArcGIS
comparison's schema for the ArcGIS side; TSN and ArcGIS rows left over pair with
each other the same way. A record is therefore never keyed differently here than
in either two-sided comparison.

Values arrive projected by the report's adapter into ONE value space — every
side through the same format-only reconciliations — and compare through
compare_core's TRIM, so equality is transitive and the classes partition every
compared cell exactly. A column is compared three ways only where BOTH two-sided
comparisons count it; every other column is shown and never counted.

The workbook is values only: Summary, the three finding lists (one row per cell,
each naming the ArcGIS layer and field the build reads the value from and where
to look it up), the record grid, the layer trace and the Notes. The counts are
published in the outcome sidecar beside it (`SIDECAR_KEY`) — producer-owned truth
the matrix reads; its cache only records whether the inputs moved. Console-free:
progress via the Events sink, exceptions raised.
"""
from __future__ import annotations

import logging
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Mapping

import artifact_store
import compare_core as core
import consolidation_meta
import outcome
from errors import RunCancelled
from events import ConsolidateResult, Events

log = logging.getLogger("tsmis.three_way")

SIDES = ("TSMIS", "TSN", "ArcGIS")
SIDECAR_KEY = "three_way"
COUNTS_VERSION = 1

# Cell classes of a record present on all three sides, in the order a reader
# should look at them.
REPORT = "report"
ARCGIS = "arcgis"
ALL = "all"
TSN = "tsn"
MATCH = "match"
UNDECIDED = "undecided"
CLASSES = (REPORT, ARCGIS, ALL, TSN, MATCH, UNDECIDED)
CLASS_LABELS = {
    REPORT: "Report differs — TSN = ArcGIS ≠ TSMIS",
    ARCGIS: "ArcGIS differs — TSMIS = TSN ≠ ArcGIS",
    ALL: "All three differ",
    TSN: "TSN differs — TSMIS = ArcGIS ≠ TSN",
    MATCH: "All three agree",
    UNDECIDED: "Undecided — the ArcGIS build has no placeable value",
}
CLASS_MEANING = {
    REPORT: "two independent sources agree and the report does not — look here first",
    ARCGIS: "our layer build disagrees with both — a build gap, or the report reads "
            "another source than the traced layer",
    ALL: "three different values — read the three side by side",
    TSN: "expected: the report follows the layers and TSN is out of date (TSN is "
         "frozen at the September-2025 cutover)",
    MATCH: "",
    UNDECIDED: "the build could not place its source span here (HF-01), so only "
               "TSMIS vs TSN is known",
}
_CODE = {MATCH: "=", REPORT: "R", ARCGIS: "A", ALL: "X", TSN: "T", UNDECIDED: "U"}
_CLASS_OF = {v: k for k, v in _CODE.items()}

# Which sides a record has.
PRESENCE = ("all", "tsmis_tsn", "tsmis_arcgis", "tsn_arcgis", "tsmis", "tsn",
            "arcgis")
PRESENCE_LABELS = {
    "all": "On all three",
    "tsmis_tsn": "TSMIS + TSN, not in the ArcGIS build",
    "tsmis_arcgis": "TSMIS + ArcGIS, not in TSN",
    "tsn_arcgis": "TSN + ArcGIS, missing from the TSMIS report",
    "tsmis": "TSMIS only",
    "tsn": "TSN only",
    "arcgis": "ArcGIS only",
}
_PRESENCE_SHORT = {
    "all": "all three", "tsmis_tsn": "TSMIS + TSN", "tsmis_arcgis": "TSMIS + ArcGIS",
    "tsn_arcgis": "TSN + ArcGIS (not in TSMIS)", "tsmis": "TSMIS only",
    "tsn": "TSN only", "arcgis": "ArcGIS only",
}
# The records two sources share, and which two (side indices).
_TWO_SIDE = {"tsmis_tsn": (0, 1), "tsmis_arcgis": (0, 2), "tsn_arcgis": (1, 2)}

_BLANK = core._BLANK_MARK
_MAX_LIST_ROWS = 1_048_000        # one finding list fits one sheet
_GRID_FIRST_ROW = 3               # the grid's legend row, its header, then records
_LIST_SHEETS = ((REPORT, "Report differs"), (ARCGIS, "ArcGIS differs"),
                (ALL, "All differ"))

# Presentation — the class fills match the Summary key everywhere.
_DARK = "1F3864"
_FILLS = {REPORT: "FFC7CE", ARCGIS: "FFEB9C", ALL: "E4DFEC", TSN: "DDEBF7",
          UNDECIDED: "EDEDED", "two": "FFF2CC"}
_FONTS = {REPORT: "9C0006", ARCGIS: "7F6000", ALL: "5B2C6F", TSN: "1F4E79",
          UNDECIDED: "595959", "two": "7F6000"}


# --------------------------------------------------------------------------- #
# the schema a report adapter supplies
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class ColumnTrace:
    """Where the ArcGIS build gets one column — what a reader needs to look the
    value up in the layer themselves. `layers` are library layer names (empty
    for a constant or a value the build derives without one); `fields` the
    layer field(s); `rule` how the build turns them into the value."""
    layers: tuple = ()
    fields: str = ""
    rule: str = ""

    @property
    def source(self):
        if not self.layers:
            return self.rule or "(no layer)"
        names = " + ".join(self.layers)
        return f"{names} · {self.fields}" if self.fields else names


@dataclass(frozen=True)
class ThreeWaySchema:
    """One report's three-way contract.

    `header` is the projected rows' data columns (each row is [route, *header]),
    `key_field` the index of the typed physical key. `pair_tsn` / `pair_arcgis`
    are the two-sided comparisons' own CompareSchemas — they decide record
    pairing exactly as those comparisons do. `counted` names the columns compared
    three ways. `trace` maps every column to its ColumnTrace. `locate(route,
    key, sides)` says where to look a record up ("route 071 · county SBD · PM
    R7.727"): `key` is the record's postmile text, `sides` its (TSMIS, TSN,
    ArcGIS) values as {column: text} dicts, None for an absent side."""
    report_name: str
    noun: str
    noun_plural: str
    header: tuple
    key_field: int
    counted: frozenset
    pair_tsn: object
    pair_arcgis: object
    trace: Mapping
    locate: Callable
    key_label: str = "Location"
    unavailable_token: str = ""
    text_columns: tuple = ()

    def counted_indices(self):
        return tuple(i for i, c in enumerate(self.header)
                     if c in self.counted and i != self.key_field)


@dataclass
class ThreeWayCounts:
    """The published truth of one three-way comparison."""
    records: dict = field(default_factory=lambda: {p: 0 for p in PRESENCE})
    cells: dict = field(default_factory=lambda: {c: 0 for c in CLASSES})
    columns: dict = field(default_factory=dict)       # column -> {class: n}
    two_side: dict = field(default_factory=lambda: {p: 0 for p in _TWO_SIDE})
    counted_columns: int = 0
    pairing_quality: str = "exact"
    completion: str = outcome.COMPLETE

    def to_dict(self):
        return {"version": COUNTS_VERSION, "records": dict(self.records),
                "cells": dict(self.cells),
                "columns": {c: dict(v) for c, v in self.columns.items()},
                "two_side": dict(self.two_side),
                "counted_columns": self.counted_columns,
                "pairing_quality": self.pairing_quality,
                "completion": self.completion}

    @classmethod
    def from_dict(cls, data):
        """The published counts, or None when the record is not this version's
        complete shape (a reader then treats the cell as needing a rebuild)."""
        if not isinstance(data, dict) or data.get("version") != COUNTS_VERSION:
            return None
        recs, cells = data.get("records"), data.get("cells")
        two = data.get("two_side")
        if not all(isinstance(d, dict) for d in (recs, cells, two)):
            return None
        if (set(recs) != set(PRESENCE) or set(cells) != set(CLASSES)
                or set(two) != set(_TWO_SIDE)):
            return None
        if not all(_is_count(v) for d in (recs, cells, two) for v in d.values()):
            return None
        completion = data.get("completion")
        if completion not in outcome.COMPLETIONS:
            return None
        cols = data.get("columns") if isinstance(data.get("columns"), dict) else {}
        return cls(records=dict(recs), cells=dict(cells),
                   columns={c: dict(v) for c, v in cols.items() if isinstance(v, dict)},
                   two_side=dict(two),
                   counted_columns=int(data.get("counted_columns") or 0),
                   pairing_quality=str(data.get("pairing_quality") or "exact"),
                   completion=completion)

    def headline(self):
        """One line for the log and the result summary."""
        c, r = self.cells, self.records
        return (f"{c[REPORT]:,} report differs · {c[ARCGIS]:,} ArcGIS differs · "
                f"{c[ALL]:,} all differ · {c[TSN]:,} TSN differs (expected) · "
                f"{c[MATCH]:,} agree — {r['all']:,} records on all three, "
                f"{r['tsn_arcgis']:,} missing from the report")


def _is_count(v):
    return isinstance(v, int) and not isinstance(v, bool) and v >= 0


# --------------------------------------------------------------------------- #
# pairing
# --------------------------------------------------------------------------- #
def _cancelled(events):
    if events is not None and events.is_cancelled():
        raise RunCancelled()


def _pairs(sc, rows_a, rows_b, idx_a, idx_b, events):
    """{index into rows_a: index into rows_b} for the sub-lists `idx_a` /
    `idx_b`, paired exactly as compare_core pairs two files: the same (route,
    physical key), duplicates assigned by similarity under `sc`. Returns the
    map and the pairing quality ("exact" / "capped")."""
    sub_a = [rows_a[i] for i in idx_a]
    sub_b = [rows_b[j] for j in idx_b]
    if not sub_a or not sub_b:
        return {}, "exact"
    is_cancelled = events.is_cancelled if events is not None else None
    keys_a = core.keys_for(sub_a, True, sc.key_field, sc.key_normalizer, is_cancelled)
    keys_b = core.keys_for(sub_b, True, sc.key_field, sc.key_normalizer, is_cancelled)
    pairing = core.pair_occurrences_by_similarity(
        sc, sub_a, sub_b, keys_a, keys_b, True, events)
    where = {k: i for i, k in enumerate(pairing.keys_a)}
    out = {}
    for j, k in enumerate(pairing.keys_b):
        i = where.get(k)
        if i is not None:
            out[idx_a[i]] = idx_b[j]
    return out, pairing.pairing_quality


def align(schema, rows_m, rows_n, rows_a, events=None):
    """([(m, n, a)], quality) — every record as its row index on each side (None
    where the side has no row), in the report's own order.

    The report anchors: its rows pair with TSN under the TSMIS-vs-TSN schema and
    with ArcGIS under the Reports-vs-ArcGIS schema, independently, so each pair
    is the one its own comparison would make. TSN and ArcGIS rows that paired
    with no report row pair with each other. A record with no report row sits
    right after the report row its nearest preceding same-route row (on its own
    side) paired with, so the grid reads in the report's order."""
    _cancelled(events)
    all_m = list(range(len(rows_m)))
    mn, q1 = _pairs(schema.pair_tsn, rows_m, rows_n, all_m,
                    list(range(len(rows_n))), events)
    _cancelled(events)
    ma, q2 = _pairs(schema.pair_arcgis, rows_m, rows_a, all_m,
                    list(range(len(rows_a))), events)
    _cancelled(events)
    used_n, used_a = set(mn.values()), set(ma.values())
    left_n = [j for j in range(len(rows_n)) if j not in used_n]
    left_a = [j for j in range(len(rows_a)) if j not in used_a]
    na, q3 = _pairs(schema.pair_tsn, rows_n, rows_a, left_n, left_a, events)
    _cancelled(events)
    quality = "capped" if "capped" in (q1, q2, q3) else "exact"

    route_rank = {}
    for rows in (rows_m, rows_n, rows_a):
        for r in rows:
            route_rank.setdefault(r[0], len(route_rank))

    def anchors(rows, report_row_of):
        """Per row: the report row its nearest preceding same-route row (this
        one included) pairs with, or -1."""
        last, out = {}, []
        for j, r in enumerate(rows):
            m = report_row_of.get(j)
            if m is not None:
                last[r[0]] = m
            out.append(last.get(r[0], -1))
        return out

    anchor_n = anchors(rows_n, {j: i for i, j in mn.items()})
    anchor_a = anchors(rows_a, {j: i for i, j in ma.items()})
    ordered = [((route_rank[r[0]], i, 0, 0), (i, mn.get(i), ma.get(i)))
               for i, r in enumerate(rows_m)]
    for j in left_n:
        ordered.append(((route_rank[rows_n[j][0]], anchor_n[j], 1, j),
                        (None, j, na.get(j))))
    a_with_n = set(na.values())
    for j in left_a:
        if j not in a_with_n:
            ordered.append(((route_rank[rows_a[j][0]], anchor_a[j], 2, j),
                            (None, None, j)))
    ordered.sort(key=lambda rec: rec[0])
    return [rec[1] for rec in ordered], quality


def presence_of(rec):
    m, n, a = (x is not None for x in rec)
    if m and n and a:
        return "all"
    if m and n:
        return "tsmis_tsn"
    if m and a:
        return "tsmis_arcgis"
    if n and a:
        return "tsn_arcgis"
    return "tsmis" if m else "tsn" if n else "arcgis"


# --------------------------------------------------------------------------- #
# classification
# --------------------------------------------------------------------------- #
def classify(m, n, a, token=""):
    """The class of one compared cell present on all three sides (TRIMmed
    texts). Equality is plain string equality in the one value space, so the
    five value classes partition every cell."""
    if token and token in (m, n, a):
        return UNDECIDED
    if m == n == a:
        return MATCH
    if n == a:
        return REPORT
    if m == a:
        return TSN
    if m == n:
        return ARCGIS
    return ALL


def _trimmed(rows):
    """Each row's cells as compare_core's TRIMmed comparison text — the text the
    classes are decided on and the grid shows."""
    return [[core._xl_trim(v) for v in r] for r in rows]


def _texts(trim, rec, f):
    """Field `f`'s three texts for one record (None for an absent side)."""
    return [None if idx is None else trim[side][idx][f + 1]
            for side, idx in enumerate(rec)]


def compute(schema, rows_m, rows_n, rows_a, events=None):
    """(records, states, counts, trimmed rows): every record, its per-column
    state string (one character per header column — the class code for a
    compared cell of a record on all three, E/D for a two-side record's compared
    cells, '.' for nothing to compare, '-' for a column shown only, 'k' for the
    key) and the counts."""
    records, quality = align(schema, rows_m, rows_n, rows_a, events)
    trim = (_trimmed(rows_m), _trimmed(rows_n), _trimmed(rows_a))
    counted = schema.counted_indices()
    token = schema.unavailable_token
    counts = ThreeWayCounts(counted_columns=len(counted), pairing_quality=quality)
    counts.columns = {schema.header[f]: {c: 0 for c in CLASSES} for f in counted}
    base = ["k" if f == schema.key_field else "-" for f in range(len(schema.header))]
    for f in counted:
        base[f] = "."
    states = []
    for k, rec in enumerate(records):
        if k % 2000 == 0:
            _cancelled(events)
        presence = presence_of(rec)
        counts.records[presence] += 1
        state = list(base)
        if presence == "all":
            for f in counted:
                m, n, a = _texts(trim, rec, f)
                c = classify(m, n, a, token)
                state[f] = _CODE[c]
                counts.cells[c] += 1
                counts.columns[schema.header[f]][c] += 1
        elif presence in _TWO_SIDE:
            x, y = _TWO_SIDE[presence]
            for f in counted:
                t = _texts(trim, rec, f)
                if token and token in (t[x], t[y]):
                    continue
                if t[x] == t[y]:
                    state[f] = "E"
                else:
                    state[f] = "D"
                    counts.two_side[presence] += 1
        states.append("".join(state))
    if quality != "exact":
        counts.completion = outcome.PARTIAL
    return records, states, counts, trim


# --------------------------------------------------------------------------- #
# the workbook
# --------------------------------------------------------------------------- #
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


def _link(ws, st, sheet, row):
    """A live HYPERLINK to a grid row. Built here from a sheet name and a row
    number, never from data, so it is the one cell that must NOT be guarded."""
    return core._styled_cell(ws, f'=HYPERLINK("#\'{sheet}\'!{row}:{row}",{row})',
                             font=st.link)


def _show(text):
    return text if text else _BLANK


def _display(code, t):
    """The grid text of one cell. The TSMIS value always leads — it is what the
    report printed — followed by whichever source disagrees with it."""
    m, n, a = t
    if code == "=":
        return m
    if code == "R":
        return f"{_show(m)} ≠ {_show(n)} (TSN & ArcGIS)"
    if code == "T":
        return f"{_show(m)} (TSN: {_show(n)})"
    if code == "A":
        return f"{_show(m)} (ArcGIS: {_show(a)})"
    if code == "X":
        return f"{_show(m)} ≠ TSN {_show(n)} ≠ ArcGIS {_show(a)}"
    if code == "U":
        return (f"{_show(m)} (ArcGIS: unplaced)" if m == n
                else f"{_show(m)} (TSN: {_show(n)}; ArcGIS: unplaced)")
    present = [(side, v) for side, v in zip(SIDES, t) if v is not None]
    if not present:
        return ""
    if len({v for _side, v in present}) == 1:
        return present[0][1]
    if code == "D":
        (s1, v1), (s2, v2) = present[0], present[1]
        if s1 == "TSMIS":
            return f"{_show(v1)} ({s2}: {_show(v2)})"
        return f"TSN {_show(v1)} · ArcGIS {_show(v2)}"
    # shown only, or a cell with nothing decidable: every value, labelled
    return " · ".join(f"{side} {_show(v)}" for side, v in present)


def _first_side(rec):
    return next(side for side, idx in enumerate(rec) if idx is not None)


def _key_text(rows, rec, schema):
    side = _first_side(rec)
    return core._visible_key(rows[side][rec[side]][schema.key_field + 1])


def _route_of(rows, rec):
    side = _first_side(rec)
    return rows[side][rec[side]][0]


def _record_sides(trim, rec, schema):
    """The record's (TSMIS, TSN, ArcGIS) values as {column: text}, None for an
    absent side — what `locate` reads."""
    return tuple(None if idx is None else dict(zip(schema.header, trim[side][idx][1:]))
                 for side, idx in enumerate(rec))


def _widths(ws, widths):
    from openpyxl.utils import get_column_letter
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _header(ws, st, labels, grey=()):
    return [_cell(ws, st, label, st.head,
                  st.grey_head_fill if label in grey else st.head_fill, st.center)
            for label in labels]


def _summary_lines(schema, counts, meta):
    """The Summary sheet's content as (kind, cells) lines — kept apart from the
    writer so a check can read what the sheet says without a workbook."""
    out = [("title", (f"{schema.report_name} — TSMIS vs TSN vs ArcGIS (three-way)",)),
           ("blank", ())]
    for side, line in zip(SIDES, meta.get("inputs", ())):
        out.append(("input", (side, line)))
    for line in meta.get("warnings", ()):
        out.append(("warn", (line,)))
    if counts.completion != outcome.COMPLETE:
        reasons = list(meta.get("partial_reasons", ()))
        if counts.pairing_quality == "capped":
            reasons.append("duplicate pairing was capped")
        out.append(("warn", ("PARTIAL — " + ("; ".join(reasons)
                                             or "an input is incomplete"),)))
    out += [("blank", ()), ("head", ("Records", schema.noun_plural))]
    for p in PRESENCE:
        out.append(("count", (PRESENCE_LABELS[p], counts.records[p])))
    out += [("blank", ()),
            ("head", (f"Compared cells — the {counts.counted_columns} columns compared "
                      "three ways, on the records on all three", "cells"))]
    for c in CLASSES:
        out.append(("class", (CLASS_LABELS[c], counts.cells[c], c)))
    out += [("blank", ()),
            ("head", ("Records only two sources have", "differing cells"))]
    for p in _TWO_SIDE:
        out.append(("count", (PRESENCE_LABELS[p], counts.two_side[p])))
    out += [("blank", ()),
            ("colhead", ("Column", "ArcGIS layer · field", "Report differs",
                         "ArcGIS differs", "All differ", "TSN differs", "Agree",
                         "Undecided"))]
    for f in schema.counted_indices():
        name = schema.header[f]
        tr = schema.trace.get(name) or ColumnTrace()
        c = counts.columns.get(name, {})
        out.append(("column", (name, tr.source) + tuple(
            c.get(k, 0) for k in (REPORT, ARCGIS, ALL, TSN, MATCH, UNDECIDED))))
    return out


def _write_summary(wb, st, schema, counts, meta):
    ws = wb.create_sheet("Summary")
    ws.sheet_properties.tabColor = _DARK
    _widths(ws, [56, 46, 14, 14, 12, 14, 12, 12])
    for kind, cells in _summary_lines(schema, counts, meta):
        if kind == "title":
            ws.append([_cell(ws, st, cells[0], st.title)])
        elif kind == "blank":
            ws.append([])
        elif kind == "input":
            ws.append([_cell(ws, st, cells[0], st.bold), _plain(ws, cells[1])])
        elif kind == "warn":
            ws.append([_cell(ws, st, cells[0], st.warn)])
        elif kind == "head":
            ws.append([_cell(ws, st, cells[0], st.head, st.head_fill),
                       _cell(ws, st, cells[1], st.head, st.head_fill)])
        elif kind == "count":
            ws.append([_plain(ws, cells[0]), cells[1]])
        elif kind == "class":
            label, n, cls = cells
            fill = st.fill.get(cls)
            ws.append([_cell(ws, st, label, st.font_bold.get(cls, st.bold), fill),
                       _cell(ws, st, n, st.bold, fill),
                       _cell(ws, st, CLASS_MEANING[cls], st.muted)])
        elif kind == "colhead":
            ws.append(_header(ws, st, cells))
        elif kind == "column":
            name, src, r, a, x, t, m, u = cells
            ws.append([_plain(ws, name), _plain(ws, src),
                       _cell(ws, st, r, st.font_bold[REPORT] if r else st.body,
                             st.fill[REPORT] if r else None),
                       _cell(ws, st, a, st.body, st.fill[ARCGIS] if a else None),
                       _cell(ws, st, x, st.body, st.fill[ALL] if x else None),
                       t, m, u])


_LIST_HEAD = ("Route", None, "Column", "TSMIS", "TSN", "ArcGIS", "ArcGIS layer",
              "Layer field", "How the build gets it", "Look it up", "Grid row")


def _write_list(wb, st, schema, rows, trim, records, states, cls, title, events):
    """One finding list: a row per cell of class `cls` (records on all three),
    naming the layer the ArcGIS value comes from and where to look it up."""
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = _FILLS[cls]
    ws.freeze_panes = "A3"
    _widths(ws, [7, 26, 30, 24, 24, 24, 30, 26, 44, 46, 9])
    ws.append([_cell(ws, st, f"{CLASS_LABELS[cls]} — {CLASS_MEANING[cls]}",
                     st.font_bold[cls], st.fill[cls])])
    ws.append(_header(ws, st, [schema.key_label if h is None else h
                               for h in _LIST_HEAD]))
    code = _CODE[cls]
    fill = st.fill[cls]
    written = 0
    for k, rec in enumerate(records):
        if k % 2000 == 0:
            _cancelled(events)
        state = states[k]
        if code not in state:
            continue
        route = _route_of(rows, rec)
        side = _first_side(rec)
        where = schema.locate(route, trim[side][rec[side]][schema.key_field + 1],
                              _record_sides(trim, rec, schema))
        key = _key_text(rows, rec, schema)
        for f, s in enumerate(state):
            if s != code:
                continue
            if written >= _MAX_LIST_ROWS:
                ws.append([_cell(ws, st, f"… the list stops at {_MAX_LIST_ROWS:,} "
                                 "rows — filter the Three-way sheet for the rest",
                                 st.warn)])
                return written
            name = schema.header[f]
            tr = schema.trace.get(name) or ColumnTrace()
            m, n, a = _texts(trim, rec, f)
            ws.append([
                _plain(ws, route), _plain(ws, key), _plain(ws, name),
                _cell(ws, st, _show(m), st.font_bold[cls] if cls == REPORT else st.body,
                      fill if cls in (REPORT, ALL) else None),
                _cell(ws, st, _show(n), st.body, fill if cls == ALL else None),
                _cell(ws, st, _show(a), st.body, fill if cls in (ARCGIS, ALL) else None),
                _plain(ws, " + ".join(tr.layers) if tr.layers else "(no layer)"),
                _plain(ws, tr.fields), _plain(ws, tr.rule), _plain(ws, where),
                _link(ws, st, "Three-way", _GRID_FIRST_ROW + k),
            ])
            written += 1
    if not written:
        ws.append([_cell(ws, st, "None.", st.muted)])
    else:
        ws.auto_filter.ref = f"A2:K{written + 2}"
    return written


_GRID_COUNTS = (("Report differs", "R"), ("ArcGIS differs", "A"),
                ("All differ", "X"), ("TSN differs", "T"))


def _write_grid(wb, st, schema, rows, trim, records, states, events):
    """The record grid: one row per record, every column, each compared cell
    filled by its class, so the whole picture filters like any sheet."""
    ws = wb.create_sheet("Three-way")
    ws.sheet_properties.tabColor = _DARK
    ws.freeze_panes = "C3"
    counted = set(schema.counted_indices())
    labels = (["Route", schema.key_label, "Presence"]
              + [label for label, _code in _GRID_COUNTS] + list(schema.header))
    _widths(ws, [7, 26, 22, 9, 9, 9, 9] + [
        30 if c in schema.text_columns else 15 for c in schema.header])
    grey = [c for i, c in enumerate(schema.header) if i not in counted]
    legend = [_cell(ws, st, "Each compared cell leads with the TSMIS value, then "
                    "the source that disagrees:", st.bold), None]
    for c in (REPORT, ARCGIS, ALL, TSN):
        legend.append(_cell(ws, st, CLASS_LABELS[c].split(" — ")[0],
                            st.font_bold[c], st.fill[c]))
    legend.append(_cell(ws, st, "two sources only", st.font["two"], st.fill["two"]))
    legend.append(_cell(ws, st, "grey header = shown, never counted", st.muted))
    ws.append(legend)
    ws.append(_header(ws, st, labels, grey=grey))
    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = (f"A{_GRID_FIRST_ROW - 1}:{get_column_letter(len(labels))}"
                          f"{len(records) + _GRID_FIRST_ROW - 1}")
    for k, rec in enumerate(records):
        if k % 2000 == 0:
            _cancelled(events)
        state = states[k]
        tally = Counter(state)
        key = _key_text(rows, rec, schema)
        out = [_plain(ws, _route_of(rows, rec)), _plain(ws, key),
               _PRESENCE_SHORT[presence_of(rec)]]
        out += [tally.get(code, 0) for _label, code in _GRID_COUNTS]
        for f, code in enumerate(state):
            if code == "k":
                out.append(_plain(ws, key))
                continue
            t = _texts(trim, rec, f)
            text = _display(code, t)
            if code in _CLASS_OF and code != "=":
                cls = _CLASS_OF[code]
                out.append(_cell(ws, st, text, st.font[cls], st.fill[cls]))
            elif code == "D":
                out.append(_cell(ws, st, text, st.font["two"], st.fill["two"]))
            elif code in ("-", ".") and len({v for v in t if v is not None}) > 1:
                out.append(_cell(ws, st, text, st.muted))
            else:
                out.append(_plain(ws, text))
        ws.append(out)


def _write_trace(wb, st, schema, layer_index):
    """Every column: whether it is compared three ways, and exactly where the
    ArcGIS build gets it — layer, field, rule, the library file and the feature
    service it was exported from."""
    ws = wb.create_sheet("Layer trace")
    ws.sheet_properties.tabColor = "548235"
    ws.freeze_panes = "B3"
    _widths(ws, [34, 14, 36, 34, 56, 38, 70])
    ws.append([_cell(ws, st, "Where the ArcGIS side of every column comes from. To "
                     "check a finding yourself, open the layer (ArcGIS ▸ Layers ▸ "
                     "Open layers folder) and filter it by the finding's “Look it "
                     "up” location.", st.bold)])
    ws.append(_header(ws, st, ("Column", "Compared", "ArcGIS layer(s)",
                               "Layer field(s)", "How the build gets it",
                               "Library file", "Exported from (feature service)")))
    counted = set(schema.counted_indices())
    for f, name in enumerate(schema.header):
        tr = schema.trace.get(name) or ColumnTrace()
        known = [layer_index[layer] for layer in tr.layers if layer in layer_index]
        compared = ("the key" if f == schema.key_field
                    else "three ways" if f in counted else "shown only")
        ws.append([_plain(ws, name),
                   _cell(ws, st, compared, st.body if f in counted else st.muted),
                   _plain(ws, " + ".join(tr.layers) if tr.layers else "(no layer)"),
                   _plain(ws, tr.fields),
                   _cell(ws, st, tr.rule, st.body, None, st.wrap),
                   _plain(ws, "; ".join(e.get("file", "") for e in known)),
                   _plain(ws, "; ".join(e.get("source", "") for e in known))])


def _write_notes(wb, st, title, lines):
    ws = wb.create_sheet("Notes")
    ws.sheet_properties.tabColor = "ED7D31"
    ws.column_dimensions["A"].width = 120
    ws.append([_cell(ws, st, title, st.head, st.head_fill)])
    for line in lines:
        ws.append([_cell(ws, st, line, st.body, None, st.wrap)])


def write_workbook(schema, rows, trim, records, states, counts, out_path, meta,
                   events=None, proceed=lambda: True):
    """Write the values workbook to a temp sibling and move it onto `out_path`
    (gated on `proceed()`), so a failed or cancelled write keeps the last good
    one. Sheets in reading order: Summary, the report-differs list, the grid,
    the other two lists, the layer trace, the Notes. True iff it landed."""
    from openpyxl import Workbook

    wb = Workbook(write_only=True)
    st = _Styles()
    _write_summary(wb, st, schema, counts, meta)
    lists = dict(_LIST_SHEETS)
    _write_list(wb, st, schema, rows, trim, records, states, REPORT, lists[REPORT],
                events)
    _write_grid(wb, st, schema, rows, trim, records, states, events)
    for cls in (ARCGIS, ALL):
        _write_list(wb, st, schema, rows, trim, records, states, cls, lists[cls],
                    events)
    _write_trace(wb, st, schema, meta.get("layer_index") or {})
    _write_notes(wb, st, f"{schema.report_name} — three-way comparison notes",
                 meta.get("notes", ()))
    _cancelled(events)
    return artifact_store.atomic_save_if(wb, out_path, proceed)


# --------------------------------------------------------------------------- #
# the entry point
# --------------------------------------------------------------------------- #
def read_counts(path):
    """The published counts of a three-way workbook, or None when its outcome
    sidecar is missing, untrusted or no longer describes the file."""
    record = consolidation_meta.read_outcome(path)
    if record is None or not record.trusted or not record.current:
        return None
    return ThreeWayCounts.from_dict(consolidation_meta.read_extra(path, SIDECAR_KEY))


def run(schema, rows_m, rows_n, rows_a, out_path, *, meta=None, events=None,
        commit_guard=None, input_completion=None):
    """Compare, write the workbook and publish its counts beside it. Returns a
    ConsolidateResult (ok / cancelled / error); a consumer reads the counts back
    through `read_counts`, the published truth."""
    events = events or Events()
    meta = dict(meta or {})
    out_path = Path(out_path)
    try:
        events.on_log(f"  pairing {len(rows_m):,} TSMIS · {len(rows_n):,} TSN · "
                      f"{len(rows_a):,} ArcGIS {schema.noun_plural}…")
        records, states, counts, trim = compute(schema, rows_m, rows_n, rows_a, events)
        if input_completion == outcome.PARTIAL:
            counts.completion = outcome.PARTIAL
        events.on_log("  " + counts.headline())
        events.on_log(f"  writing {out_path.name}…")
        landed = write_workbook(
            schema, (rows_m, rows_n, rows_a), trim, records, states, counts,
            out_path, meta, events,
            proceed=lambda: consolidation_meta.guard_allows(commit_guard, out_path))
    except RunCancelled:
        return ConsolidateResult(status="cancelled", message="Cancelled.",
                                 completion=outcome.CANCELLED)
    if not landed:
        return ConsolidateResult(
            status="error",
            message="The three-way workbook's destination changed while it was "
                    "being written — refresh the comparison.")
    result = ConsolidateResult(
        status="ok",
        message=f"{schema.report_name} three-way: {counts.headline()}",
        output_path=str(out_path),
        summary_lines=[f"{schema.report_name} — TSMIS vs TSN vs ArcGIS: "
                       + counts.headline()],
        completion=counts.completion,
        producer_extra={SIDECAR_KEY: counts.to_dict()})
    if not consolidation_meta.write_outcome(
            out_path, result, extra=result.producer_extra,
            commit_guard=commit_guard):
        return ConsolidateResult(
            status="error",
            message="The three-way workbook was written but its counts could not "
                    "be published beside it — rebuild the comparison.")
    result.sidecar_published = True
    return result


__all__ = ["ALL", "ARCGIS", "CLASSES", "CLASS_LABELS", "ColumnTrace", "MATCH",
           "PRESENCE", "PRESENCE_LABELS", "REPORT", "SIDECAR_KEY", "SIDES", "TSN",
           "ThreeWayCounts", "ThreeWaySchema", "UNDECIDED", "align", "classify",
           "compute", "read_counts", "run", "write_workbook"]
