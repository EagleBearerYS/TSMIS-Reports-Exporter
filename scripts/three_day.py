"""Three-day comparisons (v0.52.0) — one report's records on THREE export days of
the same source, classified cell by cell.

A two-day comparison (the vs Baseline matrix) says what differs between two
pulls. A third pull turns each difference into a history, and the history is the
finding: a value that changed on the middle day and stayed is an edit that stuck;
one that changed only on the newest day is fresh; one that changed and went back
is a blip (or a correction of one); one that changed both times is still moving.
Records sort the same way: added, removed, there on the middle day only, or
missing on the middle day only.

  * CHANGED ON THE NEWEST DAY   (day 1 = day 2 ≠ day 3)
  * CHANGED ON THE MIDDLE DAY   (day 1 ≠ day 2 = day 3) — and still so
  * CHANGED AND BACK            (day 1 = day 3 ≠ day 2)
  * CHANGED BOTH TIMES          (three different values)
  * SAME                        (day 1 = day 2 = day 3)

Nothing here re-decides what a two-day comparison decides. The three sides load
through the report's own cross-environment adapter (`compare_env.EnvCompare`, the
loader the vs Baseline matrix uses), records key and pair through compare_core's
own `keys_for` + `pair_occurrences_by_similarity` under that adapter's schema,
and a cell is equal exactly when `compare_core.compared_cell` says so (the TRIMmed
text; a context column is shown and never counted). The middle day anchors the
pairing: day 2 pairs with day 1 as "day 2 vs baseline day 1" would, day 3 pairs
with day 2 as "day 3 vs baseline day 2" would, and what is on neither side of
day 2 pairs day 3 with day 1. So the counts of each consecutive pair of days —
paired records, records only one of the two has, differing cells in total and
per column — are the two-day comparison's own numbers, published here as `pairs`
so a reader (and `build/check_three_day.py`) can hold the two side by side.

Equality has to be ONE relation across the three days for the five classes to
partition every cell, so a schema whose equality is not plain (Highway Log's
non-asserting dittos, the Med Wid rule, an unavailable-anchor token) is refused
rather than classified on faith.

The workbook (`three_day_workbook`) is values only. The counts are published in
the outcome sidecar beside it (`SIDECAR_KEY`) — producer-owned truth the matrix
reads. On request a live-formulas copy (`three_day_formulas`) is written beside
it from the SAME computed comparison — an extra that never decides the result.
Console-free: progress via the Events sink, exceptions raised.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

import artifact_store
import compare_core as core
import consolidation_meta
import outcome
from errors import RunCancelled
from events import ConsolidateResult, Events

log = logging.getLogger("tsmis.three_day")

SIDECAR_KEY = "three_day"
COUNTS_VERSION = 1
# The live-formulas copy holds one formula per grid cell and per working column.
# Past this many it is a file Excel takes minutes to open, so the copy is skipped
# and the skip is said; the values workbook holds every value regardless.
FORMULAS_MAX_CELLS = 8_000_000

# Cell classes of a record present on all three days.
NEWEST = "newest"
MIDDLE = "middle"
BACK = "back"
BOTH = "both"
SAME = "same"
CLASSES = (NEWEST, MIDDLE, BACK, BOTH, SAME)
CHANGED = (NEWEST, MIDDLE, BACK, BOTH)
# One character per header column in a record's state string: a class code for a
# compared cell of a record on all three days, E / D for the compared cells of a
# record only two days have, '.' where there is nothing to compare (a one-day
# record), '-' for a column that is shown only and 'k' for the key column.
CODE = {SAME: "=", NEWEST: "N", MIDDLE: "M", BACK: "B", BOTH: "X"}
CLASS_OF = {v: k for k, v in CODE.items()}
CHANGED_CODES = "NMBX"
DIFF_CODES = "NMBXD"
# How the workbook writes a cell's three days: 'a → b → c', an empty cell as
# (blank), a day the record is not on as —. An export value that already reads
# like one of them would be told apart from none, so `compute` lists them.
BLANK_MARK = core._BLANK_MARK
ABSENT_MARK = "—"
ARROW = " → "

# Which days a record is on (day 1 = oldest).
PRESENCE = ("all", "d12", "d23", "d13", "d1", "d2", "d3")
TWO_DAY = {"d12": (0, 1), "d23": (1, 2), "d13": (0, 2)}
# The three pairs of days, as (earlier, later) day indices.
PAIRS = {"1-2": (0, 1), "2-3": (1, 2), "1-3": (0, 2)}
# Per pair: the classes whose cells differ between its two days, the presence
# whose records only those two days have, and the presences only its earlier /
# only its later day has.
_PAIR_DIFFERS = {"1-2": (MIDDLE, BACK, BOTH), "2-3": (NEWEST, BACK, BOTH),
                 "1-3": (NEWEST, MIDDLE, BOTH)}
_PAIR_TWO = {"1-2": "d12", "2-3": "d23", "1-3": "d13"}
_PAIR_ONLY = {"1-2": (("d13", "d1"), ("d23", "d2")),
              "2-3": (("d12", "d2"), ("d13", "d3")),
              "1-3": (("d12", "d1"), ("d23", "d3"))}
# The counts one pair of days carries (beside its per-column `column_diffs`).
_PAIR_FIELDS = ("paired", "only_earlier", "only_later", "differing_records",
                "differing_cells")


# --------------------------------------------------------------------------- #
# labels — every one names the days it is about
# --------------------------------------------------------------------------- #
def class_label(cls, days):
    d1, d2, d3 = days
    return {
        NEWEST: f"Changed on {d3} — the same on {d1} and {d2}",
        MIDDLE: f"Changed on {d2} — and still so on {d3}",
        BACK: f"Changed on {d2}, back on {d3}",
        BOTH: f"Changed on {d2} and again on {d3}",
        SAME: "Same on all three days",
    }[cls]


def class_short(cls, days):
    _d1, d2, d3 = days
    return {
        NEWEST: f"changed {d3}",
        MIDDLE: f"changed {d2}",
        BACK: f"changed {d2}, back {d3}",
        BOTH: f"changed {d2} and {d3}",
        SAME: "same",
    }[cls]


def class_meaning(cls):
    return {
        NEWEST: "new since the middle day — the freshest change",
        MIDDLE: "an edit that stuck: both later days carry the new value",
        BACK: "the middle day stands alone: the value was changed and put back, "
              "or the middle day's file had it wrong",
        BOTH: "three different values — still moving",
        SAME: "",
    }[cls]


def presence_label(presence, days):
    d1, d2, d3 = days
    return {
        "all": "On all three days",
        "d12": f"Removed on {d3} (on {d1} and {d2})",
        "d23": f"Added on {d2} (on {d2} and {d3})",
        "d13": f"Missing on {d2} only (on {d1} and {d3})",
        "d1": f"Removed on {d2} (only on {d1})",
        "d2": f"Only on {d2} (not on {d1}, gone on {d3})",
        "d3": f"Added on {d3} (only on {d3})",
    }[presence]


def presence_short(presence, days):
    _d1, d2, d3 = days
    return {
        "all": "all three days",
        "d12": f"removed {d3}",
        "d23": f"added {d2}",
        "d13": f"missing {d2} only",
        "d1": f"removed {d2}",
        "d2": f"only on {d2}",
        "d3": f"added {d3}",
    }[presence]


# --------------------------------------------------------------------------- #
# the published counts
# --------------------------------------------------------------------------- #
@dataclass
class ThreeDayCounts:
    """The published truth of one three-day comparison.

    `records` counts the records by the days they are on; `cells` the compared
    cells of the records on all three days, by class; `columns` the same per
    column; `two_day` the differing compared cells of the records only two days
    have (and `two_day_columns` those per column); `differing_records` the
    records with at least one differing cell between each pair of days, read
    off the classes. `first_last` is the first day against the last day paired
    DIRECTLY, as a two-day comparison of those two days pairs them."""
    records: dict = field(default_factory=lambda: {p: 0 for p in PRESENCE})
    cells: dict = field(default_factory=lambda: {c: 0 for c in CLASSES})
    columns: dict = field(default_factory=dict)       # column -> {class: n}
    two_day: dict = field(default_factory=lambda: {p: 0 for p in TWO_DAY})
    two_day_columns: dict = field(
        default_factory=lambda: {p: {} for p in TWO_DAY})
    differing_records: dict = field(default_factory=lambda: {p: 0 for p in PAIRS})
    first_last: dict = field(default_factory=dict)
    counted_columns: int = 0
    pairing_quality: str = "exact"
    completion: str = outcome.COMPLETE

    def through_middle(self):
        """Each pair of days read off the classes: the records both days have,
        the records only one has, the records and the cells that differ, and
        `column_diffs` per column. The two consecutive pairs are paired exactly
        as their own two-day comparisons pair them, so these ARE those
        comparisons' counts. The first-vs-last pair is reached through the
        middle day, which can pair a repeated key differently from a direct
        comparison of the two days — `pairs()` reports that one directly."""
        r, c, t = self.records, self.cells, self.two_day
        out = {}
        for pair in PAIRS:
            two = _PAIR_TWO[pair]
            differs = _PAIR_DIFFERS[pair]
            two_cols = self.two_day_columns.get(two, {})
            out[pair] = {
                "paired": r["all"] + r[two],
                "only_earlier": sum(r[p] for p in _PAIR_ONLY[pair][0]),
                "only_later": sum(r[p] for p in _PAIR_ONLY[pair][1]),
                "differing_records": self.differing_records[pair],
                "differing_cells": sum(c[k] for k in differs) + t[two],
                "column_diffs": {
                    col: sum(by_class.get(k, 0) for k in differs)
                    + two_cols.get(col, 0)
                    for col, by_class in self.columns.items()},
            }
        return out

    def pairs(self):
        """Each pair of days as a two-day comparison of those two days counts
        it — the numbers the vs Baseline matrix shows for the later day against
        the earlier one."""
        out = self.through_middle()
        if self.first_last:
            out["1-3"] = {k: (dict(v) if isinstance(v, dict) else v)
                          for k, v in self.first_last.items()}
        return out

    def first_last_differs(self):
        """True when the first-vs-last pair reads differently through the
        middle day than paired directly (only a repeated key can do that)."""
        return bool(self.first_last) and self.through_middle()["1-3"] != self.pairs()["1-3"]

    def changed_cells(self):
        """Compared cells that are not the same on all three days."""
        return sum(self.cells[c] for c in CHANGED)

    def added(self):
        return self.records["d23"] + self.records["d3"]

    def removed(self):
        return self.records["d12"] + self.records["d1"]

    def middle_only(self):
        """Records on, or off, the middle day alone."""
        return self.records["d13"] + self.records["d2"]

    def to_dict(self):
        return {"version": COUNTS_VERSION, "records": dict(self.records),
                "cells": dict(self.cells),
                "columns": {c: dict(v) for c, v in self.columns.items()},
                "two_day": dict(self.two_day),
                "two_day_columns": {p: dict(v)
                                    for p, v in self.two_day_columns.items()},
                "differing_records": dict(self.differing_records),
                "first_last": {k: (dict(v) if isinstance(v, dict) else v)
                               for k, v in self.first_last.items()},
                "pairs": self.pairs(),
                "counted_columns": self.counted_columns,
                "pairing_quality": self.pairing_quality,
                "completion": self.completion}

    @classmethod
    def from_dict(cls, data):
        """The published counts, or None when the record is not this version's
        complete, well-formed shape (a reader then treats the cell as needing a
        rebuild). Every field is required and checked: a damaged sidecar must
        never raise into a snapshot, and never read as a smaller comparison."""
        if not isinstance(data, dict) or data.get("version") != COUNTS_VERSION:
            return None
        recs, cells, two = data.get("records"), data.get("cells"), data.get("two_day")
        differing, first_last = data.get("differing_records"), data.get("first_last")
        cols, two_cols = data.get("columns"), data.get("two_day_columns")
        if not all(isinstance(d, dict) for d in (recs, cells, two, differing,
                                                 first_last, cols, two_cols)):
            return None
        if (set(recs) != set(PRESENCE) or set(cells) != set(CLASSES)
                or set(two) != set(TWO_DAY) or set(differing) != set(PAIRS)
                or set(two_cols) != set(TWO_DAY)):
            return None
        if not all(_is_count(v) for d in (recs, cells, two, differing)
                   for v in d.values()):
            return None
        if not all(_count_map(v, CLASSES) for v in cols.values()):
            return None
        if not all(_count_map(v) for v in two_cols.values()):
            return None
        if not (set(_PAIR_FIELDS) <= set(first_last)
                and all(_is_count(first_last[k]) for k in _PAIR_FIELDS)
                and _count_map(first_last.get("column_diffs"))):
            return None
        counted, quality = data.get("counted_columns"), data.get("pairing_quality")
        completion = data.get("completion")
        if (not _is_count(counted) or counted != len(cols)
                or quality not in ("exact", "capped")
                or completion not in (outcome.COMPLETE, outcome.PARTIAL)):
            return None
        return cls(records=dict(recs), cells=dict(cells),
                   columns={c: dict(v) for c, v in cols.items()},
                   two_day=dict(two),
                   two_day_columns={p: dict(two_cols[p]) for p in TWO_DAY},
                   differing_records=dict(differing),
                   first_last={k: (dict(v) if isinstance(v, dict) else v)
                               for k, v in first_last.items()},
                   counted_columns=counted, pairing_quality=quality,
                   completion=completion)

    def headline(self, days):
        """One line for the log and the result summary."""
        c, r = self.cells, self.records
        _d1, d2, d3 = days
        return (f"{c[NEWEST]:,} cells changed on {d3} · {c[MIDDLE]:,} changed on {d2} "
                f"and kept · {c[BACK]:,} changed and back · {c[BOTH]:,} changed both "
                f"times — {r['all']:,} records on all three days, "
                f"{self.added():,} added, {self.removed():,} removed, "
                f"{self.middle_only():,} on or off the middle day only")


def _is_count(v):
    return isinstance(v, int) and not isinstance(v, bool) and v >= 0


def _count_map(v, keys=None):
    """True for a {name: count} map (with exactly `keys`, when given)."""
    return (isinstance(v, dict) and all(isinstance(k, str) for k in v)
            and all(_is_count(n) for n in v.values())
            and (keys is None or set(v) == set(keys)))


# --------------------------------------------------------------------------- #
# the schema gate
# --------------------------------------------------------------------------- #
def schema_problem(sc):
    """Why `sc` cannot be compared across three days, or "" when it can.

    The five classes partition every cell only when equality is ONE relation
    across the three days. compare_core's equality is the TRIMmed text except
    where a schema opts into more: a ditto marker asserts nothing (Highway Log),
    Med Wid compares as a number-with-suffix, an unavailable anchor asserts
    nothing. None of those is transitive across three sides. The per-column
    counts are published by column name, so the names must be distinct too."""
    if sc.ditto_nonasserting:
        return "its ditto markers assert nothing, so three days cannot be classified"
    if sc.unavailable_token:
        return "its unavailable-anchor cells assert nothing"
    if any(sc.is_medwid(f) for f in sc.field_indices):
        return "its Med Wid column compares by a rule of its own"
    if len(set(sc.header)) != len(sc.header):
        return "two of its columns share a name"
    return ""


# --------------------------------------------------------------------------- #
# pairing
# --------------------------------------------------------------------------- #
def _cancelled(events):
    if events is not None and events.is_cancelled():
        raise RunCancelled()


def day_keys(sc, day_rows, events=None):
    """One day's (route, key, occurrence) keys, exactly as the two-day
    comparison derives them (`compare_core.keys_for` under `sc`)."""
    is_cancelled = events.is_cancelled if events is not None else None
    return core.keys_for(day_rows, True, sc.key_field, sc.key_normalizer,
                         is_cancelled)


def _pair_keyed(sc, rows_a, rows_b, keys_a, keys_b, events):
    """{index into rows_a: index into rows_b}, paired exactly as compare_core
    pairs two folders: the same (route, physical key), duplicates assigned by
    similarity under `sc`, `rows_a` as the comparison's first side. Returns the
    map and the pairing quality."""
    if not rows_a or not rows_b:
        return {}, "exact"
    pairing = core.pair_occurrences_by_similarity(
        sc, rows_a, rows_b, keys_a, keys_b, True, events)
    where = {k: i for i, k in enumerate(pairing.keys_a)}
    out = {}
    for j, k in enumerate(pairing.keys_b):
        i = where.get(k)
        if i is not None:
            out[i] = j
    return out, pairing.pairing_quality


def _pair_subsets(sc, rows_a, rows_b, idx_a, idx_b, events):
    """`_pair_keyed` over the sub-lists `idx_a` / `idx_b`, in the full lists'
    indices. The sub-lists are keyed afresh: an occurrence number counts a key's
    repeats within the rows being paired."""
    sub_a = [rows_a[i] for i in idx_a]
    sub_b = [rows_b[j] for j in idx_b]
    if not sub_a or not sub_b:
        return {}, "exact"
    paired, quality = _pair_keyed(sc, sub_a, sub_b, day_keys(sc, sub_a, events),
                                  day_keys(sc, sub_b, events), events)
    return {idx_a[i]: idx_b[j] for i, j in paired.items()}, quality


def align(sc, rows, keys, events=None):
    """([(i1, i2, i3)], quality) — every record as its row index on each day
    (None where that day has no row), in the middle day's order. `keys` is each
    day's `day_keys`.

    Day 2 anchors. It pairs with day 1 as the two-day comparison "day 2 vs
    baseline day 1" pairs them, and day 3 pairs with it as "day 3 vs baseline day
    2" does, independently, so each consecutive pair is the one its own two-day
    comparison would make. Day-1 and day-3 rows that paired with no day-2 row
    pair with each other, day 3 first, the same way. A record with no day-2 row
    sits right after the day-2 row its nearest preceding same-route row (on its
    own day) paired with."""
    rows_1, rows_2, rows_3 = rows
    _cancelled(events)
    all_1, all_3 = range(len(rows_1)), range(len(rows_3))
    p21, q1 = _pair_keyed(sc, rows_2, rows_1, keys[1], keys[0], events)   # {i2: i1}
    _cancelled(events)
    p32, q2 = _pair_keyed(sc, rows_3, rows_2, keys[2], keys[1], events)   # {i3: i2}
    _cancelled(events)
    of_2_in_3 = {i2: i3 for i3, i2 in p32.items()}
    used_1, used_3 = set(p21.values()), set(p32)
    left_1 = [j for j in all_1 if j not in used_1]
    left_3 = [j for j in all_3 if j not in used_3]
    p31, q3 = _pair_subsets(sc, rows_3, rows_1, left_3, left_1, events)   # {i3: i1}
    _cancelled(events)
    quality = "capped" if "capped" in (q1, q2, q3) else "exact"

    route_rank = {}
    for day_rows in (rows_2, rows_1, rows_3):
        for r in day_rows:
            route_rank.setdefault(r[0], len(route_rank))

    def anchors(day_rows, day2_row_of):
        """Per row: the day-2 row its nearest preceding same-route row (this one
        included) pairs with, or -1."""
        last, out = {}, []
        for j, r in enumerate(day_rows):
            m = day2_row_of.get(j)
            if m is not None:
                last[r[0]] = m
            out.append(last.get(r[0], -1))
        return out

    anchor_1 = anchors(rows_1, {i1: i2 for i2, i1 in p21.items()})
    anchor_3 = anchors(rows_3, dict(p32))
    ordered = [((route_rank[r[0]], i2, 0, 0), (p21.get(i2), i2, of_2_in_3.get(i2)))
               for i2, r in enumerate(rows_2)]
    of_1_in_3 = {i1: i3 for i3, i1 in p31.items()}
    for i1 in left_1:
        ordered.append(((route_rank[rows_1[i1][0]], anchor_1[i1], 1, i1),
                        (i1, None, of_1_in_3.get(i1))))
    for i3 in left_3:
        if i3 not in p31:
            ordered.append(((route_rank[rows_3[i3][0]], anchor_3[i3], 2, i3),
                            (None, None, i3)))
    ordered.sort(key=lambda rec: rec[0])
    return [rec[1] for rec in ordered], quality


def presence_of(rec):
    a, b, c = (x is not None for x in rec)
    if a and b and c:
        return "all"
    if a and b:
        return "d12"
    if b and c:
        return "d23"
    if a and c:
        return "d13"
    return "d1" if a else "d2" if b else "d3"


# --------------------------------------------------------------------------- #
# classification
# --------------------------------------------------------------------------- #
def classify(a, b, c):
    """The class of one compared cell of a record on all three days, from its
    three TRIMmed texts (day 1, day 2, day 3). Equality is plain string equality
    — the relation compare_core's `compared_cell` decides a plain cell by — so
    the five classes partition every cell."""
    if a == b:
        return SAME if b == c else NEWEST
    if b == c:
        return MIDDLE
    return BACK if a == c else BOTH


def _trimmed(day_rows, memo):
    """Each row's cells as compare_core's TRIMmed comparison text — the text the
    classes are decided on and the workbook shows. `memo` shares one string per
    distinct text across the three days (a statewide report repeats a few
    thousand values over millions of cells)."""
    trim, shared = core._xl_trim, memo.setdefault
    return [[shared(t, t) for t in map(trim, r)] for r in day_rows]


def texts_of(trim, rec, f):
    """Field `f`'s three texts for one record (None for a day without it)."""
    return [None if idx is None else trim[day][idx][f + 1]
            for day, idx in enumerate(rec)]


@dataclass
class Compared:
    """One three-day comparison, computed: every record as its row index on each
    day, its per-column state string (see `CODE`), its location as the two-day
    comparison's key column shows it, the three days' rows as TRIMmed text, and
    the counts."""
    records: list
    states: list
    locations: list
    trim: tuple
    counts: ThreeDayCounts
    lookalikes: tuple = ()       # export values that read like a workbook marker


def counted_fields(sc):
    """The header columns that are compared (every column but the key and the
    ones that are shown only)."""
    return tuple(f for f in sc.field_indices if not sc.is_context(f))


def one_day_state(sc):
    """The state string of a record only one day has: `.` under a compared
    column, `k` under the key, `-` under a column that is shown only. Every
    other record's state fills the `.` places in."""
    base = ["k" if f == sc.key_field else "-" for f in range(len(sc.header))]
    for f in counted_fields(sc):
        base[f] = "."
    return "".join(base)


def _first_last(sc, rows, keys, trim, counted, events):
    """The first day against the last day, paired DIRECTLY as a two-day
    comparison "day 3 vs baseline day 1" pairs them — its counts, and the
    pairing quality."""
    paired, quality = _pair_keyed(sc, rows[2], rows[0], keys[2], keys[0], events)
    per_col = dict.fromkeys(counted, 0)
    differing = 0
    for k, (i3, i1) in enumerate(paired.items()):
        if k % 2000 == 0:
            _cancelled(events)
        later, earlier = trim[2][i3], trim[0][i1]
        differs = False
        for f in counted:
            if later[f + 1] != earlier[f + 1]:
                per_col[f] += 1
                differs = True
        differing += differs
    header = sc.header
    return {"paired": len(paired),
            "only_earlier": len(rows[0]) - len(paired),
            "only_later": len(rows[2]) - len(paired),
            "differing_records": differing,
            "differing_cells": sum(per_col.values()),
            "column_diffs": {header[f]: n for f, n in per_col.items()}}, quality


def compute(sc, rows, events=None):
    """The three days' rows (oldest first, each row `[route, *header values]`)
    aligned, classified and counted — a `Compared`."""
    problem = schema_problem(sc)
    if problem:
        raise ValueError(f"{sc.report_name} cannot be compared across three days: "
                         f"{problem}.")
    keys = tuple(day_keys(sc, r, events) for r in rows)
    records, quality = align(sc, rows, keys, events)
    memo = {}
    trim = tuple(_trimmed(r, memo) for r in rows)
    lookalikes = tuple(sorted(t for t in memo
                              if t in (BLANK_MARK, ABSENT_MARK) or ARROW in t))
    del memo
    header = list(sc.header)
    counted = counted_fields(sc)
    counts = ThreeDayCounts(counted_columns=len(counted), pairing_quality=quality)
    changed = {f: dict.fromkeys(CHANGED, 0) for f in counted}
    two = {p: dict.fromkeys(counted, 0) for p in TWO_DAY}
    pair_of = {two_day: pair for pair, two_day in _PAIR_TWO.items()}
    differs_in = {pair: frozenset(CODE[c] for c in classes)
                  for pair, classes in _PAIR_DIFFERS.items()}
    one_day = one_day_state(sc)
    base = list(one_day)
    unchanged = one_day.replace(".", CODE[SAME])
    states, locations = [], []
    for k, rec in enumerate(records):
        if k % 2000 == 0:
            _cancelled(events)
        day = next(d for d, idx in enumerate(rec) if idx is not None)
        locations.append(core._visible_key(keys[day][rec[day]][1]))
        presence = presence_of(rec)
        counts.records[presence] += 1
        if presence == "all":
            row_a, row_b, row_c = (trim[d][rec[d]] for d in range(3))
            state = None
            for f in counted:
                a, b, c = row_a[f + 1], row_b[f + 1], row_c[f + 1]
                if a == b and b == c:
                    continue
                cls = classify(a, b, c)
                if state is None:
                    state = list(unchanged)
                state[f] = CODE[cls]
                changed[f][cls] += 1
            if state is None:
                states.append(unchanged)
                continue
            seen = set(state)
            for pair, codes in differs_in.items():
                if not seen.isdisjoint(codes):
                    counts.differing_records[pair] += 1
            states.append("".join(state))
        elif presence in TWO_DAY:
            x, y = TWO_DAY[presence]
            row_x, row_y = trim[x][rec[x]], trim[y][rec[y]]
            per_col = two[presence]
            state = list(base)
            differs = False
            for f in counted:
                if row_x[f + 1] == row_y[f + 1]:
                    state[f] = "E"
                else:
                    state[f] = "D"
                    per_col[f] += 1
                    differs = True
            counts.differing_records[pair_of[presence]] += differs
            states.append("".join(state))
        else:
            states.append(one_day)
    on_all = counts.records["all"]
    for f in counted:
        by_class = dict(changed[f])
        by_class[SAME] = on_all - sum(by_class.values())
        counts.columns[header[f]] = {c: by_class[c] for c in CLASSES}
        for c in CLASSES:
            counts.cells[c] += by_class[c]
    for p in TWO_DAY:
        counts.two_day_columns[p] = {header[f]: n for f, n in two[p].items()}
        counts.two_day[p] = sum(two[p].values())
    counts.first_last, direct_quality = _first_last(sc, rows, keys, trim, counted,
                                                    events)
    if "capped" in (quality, direct_quality):
        counts.pairing_quality = "capped"
        counts.completion = outcome.PARTIAL
    return Compared(records, states, locations, trim, counts, lookalikes)


# --------------------------------------------------------------------------- #
# the entry points
# --------------------------------------------------------------------------- #
def read_counts(path):
    """The published counts of a three-day workbook, or None when its outcome
    sidecar is missing, untrusted or no longer describes the file."""
    record = consolidation_meta.read_outcome(path)
    if record is None or not record.trusted or not record.current:
        return None
    return ThreeDayCounts.from_dict(consolidation_meta.read_extra(path, SIDECAR_KEY))


def formulas_sibling(out_path):
    """Where a workbook's live-formulas copy sits: beside it, under the name
    every matrix gives that copy ('<name> (formulas).xlsx')."""
    out_path = Path(out_path)
    return out_path.with_name(f"{out_path.stem} (formulas){out_path.suffix}")


def _write_formulas_copy(sc, rows, compared, days, out_path, meta, events,
                         commit_guard):
    """Write the live-formulas copy beside the values workbook that just landed.
    True iff a fresh copy is in place. Never raises: the copy is an extra, and
    the values workbook and its published counts stand whatever happens here."""
    twin = formulas_sibling(out_path)
    try:
        import three_day_formulas as live    # lazy: the openpyxl writer

        cells = live.formula_cells(sc, len(compared.records))
        if cells > FORMULAS_MAX_CELLS:
            events.on_log(f"  Skipping the live-formulas copy of {out_path.name}: "
                          f"{cells:,} formulas, over the {FORMULAS_MAX_CELLS:,} this "
                          "matrix writes. The values workbook has every value.")
            return False
        events.on_log(f"  writing {twin.name}…")
        return bool(live.write_workbook(
            sc, rows, compared, days, twin, meta, events,
            proceed=lambda: consolidation_meta.guard_allows(commit_guard, twin)))
    except RunCancelled:
        events.on_log("  The live-formulas copy was cancelled. The values "
                      "workbook is complete.")
    except Exception as e:  # noqa: BLE001 - an optional copy must never fail the comparison
        log.warning("three-day: live-formulas copy %s not written (%s: %s)",
                    twin.name, type(e).__name__, e, exc_info=True)
        events.on_log(f"  ⚠ The live-formulas copy was not written: "
                      f"{str(e) or type(e).__name__}")
    return False


def _remove_stale_copy(out_path, events, commit_guard):
    """Remove a live-formulas copy this build did not refresh: left in place it
    would sit beside a newer values workbook and look current. One that is still
    open in Excel cannot be removed, and that is said."""
    twin = formulas_sibling(out_path)
    try:
        if not twin.exists():
            return
    except OSError:  # silent-ok: a path that cannot be read is treated as absent
        return
    if not consolidation_meta.guard_allows(commit_guard, twin):
        log.warning("three-day: an earlier live-formulas copy %s was left (the "
                    "destination changed owner)", twin.name)
        return
    try:
        twin.unlink()
    except FileNotFoundError:  # silent-ok: already gone is the goal
        return
    except OSError as e:
        log.warning("three-day: could not remove the earlier live-formulas copy "
                    "%s (%s: %s)", twin.name, type(e).__name__, e)
        events.on_log(f"  ⚠ {twin.name} is from an earlier build and could not be "
                      "removed (is it open in Excel?). It does not match this "
                      "comparison.")
        return
    events.on_log(f"  Removed the earlier live-formulas copy {twin.name}: it was "
                  "not refreshed and no longer matches.")


def run(sc, rows, days, out_path, *, meta=None, events=None, commit_guard=None,
        input_completion=None, also_formulas=False):
    """Compare three days' rows (oldest first), write the workbook and publish
    its counts beside it. Returns a ConsolidateResult (ok / cancelled / error); a
    consumer reads the counts back through `read_counts`, the published truth.

    `also_formulas` also writes the live-formulas copy beside the workbook, from
    the same computed comparison. Without it, or when the copy is skipped or
    fails, a copy left from an earlier build is removed."""
    import three_day_workbook as book        # lazy: the openpyxl writer

    events = events or Events()
    meta = dict(meta or {})
    out_path = Path(out_path)
    days = tuple(days)
    try:
        events.on_log(f"  pairing {len(rows[0]):,} · {len(rows[1]):,} · "
                      f"{len(rows[2]):,} {sc.id_noun_plural} ({' · '.join(days)})…")
        compared = compute(sc, rows, events)
        counts = compared.counts
        if input_completion == outcome.PARTIAL:
            counts.completion = outcome.PARTIAL
        events.on_log("  " + counts.headline(days))
        events.on_log(f"  writing {out_path.name}…")
        landed = book.write_workbook(
            sc, rows, compared, days, out_path, meta, events,
            proceed=lambda: consolidation_meta.guard_allows(commit_guard, out_path))
    except RunCancelled:
        return ConsolidateResult(status="cancelled", message="Cancelled.",
                                 completion=outcome.CANCELLED)
    except ValueError as e:
        return ConsolidateResult(status="error", message=str(e))
    except OSError as e:
        return ConsolidateResult(
            status="error",
            message=f"Could not write {out_path.name} ({type(e).__name__}: {e}). "
                    "If it is open in Excel, close it and rebuild the comparison.")
    except book.UnwritableValue as e:
        return ConsolidateResult(status="error", message=str(e))
    if not landed:
        return ConsolidateResult(
            status="error",
            message="The three-day workbook's destination changed while it was "
                    "being written — refresh the comparison.")
    result = ConsolidateResult(
        status="ok",
        message=f"{sc.report_name} three days: {counts.headline(days)}",
        output_path=str(out_path),
        summary_lines=[f"{sc.report_name} — {' → '.join(days)}: "
                       + counts.headline(days)],
        completion=counts.completion,
        producer_extra={SIDECAR_KEY: counts.to_dict()})
    if not consolidation_meta.write_outcome(
            out_path, result, extra=result.producer_extra,
            commit_guard=commit_guard):
        # The new values workbook has landed; a copy from the build before it
        # must not stay beside it.
        _remove_stale_copy(out_path, events, commit_guard)
        return ConsolidateResult(
            status="error",
            message="The three-day workbook was written but its counts could not "
                    "be published beside it — rebuild the comparison.")
    result.sidecar_published = True
    if also_formulas and _write_formulas_copy(sc, rows, compared, days, out_path,
                                              meta, events, commit_guard):
        result.summary_lines.append(
            f"Live-formulas copy: {formulas_sibling(out_path).name}")
    else:
        _remove_stale_copy(out_path, events, commit_guard)
    return result


def is_flat(adapter):
    """True for an adapter whose sides load as flat `[route, *row]` records (the
    per-route Excel exports and the flat print editions) — the shape this engine
    reads. The one-row-per-route summary adapters load a different shape."""
    return adapter.side_loader is None and adapter.sheet_name is not None


def _coverage(side, label):
    """(notes, incomplete) of one loaded side — the two-day comparison's own
    fail-closed reading of a loader's coverage claims
    (`compare_env._side_coverage_items`)."""
    notes = tuple(str(x) for x in side.warnings) + tuple(str(x) for x in side.failures)
    incomplete = (side.completion != "complete" or bool(notes)
                  or side.skipped_inputs > 0 or side.failed_inputs > 0)
    if incomplete and not notes:
        notes = (f"{label}: loaded-side completion was {side.completion!r}, not "
                 "complete.",)
    return notes, incomplete


def _error(message):
    return ConsolidateResult(status="error", message=message)


def compare_days(adapter, folders, days, out_path, *, events=None,
                 commit_guard=None, also_formulas=False):
    """Build the three-day workbook for one report from three run folders.

    `adapter` is the report's `compare_env.EnvCompare` — the vs Baseline matrix's
    adapter. Each day loads through the loader `compare_folders` itself resolves
    (`_side_loader_fn`) and records key and compare under the schema it builds
    (`_schema`), so this cannot key or read a file differently from the two-day
    comparison. `folders` / `days` are the three run folders and their dates,
    OLDEST FIRST. `also_formulas` also writes the live-formulas copy (see `run`).
    Returns a ConsolidateResult (ok / cancelled / error)."""
    events = events or Events()
    folders = [Path(f) for f in folders]
    days = tuple(days)
    name = adapter.REPORT_NAME
    if len(folders) != 3 or len(days) != 3:
        return _error("A three-day comparison needs three days.")
    if len(set(days)) != 3 or list(days) != sorted(days):
        return _error("Pick three different days, oldest first: "
                      + ", ".join(days) + ".")
    if not is_flat(adapter):
        return _error(f"{name} is a one-row-per-route summary — a three-day "
                      "comparison is not offered for it yet.")
    for day, folder in zip(days, folders):
        if not folder.is_dir():
            return _error(f"The {day} folder doesn't exist:\n{folder}")
    try:
        files = [adapter._effective_source_files(f) for f in folders]
        effective = [adapter._effective_input_dir(f) for f in folders]
        # Two days must not be one dataset: the same resolved folder, or (through
        # a junction or hardlinked tree) the same discovered files — the two-day
        # comparison's own guard, across all three.
        identities = [artifact_store.canonical_path_identities(found)
                      for found in files if found]
        if (len(set(effective)) != 3
                or any(a == b for i, a in enumerate(identities)
                       for b in identities[i + 1:])):
            return _error("Pick three DIFFERENT days — two of them resolve to "
                          f"the same {name} files.")
        # The copy's path is guarded whether or not this build writes it: an
        # earlier copy at that path is removed when it is not refreshed.
        artifact_store.ensure_outputs_do_not_alias_sources(
            (Path(out_path), formulas_sibling(out_path)),
            (*folders, *(p for found in files for p in found)))
    except ValueError as e:
        return _error(str(e))

    events.on_log("=" * 60)
    events.on_log(f"{name} — three days: {' → '.join(days)}")
    for day, folder in zip(days, folders):
        events.on_log(f"  {day}: {folder}")
    events.on_log("=" * 60)

    load_side = adapter._side_loader_fn()
    try:
        # A day with nothing exported refuses through its own loader BEFORE any
        # other day is parsed (a statewide print side takes minutes to read).
        for day, folder, found in zip(days, folders, files):
            if not found:
                load_side(folder, day, events)
        loaded = [load_side(folder, day, events)
                  for day, folder in zip(days, folders)]
    except ValueError as e:
        msg = str(e)
        if msg == "Cancelled by user.":
            return ConsolidateResult(status="cancelled", message=msg,
                                     completion=outcome.CANCELLED)
        return _error(msg)
    headers = [list(side.declared_schema) for side in loaded]
    if adapter.header_canonicalizer is not None:
        canon = [adapter.header_canonicalizer(h) for h in headers]
        for day, c in zip(days, canon):
            if c is None:
                return _error(f"The {day} {name} files do not use a recognized "
                              "column layout — re-export them with a supported "
                              "app version.")
        headers = [list(c) for c in canon]
    if not (headers[0] == headers[1] == headers[2]):
        return _error(f"The three days' {name} files have different column "
                      "layouts — a three-day comparison needs exports of the "
                      "same layout.")
    try:
        # The later day is the comparison's first side, as on the vs Baseline
        # matrix (the day vs its baseline). The side names only label a two-day
        # workbook; nothing here reads them.
        sc = adapter._schema(headers[0], days[1], days[0])
    except ValueError as e:
        return _error(str(e))
    if adapter.physical_key_builder is not None and sc.key_normalizer is None:
        return _error(f"The {name} files do not carry the columns this "
                      "comparison keys a record on, so the three days cannot be "
                      "reliably matched.")
    problem = schema_problem(sc)
    if problem:
        return _error(f"{name} cannot be compared across three days yet: "
                      f"{problem}.")

    rows = tuple(side.rows for side in loaded)
    for day, day_rows in zip(days, rows):
        if not day_rows:
            return _error(f"The {day} {name} files have no data rows — nothing "
                          "to compare.")
    warnings, incomplete = [], False
    for day, side in zip(days, loaded):
        notes, bad = _coverage(side, day)
        warnings.extend(notes)
        incomplete = incomplete or bad
    if incomplete:
        events.on_log("⚠ Incomplete input coverage — the comparison will be flagged "
                      "incomplete.")
    meta = {
        "inputs": tuple(
            f"{folder}  —  {len(found):,} file{'' if len(found) == 1 else 's'}, "
            f"{len(side.rows):,} {sc.id_noun_plural}"
            for folder, found, side in zip(folders, files, loaded)),
        "warnings": tuple(warnings),
    }
    if incomplete:
        meta["partial_reasons"] = tuple(warnings)
    return run(sc, rows, days, out_path, meta=meta, events=events,
               commit_guard=commit_guard,
               input_completion=outcome.PARTIAL if incomplete else None,
               also_formulas=also_formulas)


__all__ = ["BACK", "BOTH", "CHANGED", "CLASSES", "Compared", "MIDDLE", "NEWEST",
           "PAIRS", "PRESENCE", "SAME", "SIDECAR_KEY", "TWO_DAY", "ThreeDayCounts",
           "align", "class_label", "class_short", "classify", "compare_days",
           "compute", "counted_fields", "day_keys", "formulas_sibling", "is_flat",
           "one_day_state", "presence_label", "presence_of", "presence_short",
           "read_counts", "run", "schema_problem", "texts_of"]
