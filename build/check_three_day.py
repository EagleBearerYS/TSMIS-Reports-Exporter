"""build/check_three_day.py — the three-day comparison (v0.52.0): one report's
records on three export days, classified cell by cell.

Drives the SHIPPED path on synthetic data (no corpus files):

  * the five classes partition every cell, and a record's presence names the
    days it is on;
  * a hand fixture with every class and every presence: the published counts,
    every sheet of the workbook, the Changes and Added & removed lists, the
    grid's cell texts and links, formula-looking text kept as text, and the
    counts bound to the workbook (a rewrite reads as unknown);
  * THE SAME NUMBERS AS THE TWO-DAY COMPARISON — on random data, with and
    without repeated keys, each pair of days' records on both, records only one
    day has, records that differ, differing cells and per-column counts equal
    `compare_core.run_compare`'s own (counts-only preview), and every cell's
    class agrees with `compare_core.compared_cell`;
  * the schema gate (a non-plain equality rule is refused, never classified);
  * the edges: a formula-looking location on a record with several changed
    cells, an export value that reads like one of the workbook's own marks
    (disclosed, never silently ambiguous), a damaged counts record (unknown,
    never an exception), a value or a locked file the workbook cannot take (a
    result, not a crash), two days that are one hardlinked dataset (refused);
  * Clean Road Ramp end to end through the matrix: three day folders of
    per-route exports, the real cross-environment adapter, the cell states, the
    workbook in its store, the counts, and each pair again equal to the app's
    own two-day comparison of the same folders (`compare_folders`);
  * the independent raw-file check (`tools/three_day_check.py`, which shares
    none of the comparison's code) passes on that workbook and goes red when a
    raw file no longer says what the workbook says;
  * staleness when any of the three days' exports moves, when the results
    cache is lost, and when the workbook is rewritten; the scoped rebuild list;
  * the matrix store: the cell's file and cache record per source.

Run with the build venv:
    build\\.venv\\Scripts\\python.exe build\\check_three_day.py
"""
import contextlib
import io
import os
import random
import subprocess
import sys
import time
from pathlib import Path

from _checklib import ROOT, Checker, patch, pinned_site, scripts_path, temp_dir

scripts_path()
sys.path.insert(0, str(ROOT / "tools"))

from openpyxl import Workbook, load_workbook  # noqa: E402

import clean_road_columns as crc  # noqa: E402
import compare_core as core  # noqa: E402
import compare_env_editions as cee  # noqa: E402
import matrix  # noqa: E402
import outcome  # noqa: E402
import paths  # noqa: E402
import reports  # noqa: E402
import three_day as td  # noqa: E402
import three_day_check  # noqa: E402  (tools/: the independent raw-file check)
import three_day_matrix as tdm  # noqa: E402
import three_day_workbook as book  # noqa: E402
from events import Events  # noqa: E402

c = Checker()
DAYS = ("2026-10-30", "2026-11-03", "2026-11-06")
SRC = "ssor-prod dev-site"
SET = "_".join(DAYS)


# --------------------------------------------------------------------------- #
# 1. the classes
# --------------------------------------------------------------------------- #
def test_classes():
    print("the classes:")
    table = {("a", "a", "a"): td.SAME, ("a", "a", "b"): td.NEWEST,
             ("a", "b", "b"): td.MIDDLE, ("a", "b", "a"): td.BACK,
             ("a", "b", "c"): td.BOTH, ("", "", ""): td.SAME, ("", "", "x"): td.NEWEST}
    c.check("classify: same / newest / middle / back / both",
            all(td.classify(*k) == v for k, v in table.items()))
    presence = {(0, 0, 0): "all", (0, 0, None): "d12", (None, 0, 0): "d23",
                (0, None, 0): "d13", (0, None, None): "d1", (None, 0, None): "d2",
                (None, None, 0): "d3"}
    c.check("presence_of names the days a record is on",
            all(td.presence_of(k) == v for k, v in presence.items()))
    c.check("every label names its days",
            all(DAYS[2] in td.class_label(td.NEWEST, DAYS)
                and DAYS[1] in td.class_label(cls, DAYS)
                for cls in (td.MIDDLE, td.BACK, td.BOTH))
            and all(td.presence_label(p, DAYS) for p in td.PRESENCE))


# --------------------------------------------------------------------------- #
# 2. a hand fixture: every class, every presence, the workbook
# --------------------------------------------------------------------------- #
HEADER = ["PM", "A", "B", "X_EXTRACT_DATE"]


def _schema(header=HEADER, **extra):
    return core.CompareSchema(
        report_name="Demo", header=list(header), id_noun="segment",
        id_noun_plural="segments", key_field=0,
        context_fields=("X_EXTRACT_DATE",), **extra)


def _fixture_days():
    d = DAYS
    day1 = [["001", "0.0", "a", "1", d[0]],      # same (spaces only differ)
            ["001", "1.0", "a", "1", d[0]],      # A changed on the newest day
            ["001", "2.0", "a", "1", d[0]],      # A changed on the middle day, kept
            ["001", "3.0", "a", "1", d[0]],      # A changed and back
            ["001", "4.0", "a", "1", d[0]],      # A changed both times; B blank on day 3
            ["001", "5.0", "a", "1", d[0]],      # removed on day 3, B differs
            ["001", "6.0", "a", "1", d[0]],      # removed on day 2
            ["001", "7.0", "a", "1", d[0]],      # missing on day 2 only, A differs
            ["002", "0.0", "=1+1", "1", d[0]]]   # formula-looking text, unchanged
    day2 = [["001", "0.0", "a ", "1", d[1]],
            ["001", "1.0", "a", "1", d[1]],
            ["001", "2.0", "b", "1", d[1]],
            ["001", "3.0", "b", "1", d[1]],
            ["001", "4.0", "b", "1", d[1]],
            ["001", "5.0", "a", "2", d[1]],
            ["001", "8.0", "a", "1", d[1]],      # added on day 2
            ["001", "9.0", "a", "1", d[1]],      # only on day 2
            ["002", "0.0", "=1+1", "1", d[1]]]
    day3 = [["001", "0.0", " a", "1", d[2]],
            ["001", "1.0", "b", "1", d[2]],
            ["001", "2.0", "b", "1", d[2]],
            ["001", "3.0", "a", "1", d[2]],
            ["001", "4.0", "c", "", d[2]],
            ["001", "7.0", "z", "1", d[2]],
            ["001", "8.0", "a", "1", d[2]],
            ["001", "10.0", "a", "1", d[2]],     # added on day 3
            ["002", "0.0", "=1+1", "1", d[2]]]
    return day1, day2, day3


def _sheet_rows(wb, name):
    return [list(r) for r in wb[name].iter_rows(values_only=True)]


def test_fixture(tmp):
    print("a hand fixture — every class and every presence:")
    sc = _schema()
    out = tmp / "demo three days.xlsx"
    res = td.run(sc, _fixture_days(), DAYS, out, events=Events(),
                 meta={"inputs": ("one", "two", "three")})
    c.check("the comparison runs and writes its workbook",
            res.status == "ok" and out.exists() and res.completion == outcome.COMPLETE,
            res.message)
    counts = td.read_counts(out)
    c.check("the counts are published beside the workbook", counts is not None)
    if counts is None:
        return
    c.check("records by the days they are on",
            counts.records == {"all": 6, "d12": 1, "d23": 1, "d13": 1, "d1": 1,
                               "d2": 1, "d3": 1}, counts.records)
    c.check("compared cells by class (a trailing or leading space is not a change)",
            counts.cells == {td.NEWEST: 2, td.MIDDLE: 1, td.BACK: 1, td.BOTH: 1,
                             td.SAME: 7}, counts.cells)
    c.check("the context column is never counted",
            "X_EXTRACT_DATE" not in counts.columns and counts.counted_columns == 2)
    pairs = counts.pairs()
    want = {"1-2": (7, 2, 2, 4, 4), "2-3": (7, 2, 2, 3, 4), "1-3": (7, 2, 2, 4, 5)}
    got = {k: (p["paired"], p["only_earlier"], p["only_later"], p["differing_records"],
               p["differing_cells"]) for k, p in pairs.items()}
    c.check("each pair of days: records on both, only earlier, only later, records "
            "that differ, differing cells", got == want, got)
    c.check("per-column counts by class",
            counts.columns["A"] == {td.NEWEST: 1, td.MIDDLE: 1, td.BACK: 1, td.BOTH: 1,
                                    td.SAME: 2}
            and counts.columns["B"][td.NEWEST] == 1 and counts.columns["B"][td.SAME] == 5,
            counts.columns)
    c.check("added / removed / on or off the middle day only",
            (counts.added(), counts.removed(), counts.middle_only()) == (2, 2, 2))
    c.check("the published counts round-trip",
            td.ThreeDayCounts.from_dict(counts.to_dict()).to_dict() == counts.to_dict())
    c.check("a record of another version reads as unknown",
            td.ThreeDayCounts.from_dict({**counts.to_dict(), "version": 99}) is None)

    wb = load_workbook(out)
    c.check("sheets in reading order", tuple(wb.sheetnames) == book.SHEETS, wb.sheetnames)
    summary = {r[0]: r[1:] for r in _sheet_rows(wb, "Summary") if r and r[0]}
    c.check("Summary: the records and the classes carry the counts",
            summary[td.presence_label("all", DAYS)][0] == 6
            and summary[td.presence_label("d13", DAYS)][0] == 1
            and summary[td.class_label(td.NEWEST, DAYS)][0] == 2
            and summary[td.class_label(td.BACK, DAYS)][0] == 1
            and summary[td.class_label(td.SAME, DAYS)][0] == 7)
    c.check("Summary: each pair of days reads as its two-day comparison",
            summary[f"{DAYS[0]} vs {DAYS[1]}"][:5] == [7, 2, 2, 4, 4]
            and summary[f"{DAYS[1]} vs {DAYS[2]}"][:5] == [7, 2, 2, 3, 4]
            and summary[f"{DAYS[0]} vs {DAYS[2]}"][:5] == [7, 2, 2, 4, 5], summary)
    c.check("Summary: per-column classes", summary["A"][:6] == [1, 1, 1, 1, 2, 1]
            and summary["B"][:6] == [1, 0, 0, 0, 5, 1], (summary["A"], summary["B"]))

    changes = _sheet_rows(wb, "Changes")[2:]
    c.check("Changes: one row per changed cell, with its value on each day",
            [r[:7] for r in changes] == [
                ["001", "1.0", "A", "a", "a", "b", td.class_short(td.NEWEST, DAYS)],
                ["001", "2.0", "A", "a", "b", "b", td.class_short(td.MIDDLE, DAYS)],
                ["001", "3.0", "A", "a", "b", "a", td.class_short(td.BACK, DAYS)],
                ["001", "4.0", "A", "a", "b", "c", td.class_short(td.BOTH, DAYS)],
                ["001", "4.0", "B", "1", "1", "(blank)", td.class_short(td.NEWEST, DAYS)],
                ["001", "5.0", "B", "1", "2", "—", book.two_day_change("d12", DAYS)],
                ["001", "7.0", "A", "a", "—", "z", book.two_day_change("d13", DAYS)]],
            changes)
    grid = _sheet_rows(wb, book.GRID_SHEET)
    by_loc = {r[1]: r for r in grid[2:] if r[0] == "001"}
    c.check("Changes: the Grid row link points at that record's grid row",
            all(grid[int(str(r[7]).rsplit(",", 1)[1].rstrip(")")) - 1][1] == r[1]
                for r in changes), [r[7] for r in changes])
    records = _sheet_rows(wb, "Added & removed")[2:]
    c.check("Added & removed: one row per record not on all three days",
            sorted((r[1], r[2], r[3], r[4], r[5]) for r in records) == sorted([
                ("5.0", td.presence_label("d12", DAYS), "✓", "✓", "—"),
                ("6.0", td.presence_label("d1", DAYS), "✓", "—", "—"),
                ("7.0", td.presence_label("d13", DAYS), "✓", "—", "✓"),
                ("8.0", td.presence_label("d23", DAYS), "—", "✓", "✓"),
                ("9.0", td.presence_label("d2", DAYS), "—", "✓", "—"),
                ("10.0", td.presence_label("d3", DAYS), "—", "—", "✓")]), records)
    c.check("grid: a changed cell shows its three days in order",
            by_loc["1.0"][7] == "a → a → b" and by_loc["3.0"][7] == "a → b → a"
            and by_loc["4.0"][8] == "1 → 1 → (blank)" and by_loc["5.0"][8] == "1 → 2 → —"
            and by_loc["7.0"][7] == "a → — → z", by_loc)
    c.check("grid: an unchanged cell shows its value; the context column the newest day",
            by_loc["0.0"][7] == "a" and by_loc["0.0"][9] == DAYS[2]
            and by_loc["6.0"][9] == DAYS[0] and by_loc["5.0"][9] == DAYS[1])
    c.check("grid: the row tallies (changed cells, changed on each day)",
            by_loc["4.0"][3:6] == [2, 1, 2] and by_loc["3.0"][3:6] == [1, 1, 1]
            and by_loc["1.0"][3:6] == [1, 0, 1] and by_loc["5.0"][3:6] == [1, 1, None]
            and by_loc["7.0"][3:6] == [1, None, None], by_loc["4.0"])
    c.check("grid: records keep the middle day's order, the others slot in after",
            [r[1] for r in grid[2:] if r[0] == "001"] ==
            ["0.0", "1.0", "2.0", "3.0", "4.0", "5.0", "6.0", "7.0", "8.0", "10.0", "9.0"],
            [r[1] for r in grid[2:]])
    ws = wb[book.GRID_SHEET]
    formula_cell = next(cell for row in ws.iter_rows(min_row=3) for cell in row
                        if cell.value == "=1+1")
    c.check("formula-looking source text stays text", formula_cell.data_type == "s",
            formula_cell.data_type)
    c.check("grid: frozen header and a filter over every record",
            ws.freeze_panes == "C3" and ws.auto_filter.ref == f"A2:J{len(grid)}",
            (ws.freeze_panes, ws.auto_filter.ref))
    routes = {r[0]: r[1:] for r in _sheet_rows(wb, "Routes")[2:]}
    c.check("Routes: each route's records and changed cells",
            routes["001"] == [8, 8, 8, 5, 2, 2, 2, 2, 1, 1, 1, 2]
            and routes["002"] == [1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0], routes)
    wb.close()

    os.utime(out, (time.time() + 5, time.time() + 5))
    c.check("the counts are bound to the workbook (one rewritten after publication "
            "reads as unknown)", td.read_counts(out) is None)


# --------------------------------------------------------------------------- #
# 3. the same numbers as the two-day comparison
# --------------------------------------------------------------------------- #
P_HEADER = ["PM", "A", "B", "C", "D", "X_EXTRACT_DATE"]
_VALUES = ["", "0", "1", "2", " 2", "2 ", "a", "A", "a  b", "a b", 5, 5.0, 5.5, None]


def _rand_row(rng, route, pm, day):
    return [route, pm] + [rng.choice(_VALUES) for _ in range(4)] + [f"2026-0{day}"]


def _mutate(rng, row, day, p_edit):
    out = list(row)
    for i in range(2, 6):
        if rng.random() < p_edit:
            out[i] = rng.choice(_VALUES)
    out[6] = f"2026-0{day}"
    return out


def _key(row):
    return (row[0], core._key_text(row[1]))


def _random_days(rng, dup, n_routes=3, n_keys=8):
    """Three days of rows: edits, removals, additions, a few rows that come back
    on day 3; with `dup`, keys that repeat within a route, in shuffled order."""
    day1 = []
    for r in range(n_routes):
        route = f"{r + 1:03d}"
        for k in range(n_keys):
            pm = f"{k * 1.5:.1f}" if k % 3 else k        # text and numeric keys
            for _ in range(rng.choice([1, 1, 1, 2, 3]) if dup else 1):
                day1.append(_rand_row(rng, route, pm, 1))

    def step(prev, day):
        out = []
        for row in prev:
            x = rng.random()
            if x < 0.12:
                continue
            out.append(_mutate(rng, row, day, 0.25))
            if dup and x > 0.93:
                out.append(_mutate(rng, row, day, 0.6))
        for _ in range(rng.randint(0, 5)):
            row = _rand_row(rng, f"{rng.randint(1, n_routes + 1):03d}",
                            rng.choice([f"{rng.randint(0, 12) * 1.5:.1f}",
                                        rng.randint(0, 9)]), day)
            if dup or _key(row) not in {_key(r) for r in out}:
                out.append(row)
        out.sort(key=lambda r: r[0])                     # routes stay contiguous
        if dup and rng.random() < 0.5:
            by = {}
            for r in out:
                by.setdefault(r[0], []).append(r)
            out = []
            for route in sorted(by):
                rng.shuffle(by[route])
                out.extend(by[route])
        return out

    day2 = step(day1, 2)
    day3 = step(day2, 3)
    have = {_key(r) for r in day3}
    for r in day1:
        if rng.random() < 0.08 and (dup or _key(r) not in have):
            have.add(_key(r))
            day3.append(_mutate(rng, r, 3, 0.2))
    return day1, day2, sorted(day3, key=lambda r: r[0])


def _two_day(sc, later, earlier, tmp):
    res = core.run_compare(sc, later, earlier, True, tmp / "x.xlsx", events=Events(),
                           mode="preview")
    assert res.status == "ok", res.message
    return res.comparison_outcome.counts


def _pair_equal(sc, pair, counts2):
    """The pair's totals and per-column counts against a two-day comparison's
    (side A = the later day, side B = the earlier one)."""
    want = (counts2.paired_rows, counts2.side_b_only_rows, counts2.side_a_only_rows,
            counts2.differing_rows, counts2.differing_cells)
    got = (pair["paired"], pair["only_earlier"], pair["only_later"],
           pair["differing_records"], pair["differing_cells"])
    per = {name.split(":", 1)[1]: n for name, n in counts2.per_field_counts.items()}
    cols = all(pair["column_diffs"][col] == per.get(col, 0)
               for col in pair["column_diffs"])
    context_free = all(per.get(col, 0) == 0 for col in sc.context_fields)
    return got == want and cols and context_free, (got, want)


def _brute_classes(sc, rows, compared):
    """Every cell's class re-decided through compare_core.compared_cell."""
    counted = [f for f in sc.field_indices if not sc.is_context(f)]
    for rec, state in zip(compared.records, compared.states):
        if td.presence_of(rec) != "all":
            continue
        r1, r2, r3 = rows[0][rec[0]], rows[1][rec[1]], rows[2][rec[2]]
        for f in counted:
            e12 = core.compared_cell(sc, f, r2, r1, 1).verdict
            e23 = core.compared_cell(sc, f, r3, r2, 1).verdict
            e13 = core.compared_cell(sc, f, r3, r1, 1).verdict
            cls = (td.SAME if e12 and e23 else td.NEWEST if e12 else td.MIDDLE if e23
                   else td.BACK if e13 else td.BOTH)
            if None in (e12, e23, e13) or state[f] != td.CODE[cls]:
                return False
            if (cls == td.SAME) != bool(e12 and e23 and e13):
                return False
    return True


def test_two_day_equivalence(tmp):
    print("the same numbers as the two-day comparison (random data):")
    sc = core.CompareSchema(report_name="Demo", header=list(P_HEADER), id_noun="row",
                            id_noun_plural="rows", key_field=0,
                            context_fields=("X_EXTRACT_DATE",))
    bad = {"partition": [], "classes": [], "1-2": [], "2-3": [], "1-3": [],
           "middle": [], "nodup": []}
    seen_class = dict.fromkeys(td.CLASSES, 0)
    seen_presence = dict.fromkeys(td.PRESENCE, 0)
    diverged = runs = 0
    for seed in range(60):
        for dup in (False, True):
            rows = _random_days(random.Random(seed * 2 + dup), dup)
            if not all(rows):
                continue
            runs += 1
            tag = (seed, dup)
            compared = td.compute(sc, rows)
            counts = compared.counts
            for day in range(3):
                used = sorted(r[day] for r in compared.records if r[day] is not None)
                if used != list(range(len(rows[day]))):
                    bad["partition"].append(tag)
            if sum(counts.cells.values()) != counts.records["all"] * counts.counted_columns:
                bad["partition"].append(tag)
            if not _brute_classes(sc, rows, compared):
                bad["classes"].append(tag)
            pairs, thru = counts.pairs(), counts.through_middle()
            for pair, (x, y) in td.PAIRS.items():
                ok, detail = _pair_equal(sc, pairs[pair],
                                         _two_day(sc, rows[y], rows[x], tmp))
                if not ok:
                    bad[pair].append((tag, detail))
            if thru["1-2"] != pairs["1-2"] or thru["2-3"] != pairs["2-3"]:
                bad["middle"].append(tag)
            if not dup and (thru["1-3"] != pairs["1-3"] or counts.first_last_differs()):
                bad["nodup"].append(tag)
            diverged += bool(dup and counts.first_last_differs())
            for cls in td.CLASSES:
                seen_class[cls] += counts.cells[cls]
            for p in td.PRESENCE:
                seen_presence[p] += counts.records[p]
    c.check(f"{runs} random datasets exercised every class and every presence",
            runs >= 100 and all(seen_class.values()) and all(seen_presence.values()),
            (runs, seen_class, seen_presence))
    c.check("every day's rows land in exactly one record; the classes partition the "
            "compared cells", not bad["partition"], bad["partition"][:3])
    c.check("every cell's class agrees with compare_core.compared_cell",
            not bad["classes"], bad["classes"][:3])
    c.check("day 1 vs day 2 equals the two-day comparison (totals and per column)",
            not bad["1-2"], bad["1-2"][:2])
    c.check("day 2 vs day 3 equals the two-day comparison (totals and per column)",
            not bad["2-3"], bad["2-3"][:2])
    c.check("day 1 vs day 3 equals the two-day comparison (totals and per column)",
            not bad["1-3"], bad["1-3"][:2])
    c.check("the two consecutive pairs are read straight off the classes",
            not bad["middle"], bad["middle"][:3])
    c.check("without a repeated key, first-vs-last through the middle day IS the "
            "direct comparison", not bad["nodup"], bad["nodup"][:3])
    c.check("with repeated keys the two readings can part, and the engine says so",
            diverged > 0, diverged)


# --------------------------------------------------------------------------- #
# 4. the gate
# --------------------------------------------------------------------------- #
def test_gate(tmp):
    print("the gate:")
    rows = _fixture_days()
    plain = _schema()
    c.check("a plain schema passes", td.schema_problem(plain) == "")
    for name, sc in (("ditto markers", _schema(ditto_nonasserting=True)),
                     ("a Med Wid column", _schema(medwid_fields=("A",))),
                     ("an unavailable-anchor token", _schema(unavailable_rule=("N/A", "n"))),
                     ("two columns of one name", _schema(header=["PM", "A", "A", "B"]))):
        refused = bool(td.schema_problem(sc))
        res = td.run(sc, rows, DAYS, tmp / "refused.xlsx", events=Events())
        c.check(f"{name}: refused, nothing written",
                refused and res.status == "error"
                and not (tmp / "refused.xlsx").exists(), res.message)
    ramp = cee.CLEAN_RAMP
    folders = [tmp / d for d in DAYS]
    for f in folders:
        f.mkdir()
    out = tmp / "gate.xlsx"
    c.check("days out of order are refused",
            td.compare_days(ramp, folders, (DAYS[1], DAYS[0], DAYS[2]), out).status
            == "error")
    c.check("a repeated day is refused",
            td.compare_days(ramp, folders, (DAYS[0], DAYS[0], DAYS[2]), out).status
            == "error")
    c.check("a missing folder is refused",
            td.compare_days(ramp, [folders[0], folders[1], tmp / "nope"], DAYS,
                            out).status == "error")
    c.check("the same folder twice is refused",
            td.compare_days(ramp, [folders[0], folders[0], folders[2]], DAYS,
                            out).status == "error")
    summary = {k: a for k, _l, _s, _i, a in reports.matrix_rows()}["ramp_summary"]
    c.check("a one-row-per-route summary report is refused (it loads another shape)",
            not td.is_flat(summary)
            and td.compare_days(summary, folders, DAYS, out).status == "error")
    res = td.compare_days(ramp, folders, DAYS, out)
    c.check("a day with nothing exported refuses through its own loader",
            res.status == "error" and "No Clean Road Ramp files" in res.message,
            res.message)
    c.check("no refusal wrote a workbook", not out.exists())


# --------------------------------------------------------------------------- #
# 4b. the edges
# --------------------------------------------------------------------------- #
def _make_dir_link(link, target):
    """A directory link `link` -> `target`: a symlink, else (Windows, no admin
    needed) a junction. False where neither can be made."""
    try:
        os.symlink(target, link, target_is_directory=True)
        return True
    except (OSError, NotImplementedError):
        if os.name != "nt":
            return False
    made = subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(target)],
                          capture_output=True, text=True, check=False)
    return made.returncode == 0 and link.exists()


def test_edges(tmp):
    print("the edges:")
    d = DAYS
    sc = _schema()

    # a formula-looking LOCATION on a record with two changed cells: each row of
    # the Changes sheet needs its own cell (one guarded cell cannot sit in two)
    days = ([["001", "-5", "a", "1", d[0]], ["001", "@x", "=A1", "1", d[0]]],
            [["001", "-5", "b", "2", d[1]], ["001", "@x", "=A2", "2", d[1]]],
            [["001", "-5", "b", "2", d[2]], ["001", "@x", "=A2", "3", d[2]]])
    out = tmp / "guarded.xlsx"
    res = td.run(sc, days, DAYS, out, events=Events())
    wb = load_workbook(out) if res.status == "ok" else None
    rows = _sheet_rows(wb, "Changes")[2:] if wb else []
    c.check("a formula-looking location with several changed cells: the workbook "
            "opens and every Changes row carries its location",
            res.status == "ok" and [r[:6] for r in rows] == [
                ["001", "-5", "A", "a", "b", "b"], ["001", "-5", "B", "1", "2", "2"],
                ["001", "@x", "A", "=A1", "=A2", "=A2"],
                ["001", "@x", "B", "1", "2", "3"]], rows or res.message)
    if wb:
        texts = [cell for row in wb["Changes"].iter_rows(min_row=3) for cell in row[:6]]
        c.check("...and none of that text became a formula",
                all(cell.data_type == "s" for cell in texts),
                [(cell.coordinate, cell.data_type) for cell in texts
                 if cell.data_type != "s"])
        wb.close()

    # an export value that reads like one of the workbook's own marks
    days = ([["001", "1", "(blank)", "x", d[0]], ["001", "2", "a → b", "—", d[0]]],
            [["001", "1", "", "x", d[1]], ["001", "2", "a → b", "—", d[1]]],
            [["001", "1", "(blank)", "x", d[2]], ["001", "2", "a → b", "y", d[2]]])
    compared = td.compute(sc, days)
    c.check("a value that reads like a mark is still compared as the text it is",
            compared.counts.cells[td.BACK] == 1 and compared.counts.cells[td.NEWEST] == 1
            and compared.lookalikes == ("(blank)", "a → b", "—"),
            (compared.counts.cells, compared.lookalikes))
    out = tmp / "marks.xlsx"
    res = td.run(sc, days, DAYS, out, events=Events())
    said = " ".join(str(r[0]) for r in _sheet_rows(load_workbook(out), "Summary")
                    if r and r[0])
    c.check("...and the Summary says which values, so no cell misleads in silence",
            res.status == "ok" and "read like this workbook's own marks" in said
            and "'(blank)'" in said and "'a → b'" in said, said[:500])
    plain = tmp / "plain.xlsx"
    td.run(sc, _fixture_days(), DAYS, plain, events=Events())
    c.check("an ordinary export carries no such warning",
            "own marks" not in " ".join(
                str(r[0]) for r in _sheet_rows(load_workbook(plain), "Summary")
                if r and r[0]))

    # a damaged counts record reads as unknown and never raises
    good = td.read_counts(plain).to_dict()
    damaged = {
        "counted_columns a word": {**good, "counted_columns": "x"},
        "counted_columns not the columns": {**good, "counted_columns": 99},
        "no columns": {k: v for k, v in good.items() if k != "columns"},
        "no two_day_columns": {k: v for k, v in good.items() if k != "two_day_columns"},
        "a column gone": {**good, "columns": {"A": good["columns"]["A"]}},
        "a class missing from a column": {
            **good, "columns": {**good["columns"], "A": {td.SAME: 1}}},
        "a column count that is not a number": {
            **good, "columns": {**good["columns"],
                                "A": {**good["columns"]["A"], td.SAME: "x"}}},
        "a two-day column count that is not a number": {
            **good, "two_day_columns": {**good["two_day_columns"], "d12": {"A": "x"}}},
        "first-vs-last column counts that are not numbers": {
            **good, "first_last": {**good["first_last"], "column_diffs": {"A": None}}},
        "a negative count": {**good, "cells": {**good["cells"], td.SAME: -1}},
        "an unknown pairing quality": {**good, "pairing_quality": "fine"},
        "a completion that never publishes": {**good, "completion": "failed"},
        "first_last without its counts": {**good, "first_last": {"paired": 1}},
        "not a mapping": [1, 2, 3],
    }
    bad = []
    for name, record in damaged.items():
        try:
            if td.ThreeDayCounts.from_dict(record) is not None:
                bad.append(f"{name}: accepted")
        except Exception as e:                       # noqa: BLE001
            bad.append(f"{name}: raised {type(e).__name__}")
    c.check(f"{len(damaged)} damaged counts records each read as unknown",
            not bad and td.ThreeDayCounts.from_dict(good) is not None, bad)

    # what the workbook cannot take is a result, not a crash
    ctrl = ([["001", "1", "a\x0bb", "1", d[0]]], [["001", "1", "a", "1", d[1]]],
            [["001", "1", "a", "1", d[2]]])
    out = tmp / "control.xlsx"
    res = td.run(sc, ctrl, DAYS, out, events=Events())
    c.check("a control character Excel cannot store: an error naming it, no workbook",
            res.status == "error" and "control character" in res.message
            and not out.exists(), res.message)

    def locked(*_a, **_k):
        raise PermissionError(13, "Permission denied")

    out = tmp / "locked.xlsx"
    with patch(book.artifact_store, "atomic_save_if", locked):
        res = td.run(sc, _fixture_days(), DAYS, out, events=Events())
    c.check("a workbook that cannot be written (open in Excel): an error saying so",
            res.status == "error" and "Excel" in res.message and not out.exists(),
            res.message)

    # two days that are one dataset under two names (a link to the other day)
    folders = _write_days(tmp / "linked")
    alias = tmp / "linked" / "alias"
    if _make_dir_link(alias, folders[0]):
        try:
            res = td.compare_days(cee.CLEAN_RAMP, [folders[0], alias, folders[2]],
                                  DAYS, tmp / "linked.xlsx")
            c.check("a day that is a link to another day's folder is refused, not "
                    "compared as 'no change'",
                    res.status == "error" and "DIFFERENT" in res.message
                    and not (tmp / "linked.xlsx").exists(), res.message)
        finally:
            if alias.is_symlink():
                alias.unlink()
            else:
                alias.rmdir()
    else:
        print("  (skipped: this volume cannot make a directory link)")


# --------------------------------------------------------------------------- #
# 5. Clean Road Ramp end to end, through the matrix
# --------------------------------------------------------------------------- #
SPEC = crc.RAMP


def _ramp(pm, desc, design="D", adt=100, day=DAYS[0], route="001"):
    row = dict.fromkeys(SPEC.header, "")
    row.update({"RAM_ROUTE_NAME": route, "RAM_ROUTE_SUFFIX_CODE": ".",
                "RAM_COUNTY_CODE": "ORA.", "RAM_PM_PREFIX_CODE": ".",
                "RAM_PM_LOC_AMT": pm, "RAM_PM_SUFFIX_CODE": ".",
                "RAM_DESCRIPTION": desc, "RAM_DESIGN_CODE": design, "RAM_ADT": adt,
                "RAM_DISTRICT_CODE": 12, "RAM_ON_OFF_CODE": "0",
                "RAM_EXTRACT_DATE": day})
    return row


def _write_route(folder, run, route, rows):
    folder.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = SPEC.sheet
    ws.append(list(SPEC.header))
    for r in rows:
        ws.append([r[col] for col in SPEC.header])
    path = folder / f"{run} {SPEC.key}_route_{route}.xlsx"
    wb.save(path)
    return path


def _ramp_days():
    """Per day: {route: rows}. Route 005 carries a key that repeats, its two rows
    in the opposite order on day 2 — duplicate pairing must see them unchanged."""
    d1, d2, d3 = DAYS
    twin_a, twin_b = "SB ON FR MAIN ST", "SB OFF TO MAIN ST"
    return (
        {"001": [_ramp(0.606, "NB OFF TO A ST", day=d1),
                 _ramp(0.9, "NB ON FR B ST", day=d1),
                 _ramp(1.2, "NB OFF TO C ST", day=d1),
                 _ramp(1.5, "NB ON FR D ST", day=d1),
                 _ramp(1.8, "NB OFF TO E ST", day=d1),
                 _ramp(2.1, "NB ON FR F ST", day=d1),
                 _ramp(2.4, "NB OFF TO G ST", day=d1),
                 _ramp(2.7, "NB ON FR H ST", day=d1)],
         "005": [_ramp(3.0, twin_a, day=d1, route="005"),
                 _ramp(3.0, twin_b, day=d1, route="005")]},
        {"001": [_ramp(0.606, "NB OFF TO A ST", day=d2),
                 _ramp(0.9, "NB ON FR B ST", day=d2),
                 _ramp(1.2, "NB OFF TO C STREET", day=d2),           # middle, kept
                 _ramp(1.5, "NB ON FR D ST", adt=250, day=d2),       # back
                 _ramp(1.8, "NB OFF TO E AVE", day=d2),              # both
                 _ramp(2.1, "NB ON FR F ST", day=d2),
                 _ramp(3.3, "NB OFF TO J ST", day=d2),               # added day 2
                 _ramp(3.6, "NB ON FR K ST", day=d2)],               # only day 2
         "005": [_ramp(3.0, twin_b, day=d2, route="005"),
                 _ramp(3.0, twin_a, day=d2, route="005")]},
        {"001": [_ramp(0.606, "NB OFF TO A ST", day=d3),
                 _ramp(0.9, "NB ON FR B ST", design="Z", day=d3),    # newest
                 _ramp(1.2, "NB OFF TO C STREET", day=d3),
                 _ramp(1.5, "NB ON FR D ST", day=d3),
                 _ramp(1.8, "NB OFF TO E BLVD", day=d3),
                 _ramp(2.7, "NB ON FR H STREET", day=d3),            # missing day 2 only
                 _ramp(3.3, "NB OFF TO J ST", day=d3),
                 _ramp(3.9, "NB ON FR L ST", day=d3)],               # added day 3
         "005": [_ramp(3.0, twin_a, day=d3, route="005"),
                 _ramp(3.0, twin_b, day=d3, route="005")]})


def _write_days(root):
    folders = []
    for day, routes in zip(DAYS, _ramp_days()):
        run = f"{day} {SRC}"
        for route, rows in routes.items():
            _write_route(root / run / SPEC.key, run, route, rows)
        folders.append(root / run)
    return folders


def _cell(snap, row="clean_ramp", sid=SET):
    return snap["cells"][row][sid]["cmp"]


def _snap():
    return tdm.matrix_snapshot(SRC, [SET], today="2099-01-01")


def test_matrix(tmp):
    print("Clean Road Ramp end to end, through the Three Days matrix:")
    with patch(paths, "OUTPUT_ROOT", tmp / "output"):
        folders = _write_days(paths.OUTPUT_ROOT)
        c.check("rows: the six Clean Road editions, Excel then PDF per report",
                [r.key for r in tdm._rows()] == list(tdm.ROW_KEYS))
        c.check("a set of days: any order in, oldest first out; bad ones refused",
                tdm.make_set([DAYS[2], DAYS[0], DAYS[1]]) == DAYS
                and tdm.parse_set(SET) == DAYS
                and tdm.parse_set("_".join(reversed(DAYS))) is None
                and tdm.make_set([DAYS[0], DAYS[0], DAYS[2]]) is None
                and tdm.parse_set("2026-13-01_2026-11-03_2026-11-06") is None
                and tdm.clean_sets([SET, "junk", SET]) == [SET])
        c.check("the day pickers offer the three exported days, newest first",
                tdm.available_days(SRC) == list(reversed(DAYS))
                and tdm.available_day_reports(SRC)[DAYS[0]] == ["CR-RMP"],
                tdm.available_days(SRC))
        snap = _snap()
        cell = _cell(snap)
        c.check("a cell with all three days exported is buildable, not built",
                matrix.cell_buildable(cell) and not cell["built"]
                and cell["missing_days"] == [], cell)
        pdf = _cell(snap, "clean_ramp_pdf")
        c.check("an edition exported on none of the days names every missing day",
                pdf["missing_side"] == "day1" and pdf["missing_days"] == list(DAYS)
                and not matrix.cell_buildable(pdf), pdf)

        res = tdm.build_cell(SRC, SET, "clean_ramp", Events())
        out = tdm.set_out_path(DAYS, SRC, "clean_ramp")
        c.check("build_cell compares the three days", res.status == "ok"
                and res.completion == outcome.COMPLETE, res.message)
        c.check("...into the three-day store, filed under its newest day, under a "
                "name that carries all three days and the source",
                out.exists() and out.parent.parent.name == "three-days"
                and out.parent.name == DAYS[2]
                and out.name == f"clean_ramp_three_days {DAYS[0]} {DAYS[1]} "
                                f"{DAYS[2]} {SRC}.xlsx", out)
        other = "ssor-prod main-site"
        c.check("the same days read as another source are another file and record",
                tdm.set_out_path(DAYS, other, "clean_ramp") != out
                and tdm._result_key(DAYS, other, "clean_ramp")
                != tdm._result_key(DAYS, SRC, "clean_ramp"))
        c.check("the cell's sidecar path stays short (the longest row and a year of "
                "days: under 140 characters below output/comparisons/)",
                len(str(tdm.set_out_path(DAYS, "ssor-prod other-site",
                                         "clean_intersection_pdf")
                        .relative_to(tdm.td_root().parent))) + len("_state/")
                + len(".outcome.json") < 140)
        counts = td.read_counts(out)
        c.check("records: seven on all three days (the repeated key pairs by "
                "likeness, not by file order), one of each other presence",
                counts is not None and counts.records == {
                    "all": 7, "d12": 1, "d23": 1, "d13": 1, "d1": 1, "d2": 1, "d3": 1},
                counts and counts.records)
        c.check("cells: one changed on the newest day, one on the middle day and "
                "kept, one changed and back, one changed both times",
                counts.cells[td.NEWEST] == 1 and counts.cells[td.MIDDLE] == 1
                and counts.cells[td.BACK] == 1 and counts.cells[td.BOTH] == 1,
                counts.cells)
        c.check("the changed columns are the ones the fixture edited",
                counts.columns["RAM_DESIGN_CODE"][td.NEWEST] == 1
                and counts.columns["RAM_DESCRIPTION"][td.MIDDLE] == 1
                and counts.columns["RAM_ADT"][td.BACK] == 1
                and counts.columns["RAM_DESCRIPTION"][td.BOTH] == 1
                and "RAM_EXTRACT_DATE" not in counts.columns
                and counts.counted_columns == len(SPEC.header) - 2, counts.columns)

        # the app's own two-day comparison of the same folders, every pair
        adapter = cee.CLEAN_RAMP
        pairs = counts.pairs()
        for pair, (x, y) in td.PAIRS.items():
            two = adapter.compare_folders(folders[y], folders[x], tmp / "two.xlsx",
                                          events=Events(), mode="preview")
            ok, detail = (False, two.message) if two.status != "ok" else _pair_equal(
                adapter.base_schema, pairs[pair], two.comparison_outcome.counts)
            c.check(f"{DAYS[x]} vs {DAYS[y]} equals the vs Baseline comparison of "
                    "those two folders", ok, detail)

        wb = load_workbook(out, read_only=True)
        changes = [list(r) for r in wb["Changes"].iter_rows(values_only=True)][2:]
        c.check("Changes lists the five changed cells with their three values",
                sorted((r[2], r[3], r[4], r[5]) for r in changes) == sorted([
                    ("RAM_DESIGN_CODE", "D", "D", "Z"),
                    ("RAM_DESCRIPTION", "NB OFF TO C ST", "NB OFF TO C STREET",
                     "NB OFF TO C STREET"),
                    ("RAM_ADT", "100", "250", "100"),
                    ("RAM_DESCRIPTION", "NB OFF TO E ST", "NB OFF TO E AVE",
                     "NB OFF TO E BLVD"),
                    ("RAM_DESCRIPTION", "NB ON FR H ST", "—", "NB ON FR H STREET")]),
                changes)
        wb.close()

        raw_folders = [str(f / SPEC.key) for f in folders]

        def raw_check():
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                code = three_day_check.main([str(out), *raw_folders, "--spot", "3"])
            return code, buf.getvalue()

        code, said = raw_check()
        c.check("the independent raw-file check passes: every raw row once, the "
                "Summary and both lists re-derived", code == 0
                and "ALL CHECKS PASSED" in said and said.count("[OK ]") >= 16
                and "[FAIL]" not in said, said[-1200:])
        c.check("...and it tells the forced records from the repeated key's twins",
                "11 records at a location with one row a day" in said
                and "2 records at a location that repeats" in said,
                [ln for ln in said.splitlines() if "records at a location" in ln])
        victim = folders[0] / SPEC.key / f"{DAYS[0]} {SRC} {SPEC.key}_route_001.xlsx"
        original, stamp = victim.read_bytes(), victim.stat().st_mtime
        wb_raw = load_workbook(victim)
        wb_raw.active.cell(row=2, column=SPEC.header.index("RAM_DESCRIPTION") + 1,
                           value="NOT WHAT THE WORKBOOK SAYS")
        wb_raw.save(victim)
        code, said = raw_check()
        c.check("...and goes red when a raw file no longer says what the workbook says",
                code == 1 and "raw rows not in the grid" in said, said[-600:])
        wb_raw = load_workbook(victim)
        wb_raw.active.cell(row=2, column=SPEC.header.index("RAM_CREATE_USER_NAME") + 1,
                           value="(blank)")
        wb_raw.save(victim)
        code, said = raw_check()
        c.check("...and will not vouch for an export value that reads like a mark",
                code == 1 and "reads like one of the workbook's own marks" in said,
                said[:700])
        victim.write_bytes(original)
        os.utime(victim, (stamp, stamp))

        snap = _snap()
        cell = _cell(snap)
        c.check("the cell reads fresh with the published counts",
                cell["built"] and not cell["stale"] and cell["reason"] == "fresh"
                and cell["three_day"]["newest"] == 1 and cell["three_day"]["middle"] == 1
                and cell["three_day"]["added"] == 2 and cell["three_day"]["removed"] == 2
                and cell["three_day"]["middle_only"] == 2
                and cell["three_day"]["pairs"]["1-2"]["differing_cells"] == 3, cell)
        c.check("nothing to rebuild while it is fresh; 'all' still lists the cell",
                tdm.cells_to_rebuild(snap) == []
                and tdm.cells_to_rebuild(snap, scope="all") == [(SET, "clean_ramp")])

        for i, day in enumerate(DAYS):
            run = f"{day} {SRC}"
            added = _write_route(paths.OUTPUT_ROOT / run / SPEC.key, run, "099",
                                 [_ramp(9.9, "NEW", day=day, route="099")])
            os.utime(added, (time.time() + 5, time.time() + 5))
            cell = _cell(_snap())
            c.check(f"a new route in the {day} export reads the cell stale",
                    cell["stale"] and cell["reason"] in (f"day{i + 1}_newer",
                                                         "inputs_changed"), cell)
            added.unlink()
            cell = _cell(_snap())
            c.check(f"...and fresh again once the {day} folder is as it was",
                    not cell["stale"], cell)
        gone = paths.OUTPUT_ROOT / f"{DAYS[1]} {SRC}" / SPEC.key / \
            f"{DAYS[1]} {SRC} {SPEC.key}_route_005.xlsx"
        kept = gone.read_bytes()
        stamp = gone.stat().st_mtime
        gone.unlink()
        cell = _cell(_snap())
        c.check("a route DELETED from the middle day reads the cell stale (no file "
                "is newer — the folder's fingerprint moved)",
                cell["stale"] and cell["reason"] == "inputs_changed", cell)
        c.check("the scoped rebuild list picks the stale cell",
                tdm.cells_to_rebuild(_snap()) == [(SET, "clean_ramp")])
        gone.write_bytes(kept)
        os.utime(gone, (stamp, stamp))

        results = tdm._results_path()
        saved = results.read_bytes()
        results.unlink()
        cell = _cell(_snap())
        c.check("a lost results cache reads the cell stale, never fresh",
                cell["stale"] and cell["reason"] == "cache_missing_or_mismatched", cell)
        results.write_bytes(saved)
        c.check("...and fresh with the cache back", not _cell(_snap())["stale"])

        res = tdm.build_cell(SRC, SET, "clean_ramp", Events())
        c.check("rebuilding reads fresh again",
                res.status == "ok" and not _cell(_snap())["stale"])
        os.utime(out, (time.time() + 5, time.time() + 5))
        cell = _cell(_snap())
        c.check("a workbook rewritten after publication reads stale with no counts",
                cell["stale"] and "three_day" not in cell, cell)

        try:
            tdm.build_cell("not a source", SET, "clean_ramp", Events())
            refused = False
        except ValueError:
            refused = True
        c.check("build_cell refuses something that is not a source", refused)
        res = tdm.build_cell("ssor-prod", SET, "clean_ramp", Events())
        c.check("a bare source builds the SAME cell as its explicit site form",
                res.status == "ok" and not _cell(_snap())["stale"]
                and [f.name for f in out.parent.glob("*.xlsx")] == [out.name],
                [f.name for f in out.parent.glob("*.xlsx")])
        for bad, why in ((("clean_ramp_pdf", SET), "an edition with no export"),
                         (("clean_ramp", "2026-10-30_2026-11-03"), "two days"),
                         (("highway_log", SET), "a report the matrix does not offer")):
            try:
                tdm.build_cell(SRC, bad[1], bad[0], Events())
                refused = False
            except ValueError:
                refused = True
            c.check(f"build_cell refuses {why}", refused)


def main():
    with pinned_site("dev"), temp_dir("tsmis_td_") as tmp:
        tmp = Path(tmp)
        test_classes()
        for name, test in (("fx", test_fixture), ("eq", test_two_day_equivalence),
                           ("gate", test_gate), ("edge", test_edges),
                           ("mx", test_matrix)):
            sub = tmp / name
            sub.mkdir()
            test(sub)
    return c.summary()


if __name__ == "__main__":
    raise SystemExit(main())
