"""Check a three-day comparison workbook (Compare ▸ Three Days Matrix) against the
raw per-route export files directly, the way a person checks it by hand: open
the three days' files, find the row at a location on each day, and read the
cell. It shares none of the comparison's loading, keying, pairing, classing or
counting code — it reads the files with openpyxl and the workbook's own sheets —
so it can say whether the workbook tells the truth about the files.

What it proves, for the Clean Road reports (Highway, Intersection, Ramp):

  1. every raw row of each day is in the workbook's grid exactly once, with its
     compared values unaltered (and no grid row that is in no file);
  2. a location whose key is on one row a day is ONE grid record (no pairing
     was missed), and its key and shown-only columns are the newest day's own;
  3. the Summary re-decided here from the grid's values: the cells by class and
     per column, the records by the days they are on, and each pair of days
     (records on both, only on each day, records that differ, differing cells);
  4. the Routes sheet, the grid's own tallies and Location text, the Changes
     sheet (exactly the changed cells, with the same three values) and Added &
     removed (exactly the records not on all three days, with their differing
     cells), each list row naming the record its link points at.

A location whose key repeats within a day (two rows at one postmile) is still
fully accounted for by (1); WHICH of its twins pairs with which is the
comparison's choice (the rows most alike pair), and is reported as a count.

It fails closed on what it cannot read: a route file the comparison skipped
(the workbook is then partial anyway) shows as raw rows missing from the grid,
and an export value that reads like one of the workbook's own marks ("(blank)",
"—", " → ") is reported rather than guessed at. It reads the per-route Excel
exports; a print edition has no raw rows to read, so its workbook rests on the
comparison's own checks.

Usage (from the repo root):

  python tools/three_day_check.py <three-day workbook.xlsx>
         <day 1 folder> <day 2 folder> <day 3 folder> [--spot 12] [--seed 1]
         [--out report.txt]

Each day folder is that day's per-route export folder of the report, oldest
first (e.g. "output/2026-10-30 ssor-prod dev-site/clean_highway"). `--spot N`
also prints N random changed cells with, for each day, the file, the Excel row
and the raw cell, so a person can open the three files and see for themselves.
Exit code 0 when every check passes.
"""
import argparse
import collections
import datetime as dt
import hashlib
import random
import re
import sys
from pathlib import Path

import openpyxl

# The identity columns of each report — names only (route comes from the file
# name): county, postmile prefix, postmile, postmile suffix, odometer.
KEY_COLUMNS = {
    "THY_": ("THY_COUNTY_CODE", "THY_PM_PREFIX_CODE", "THY_BEGIN_PM_AMT",
             "THY_PM_SUFFIX_CODE", "THY_BEGIN_OFFSET_AMT"),
    "INX_": ("INX_COUNTY_CODE", "INX_PM_PREFIX_CODE", "INX_BEGIN_PM_AMT",
             "INX_PM_SUFFIX_CODE", ""),
    "RAM_": ("RAM_COUNTY_CODE", "RAM_PM_PREFIX_CODE", "RAM_PM_LOC_AMT",
             "RAM_PM_SUFFIX_CODE", "RAM_BEGIN_OFFSET_AMT"),
}
ARROW = " → "
ABSENT = "—"
BLANK = "(blank)"
TICK = "✓"
GRID = "Three days"
ALL_THREE = (True, True, True)
PRESENCES = (ALL_THREE, (True, True, False), (False, True, True), (True, False, True),
             (True, False, False), (False, True, False), (False, False, True))
PAIRS = ((0, 1), (1, 2), (0, 2))
_ROUTE = re.compile(r"_route_(\w+)\.xlsx$", re.IGNORECASE)
_LINK = re.compile(r"!(\d+):\d+\"")
_DAY = re.compile(r"\d{4}-\d{2}-\d{2}")
_PAIR_ROW = re.compile(r"^(\d{4}-\d{2}-\d{2}) vs (\d{4}-\d{2}-\d{2})$")
_RECORD_ROWS = ("On all three", "Removed on", "Added on", "Missing on", "Only on")
# The Summary's note when first-vs-last reads differently through the middle day.
_THROUGH = re.compile(r"comes to ([\d,]+) .*? on both and ([\d,]+) differing cells")


def text(v):
    """A cell as the text two exports are compared on: dates as ISO, a whole
    float as its integer, edge spaces dropped and runs of spaces read as one."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dt.datetime):
        return (v.date().isoformat() if v.time() == dt.time(0)
                else v.isoformat(sep=" "))
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(" +", " ", str(v)).strip(" ")


def number(s):
    """One spelling for a number in any spelling ('009.600' = '9.6' = 9.6)."""
    try:
        f = float(s)
    except ValueError:
        return s
    out = f"{f:.10f}".rstrip("0").rstrip(".")
    return "0" if out in ("", "-0") else out


def route_of(file_name):
    """The route a per-route file is for: '..._route_5s.xlsx' -> '005S'."""
    token = _ROUTE.search(file_name).group(1)
    m = re.match(r"^(\d+)(\w*)$", token)
    return m.group(1).zfill(3) + m.group(2).upper() if m else token.upper()


def location(route, county, prefix, pm, suffix, odometer):
    """The physical location a row is at: route, county, postmile prefix,
    postmile and roadbed; a row with no postmile is at its odometer."""
    county = county.upper().rstrip(".")
    roadbed = suffix.upper() if suffix.upper() in ("R", "L", "X") else ""
    if not pm:
        return (route, county, "no postmile", number(odometer) if odometer else "")
    return (route, county, prefix.upper(), number(pm), roadbed)


def location_text(loc):
    """A location as the workbook's Location column writes it."""
    if loc[2] == "no postmile":
        return (f"{loc[0]} / {loc[1]} / no postmile"
                + (f" @ odometer {loc[3]}" if loc[3] else ""))
    return f"{loc[0]} / {loc[1]} / {loc[2]}{loc[3]}{loc[4]}"


def digest(parts):
    return hashlib.blake2b("\x1f".join(parts).encode("utf-8"),
                           digest_size=16).digest()


def looks_like_a_mark(value):
    return value in (BLANK, ABSENT) or ARROW in value


class Report:
    def __init__(self):
        self.lines, self.failed = [], 0

    def say(self, line=""):
        self.lines.append(line)
        print(line)

    def check(self, name, ok, detail=""):
        self.failed += not ok
        self.say(f"  [{'OK ' if ok else 'FAIL'}] {name}"
                 + (f"\n         {detail}" if detail and not ok else ""))
        return ok


class Day:
    """One day's per-route files: every data row by the digest of its location
    and compared values."""

    def __init__(self, folder, header, key_cols, compared):
        self.rows = collections.defaultdict(list)   # digest -> [(file, excel row)]
        self.at = collections.Counter()             # location -> rows there
        self.texts = {}                             # digest -> the row's texts
        self.marks = []                             # values that read like a mark
        files = sorted(p for p in Path(folder).iterdir()
                       if _ROUTE.search(p.name) and not p.name.startswith("~$"))
        if not files:
            raise SystemExit(f"no per-route files in {folder}")
        self.files = len(files)
        idx = {name: i for i, name in enumerate(header)}
        for p in files:
            route = route_of(p.name)
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            try:
                it = wb.worksheets[0].iter_rows(values_only=True)
                head = [v for v in next(it, [])]
                while head and head[-1] in (None, ""):
                    head.pop()
                if head != list(header):
                    raise SystemExit(f"{p.name}: its columns are not the report's")
                for n, r in enumerate(it, start=2):
                    r = (list(r) + [None] * len(header))[:len(header)]
                    if all(v is None or str(v).strip() == "" for v in r):
                        continue                    # an empty row is no record
                    t = [text(v) for v in r]
                    for name, value in zip(header, t):
                        if looks_like_a_mark(value):
                            self.marks.append((p.name, n, name, value))
                    loc = location(route, *(t[idx[c]] if c else "" for c in key_cols))
                    d = digest([*loc, *(t[idx[c]] for c in compared)])
                    self.rows[d].append((p.name, n))
                    self.at[loc] += 1
                    self.texts[d] = t
            finally:
                wb.close()

    def count(self):
        return sum(len(v) for v in self.rows.values())


def parts_of(cell, present):
    """A grid cell's value on each day (None where the record is not on that
    day): one value on every day it is on, or 'a → b → c'."""
    s = "" if cell is None else str(cell)
    if ARROW in s:
        bits = s.split(ARROW)
        if len(bits) == 3:
            return [None if b == ABSENT else "" if b == BLANK else b for b in bits]
    return [s if on else None for on in present]


def klass(a, b, c):
    if a == b:
        return "same" if b == c else "newest"
    if b == c:
        return "middle"
    return "back" if a == c else "both"


def shown(v):
    """A day's value as the Changes sheet prints it."""
    return ABSENT if v is None else v or BLANK


def read_summary(rows):
    """The Summary sheet's numbers: per-column classes, records by presence (in
    sheet order), each pair of days, and — where the sheet says first-vs-last
    reads differently through the middle day — that reading's records on both
    and differing cells."""
    start = next(i for i, r in enumerate(rows) if r and r[0] == "Column")
    per_column = {}
    for r in rows[start + 1:]:
        if not r or r[0] is None:
            break
        per_column[r[0]] = [int(x or 0) for x in r[1:7]]
    records, pairs, through = [], {}, None
    for r in rows:
        if not r or not isinstance(r[0], str):
            continue
        if len(r) > 1 and isinstance(r[1], int) and r[0].startswith(_RECORD_ROWS):
            records.append(r[1])
        m = _PAIR_ROW.match(r[0])
        if m and len(r) > 5:
            pairs[m.groups()] = [int(x) for x in r[1:6]]
        m = _THROUGH.search(r[0])
        if m:
            through = [int(x.replace(",", "")) for x in m.groups()]
    return per_column, records, pairs, through


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("workbook")
    ap.add_argument("day1")
    ap.add_argument("day2")
    ap.add_argument("day3")
    ap.add_argument("--spot", type=int, default=0)
    ap.add_argument("--seed", type=int, default=1)
    ap.add_argument("--out")
    a = ap.parse_args(argv)
    rep = Report()
    folders = [Path(a.day1), Path(a.day2), Path(a.day3)]

    first = next(p for p in sorted(folders[0].iterdir()) if _ROUTE.search(p.name))
    wb = openpyxl.load_workbook(first, read_only=True, data_only=True)
    header = [v for v in next(wb.worksheets[0].iter_rows(values_only=True))]
    wb.close()
    while header and header[-1] in (None, ""):
        header.pop()
    family = next((k for k in KEY_COLUMNS if str(header[0]).startswith(k)), None)
    if family is None:
        raise SystemExit("not a Clean Road export (THY_ / INX_ / RAM_ columns)")
    key_cols = KEY_COLUMNS[family]
    pm_col = key_cols[2]

    book = openpyxl.load_workbook(a.workbook, read_only=True, data_only=False)
    summary = [list(r) for r in book["Summary"].iter_rows(values_only=True)]
    title = str(summary[0][0])
    dates = _DAY.findall(title)[-3:]
    rep.say(title)
    rep.say(f"workbook: {a.workbook}")
    per_column, sum_records, sum_pairs, through = read_summary(summary)
    compared = [c for c in header if c in per_column]
    shown_only = [c for c in header if c not in per_column]
    if (len(dates) != 3 or pm_col in compared
            or not set(key_cols[:2] + key_cols[3:4]) <= set(compared)):
        raise SystemExit("the workbook is not a three-day comparison of this report")
    rep.say(f"{len(compared)} compared columns; shown only: {', '.join(shown_only)}")

    days = []
    for i, folder in enumerate(folders, start=1):
        day = Day(folder, header, key_cols, compared)
        days.append(day)
        rep.say(f"day {i} ({dates[i - 1]}): {folder}  —  {day.files} files, "
                f"{day.count():,} rows")
    marks = [m for day in days for m in day.marks]
    rep.say()
    rep.check("no export value reads like one of the workbook's own marks "
              f"({BLANK!r}, {ABSENT!r}, {ARROW!r})", not marks,
              f"{len(marks)} such cells, e.g. {marks[:3]} — a cell showing one "
              "cannot be checked from the workbook; read it in the export file")

    # ---- the lists ---------------------------------------------------------
    listed_records = {}          # grid row -> (the ✓ / — of its days, differing cells)
    listed_names = set()         # (grid row, route, location) as a list row says it
    for r in list(book["Added & removed"].iter_rows(values_only=True))[2:]:
        m = _LINK.search(str(r[7] or "")) if len(r) > 7 else None
        if m:
            listed_records[int(m.group(1))] = (tuple(x == TICK for x in r[3:6]), r[6])
            listed_names.add((int(m.group(1)), str(r[0]), str(r[1])))
    listed_changes = set()       # (grid row, column, v1, v2, v3)
    for r in list(book["Changes"].iter_rows(values_only=True))[2:]:
        m = _LINK.search(str(r[7] or "")) if len(r) > 7 else None
        if m:
            listed_changes.add((int(m.group(1)), r[2], *("" if v is None else str(v)
                                                         for v in r[3:6])))
            listed_names.add((int(m.group(1)), str(r[0]), str(r[1])))
    routes_sheet = {str(r[0]): [int(x or 0) for x in r[1:13]]
                    for r in list(book["Routes"].iter_rows(values_only=True))[2:]
                    if r and r[0] is not None}

    # ---- the grid ----------------------------------------------------------
    grid = book[GRID].iter_rows(values_only=True)
    next(grid)
    head = list(next(grid))
    off = head.index(header[0])
    if head[off:off + len(header)] != list(header):
        raise SystemExit("the grid's columns are not the export's")
    col = {name: off + i for i, name in enumerate(header)}
    used = [collections.Counter() for _ in range(3)]
    records_at = collections.Counter()           # location -> grid records
    forced = twins = 0
    presence = collections.Counter()
    cells = collections.Counter()
    by_column = {c: collections.Counter() for c in compared}
    pair_counts = {p: [0, 0, 0, 0, 0] for p in PAIRS}
    by_route = {}
    changed, changed_rows = set(), {}
    unreadable, shown_wrong, bad_tally, bad_listing = [], [], [], []
    named, bad_location = {}, []                 # grid row -> (route, Location)
    for n, r in enumerate(grid, start=3):
        present, listed_diff = listed_records.get(n, (ALL_THREE, None))
        presence[present] += 1
        vals = {c: parts_of(r[col[c]], present) for c in compared}
        on = [d for d in range(3) if present[d]]
        if any(vals[c][d] is None for c in compared for d in on):
            unreadable.append(n)
            continue
        pm = "" if r[col[pm_col]] is None else str(r[col[pm_col]])
        route = str(r[0])
        locs = set()
        for d in on:
            loc = location(route, vals[key_cols[0]][d], vals[key_cols[1]][d], pm,
                           vals[key_cols[3]][d],
                           vals[key_cols[4]][d] if key_cols[4] else "")
            locs.add(loc)
            dg = digest([*loc, *(vals[c][d] for c in compared)])
            used[d][dg] += 1
            if d == on[-1] and len(days[d].rows.get(dg, ())) == 1:
                raw_texts = days[d].texts[dg]
                for c in shown_only:
                    cell = "" if r[col[c]] is None else str(r[col[c]])
                    if cell != raw_texts[header.index(c)]:
                        shown_wrong.append((n, c, cell, raw_texts[header.index(c)]))
        for loc in locs:
            records_at[loc] += 1
        named[n] = (route, str(r[1]))
        if str(r[1]) not in {location_text(loc) for loc in locs}:
            bad_location.append((n, r[1], sorted(location_text(loc) for loc in locs)))
        unique = all(days[d].at.get(loc, 0) <= 1 for loc in locs for d in range(3))
        forced += unique
        twins += not unique

        tally = collections.Counter()
        for c in compared:
            v = vals[c]
            if len(on) == 3:
                k = klass(*v)
                cells[k] += 1
                by_column[c][k] += 1
                tally[k] += 1
                if k != "same":
                    changed.add((n, c, *(shown(x) for x in v)))
                    changed_rows.setdefault(n, []).append((c, v, k))
            elif len(on) == 2 and v[on[0]] != v[on[1]]:
                by_column[c]["two"] += 1
                tally["two"] += 1
                changed.add((n, c, *(shown(x) for x in v)))
                changed_rows.setdefault(n, []).append((c, v, "two"))
        for (x, y), pc in pair_counts.items():
            if present[x] and present[y]:
                differs = sum(vals[c][x] != vals[c][y] for c in compared)
                pc[0] += 1
                pc[3] += bool(differs)
                pc[4] += differs
            elif present[x]:
                pc[1] += 1
            elif present[y]:
                pc[2] += 1

        # the grid's own tallies, and the Added & removed row
        if len(on) == 3:
            moved = tally["newest"] + tally["middle"] + tally["back"] + tally["both"]
            want = [moved, tally["middle"] + tally["back"] + tally["both"],
                    tally["newest"] + tally["back"] + tally["both"]]
        else:
            want = [tally["two"], tally["two"] if present == PRESENCES[1] else None,
                    tally["two"] if present == PRESENCES[2] else None]
        if [r[3], r[4], r[5]] != want:
            bad_tally.append((n, [r[3], r[4], r[5]], want))
        if len(on) == 2 and listed_diff != tally["two"]:
            bad_listing.append((n, listed_diff, tally["two"]))
        if (present == ALL_THREE) != (str(r[2]) == "all three days"):
            bad_listing.append((n, r[2], present))

        t = by_route.setdefault(route, [0] * 12)
        for d in on:
            t[d] += 1
        t[3] += present == ALL_THREE
        t[4] += present in (PRESENCES[2], PRESENCES[6])          # added
        t[5] += present in (PRESENCES[1], PRESENCES[4])          # removed
        t[6] += present in (PRESENCES[3], PRESENCES[5])          # the middle day only
        for i, k in enumerate(("newest", "middle", "back", "both", "two"), start=7):
            t[i] += tally[k]
    book.close()

    # ---- 1. every raw row once ---------------------------------------------
    rep.say()
    rep.say("1. Every raw row of each day is in the grid exactly once:")
    rep.check("every grid cell reads as a value on each day its record is on",
              not unreadable, f"{len(unreadable)} grid rows, e.g. {unreadable[:5]}")
    for d in range(3):
        raw_count = collections.Counter({k: len(v) for k, v in days[d].rows.items()})
        missing = raw_count - used[d]
        extra = used[d] - raw_count
        where = [days[d].rows[k][0] for k in list(missing)[:3]]
        rep.check(f"day {d + 1}: {sum(raw_count.values()):,} raw rows = "
                  f"{sum(used[d].values()):,} grid rows, values unaltered",
                  not missing and not extra,
                  f"{sum(missing.values())} raw rows not in the grid (e.g. {where}); "
                  f"{sum(extra.values())} grid rows in no file")

    # ---- 2. one record per location ----------------------------------------
    rep.say()
    rep.say("2. A location on one row a day is one grid record:")
    split = [loc for loc, n in records_at.items()
             if n > 1 and all(days[d].at.get(loc, 0) <= 1 for d in range(3))]
    rep.check(f"{forced:,} records at a location with one row a day — none split "
              "across two grid rows", not split, f"split: {split[:5]}")
    rep.say(f"       {twins:,} records at a location that repeats within a day "
            "(their twins pair by likeness)")
    rep.check("the key and shown-only columns are the newest day's own",
              not shown_wrong, f"{len(shown_wrong)} differ, e.g. {shown_wrong[:3]}")

    # ---- 3. the Summary -----------------------------------------------------
    rep.say()
    rep.say("3. The Summary's counts, re-decided from the grid's own values:")
    want_cells = dict.fromkeys(("newest", "middle", "back", "both", "same"), 0)
    for vals in per_column.values():
        for k, n in zip(want_cells, vals[:5]):
            want_cells[k] += n
    rep.check(f"cells by class: {dict(cells)}",
              all(cells[k] == want_cells[k] for k in want_cells), f"Summary: {want_cells}")
    bad_cols = [c for c in compared
                if [by_column[c][k] for k in ("newest", "middle", "back", "both", "same",
                                              "two")] != per_column[c]]
    rep.check("every column's classes", not bad_cols, f"differ: {bad_cols[:5]}")
    rep.check(f"records by the days they are on: {[presence[p] for p in PRESENCES]}",
              [presence[p] for p in PRESENCES] == sum_records,
              f"Summary: {sum_records}")
    for (x, y), got in pair_counts.items():
        want = sum_pairs.get((dates[x], dates[y]))
        name = (f"{dates[x]} vs {dates[y]}: {got[0]:,} on both, {got[1]:,} / {got[2]:,} "
                f"one-sided, {got[3]:,} records and {got[4]:,} cells differ")
        if (x, y) == (0, 2) and through is not None:
            # The Summary pairs first-vs-last directly and says, in a note, what
            # it comes to through the middle day (a repeated key can pair
            # differently). The grid IS the through-the-middle reading.
            rep.check(name + " — read through the middle day, as the Summary's note "
                      "says", through == [got[0], got[4]], f"Summary's note: {through}")
            continue
        rep.check(name, want == got, f"Summary: {want}")

    # ---- 4. the other sheets ---------------------------------------------------
    rep.say()
    rep.say("4. The other sheets:")
    rep.check(f"Routes: {len(by_route):,} routes, every count",
              by_route == routes_sheet,
              f"differ: {[k for k in by_route if by_route[k] != routes_sheet.get(k)][:5]}")
    rep.check("the grid's own tallies (changed cells, and on each day)",
              not bad_tally, f"{len(bad_tally)} rows, e.g. {bad_tally[:3]}")
    rep.check("the grid's Location names each record's own location",
              not bad_location, f"{len(bad_location)} rows, e.g. {bad_location[:3]}")
    misnamed = [x for x in listed_names if named.get(x[0]) != (x[1], x[2])]
    rep.check("every row of the two lists names the record its link points at",
              not misnamed, f"{len(misnamed)} rows, e.g. {sorted(misnamed)[:3]}")
    rep.check(f"Changes lists exactly the {len(changed):,} changed cells, with the "
              "same three values", changed == listed_changes,
              f"only in the grid: {sorted(changed - listed_changes)[:3]}; only in "
              f"Changes: {sorted(listed_changes - changed)[:3]}")
    rep.check(f"Added & removed lists the {len(listed_records):,} records not on "
              "all three days, with their differing cells",
              not bad_listing
              and all(p != ALL_THREE for p, _n in listed_records.values()),
              f"{len(bad_listing)} rows disagree, e.g. {bad_listing[:3]}")

    # ---- spot cells -----------------------------------------------------------
    if a.spot and changed_rows:
        rep.say()
        rep.say(f"{a.spot} changed cells drawn at random (seed {a.seed}) — open the "
                "files and look:")
        rng = random.Random(a.seed)
        book = openpyxl.load_workbook(a.workbook, read_only=True, data_only=False)
        wanted = set(rng.sample(sorted(changed_rows), min(a.spot, len(changed_rows))))
        for n, r in enumerate(book[GRID].iter_rows(values_only=True), start=1):
            if n not in wanted:
                continue
            present = listed_records.get(n, (ALL_THREE, None))[0]
            vals = {c: parts_of(r[col[c]], present) for c in compared}
            c, _v, k = rng.choice(changed_rows[n])
            rep.say(f"  grid row {n}  {r[1]}  {c}  [{k}]")
            pm = "" if r[col[pm_col]] is None else str(r[col[pm_col]])
            for d in range(3):
                if not present[d]:
                    rep.say(f"     {dates[d]}: not in that day's files")
                    continue
                loc = location(str(r[0]), vals[key_cols[0]][d], vals[key_cols[1]][d], pm,
                               vals[key_cols[3]][d],
                               vals[key_cols[4]][d] if key_cols[4] else "")
                dg = digest([*loc, *(vals[x][d] for x in compared)])
                src = days[d].rows.get(dg) or [("?", "?")]
                raw = days[d].texts[dg][header.index(c)] if dg in days[d].texts else "?"
                rep.say(f"     {dates[d]}: {(raw or BLANK)!r:28} {src[0][0]} "
                        f"row {src[0][1]}")
        book.close()

    rep.say()
    rep.say("ALL CHECKS PASSED" if not rep.failed else f"{rep.failed} CHECK(S) FAILED")
    if a.out:
        Path(a.out).write_text("\n".join(rep.lines) + "\n", encoding="utf-8")
    return 1 if rep.failed else 0


if __name__ == "__main__":
    sys.exit(main())
