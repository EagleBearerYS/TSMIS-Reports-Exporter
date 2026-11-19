"""build/check_clean_road_points.py — the Clean Road Highway build's POINT rules
and how the vs-TSMIS comparison reads them.

The site lists every landmark, each side of a postmile equation and each
county's end as a row of its own. The build follows it (`site_points`):

  * a point belongs to its own roadbed: a landmark on the left roadbed of an
    independent alignment cuts and labels the left rows only;
  * a point no segment of its roadbed begins at gets a zero-length row, read at
    the point the way the site reads a row there — the span that begins at the
    point before the one that ends there, the network date by the spans that
    hold the measure and nothing more;
  * several landmarks at one postmile are a row each, in AR order, the last
    one the segment's; an equation's row takes them all, joined;
  * the names the site leaves out are left out (a realignment marker beside a
    named landmark, SELF INTERSECT);
  * a county the route runs on from ends in a zero-length row of its own;
  * a report that does not take these rules (`site_points=False`) builds what
    it built before.

And the comparison: a realignment marker reads in the site's spelling, a
zero-length row agrees on either stretch that meets at its postmile, and it is
never joined into a segment.

Run with the build venv:
    build\\.venv\\Scripts\\python.exe build\\check_clean_road_points.py
"""
import sys
from datetime import date

from _checklib import Checker, scripts_path

scripts_path()

import clean_road_layers as crl  # noqa: E402
import compare_clean_highway_arcgis as cmp  # noqa: E402
import consolidate_clean_highway as cch  # noqa: E402

c = Checker()
u = crl.pm_units
COL = cch._COL
DAY = date(2026, 11, 6)
RID, RID_L = "SHS_001._P", "SHS_001._S"
DIVIDED, UNDIVIDED = "D- Divided Highway", "U- Undivided Highway"
RIGHT, LEFT = "R- Independent Alignment - Right", "L- Independent Alignment - Left"


def span(b, e, *vals, rank=1.0, od=None):
    return cch._Span(u(b), u(e), tuple(vals), (rank, rank), 1.0, od, None)


def pt(value, roadbed="", rid=None, ar=None, item=None):
    return cch._Point(value, roadbed, rid, ar, None, item)


def cell(rec, name):
    row = rec["row"] if isinstance(rec, dict) else rec
    return row[COL[name]]


def begins(rows, kind):
    return sorted(crl.pm_float(r["b"]) for r in rows
                  if r["kind"] == kind and not r.get("point"))


# --------------------------------------------------------------------------- #
def test_roadbeds():
    print("a point belongs to its own roadbed:")
    rb = {"HG": [span(1.0, 3.0, RIGHT)]}
    lb = {"HG": [span(1.0, 3.0, LEFT)]}
    pts = {"LMK": {u(2.0): [pt("LEFT ONLY", "L")], u(2.5): [pt("RIGHT ONLY", "R")]}}
    rows = cch._build_cp("004", "CC", "R", rb, lb, pts, DAY, True)
    c.check("a left-roadbed landmark cuts the left rows only",
            begins(rows, "L") == [1.0, 2.0] and begins(rows, "R") == [1.0, 2.5],
            (begins(rows, "L"), begins(rows, "R")))
    named = {(r["kind"], crl.pm_float(r["b"])): cell(r, "THY_LANDMARK_SHORT_DESC")
             for r in rows}
    c.check("... and labels the left row alone",
            named[("L", 2.0)] == "LEFT ONLY" and named[("R", 2.5)] == "RIGHT ONLY"
            and named[("R", 1.0)] is None and named[("L", 1.0)] is None, named)
    c.check("no zero-length row where a segment of the roadbed begins at the point",
            not any(r.get("point") for r in rows))
    old = cch._build_cp("004", "CC", "R", rb, lb, pts, DAY, True, None, False)
    c.check("a report that does not take the site's point rules builds as before: "
            "every point cuts and labels both roadbeds, and makes no row",
            begins(old, "L") == begins(old, "R") == [1.0, 2.0, 2.5]
            and {cell(r, "THY_LANDMARK_SHORT_DESC") for r in old
                 if crl.pm_float(r["b"]) == 2.0} == {"LEFT ONLY"}
            and not any(r.get("point") for r in old))


def _end_of_space():
    """001 SLO: the R space ends at R7.531 where the plain space begins; the
    landmark sits 4e-6 short of the boundary, as the layer has it."""
    fill = cch._ArFill()
    spans = {
        "HG": (span(6.0, 7.531, UNDIVIDED), span(7.531, 8.61, DIVIDED)),
        "OSR": (span(6.0, 7.531, 8, 8), span(7.531, 8.61, 4, 4)),
        "NET": (span(6.0, 7.531, date(1996, 1, 17), 67400),
                span(7.531, 8.61, date(1964, 1, 1), 68400)),
    }
    for tag, (before, after) in spans.items():
        fill.add(tag, "001", "R", ("SLO", "R"), ("SLO", "R"), before.b, before.e,
                 RID, 167.4, 168.929337, before)
        fill.add(tag, "001", "R", ("SLO", ""), ("SLO", ""), after.b, after.e,
                 RID, 168.929337, 170.0, after)
    return fill, {tag: [pair[0]] for tag, pair in spans.items()}


def test_end_of_space():
    print("a landmark at the end of its postmile space is a zero-length row:")
    fill, bucket = _end_of_space()
    pts = {"LMK": {u(7.531): [pt("END REALIGNMENT", "", RID, 168.9293326)]}}
    rows = cch._build_cp("001", "SLO", "R", bucket, {}, pts, DAY, True, fill)
    point = [r for r in rows if r.get("point")]
    c.check("one segment and one point row", len(rows) == 2 and len(point) == 1,
            [(r["b"], r["e"], r.get("point")) for r in rows])
    p = point[0]
    c.check("the point row begins and ends at the landmark's postmile, length 0",
            cell(p, "THY_BEGIN_PM_AMT") == cell(p, "THY_END_PM_AMT") == 7.531
            and cell(p, "THY_LENGTH_MILES_AMT") == 0
            and cell(p, "THY_PM_PREFIX_CODE") == "R"
            and cell(p, "THY_LANDMARK_SHORT_DESC") == "END REALIGNMENT")
    c.check("it reads the stretch that BEGINS at the point (the span that begins "
            "within 1e-5 of the measure), as the site does",
            cell(p, "THY_HIGHWAY_GROUP_CODE") == "D"
            and cell(p, "THY_RT_O_SHD_TOT_WIDTH_AMT") == 4,
            (cell(p, "THY_HIGHWAY_GROUP_CODE"), cell(p, "THY_RT_O_SHD_TOT_WIDTH_AMT")))
    c.check("... but the network date by the spans that HOLD the measure alone: "
            "the stretch that ends there",
            cell(p, "THY_SEG_ORDER_ID") == 67400
            and cell(p, "THY_RECORD_DATE") == date(1996, 1, 17),
            (cell(p, "THY_SEG_ORDER_ID"), cell(p, "THY_RECORD_DATE")))
    equated = dict(pts, EQP={u(7.531): [pt("***Equates to SLO001", "")]})
    rows = cch._build_cp("001", "SLO", "R", bucket, {}, equated, DAY, True, fill)
    p = next(r for r in rows if r.get("point"))
    c.check("an equation's row is placed by its POSTMILE (the boundary itself), so "
            "its network date is the stretch that begins there",
            cell(p, "THY_EQUATE_CODE") == "E" and cell(p, "THY_SEG_ORDER_ID") == 68400,
            (cell(p, "THY_EQUATE_CODE"), cell(p, "THY_SEG_ORDER_ID")))
    alone = cch._build_cp("001", "SLO", "R", {"HG": bucket["HG"]}, {}, pts, DAY, True)
    p = next(r for r in alone if r.get("point"))
    c.check("with no AR network to place it on, the point reads the segment that "
            "ends at its postmile", cell(p, "THY_HIGHWAY_GROUP_CODE") == "U")
    dated = {"LMK": {u(7.531): [pt("END REALIGNMENT", "", RID, 168.9293326,
                                   crl.to_serial(date(2019, 10, 14)))]}}
    rows = cch._build_cp("001", "SLO", "R", {"HG": bucket["HG"]}, {}, dated, DAY, True)
    p = next(r for r in rows if r.get("point"))
    c.check("a row no network-date span holds is dated by its landmark",
            cell(p, "THY_RECORD_DATE") == date(2019, 10, 14),
            cell(p, "THY_RECORD_DATE"))


def test_several_landmarks():
    print("several landmarks at one postmile are a row each:")
    buckets = {("001", "R", "SB", "R"): {"HG": [span(30.0, 32.0, DIVIDED, od=0.0)]}}
    here = {"LMK": {u(31.042): [pt("SECOND BY MEASURE", "", RID, 145.8031820),
                                pt("FIRST BY MEASURE", "", RID, 145.8031819)]}}
    rows = cch._build_route("001", buckets, {("001", "SB", "R"): here}, DAY)
    at = [(r[COL["THY_BEGIN_PM_AMT"]], r[COL["THY_LENGTH_MILES_AMT"]],
           r[COL["THY_LANDMARK_SHORT_DESC"]]) for r in rows]
    c.check("the first by AR measure is a zero-length row set before the segment, "
            "and the segment carries the last",
            at == [(30.0, 1.042, None), (31.042, 0, "FIRST BY MEASURE"),
                   (31.042, 0.958, "SECOND BY MEASURE")], at)
    same = {"LMK": {u(31.042): [pt("TWICE", "", RID, 145.8), pt("TWICE", "", RID, 145.8)]}}
    rows = cch._build_route("001", buckets, {("001", "SB", "R"): same}, DAY)
    c.check("one name at one place counts once", len(rows) == 2, len(rows))
    joined = {"LMK": {u(31.042): [pt("ONE   ", "", RID, 145.8), pt("TWO", "", RID, 145.9)]},
              "EQP": {u(31.042): [pt("Equation", "")]}}
    rows = cch._build_route("001", buckets, {("001", "SB", "R"): joined}, DAY)
    c.check("an equation's row takes every landmark at its postmile, joined, and "
            "they make no rows of their own",
            len(rows) == 2 and rows[1][COL["THY_LANDMARK_SHORT_DESC"]] == "ONE    / TWO"
            and rows[1][COL["THY_EQUATE_CODE"]] == "E",
            [(r[COL["THY_BEGIN_PM_AMT"]], r[COL["THY_LANDMARK_SHORT_DESC"]]) for r in rows])


def test_equation_takes_both_roadbeds():
    print("an equation's row takes the landmarks of every roadbed at its postmile:")
    pts = {"LMK": {u(82.739): [pt("END INDEP ALIGN RT", "", RID, 318.6),
                               pt("END LT INDEP ALIGN", "L", RID_L, 318.5)]},
           "EQP": {u(82.739): [pt("Equation", "")]}}
    listed = cch._listed_landmarks(pts, "R")
    c.check("both are listed under the equation's roadbed, in AR order, and under "
            "no other",
            {k: [p.value for p in v] for k, v in listed.items()}
            == {("", u(82.739)): ["END LT INDEP ALIGN", "END INDEP ALIGN RT"]},
            {k: [p.value for p in v] for k, v in listed.items()})
    rb = {"HG": [span(81.505, 82.739, RIGHT), span(82.739, 84.0, DIVIDED)]}
    lb = {"HG": [span(81.505, 82.739, LEFT)]}
    rows = cch._build_cp("001", "MON", "R", rb, lb, pts, DAY, True)
    at = [(r["kind"], cell(r, "THY_LANDMARK_SHORT_DESC"), cell(r, "THY_EQUATE_CODE"))
          for r in rows if crl.pm_float(r["b"]) == 82.739]
    c.check("so the equation's row carries both names and the left roadbed has no "
            "row of its own there",
            at == [("base", "END LT INDEP ALIGN / END INDEP ALIGN RT", "E")], at)
    plain = {"LMK": pts["LMK"]}
    rows = cch._build_cp("001", "MON", "R", rb, lb, plain, DAY, True)
    c.check("with no equation there, each landmark keeps its own roadbed's row",
            sorted((r["kind"], cell(r, "THY_LANDMARK_SHORT_DESC")) for r in rows
                   if crl.pm_float(r["b"]) == 82.739)
            == [("L", "END LT INDEP ALIGN"), ("base", "END INDEP ALIGN RT")])


def test_names_left_out():
    print("the landmark names the site leaves out:")

    def listed(prefix, *names):
        pts = {"LMK": {u(5.0): [pt(n, "", RID, 10.0 + i) for i, n in enumerate(names)]}}
        return [p.value for found in cch._listed_landmarks(pts, prefix).values()
                for p in found]

    c.check("a BEGIN / END REALIGNMENT marker beside a named landmark",
            listed("R", "BEGIN REALIGNMENT", "MAIN ST UC") == ["MAIN ST UC"]
            and listed("", "END REALIGNMENT", "MAIN ST UC") == ["MAIN ST UC"])
    c.check("... but not one that stands alone, or two markers together",
            listed("R", "END REALIGNMENT") == ["END REALIGNMENT"]
            and listed("M", "END REALIGNMENT", "BEGIN REALIGNMENT")
            == ["END REALIGNMENT", "BEGIN REALIGNMENT"])
    c.check("... nor under a prefix the site's own pattern does not know (T)",
            listed("T", "BEGIN REALIGNMENT", "MAIN ST UC")
            == ["BEGIN REALIGNMENT", "MAIN ST UC"])
    c.check("SELF INTERSECT is never listed",
            listed("", "SELF INTERSECT", "MAIN ST UC") == ["MAIN ST UC"]
            and listed("", "Self Intersect") == [])
    rows = cch._build_cp("001", "SB", "", {"HG": [span(4.0, 6.0, DIVIDED)]}, {},
                         {"LMK": {u(5.0): [pt("SELF INTERSECT", "")]}}, DAY, True)
    c.check("... and cuts no row", len(rows) == 1, len(rows))


def test_county_ends():
    print("a county the route runs on from ends in a zero-length row:")

    def route(next_od):
        buckets = {
            ("001", "R", "ORA", ""): {"HG": [span(30.0, 33.719, DIVIDED, od=0.0)],
                                      "OSR": [span(30.0, 33.719, 8, 8)]},
            ("001", "R", "LA", ""): {"HG": [span(0.0, 2.0, UNDIVIDED, od=next_od)]},
        }
        points = {("001", "ORA", ""): {"LMK": {u(30.0): [pt("BEGIN HERE", "")]}}}
        return cch._build_route("001", buckets, points, DAY)

    rows = route(3.719)
    at = [(r[COL["THY_COUNTY_CODE"]], r[COL["THY_BEGIN_PM_AMT"]],
           r[COL["THY_LENGTH_MILES_AMT"]]) for r in rows]
    c.check("the departing county's last postmile, between its last segment and the "
            "next county's first",
            at == [("ORA", 30.0, 3.719), ("ORA", 33.719, 0), ("LA", 0.0, 2.0)], at)
    end = rows[1]
    c.check("it carries the last segment's values and nothing of the point that "
            "segment began at",
            end[COL["THY_HIGHWAY_GROUP_CODE"]] == "D"
            and end[COL["THY_RT_O_SHD_TOT_WIDTH_AMT"]] == 8
            and end[COL["THY_END_PM_AMT"]] == 33.719
            and end[COL["THY_LANDMARK_SHORT_DESC"]] is None
            and end[COL["THY_BREAK_DESC"]] is None,
            [end[COL[n]] for n in ("THY_HIGHWAY_GROUP_CODE", "THY_LANDMARK_SHORT_DESC",
                                   "THY_BREAK_DESC")])
    c.check("the segments keep the markers they had: the route's END stays on its "
            "last segment, the county change on the next county's first",
            rows[2][COL["THY_BREAK_DESC"]] in ("END", "CNTY", "D-C")
            and rows[0][COL["THY_BREAK_DESC"]] == "BEG",
            [r[COL["THY_BREAK_DESC"]] for r in rows])
    c.check("no such row where the route does not run on across the line",
            len(route(9.0)) == 2)
    c.check("... and the route's own end has none", rows[-1][COL["THY_COUNTY_CODE"]] == "LA"
            and rows[-1][COL["THY_LENGTH_MILES_AMT"]] == 2.0)


def test_point_read():
    print("one layer read at one point, as the site reads a row there:")
    fill = cch._ArFill()
    a, b, far = span(0.0, 1.0, "A"), span(1.0, 2.0, "B"), span(9.0, 9.5, "FAR")
    fill.add("SUR", "001", "R", ("ORA", ""), ("ORA", ""), a.b, a.e, RID, 10.0, 11.0, a)
    fill.add("SUR", "001", "R", ("ORA", ""), ("ORA", ""), b.b, b.e, RID, 11.0000001, 12.0, b)
    fill.add("SUR", "001", "R", ("ORA", ""), ("ORA", ""), far.b, far.e, RID, 20.0, 20.5, far)

    def read(ar, rid=RID, **kw):
        got = fill.at("SUR", (rid, ar), **kw)
        return got.vals[0] if got is not None else None

    c.check("inside a span: that span", read(10.5) == "A")
    c.check("at a span's end with the next beginning a hair on: the next",
            read(11.0) == "B")
    c.check("... unless only the spans that hold the measure are read (network date)",
            read(11.0, exact=True) == "A")
    c.check("past the last touch but within 0.005: the nearest", read(12.004) == "B")
    c.check("further than that: nothing", read(12.006) is None and read(15.0) is None)
    c.check("beyond everything the layer has on the RouteID, within half a mile "
            "when asked (the left roadbed's shorter calibration)",
            read(20.9, far=True) == "FAR" and read(20.9) is None
            and read(21.1, far=True) is None)
    c.check("a RouteID with nothing falls to the second place given",
            fill.at("SUR", (RID_L, 10.5), (RID, 10.5)).vals[0] == "A"
            and fill.at("SUR", (RID_L, 10.5)) is None)


def test_spelling():
    print("the comparison reads a realignment marker in the site's spelling:")
    c.check("the row's PM prefix goes into the name",
            cmp._landmark_as_site("END REALIGNMENT", "R") == "END R REALIGNMENT"
            and cmp._landmark_as_site("BEGIN REALIGNMENT", "M") == "BEGIN M REALIGNMENT")
    c.check("no prefix, no change",
            cmp._landmark_as_site("END REALIGNMENT", "") == "END REALIGNMENT")
    c.check("each marker in an equation's joined names, and only the markers",
            cmp._landmark_as_site("MERCED RIV BR 40-47     / END REALIGNMENT", "R")
            == "MERCED RIV BR 40-47     / END R REALIGNMENT")
    c.check("any other name is left as it is",
            cmp._landmark_as_site("REALIGNMENT RD UC", "R") == "REALIGNMENT RD UC"
            and cmp._landmark_as_site("", "R") == "")


def _row(key, county, value, other, landmark=""):
    row = [""] * (1 + len(cmp.SHARED_HEADER))
    row[0] = "001"
    row[cmp._KEY_AT] = key
    row[cmp._COUNTY_AT] = county
    row[cmp._at("THY_MEDIAN_TYPE_CODE")] = value
    row[cmp._at("THY_TERRAIN_CODE")] = other
    row[cmp._at("THY_LANDMARK_SHORT_DESC")] = landmark
    return row


def test_settle():
    print("a zero-length row agrees on either stretch that meets at its postmile:")
    value, other = cmp._at("THY_MEDIAN_TYPE_CODE"), cmp._at("THY_TERRAIN_CODE")
    before = _row("K0", "ORA", "ENDING", "FLAT")
    point = _row("K1", "ORA", "BEGINNING", "ROLLING", "END REALIGNMENT")
    rows, spans = [before, point], [(0, 100), (100, 100)]
    found = cmp._settle_point_rows(rows, spans, [_row("K1", "ORA", "ENDING", "MOUNTAIN")])
    c.check("the cell takes the ending stretch's value where the site shows it",
            point[value] == "ENDING", point[value])
    c.check("a value that is neither stretch's is left as built, and counts",
            point[other] == "ROLLING", point[other])
    c.check("the counts say so", found == {"rows": 1, "settled": 1, "cells": 1}, found)
    first = _row("K1", "ORA", "BEGINNING", "ROLLING", "ONE")
    second = _row("K1", "ORA", "BEGINNING", "ROLLING", "TWO")
    site = [_row("K1", "ORA", "BEGINNING", "ROLLING"), _row("K1", "ORA", "ENDING", "ROLLING")]
    cmp._settle_point_rows([_row("K0", "ORA", "ENDING", "FLAT"), first, second],
                           [(0, 100), (100, 100), (100, 100)], site)
    c.check("rows of one key with no landmark to tell them apart pair in order",
            first[value] == "BEGINNING" and second[value] == "ENDING",
            (first[value], second[value]))
    first = _row("K1", "ORA", "BEGINNING", "ROLLING", "ONE")
    second = _row("K1", "ORA", "BEGINNING", "ROLLING", "TWO")
    site = [_row("K1", "ORA", "ENDING", "ROLLING", "TWO"),
            _row("K1", "ORA", "BEGINNING", "ROLLING", "ONE")]
    cmp._settle_point_rows([_row("K0", "ORA", "ENDING", "FLAT"), first, second],
                           [(0, 100), (100, 100), (100, 100)], site)
    c.check("... and by their landmark where they carry one, as the engine pairs "
            "rows of one key by likeness",
            first[value] == "BEGINNING" and second[value] == "ENDING",
            (first[value], second[value]))
    lone = _row("K9", "ORA", "BEGINNING", "ROLLING")
    found = cmp._settle_point_rows([lone], [(100, 100)],
                                   [_row("K9", "ORA", "ENDING", "FLAT")])
    c.check("with no segment ending at the point there is nothing to take",
            lone[value] == "BEGINNING" and found["settled"] == 0, found)

    a = _row("K0", "ORA", "SAME", "FLAT")
    p = _row("K1", "ORA", "SAME", "FLAT", "A LANDMARK")
    b = _row("K2", "ORA", "SAME", "FLAT")
    kept, kept_spans, joined = cmp._join_unchanged(
        [a, p, b], [(0, 100), (100, 100), (100, 200)], set())
    c.check("a zero-length row is never joined into a segment, and the segments "
            "on either side of it still join",
            len(kept) == 2 and kept_spans == [(0, 200), (100, 100)]
            and joined == {"same": 1, "dates": 0}, (kept_spans, joined))


def main():
    test_roadbeds()
    test_end_of_space()
    test_several_landmarks()
    test_equation_takes_both_roadbeds()
    test_names_left_out()
    test_county_ends()
    test_point_read()
    test_spelling()
    test_settle()
    return c.summary()


if __name__ == "__main__":
    sys.exit(main())
