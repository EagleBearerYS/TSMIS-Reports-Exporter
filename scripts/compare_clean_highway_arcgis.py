"""Clean Road Highway — the ArcGIS-built CA HIGHWAYS workbook vs the SITE's
Clean Road Highway export (v0.50.0).

The Reports vs ArcGIS matrix's Clean Road Highway rows. Side A is OUR CA
HIGHWAYS table built from the ArcGIS layer library (`consolidate_clean_highway`,
the ArcGIS reports tab); side B is the site's own Clean Road Highway file for one
day, consolidated — the Excel export or the print read back from its PDF, both
the same shape (the PDF-vs-Excel self-check proves it), so one comparator serves
both editions.

This one is TSMIS vs TSMIS: both sides are built from the same LRS layers, so a
difference is a gap in OUR reconstruction or a finding about how the site builds
its file — never migration drift. Values are compared on the vs-TSN comparison's
FORMAT-only reconciliations, because the two sides really do spell values
differently (the site prints MM/DD/YYYY, '02' lanes and 'LA.' for the two-letter
counties; the build writes the TSN extract's shapes): dates as dates, amounts
as numbers, county codes with or without the trailing period, a blank
toll/forest code as 0, and the equation marker 'E' the site writes into the
postmile-suffix column is no roadbed for the key. Both sides are then SHOWN the
way the site's file writes them (`compare_clean_road_tsn.rows_as_written`,
v0.54.1), which moves no count. One vocabulary is read across:
THY_BREAK_DESC, where the site writes only SHS Route Break's points ('Route
Resume' as BEG, 'Route Break' as END) and the build the TSN extract's markers —
its U-BR / R-BR are those same two points, and its structural markers (a
route's BEG / END, CNTY, D-C, DIST, BEGL, ENDR) have no source on the site, so
they read as blank. Nothing else semantic is bridged.

Rows pair exactly as the site-vs-TSN comparison pairs them (route + county + PM
prefix + BEGIN postmile + roadbed), through the same `compare_clean_road_tsn`
projection, so the two comparisons of one export can never key it differently.
A site row with no postmile keys on its odometer instead and is disclosed on the
Notes sheet (`compare_clean_road_tsn.UNPLACED`).

CONTEXT (the site file's own value shown, never counted) is the vs-TSN set:
the columns the site has no
source for (blank in its file by design — the build may still carry a value, the
ADT trio for one), TSN's bookkeeping columns, the two offset columns (the site
writes the LRS odometer, the build its PM-continued cumulative),
THY_LAST_SIG_CHG_DATE, THY_FUNCTIONAL_CLASS_CODE and the extract date. The
build's unplaceable-span anchors stay NON-ASSERTING here exactly as they are
against TSN.

POINT ROWS. The site lists every landmark, each side of a postmile equation
and each county's end as a row; the build writes a zero-length row where no
segment begins at such a point (`consolidate_clean_highway._point_rows`,
`_county_end_rows`). A zero-length row sits where two stretches meet, so its
cells agree with the site's on either stretch's value (`_settle_point_rows`).
The build's BEGIN / END REALIGNMENT landmark names read with the row's PM
prefix, as the site spells them (`_landmark_as_site`).

SEGMENTS are read the way the site's file is defined. The business requirement
asks for "homogenous segments of all the highway inventory data items" that
break at the county boundary (items 8 and 10). The build cuts a segment at
every layer breakpoint (the CA HIGHWAYS table does); the site's Highway Log
pipeline — the engine the vendor's source list names for this file — keeps a
boundary only where an attribute changes, plus its landmarks, county lines,
city limits and equations (`highway_log.js`: an attribute-change row with no
changed field is dropped; the four block effective dates are not among the
fields). So a build segment that begins where nothing the file reports changes,
or where only a block effective date changes, is JOINED to the segment before
it — unless the site's file has a row at that postmile, which stays paired so
its cells still compare. A joined boundary carries no attribute change, so
nothing a reader could act on is hidden; the Notes say how many were joined.

VINTAGE and DATA SOURCE open the Notes sheet: the build is a reconstruction as
of a date from one feature service, the export is one day's file from one of
the site's data sources; across a date gap the counts include network change.
Console-free; engine in compare_core.
"""
import logging
import re
from collections import Counter, defaultdict
from dataclasses import replace
from pathlib import Path

try:
    from openpyxl import load_workbook
    _DEPS_OK = True
except ImportError:
    _DEPS_OK = False

import clean_highway_columns as chc
import clean_road_layers as crl
import clean_road_sources as crs
import compare_clean_highway_tsn as cht
import compare_clean_road_tsn as crt
import compare_tsn_common as ctc
import consolidate_clean_highway as cch
from consolidate_clean_highway import UNAVAILABLE_TOKEN
from paths import today_str
from pdf_table_lib import pdf_source_marker_state

log = logging.getLogger("tsmis.compare")

REPORT_NAME = "Clean Road Highway — ArcGIS vs TSMIS"
SIDE_A = "ArcGIS"
SIDE_B = "TSMIS"
file_a_label = SIDE_A
file_b_label = SIDE_B

_PROFILE = crt.profile_for("clean_highway")
_SPEC = _PROFILE.spec
_ARC_INDEX = {name: i for i, name in enumerate(chc.ARC_HEADER)}
SHARED_HEADER = list(_SPEC.header)
CONTEXT_FIELDS = tuple(_PROFILE.context)
# What the site's export adds to the build's own context: the columns its file
# leaves blank by design (`clean_road_columns`) and its extract date.
EXPORT_CONTEXT_FIELDS = tuple(c for c in CONTEXT_FIELDS
                              if c not in cch.CONTEXT_COLUMNS)
# The build's THY_BREAK_DESC read in the site's route-break vocabulary
# (clean_highway.js: "133 Route_Break_Type → BEG/END"); every other build
# marker is one the site never writes.
_BREAK_AS_SITE = {"U-BR": "BEG", "R-BR": "END"}
_BREAK_AT = 1 + SHARED_HEADER.index("THY_BREAK_DESC")
# A BEGIN / END REALIGNMENT landmark as the site writes it: with the row's PM
# prefix set into the name (hsl.js `queryLandmarks`: "BEGIN R REALIGNMENT").
# The layer's own name has none, and the build writes the layer's name.
_REALIGNMENT_AS_SITE = {"begin realignment": "BEGIN", "end realignment": "END"}
_LANDMARK_AT = 1 + SHARED_HEADER.index("THY_LANDMARK_SHORT_DESC")
# The build's THY_TOLL_FOREST_CODE in the site's terms, for this two-sided
# comparison only. The build writes the TSN extract's codes (1 = a toll span,
# 2 = a forest highway, toll first); the site writes SHS Forest HWY's own
# Forest_Hwy code, 1 on a forest highway, and reads no toll layer
# (highway_log.js: queryRangeLayer(lp, 115, 'Forest_Hwy')). So here a forest
# highway reads 1 and a toll span as none. The three-way keeps the TSN codes:
# there the site's 1 against TSN's and the layers' 2 is a finding. Measured on
# the 2026-11-19 export: no segment the build marks toll is a forest highway in
# the site's file, so toll-first loses nothing.
_TOLL_FOREST_AS_SITE = {"2": "1", "1": "0"}
_TOLL_FOREST_AT = 1 + SHARED_HEADER.index("THY_TOLL_FOREST_CODE")


def _tolls_as_site(rows):
    """THY_TOLL_FOREST_CODE of every build row in the site's codes (above),
    in place; returns how many cells changed."""
    changed = 0
    for row in rows:
        now = _TOLL_FOREST_AS_SITE.get(row[_TOLL_FOREST_AT])
        if now is not None:
            row[_TOLL_FOREST_AT] = now
            changed += 1
    return changed


def _at(column):
    return 1 + SHARED_HEADER.index(column)


# Joining the build's segments the way the site's file is cut (module
# docstring). A boundary starts a segment when one of _SEGMENT_COLUMNS changes
# across it or the row carries one of _POINT_MARKS: every counted column except
# where the segment ends, the four block effective dates, and the point marks
# themselves — a landmark, equation or break belongs to the postmile a segment
# BEGINS at, so the segment before a plain continuation may carry one and the
# joined segment keeps it.
_KEY_AT = _at(_SPEC.pm_col)
_END_AT = _at("THY_END_PM_AMT")
_LENGTH_AT = _at("THY_LENGTH_MILES_AMT")
_END_OFFSET_AT = _at("THY_END_OFFSET_AMT")
_COUNTY_AT = _at(_SPEC.county_col)
_PREFIX_AT = _at(_SPEC.pm_prefix_col)
_ROADBED_AT = _at(_SPEC.pm_suffix_col)
_EFF_DATES = ("THY_LEFT_ROAD_EFF_DATE", "THY_MEDIAN_EFF_DATE",
              "THY_RIGHT_ROAD_EFF_DATE", "THY_ACCESS_EFF_DATE")
_EFF_DATES_AT = tuple(_at(c) for c in _EFF_DATES)
_POINT_MARKS = ("THY_LANDMARK_SHORT_DESC", "THY_EQUATE_CODE", "THY_BREAK_DESC")
_POINT_MARKS_AT = tuple(_at(c) for c in _POINT_MARKS)
_SEGMENT_COLUMNS = tuple(
    c for c in SHARED_HEADER
    if c not in CONTEXT_FIELDS and c not in _EFF_DATES and c not in _POINT_MARKS
    and c not in (_SPEC.pm_col, "THY_END_PM_AMT", "THY_LENGTH_MILES_AMT"))
_SEGMENT_AT = tuple(_at(c) for c in _SEGMENT_COLUMNS)
# What a zero-length row READS at its point (`_settle_point_rows`): every
# counted column but the ones that say where the row is and what the point is.
_POINT_READ_COLUMNS = tuple(
    c for c in SHARED_HEADER
    if c not in CONTEXT_FIELDS and c not in _POINT_MARKS
    and c not in (_SPEC.pm_col, "THY_END_PM_AMT", "THY_LENGTH_MILES_AMT",
                  _SPEC.route_col, _SPEC.suffix_col, _SPEC.county_col,
                  _SPEC.pm_prefix_col, _SPEC.pm_suffix_col))
_POINT_READ_AT = tuple(_at(c) for c in _POINT_READ_COLUMNS)


def _s(v):
    return "" if v is None else str(v).strip()


def _landmark_as_site(text, prefix):
    """The build's landmark cell in the site's spelling: each BEGIN / END
    REALIGNMENT marker in it (an equation's row can carry several names,
    joined) takes the row's PM prefix. Any other name is returned as it is."""
    if "realignment" not in text.lower():
        return text
    names = []
    for name in text.split(cch.LANDMARK_JOIN):
        word = _REALIGNMENT_AS_SITE.get(name.strip().lower())
        names.append(name if word is None else
                     f"{word} {prefix} REALIGNMENT" if prefix else f"{word} REALIGNMENT")
    return cch.LANDMARK_JOIN.join(names)


def _site_row(row, collapse, hint):
    """One build row (in the build's column order) as this comparison reads
    it: the site export's 74 columns by name, through the projection the site
    side uses, with the build's break markers and realignment landmarks read
    in the site's vocabulary. A build-only column has no site counterpart and
    is left out; an unplaceable-span anchor keeps its reserved token verbatim
    so the schema can hold it non-asserting."""
    vals = [row[_ARC_INDEX[c]] for c in _SPEC.header]
    out = crt._row(_PROFILE, vals, "tsn", hint, collapse=collapse)
    out[_BREAK_AT] = _BREAK_AS_SITE.get(out[_BREAK_AT], "")
    out[_LANDMARK_AT] = _landmark_as_site(out[_LANDMARK_AT], out[_PREFIX_AT])
    for i, raw in enumerate(vals):
        if raw == UNAVAILABLE_TOKEN and _SPEC.header[i] != _SPEC.pm_col:
            out[1 + i] = UNAVAILABLE_TOKEN
    return out


def _open(path):
    name = Path(path).name
    try:
        return name, load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise ValueError(f"Could not open {name}: {type(e).__name__}: {e}")


def _load_build(path, collapse):
    """Side A's rows as the build wrote them (every layer breakpoint a
    segment) — what the three-way comparison reads."""
    return _read_build(path, collapse)[0]


def _read_build(path, collapse):
    """Side A: OUR CA HIGHWAYS workbook — marker-gated, the exact build header —
    each row projected by `_site_row`. Returns (rows, spans): each row's
    (begin, end) postmile in micro-units, for the join."""
    name, wb = _open(path)
    try:
        if chc.ARC_MARKER_SHEET not in wb.sheetnames:
            raise ValueError(
                f"The {SIDE_A} side is not an ArcGIS-built Clean Road Highway "
                f"workbook (no '{chc.ARC_MARKER_SHEET}' sheet):\n{path}\n\n"
                "Build it on the ArcGIS ▸ ArcGIS reports tab first.")
        if chc.ARC_SHEET not in wb.sheetnames:
            raise ValueError(f"{name} has no '{chc.ARC_SHEET}' sheet — rebuild "
                             "the Clean Road Highway report.")
        it = wb[chc.ARC_SHEET].iter_rows(values_only=True)
        header = [_s(c) for c in (next(it, None) or ())]
        if header != list(chc.ARC_HEADER):
            raise ValueError(
                f"{name} does not carry the exact {len(chc.ARC_HEADER)}-column "
                "build header — rebuild the Clean Road Highway report with this "
                "version.")
        hint = f"{name} ({chc.ARC_SHEET})"
        rows, spans = [], []
        begin_at = _ARC_INDEX[_SPEC.pm_col]
        end_at = _ARC_INDEX["THY_END_PM_AMT"]
        for r in it:
            if not ctc.row_has_data(r):
                continue
            r = list(r) + [None] * (len(chc.ARC_HEADER) - len(r))
            rows.append(_site_row(r, collapse, hint))
            spans.append((crl.pm_units(r[begin_at]), crl.pm_units(r[end_at])))
        return rows, spans
    finally:
        wb.close()


def _sum_amounts(a, b):
    """Two canonical lengths as one (a non-numeric cell keeps the first)."""
    try:
        return cht._norm_amount(round(float(a) + float(b), 5))
    except (TypeError, ValueError):  # silent-ok: a length cell holding no number (a blank, the unavailable token) has nothing to add, so the joined segment keeps the first segment's cell as written
        return a


def _join_unchanged(rows, spans, site_keys):
    """The build's segments with the boundaries the site's pipeline does not
    keep joined away (module docstring): a segment is folded into the one that
    ends where it begins when it carries no landmark / equation / break, none
    of its segment columns differs, and the site's file has no row keyed
    there. A zero-length row is a point, never a segment: it is neither
    joined nor joined onto. Returns (rows, their spans, {"same": n, "dates":
    n}) — the boundaries joined where nothing changes, and where only a block
    effective date changes."""
    kept, kept_spans, open_at = [], [], {}
    joined = {"same": 0, "dates": 0}
    for row, (begin, end) in zip(rows, spans):
        stretch = _stretch(row)
        if begin is not None and begin == end:
            kept.append(row)
            kept_spans.append((begin, end))
            continue
        i = open_at.get((stretch, begin)) if begin is not None else None
        if (i is not None and end is not None
                and (row[0], row[_KEY_AT]) not in site_keys
                and not any(row[at] for at in _POINT_MARKS_AT)
                and all(row[at] == kept[i][at] for at in _SEGMENT_AT)):
            before = kept[i]
            same_dates = all(row[at] == before[at] for at in _EFF_DATES_AT)
            joined["same" if same_dates else "dates"] += 1
            before[_LENGTH_AT] = _sum_amounts(before[_LENGTH_AT], row[_LENGTH_AT])
            before[_END_AT] = row[_END_AT]
            before[_END_OFFSET_AT] = row[_END_OFFSET_AT]
            kept_spans[i] = (kept_spans[i][0], end)
            del open_at[(stretch, begin)]
            open_at[(stretch, end)] = i
            continue
        kept.append(row)
        kept_spans.append((begin, end))
        if end is not None:
            open_at[(stretch, end)] = len(kept) - 1
    return kept, kept_spans, joined


def _stretch(row):
    """The road a row runs along: route, county, PM prefix and roadbed."""
    roadbed = row[_ROADBED_AT] if row[_ROADBED_AT] in ("R", "L", "X") else ""
    return (row[0], row[_COUNTY_AT], row[_PREFIX_AT], roadbed)


def _settle_point_rows(rows, spans, rows_b):
    """A zero-length build row — a landmark or an equation point with no
    segment of its own — sits where two stretches meet, and at that point
    either stretch's value is right. The build writes the one it reads at the
    point the way the site reads a row (`consolidate_clean_highway.
    _seg_at_point`): mostly the stretch that BEGINS there. Which of the two
    the site prints turns on a measure difference below its own 1e-5
    tolerance (the LRS's translation of the postmile against the layers'
    stored measures), which no layer file carries. So where the site's cell
    shows the value of the build's segment that ENDS at the point, the row
    takes that value. A value that is neither stretch's is left as built, and
    counts. Where a key holds several rows, a zero-length row is read against
    the site's row of the same landmark (the engine pairs rows of one key by
    likeness, and the landmark is what tells them apart), else the next one
    not yet taken. Returns {"rows": zero-length rows, "settled": rows that
    took a value, "cells": the cells taken}."""
    ending = {}
    for row, (begin, end) in zip(rows, spans):
        if begin is not None and end is not None and end != begin:
            ending[(_stretch(row), end)] = row
    site = defaultdict(list)
    for row in rows_b:
        site[(row[0], row[_KEY_AT])].append(row)
    taken = defaultdict(set)
    found = Counter(rows=0, settled=0, cells=0)
    for row, (begin, end) in zip(rows, spans):
        if begin is None or begin != end:
            continue
        found["rows"] += 1
        key = (row[0], row[_KEY_AT])
        theirs = site.get(key, ())
        free = [i for i in range(len(theirs)) if i not in taken[key]]
        if not free:
            continue
        pick = next((i for i in free
                     if theirs[i][_LANDMARK_AT] == row[_LANDMARK_AT]), free[0])
        taken[key].add(pick)
        before = ending.get((_stretch(row), begin))
        if before is None:
            continue
        other, took = theirs[pick], 0
        for at in _POINT_READ_AT:
            if (row[at] != other[at] and before[at] == other[at]
                    and row[at] != UNAVAILABLE_TOKEN):
                row[at] = before[at]
                took += 1
        if took:
            found["settled"] += 1
            found["cells"] += took
    return dict(found)


def _is_print(path):
    return pdf_source_marker_state(path) != 0


def _load_pair(build_path, export_path, found=None):
    """Both sides, with the ROLE of each enforced: side A must be the app's
    build, side B a consolidated site export (`load_site` refuses the build).
    A print-edition export folds whitespace runs on both sides, as the print's
    HTML does. The build's segments are joined the way the site's file is cut
    (`_join_unchanged`) and its zero-length rows read either stretch that
    meets at their postmile (`_settle_point_rows`); `found["joined"]` and
    `found["points"]` receive the counts for the Notes."""
    pdf = _is_print(export_path)
    rows_a, spans = _read_build(build_path, collapse=pdf)
    _tolls_as_site(rows_a)
    rows_b = crt.load_site(export_path, _PROFILE, pdf=pdf, collapse=pdf)
    rows_a, spans, joined = _join_unchanged(
        rows_a, spans, {(r[0], r[_KEY_AT]) for r in rows_b})
    points = _settle_point_rows(rows_a, spans, rows_b)
    if found is not None:
        found["joined"] = joined
        found["points"] = points
    return (crt.rows_as_written(_PROFILE, ctc.same_source_render_rows(rows_a)),
            crt.rows_as_written(_PROFILE, ctc.same_source_render_rows(rows_b)), None)


# --------------------------------------------------------------------------- #
# Notes
# --------------------------------------------------------------------------- #
def _export_day(path):
    m = re.search(r"(\d{4}-\d{2}-\d{2})", Path(path).name)
    return m.group(1) if m else ""


def _facts(build_path):
    try:
        return cch.report_facts(build_path)
    except Exception as e:   # silent-ok: the Notes sheet is documentation; an unreadable marker costs one stated date, never the comparison (the loaders gate the workbooks themselves)
        log.info("clean-road arcgis notes: build facts unreadable (%s: %s)",
                 type(e).__name__, e)
        return {}


def _joined_line(joined):
    """The Notes paragraph on the segments joined (module docstring)."""
    same, dates = joined.get("same", 0), joined.get("dates", 0)
    return (
        "SEGMENTS: the build cuts a segment at every layer breakpoint; the "
        "site's Highway Log pipeline keeps a boundary only where an attribute "
        "changes (plus its landmarks, county lines, city limits and "
        f"equations). So {same + dates:,} build segments were joined to the "
        f"segment before them: {same:,} begin where nothing this file reports "
        f"changes, {dates:,} where only a block effective date changes (the "
        "site's pipeline does not start a segment on a date). A boundary the "
        "site's file has a row at is never joined, so its cells still compare. "
        "The ArcGIS sheet here therefore has fewer rows than the build "
        "workbook.")


def _points_line(points):
    """The Notes paragraph on the zero-length rows (`_settle_point_rows`)."""
    rows, settled, cells = (points.get(k, 0) for k in ("rows", "settled", "cells"))
    return (
        "POINT ROWS: the site lists every landmark, each side of a postmile "
        "equation and the end of each county the route runs on from as a row "
        "of its own. Where no segment of that roadbed begins there the build "
        f"writes a zero-length row for it ({rows:,} here; the build before "
        "2026-11-18 had none, and listed a landmark of an independent "
        "alignment on both roadbeds). Such a row sits where two stretches "
        "meet, and either stretch's value is right at that point; which one "
        "the site prints turns on a measure difference below its own 1e-5 "
        "tolerance, which no layer file carries. So a cell of a zero-length "
        "row agrees when it shows the stretch that begins there OR the one "
        f"that ends there: {cells:,} cells on {settled:,} of these rows show "
        "the one that ends. A value that is neither stretch's still counts. "
        "The END postmile and length of these rows are compared as written: "
        "the build writes END = BEGIN and length 0.")


def _notes_lines(build_path, export_path, unplaced, joined=None, points=None):
    facts = _facts(build_path)
    asof = facts.get("asof") or "(unknown)"
    drop = facts.get("drop_exported") or "(unknown)"
    day = _export_day(export_path) or "an unrecorded day"
    edition = "print (read back from its PDF)" if _is_print(export_path) else "Excel export"
    counted = set(SHARED_HEADER) - set(CONTEXT_FIELDS)
    return (
        f"Side A ({SIDE_A}) is OUR Clean Road Highway file, rebuilt from the "
        f"ArcGIS layer library as of {asof} (layers exported {drop}). Side B "
        f"({SIDE_B}) is the site's own Clean Road Highway file — its {edition} "
        f"of {day}, consolidated.",
        "VINTAGE: these two dates must match for the counts to measure "
        "correctness. Across a gap they include network change — refresh the "
        "layers (ArcGIS ▸ Layers), rebuild the report as of the export day "
        "(ArcGIS ▸ ArcGIS reports) and compare again.",
        crs.source_note(facts.get("layer_service", ""), export_path),
    ) + crs.notes_lines("clean_highway") + (
        "Both sides are built from the same LRS layers, so a difference here is "
        "a gap in our reconstruction or a finding about how the site builds its "
        "file — never migration drift.",
        "Rows pair on route + county + PM prefix + BEGIN postmile + roadbed (the "
        "R/L/X postmile suffix). The END postmile and length are compared "
        "fields.",
        _joined_line(joined or {}),
        _points_line(points or {}),
        "Every value is shown the way the TSMIS file writes it, ours too. Format "
        "never counts: a date is the same date however it is written, an amount "
        "the same number ('02' = '2'), and a two-letter county the same with or "
        "without its period ('SD.' = 'SD'). THY_TOLL_FOREST_CODE: the site "
        "writes 1 on a forest highway and leaves the cell blank otherwise; the "
        "build's forest code 2 reads as 1 here and its toll code as a blank. "
        "The equation marker 'E' the site writes into THY_PM_SUFFIX_CODE is no "
        "roadbed for the key (the column itself still compares). A print-edition "
        "export folds runs of spaces on both sides, as its HTML does.",
        "THY_LANDMARK_SHORT_DESC: the site writes a BEGIN / END REALIGNMENT "
        "landmark with the row's PM prefix set into the name ('END R "
        "REALIGNMENT'); the layer's own name has none, so the build's reads "
        "that way here. Where a named landmark shares its postmile the site "
        "lists no realignment marker, and neither does the build. An "
        "equation's row carries all the landmarks at its postmile, joined "
        "with ' / ', on both sides.",
        "THY_BREAK_DESC: the site writes only route breaks, off SHS Route Break "
        "('Route Resume' as BEG, 'Route Break' as END). The build marks the same "
        "two points U-BR / R-BR, which read as BEG / END here, and marks other "
        "points the site's file never carries (a route's BEG / END, CNTY, D-C, "
        "DIST, BEGL, ENDR), which read as blank.",
        "CONTEXT (the TSMIS file's own value is shown, never counted; ours is "
        f"on the {SIDE_A} sheet): the 22 columns the site has no source "
        "for (blank in its file by design — the build may still carry a value, "
        "e.g. the ADT trio from Traffic Volume Segments), the two offset columns "
        "(the site writes the LRS odometer, the build its PM-continued "
        "cumulative), THY_LAST_SIG_CHG_DATE, THY_FUNCTIONAL_CLASS_CODE and the "
        "extract date.",
    ) + tuple(unplaced) + (
        "Column match table — each THY column · the ArcGIS layer (# = the "
        "number the vendor's list gives it) and field the build paints it "
        "from · whether the list covers it · counted here or shown as context:",
    ) + crs.highway_table_lines(counted)


# --------------------------------------------------------------------------- #
_BASE_SCHEMA = crt._schema(
    _PROFILE, SIDE_A, SIDE_B,
    notes_title=f"{REPORT_NAME}: comparison notes", notes=(),
    context=_PROFILE.context, representation=_PROFILE.text,
    context_side=SIDE_B, written=True)
_BASE_SCHEMA = replace(
    _BASE_SCHEMA, report_name=_SPEC.label, sides_noun="builds",
    one_sided_note_extra=(" (segments one side cuts at a postmile the other "
                          "doesn't — mostly the two segmentations, or network "
                          "change across the two dates)"))


def _schema_for(build_path, export_path, found):
    """The per-run schema: the Notes sheet reads what the loader found, and a
    build that recorded unplaceable source spans opts into the non-asserting
    anchor token (the vs-TSN comparison's HF-01 rule)."""
    def notes(wb):
        return ctc.make_notes_writer(
            f"{REPORT_NAME}: comparison notes",
            _notes_lines(build_path, export_path, found.get("unplaced", ()),
                         found.get("joined"), found.get("points")))(wb)

    schema = replace(_BASE_SCHEMA, legend_writer=notes)
    skipped, marked, _withheld = cht._build_skip_facts(build_path)
    if skipped or marked:
        schema = replace(schema, unavailable_rule=(
            UNAVAILABLE_TOKEN,
            lambda: "SOURCE COVERAGE — " + cht._coverage_sentence(
                skipped, marked, None)))
    return schema


def suggest_name(_build_path=None):
    return f"ArcGIS_vs_TSMIS_CleanRoadHighway_Comparison {today_str()}.xlsx"


def compare(build_path, export_path, out_path, events=None,
            confirm_overwrite=None, mode="formulas", commit_guard=None):
    """Build the ArcGIS-vs-TSMIS Clean Road Highway discrepancy workbook(s)."""
    found = {}

    def loader(a, b):
        rows_a, rows_b, extra = _load_pair(a, b, found)
        found["unplaced"] = crt.unplaced_note_lines(
            _SPEC, rows_a, rows_b, SIDE_A, SIDE_B)
        return rows_a, rows_b, extra

    return ctc.run_files_compare(
        _schema_for(build_path, export_path, found), build_path, export_path,
        out_path, banner=f"Clean Road Highway Comparison — {SIDE_A} vs {SIDE_B}",
        has_route=True, loader=loader, deps_ok=_DEPS_OK,
        deps_msg="Required components are missing (openpyxl).",
        side_a=SIDE_A, side_b=SIDE_B,
        events=events, confirm_overwrite=confirm_overwrite, mode=mode,
        commit_guard=commit_guard)
