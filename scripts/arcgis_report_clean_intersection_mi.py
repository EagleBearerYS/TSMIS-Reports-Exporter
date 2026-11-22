"""Build OUR Clean Road Intersection file from the Master Intersection layer
(2026-11-22): the one source the vendor's source list names for the file.

The vendor's list names a flat `Master Intersection` layer for all 32
Intersection attributes. The site's own code reads the IM tables instead, and
`arcgis_report_clean_intersection` follows the site. This build follows the
list: every value comes from the intersection's one Master Intersection row
(layer 204 of the production csms_supporting_prod service, library layer 44),
so its comparison holds the site's file to the source the vendor names. The
owner's call (2026-11-22): it sits beside the IM-tables comparison, labeled,
and it is the default for problems raised with the vendor.

  * THE ROW UNIVERSE is the layer's rows that are current at the as-of date
    (an inventory start on or before it, or none; no inventory end, or one
    after it) with a major-leg route and postmile and the geometry in effect
    (Intersection_Geometry_Eff_Date on or before it, the date the site's own
    filter reads). The layer keeps the rows a rebuild replaced, with an end
    date, so an as-of before its latest rebuild reads that day's rows;
  * a row is listed in its MAJOR leg's route file at the major postmile. When
    the minor leg is a state route too, it is listed again in that route's
    file at the minor postmile, with the major route as its crossing route,
    as the site lists a state x state intersection;
  * the leg on the file's route fills the INX_MAIN_* block and the other leg
    the INX_CROSS_* block: the leg's own effective date, signal mast arm,
    channelization, flow and override length, and its Total_Num_Lanes (the
    leg's total agrees with the site's lanes column on 99% of the 2026-11-19
    rows, Pri and Sec on 2%). Highway group and population are that leg's
    own; district and county are the major leg's on both rows (the layer has
    no minor district, and the site writes one county per intersection).

What this build does not take from Master Intersection is shown in the
comparison and never counted (`EXTRA_CONTEXT`): the two segment-order ids and
INX_BEGIN_DATE, which the layer has no field for (its inventory start date is
its own record's), and INX_CITY_CODE, which the list does not attribute to it.
Console-free; returns a ConsolidateResult.
"""
from __future__ import annotations

import logging
import sys
from dataclasses import replace

import arcgis_clean_road_build as acb
import clean_road_columns as crc
import clean_road_layers as crl
import clean_road_sources as crs
import paths

log = logging.getLogger(__name__)

REPORT_NAME = "Clean Road Intersection (Master Intersection)"
# The site file's own shape, under a key of its own: a Master Intersection
# build and an IM-tables build can never stand for each other.
SPEC = replace(crc.INTERSECTION, key="clean_intersection_mi",
               label="Clean Road Intersection (Master Intersection)")
# The export it is compared against, read through that report's profile.
PROFILE_KEY = crc.INTERSECTION.key
BUILD_VERSION = 1
FILENAME = "clean_intersection_mi_from_layers.xlsx"
OUT_DIR = paths.OUTPUT_ROOT / "arcgis_reports"
OUT_PATH = OUT_DIR / FILENAME
SIDECAR_KEY = "arcgis_report_build"
HEADER = [acb.ROUTE_COL] + list(SPEC.header)
CONTEXT_COLUMNS = tuple(SPEC.unsourced)
TEXT_COLUMNS = ("INX_INTERSECTION_NAME",)

MI_LAYER = crs.INTERSECTION_SOURCE
REQUIRED_LAYERS = (MI_LAYER,)

EXTRA_CONTEXT = ("INX_BEGIN_DATE", "INX_SEG_ORDER_ID", "INX_X_SEG_ORDER_ID",
                 "INX_CITY_CODE")
EXTRA_CONTEXT_WHY = ("the layer has no field for the first three (its inventory "
                     "start date is its own record's), and the vendor's list does "
                     "not attribute the city to it; the build writes the file "
                     "route's leg's own City")
ID_NOTE = ("INX_CONNECTION_ID is the layer's Master_Intersection_ID, the IM "
           "tables' INTERSECTION_ID, so it counts.")
TABLE_NOTE = ("Where each column comes from in the build (Master Intersection "
              "#204 is its number on the production service; the list's "
              "development copy numbers it 0):")

_LEG_FIELDS = ("RouteNum", "RouteSuffix", "PMPrefix", "PMSuffix", "PMMeasure",
               "Effective_Date", "Signal_Arm_Ind", "Left_Channel",
               "Right_Channel_Ind", "FlowCode", "Total_Num_Lanes", "N_Distance",
               "Highway_Group", "Population_Code", "City")
_COLS = (["Master_Intersection_ID", "Date_Of_Record", "Intersection_Geometry",
          "Intersection_Geometry_Eff_Date", "Intersection_Control",
          "Intersection_Control_Eff_Date", "Intersection_Lighted_Ind",
          "Intersection_Lighted_Eff_Date", "Description", "Major_District",
          "Major_County", "InventoryItemStartDate", "InventoryItemEndDate"]
         + [f"{leg}_{f}" for leg in ("Major", "Minor") for f in _LEG_FIELDS])
assert set(_COLS) <= set(crs.MASTER_INTERSECTION_HEADER), "a field the layer lacks"


def _leg_source(field, major=True):
    """'Master Intersection.Major_<field> (Minor_<field> on the row listed
    under the minor state route)': the leg on the file's route, or with
    `major` False the crossing leg."""
    first, then = ("Major", "Minor") if major else ("Minor", "Major")
    return (f"{MI_LAYER}.{first}_{field} ({then}_{field} on the row listed under "
            "the minor state route)")


_BLANK = ("(no Master Intersection field)", "blank; shown, never counted")
PROVENANCE = {
    "INX_CONNECTION_ID": (f"{MI_LAYER}.Master_Intersection_ID",
                          "verbatim (the IM tables' INTERSECTION_ID)"),
    "INX_MAIN_SEQ_ID": ("(constant)", "1, as the site writes it"),
    "INX_CROSS_SEQ_ID": ("(constant)", "2, as the site writes it"),
    "INX_DESIGN_CODE": (f"{MI_LAYER}.Intersection_Geometry", "code"),
    "INX_LIGHTED_BEGIN_DATE": (f"{MI_LAYER}.Intersection_Lighted_Eff_Date",
                               "MM/DD/YYYY"),
    "INX_LIGHTED_IND": (f"{MI_LAYER}.Intersection_Lighted_Ind", "Yes/No -> Y/N"),
    "INX_MAIN_BEGIN_DATE": (_leg_source("Effective_Date"), "MM/DD/YYYY"),
    "INX_MAIN_SIGNAL_MAST_ARM_IND": (_leg_source("Signal_Arm_Ind"), "Yes/No -> Y/N"),
    "INX_MAIN_LEFT_CHANNEL_CODE": (_leg_source("Left_Channel"), "code"),
    "INX_MAIN_RIGHT_CHANNEL_CODE": (_leg_source("Right_Channel_Ind"), "Yes/No -> Y/N"),
    "INX_MAIN_FLOW_CODE": (_leg_source("FlowCode"), "code"),
    "INX_CROSS_BEGIN_DATE": (_leg_source("Effective_Date", False), "MM/DD/YYYY"),
    "INX_CROSS_SIGNAL_MAST_ARM_IND": (_leg_source("Signal_Arm_Ind", False),
                                      "Yes/No -> Y/N"),
    "INX_CROSS_LEFT_CHANNEL_CODE": (_leg_source("Left_Channel", False), "code"),
    "INX_CROSS_RIGHT_CHANNEL_CODE": (_leg_source("Right_Channel_Ind", False),
                                     "Yes/No -> Y/N"),
    "INX_CROSS_FLOW_CODE": (_leg_source("FlowCode", False), "code"),
    "INX_BEGIN_DATE": _BLANK,
    "INX_RECORD_DATE": (f"{MI_LAYER}.Date_Of_Record", "MM/DD/YYYY"),
    "INX_INTERSECTION_NAME": (f"{MI_LAYER}.Description", "verbatim"),
    "INX_DESIGN_DATE": (f"{MI_LAYER}.Intersection_Geometry_Eff_Date", "MM/DD/YYYY"),
    "INX_CONTROL_DATE": (f"{MI_LAYER}.Intersection_Control_Eff_Date", "MM/DD/YYYY"),
    "INX_CONTROL_CODE": (f"{MI_LAYER}.Intersection_Control", "code"),
    "INX_MAIN_LANES_AMT": (_leg_source("Total_Num_Lanes"), "integer"),
    "INX_MAIN_OVERRIDE_LENGTH_AMT": (_leg_source("N_Distance"), "integer"),
    "INX_CROSS_LANES_AMT": (_leg_source("Total_Num_Lanes", False), "integer"),
    "INX_CROSS_OVERRIDE_LENGTH_AMT": (_leg_source("N_Distance", False), "integer"),
    "INX_ROUTE_NAME": (_leg_source("RouteNum"), "3 digits"),
    "INX_BEGIN_PM_AMT": (_leg_source("PMMeasure"), "natural decimal"),
    "INX_END_PM_AMT": (_leg_source("PMMeasure"), "a point: begin = end"),
    "INX_DISTRICT_CODE": (f"{MI_LAYER}.Major_District",
                          "2 digits, on both rows (the layer has no minor district)"),
    "INX_COUNTY_CODE": (f"{MI_LAYER}.Major_County",
                        "TASAS county code, on both rows: one county per "
                        "intersection, as the site writes it"),
    "INX_PM_PREFIX_CODE": (_leg_source("PMPrefix"), "code ('.' = none)"),
    "INX_PM_SUFFIX_CODE": (_leg_source("PMSuffix"), "code ('.' = none)"),
    "INX_ROUTE_SUFFIX_CODE": (_leg_source("RouteSuffix"), "code ('.' = none)"),
    "INX_SEG_ORDER_ID": _BLANK,
    "INX_HIGHWAY_GROUP": (_leg_source("Highway_Group"), "code"),
    "INX_CITY_CODE": (_leg_source("City"),
                      "verbatim; not on the vendor's list, so shown, never counted"),
    "INX_POPULATION_GROUP": (_leg_source("Population_Code"), "code"),
    "INX_X_ROUTE_NAME": (_leg_source("RouteNum", False), "3 digits"),
    "INX_X_PM_PREFIX_CODE": (_leg_source("PMPrefix", False), "code ('.' = none)"),
    "INX_X_POSTMILE_AMT": (_leg_source("PMMeasure", False), "natural decimal"),
    "INX_X_PM_SUFFIX_CODE": (_leg_source("PMSuffix", False), "code ('.' = none)"),
    "INX_X_ROUTE_SUFFIX_CODE": (_leg_source("RouteSuffix", False), "code ('.' = none)"),
    "INX_X_SEG_ORDER_ID": _BLANK,
    "INX_EXTRACT_DATE": ("(the build's as-of date)", "MM/DD/YYYY"),
}
assert set(PROVENANCE) | set(SPEC.unsourced) == set(SPEC.header), \
    "every Clean Road Intersection column is either sourced or unsourced"
assert set(EXTRA_CONTEXT) <= set(SPEC.header) - set(SPEC.unsourced)


# --------------------------------------------------------------------------- #
# one row
# --------------------------------------------------------------------------- #
def _leg(r, which):
    """One leg's fields, keyed by their name without the Major_/Minor_ prefix."""
    return {f: r[f"{which}_{f}"] for f in _LEG_FIELDS}


def _project(r, here, there, extract):
    """One row in the file of `here`'s route, crossing `there`."""
    route = acb.route_token(here["RouteNum"], here["RouteSuffix"])
    pm_text = (acb.num(here["PMMeasure"])
               if crl.pm_units(here["PMMeasure"]) is not None else "")
    x_route = acb.s(there["RouteNum"])
    v = {
        "INX_CONNECTION_ID": acb.integer(r["Master_Intersection_ID"]),
        "INX_MAIN_SEQ_ID": "1",
        "INX_CROSS_SEQ_ID": "2",
        "INX_DESIGN_CODE": acb.code(r["Intersection_Geometry"]),
        "INX_LIGHTED_BEGIN_DATE": acb.mdy(r["Intersection_Lighted_Eff_Date"]),
        "INX_LIGHTED_IND": acb.yn(r["Intersection_Lighted_Ind"]),
        "INX_MAIN_BEGIN_DATE": acb.mdy(here["Effective_Date"]),
        "INX_MAIN_SIGNAL_MAST_ARM_IND": acb.yn(here["Signal_Arm_Ind"]),
        "INX_MAIN_LEFT_CHANNEL_CODE": acb.code(here["Left_Channel"]),
        "INX_MAIN_RIGHT_CHANNEL_CODE": acb.yn(here["Right_Channel_Ind"]),
        "INX_MAIN_FLOW_CODE": acb.code(here["FlowCode"]),
        "INX_CROSS_BEGIN_DATE": acb.mdy(there["Effective_Date"]),
        "INX_CROSS_SIGNAL_MAST_ARM_IND": acb.yn(there["Signal_Arm_Ind"]),
        "INX_CROSS_LEFT_CHANNEL_CODE": acb.code(there["Left_Channel"]),
        "INX_CROSS_RIGHT_CHANNEL_CODE": acb.yn(there["Right_Channel_Ind"]),
        "INX_CROSS_FLOW_CODE": acb.code(there["FlowCode"]),
        "INX_RECORD_DATE": acb.mdy(r["Date_Of_Record"]),
        "INX_INTERSECTION_NAME": acb.s(r["Description"]),
        "INX_DESIGN_DATE": acb.mdy(r["Intersection_Geometry_Eff_Date"]),
        "INX_CONTROL_DATE": acb.mdy(r["Intersection_Control_Eff_Date"]),
        "INX_CONTROL_CODE": acb.code(r["Intersection_Control"]),
        "INX_MAIN_LANES_AMT": acb.integer(here["Total_Num_Lanes"]),
        "INX_MAIN_OVERRIDE_LENGTH_AMT": acb.integer(here["N_Distance"]),
        "INX_CROSS_LANES_AMT": acb.integer(there["Total_Num_Lanes"]),
        "INX_CROSS_OVERRIDE_LENGTH_AMT": acb.integer(there["N_Distance"]),
        "INX_ROUTE_NAME": crl.norm_route(here["RouteNum"]),
        "INX_BEGIN_PM_AMT": pm_text,
        "INX_END_PM_AMT": pm_text,
        "INX_DISTRICT_CODE": crl.norm_district(r["Major_District"]),
        "INX_COUNTY_CODE": crl.norm_county(r["Major_County"]),
        "INX_PM_PREFIX_CODE": acb.pfx(here["PMPrefix"]),
        "INX_PM_SUFFIX_CODE": acb.pfx(here["PMSuffix"]),
        "INX_ROUTE_SUFFIX_CODE": acb.pfx(here["RouteSuffix"]),
        "INX_HIGHWAY_GROUP": acb.code(here["Highway_Group"]),
        "INX_CITY_CODE": acb.s(here["City"]).upper(),
        "INX_POPULATION_GROUP": acb.code(here["Population_Code"]),
        "INX_X_ROUTE_NAME": crl.norm_route(x_route) if x_route else "",
        "INX_X_PM_PREFIX_CODE": acb.pfx(there["PMPrefix"]),
        "INX_X_POSTMILE_AMT": (acb.num(there["PMMeasure"])
                               if acb.s(there["PMMeasure"]) else ""),
        "INX_X_PM_SUFFIX_CODE": acb.pfx(there["PMSuffix"]),
        "INX_X_ROUTE_SUFFIX_CODE": acb.pfx(there["RouteSuffix"]),
        "INX_EXTRACT_DATE": extract,
    }
    return [route] + [v.get(c, "") for c in SPEC.header]


def _current(r, asof):
    """The row universe: the row current at `asof`, the geometry in effect."""
    start = crl.to_serial(r["InventoryItemStartDate"])
    end = crl.to_serial(r["InventoryItemEndDate"])
    if (start is not None and start > asof) or (end is not None and end <= asof):
        return False
    effective = crl.to_serial(r["Intersection_Geometry_Eff_Date"])
    return effective is not None and effective <= asof


def build_rows(lib, index, *, asof, asof_day, events):
    """(rows, stats): one row per intersection per state route it is on."""
    entry = index.get(MI_LAYER) or {}
    extract = f"{asof_day.month:02d}/{asof_day.day:02d}/{asof_day.year:04d}"
    rows, not_current, no_route, crossed = [], 0, 0, 0
    for r in crl.stream_layer(lib["present"][MI_LAYER], _COLS, layer_name=MI_LAYER,
                              expected_rows=entry.get("rows")):
        if events.is_cancelled():
            return [], {}
        if not _current(r, asof):
            not_current += 1
            continue
        major, minor = _leg(r, "Major"), _leg(r, "Minor")
        if not acb.s(major["RouteNum"]) or crl.pm_units(major["PMMeasure"]) is None:
            no_route += 1
            continue
        rows.append(_project(r, major, minor, extract))
        if acb.s(minor["RouteNum"]) and crl.pm_units(minor["PMMeasure"]) is not None \
                and acb.route_token(minor["RouteNum"], minor["RouteSuffix"]) \
                != acb.route_token(major["RouteNum"], major["RouteSuffix"]):
            crossed += 1
            rows.append(_project(r, minor, major, extract))

    pm_at = 1 + SPEC.header.index("INX_BEGIN_PM_AMT")
    county_at = 1 + SPEC.header.index("INX_COUNTY_CODE")
    rows.sort(key=lambda row: (row[0], row[county_at],
                               crl.pm_units(row[pm_at]) or 0))
    stats = {"rows": len(rows), "routes": len({row[0] for row in rows}),
             "not_current": not_current, "no_route": no_route,
             "crossing_rows": crossed}
    events.on_log(f"  built {len(rows):,} intersection records from {MI_LAYER} "
                  f"({stats['routes']} routes; {crossed:,} listed again under the "
                  f"minor state route); {not_current:,} layer rows are not current "
                  f"at the as-of date and {no_route:,} have no major route")
    return rows, stats


def marker_rows(stats):
    return [
        ("Source", f"{MI_LAYER} (library layer 44; layer 204 of the production "
                   "csms_supporting_prod service), the source the vendor's list "
                   "names for this file"),
        ("Layer rows not current at the as-of date", stats.get("not_current", 0)),
        ("Current rows with no major route", stats.get("no_route", 0)),
        ("Rows listed again under the minor state route",
         stats.get("crossing_rows", 0)),
        ("Row universe", f"{MI_LAYER} rows current at the as-of date (inventory "
                         "start on or before it, no end or an end after it) with a "
                         "major route and postmile and the geometry in effect; a "
                         "state x state intersection is listed in both routes' "
                         "files, as the site lists it"),
        ("Leg rule", "the leg on the file's route fills the INX_MAIN_* block and "
                     "the other leg the INX_CROSS_* block"),
        ("Columns with no source", ", ".join(SPEC.unsourced)),
        ("Not from Master Intersection (never counted)", ", ".join(EXTRA_CONTEXT)),
    ]


def is_arcgis_report(path):
    return acb.is_build(path, SPEC)


def report_facts(path):
    return acb.report_facts(path)


def consolidate(events=None, confirm_overwrite=None, day=None, *,
                out_path=None, asof=None, lib_root=None):
    """Build the Clean Road Intersection file FROM THE MASTER INTERSECTION
    LAYER as of `asof` (set it to the export day being compared)."""
    del day                       # ConsolidateWorker passes it; no run days here
    return acb.consolidate(sys.modules[__name__], events, confirm_overwrite,
                           out_path=out_path, asof=asof, lib_root=lib_root)
