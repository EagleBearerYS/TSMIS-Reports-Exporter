"""Build OUR Clean Road Intersection file (the TASAS "CA INTERSECTIONS" layout)
from the ArcGIS layer library (v0.50.0).

The site's `clean_intersection.js` re-renders its Intersection Detail record
set (`intd_queryIntersections`, the 9.1 dev capture) flat in the 55-column INX_*
layout, one route file at a time. This build projects the same record set from
the layers, following that pipeline rather than the older Intersection Detail
build (`arcgis_report_intersection_detail`, whose leg rule the site has since
replaced):

  * THE ROW UNIVERSE is the `IM Intersection Detail` rows the site's query
    returns: not retired (LRS_DATE_RETIRE empty) and with the geometry live at
    the as-of date (Int_Geometry_Begin_Date on or before it, no end or an end
    after it). A row is listed in its MAIN route's file at the Main_* postmile;
    when its cross route is a state route too, it is listed again in THAT
    route's file at the Cross_* postmile with the main route as its crossing
    route — which is why the site's file carries some INTERSECTION_IDs twice;
  * THE LEGS are picked by route, not by LEG_TYPE (a road-class flag — at a
    state x state intersection every leg is Major): the mainline leg is the one
    whose ROUTE_ID is on the file's route (`SHS_<route>.`), the cross-street
    leg any other; within each, Increasing before Decreasing, the _P roadbed
    before _S, then the lowest APPROACH_ID. One leg's attributes are printed;
  * HIGHWAY GROUP, CITY and POPULATION come from the span layers at the
    intersection's postmile on its own roadbed (`PostmileOverlay` — the site
    translates the postmile to an AR measure and looks the layer up there; an
    intersection carries no LRS measure, so the build stays in postmile space);
  * INX_SEG_ORDER_ID / INX_X_SEG_ORDER_ID are `SHS Inv Network Date.SegOrderId`
    at the intersection on the file's route and at the crossing route's own
    postmile on that route (the primary roadbed, as the site's lookup uses);
  * INX_MAIN_BEGIN_DATE and INX_CROSS_BEGIN_DATE are the layer's own
    Main_Begin_Date / Cross_Begin_Date. The site writes the GEOMETRY begin date
    into both (`ml1Eff` / `interEff` = `dateOfRecord`), so these two columns are
    where the comparison shows the site's mapping, on purpose.

The ten columns the site gives no source (placement ids, create/update, end
date, both ADTs, the LSC date) stay blank. INX_EXTRACT_DATE is the build's
as-of date. Console-free; returns a ConsolidateResult.
"""
from __future__ import annotations

import logging
import sys
from collections import defaultdict

import arcgis_clean_road_build as acb
import arcgis_report_intersection_detail as ari
import city_codes
import clean_road_columns as crc
import clean_road_layers as crl
import paths

log = logging.getLogger(__name__)

REPORT_NAME = "Clean Road Intersection (ArcGIS)"
SPEC = crc.INTERSECTION
BUILD_VERSION = 1
FILENAME = "clean_intersection_from_layers.xlsx"
OUT_DIR = paths.OUTPUT_ROOT / "arcgis_reports"
OUT_PATH = OUT_DIR / FILENAME
SIDECAR_KEY = "arcgis_report_build"
# The lane's surface (`arcgis_reports`): the built sheet's header, and the
# columns the comparison shows but never counts — the ones the site has no
# source for.
HEADER = [acb.ROUTE_COL] + list(SPEC.header)
CONTEXT_COLUMNS = tuple(SPEC.unsourced)
TEXT_COLUMNS = ("INX_INTERSECTION_NAME",)

INX_LAYER = ari.INX_LAYER
SEG_LAYER = ari.SEG_LAYER
APP_LAYER = ari.APP_LAYER
NET_LAYER = "SHS Inv Network Date"
RANGE_LAYERS = {
    "HG": ("SHS Highway Group", "Highway_Group"),
    "CITY": ("City", "City_Code"),
    "POP": ("SHS Population", "Population_Code"),
}
REQUIRED_LAYERS = tuple(sorted(
    {INX_LAYER, SEG_LAYER, APP_LAYER, NET_LAYER}
    | {layer for layer, _attr in RANGE_LAYERS.values()}))

_INX_COLS = [
    "INTERSECTION_ID", "LRS_DATE_RETIRE", "InventoryItemStartDate",
    "County_Code", "District_Code",
    "Main_RouteNum", "Main_RouteSuffix", "Main_PMPrefix", "Main_PMSuffix",
    "Main_PMMeasure", "Main_Begin_Date",
    "Cross_RouteNum", "Cross_RouteSuffix", "Cross_PMPrefix", "Cross_PMSuffix",
    "Cross_PMMeasure", "Cross_Begin_Date",
    "Intersection_Name", "Intersection_Geometry", "Int_Geometry_Begin_Date",
    "Int_Geometry_End_Date", "Intersection_Control", "Int_Control_Begin_Date",
    "Intersection_Lighted_Ind", "Int_Lighted_Ind_Begin_Date",
    "Int_Date_Of_Record",
]
_SEG_COLS = ["INTERSECTION_ID", "APPROACH_ID", "ROUTE_ID", "LEG_TRAVEL_DIR"]
_APP_COLS = ["APPROACH_ID", "Number_Thru_Lanes", "FlowCode", "N_Distance",
             "Left_Channel", "Right_Channel_Ind", "Signal_Arm_Ind"]

_INX = INX_LAYER
_MAIN = f"{APP_LAYER} (the leg on the file's route)"
_CROSS = f"{APP_LAYER} (the leg on another route)"
PROVENANCE = {
    "INX_CONNECTION_ID": (f"{_INX}.INTERSECTION_ID", "verbatim"),
    "INX_MAIN_SEQ_ID": ("(constant)", "1, as the site writes it"),
    "INX_CROSS_SEQ_ID": ("(constant)", "2, as the site writes it"),
    "INX_DESIGN_CODE": (f"{_INX}.Intersection_Geometry", "code"),
    "INX_LIGHTED_BEGIN_DATE": (f"{_INX}.Int_Lighted_Ind_Begin_Date", "MM/DD/YYYY"),
    "INX_LIGHTED_IND": (f"{_INX}.Intersection_Lighted_Ind", "Yes/No -> Y/N"),
    "INX_MAIN_BEGIN_DATE": (f"{_INX}.Main_Begin_Date (Cross_Begin_Date on the "
                            "cross route's row)",
                            "MM/DD/YYYY — the site writes the geometry begin date"),
    "INX_MAIN_SIGNAL_MAST_ARM_IND": (f"{_MAIN}.Signal_Arm_Ind", "Yes/No -> Y/N"),
    "INX_MAIN_LEFT_CHANNEL_CODE": (f"{_MAIN}.Left_Channel", "code"),
    "INX_MAIN_RIGHT_CHANNEL_CODE": (f"{_MAIN}.Right_Channel_Ind", "Yes/No -> Y/N"),
    "INX_MAIN_FLOW_CODE": (f"{_MAIN}.FlowCode", "code"),
    "INX_CROSS_BEGIN_DATE": (f"{_INX}.Cross_Begin_Date (Main_Begin_Date on the "
                             "cross route's row)",
                             "MM/DD/YYYY — the site writes the geometry begin date"),
    "INX_CROSS_SIGNAL_MAST_ARM_IND": (f"{_CROSS}.Signal_Arm_Ind", "Yes/No -> Y/N"),
    "INX_CROSS_LEFT_CHANNEL_CODE": (f"{_CROSS}.Left_Channel", "code"),
    "INX_CROSS_RIGHT_CHANNEL_CODE": (f"{_CROSS}.Right_Channel_Ind", "Yes/No -> Y/N"),
    "INX_CROSS_FLOW_CODE": (f"{_CROSS}.FlowCode", "code"),
    "INX_BEGIN_DATE": (f"{_INX}.InventoryItemStartDate", "MM/DD/YYYY"),
    "INX_RECORD_DATE": (f"{_INX}.Int_Date_Of_Record", "MM/DD/YYYY"),
    "INX_INTERSECTION_NAME": (f"{_INX}.Intersection_Name", "verbatim"),
    "INX_DESIGN_DATE": (f"{_INX}.Int_Geometry_Begin_Date", "MM/DD/YYYY"),
    "INX_CONTROL_DATE": (f"{_INX}.Int_Control_Begin_Date", "MM/DD/YYYY"),
    "INX_CONTROL_CODE": (f"{_INX}.Intersection_Control", "code"),
    "INX_MAIN_LANES_AMT": (f"{_MAIN}.Number_Thru_Lanes", "integer"),
    "INX_MAIN_OVERRIDE_LENGTH_AMT": (f"{_MAIN}.N_Distance", "integer"),
    "INX_CROSS_LANES_AMT": (f"{_CROSS}.Number_Thru_Lanes", "integer"),
    "INX_CROSS_OVERRIDE_LENGTH_AMT": (f"{_CROSS}.N_Distance", "integer"),
    "INX_ROUTE_NAME": (f"{_INX}.Main_RouteNum (Cross_RouteNum on the cross "
                       "route's row)", "3 digits"),
    "INX_BEGIN_PM_AMT": (f"{_INX}.Main_PMMeasure (Cross_PMMeasure on the cross "
                         "route's row)", "natural decimal"),
    "INX_END_PM_AMT": (f"{_INX}.Main_PMMeasure", "a point: begin = end"),
    "INX_DISTRICT_CODE": (f"{_INX}.District_Code", "2 digits"),
    "INX_COUNTY_CODE": (f"{_INX}.County_Code", "TASAS county code"),
    "INX_PM_PREFIX_CODE": (f"{_INX}.Main_PMPrefix", "code ('.' = none)"),
    "INX_PM_SUFFIX_CODE": (f"{_INX}.Main_PMSuffix", "code ('.' = none)"),
    "INX_ROUTE_SUFFIX_CODE": (f"{_INX}.Main_RouteSuffix", "code ('.' = none)"),
    "INX_SEG_ORDER_ID": (f"{NET_LAYER}.SegOrderId",
                         "span at the intersection's postmile, primary roadbed"),
    "INX_HIGHWAY_GROUP": ("SHS Highway Group.Highway_Group",
                          "span at the postmile on the intersection's roadbed"),
    "INX_CITY_CODE": ("City.City_Code",
                      "span at the postmile on the intersection's roadbed, TASAS code"),
    "INX_POPULATION_GROUP": ("SHS Population.Population_Code",
                             "span at the postmile on the intersection's roadbed"),
    "INX_X_ROUTE_NAME": (f"{_INX}.Cross_RouteNum (Main_RouteNum on the cross "
                         "route's row)", "3 digits"),
    "INX_X_PM_PREFIX_CODE": (f"{_INX}.Cross_PMPrefix", "code ('.' = none)"),
    "INX_X_POSTMILE_AMT": (f"{_INX}.Cross_PMMeasure", "natural decimal"),
    "INX_X_PM_SUFFIX_CODE": (f"{_INX}.Cross_PMSuffix", "code ('.' = none)"),
    "INX_X_ROUTE_SUFFIX_CODE": (f"{_INX}.Cross_RouteSuffix", "code ('.' = none)"),
    "INX_X_SEG_ORDER_ID": (f"{NET_LAYER}.SegOrderId",
                           "span at the crossing route's own postmile"),
    "INX_EXTRACT_DATE": ("(the build's as-of date)", "MM/DD/YYYY"),
}
assert set(PROVENANCE) | set(SPEC.unsourced) == set(SPEC.header), \
    "every Clean Road Intersection column is either sourced or unsourced"


# --------------------------------------------------------------------------- #
# the legs
# --------------------------------------------------------------------------- #
def _read_legs(lib, index, events):
    """({intersection id: [segment rows]}, {approach id: detail row})."""
    entry = index.get(APP_LAYER) or {}
    detail = {}
    for r in crl.stream_layer(lib["present"][APP_LAYER], _APP_COLS,
                              layer_name=APP_LAYER, expected_rows=entry.get("rows")):
        detail.setdefault(acb.s(r["APPROACH_ID"]), r)
    entry = index.get(SEG_LAYER) or {}
    legs = defaultdict(list)
    for r in crl.stream_layer(lib["present"][SEG_LAYER], _SEG_COLS,
                              layer_name=SEG_LAYER, expected_rows=entry.get("rows")):
        legs[acb.s(r["INTERSECTION_ID"])].append(r)
    events.on_log(f"  {APP_LAYER}: {len(detail):,} approaches on "
                  f"{len(legs):,} intersections")
    return legs, detail


def _leg_rank(leg):
    """The site's pick order: Increasing first, the _P roadbed before _S before
    any other, then the lowest APPROACH_ID."""
    rid = acb.s(leg["ROUTE_ID"])
    align = 0 if rid.endswith("_P") else 1 if rid.endswith("_S") else 2
    try:
        approach = int(acb.s(leg["APPROACH_ID"]))
    except ValueError:  # silent-ok: a non-numeric approach id sorts after the numeric ones, as Number() → NaN never wins the site's comparison
        approach = float("inf")
    return (0 if acb.s(leg["LEG_TRAVEL_DIR"]) == "Increasing" else 1, align, approach)


def _pick_legs(legs, route):
    """(mainline leg detail, cross-street leg detail) for the file of `route`
    (the 3-digit token plus any suffix, as the site pads it)."""
    on = f"SHS_{route}."
    mine = [leg for leg in legs if acb.s(leg["ROUTE_ID"]).startswith(on)]
    other = [leg for leg in legs if not acb.s(leg["ROUTE_ID"]).startswith(on)]
    pick = (lambda group: min(group, key=_leg_rank) if group else None)
    return pick(mine), pick(other)


def _leg_values(leg, detail):
    d = detail.get(acb.s(leg["APPROACH_ID"])) if leg is not None else None
    if d is None:
        return {"sm": "", "lc": "", "rc": "", "tf": "", "nl": "", "len": ""}
    return {"sm": acb.yn(d["Signal_Arm_Ind"]), "lc": acb.code(d["Left_Channel"]),
            "rc": acb.yn(d["Right_Channel_Ind"]), "tf": acb.code(d["FlowCode"]),
            "nl": acb.integer(d["Number_Thru_Lanes"]),
            "len": acb.integer(d["N_Distance"])}


# --------------------------------------------------------------------------- #
# one row
# --------------------------------------------------------------------------- #
def _side(r, which):
    """The location fields of one end of the intersection row ("Main"/"Cross")."""
    return {"num": r[f"{which}_RouteNum"], "sfx": r[f"{which}_RouteSuffix"],
            "pfx": r[f"{which}_PMPrefix"], "pmsfx": r[f"{which}_PMSuffix"],
            "pm": r[f"{which}_PMMeasure"], "begin": r[f"{which}_Begin_Date"]}


def _project(r, here, there, legs, detail, overlays, net, extract):
    """One row in the file of `here`'s route, crossing `there`."""
    route = acb.route_token(here["num"], here["sfx"])
    county = crl.norm_county(r["County_Code"])
    prefix = acb.pfx(here["pfx"])
    pm_units = crl.pm_units(here["pm"])
    pm = None if pm_units is None else crl.pm_float(pm_units)
    left = acb.pfx(here["pmsfx"]) == "L"
    main_leg, cross_leg = _pick_legs(legs, route)
    ml, cs = _leg_values(main_leg, detail), _leg_values(cross_leg, detail)

    def at(tag):
        return overlays[tag].value_at(route, county, prefix, pm, left)

    x_route = acb.s(there["num"])
    x_pm_units = crl.pm_units(there["pm"])
    x_seg = ""
    if x_route and x_pm_units is not None:
        x_seg = net.value_at(acb.route_token(there["num"], there["sfx"]), county,
                             acb.pfx(there["pfx"]), crl.pm_float(x_pm_units))
    pm_text = acb.num(here["pm"]) if pm_units is not None else ""
    v = {
        "INX_CONNECTION_ID": acb.integer(r["INTERSECTION_ID"]),
        "INX_MAIN_SEQ_ID": "1",
        "INX_CROSS_SEQ_ID": "2",
        "INX_DESIGN_CODE": acb.code(r["Intersection_Geometry"]),
        "INX_LIGHTED_BEGIN_DATE": acb.mdy(r["Int_Lighted_Ind_Begin_Date"]),
        "INX_LIGHTED_IND": acb.yn(r["Intersection_Lighted_Ind"]),
        "INX_MAIN_BEGIN_DATE": acb.mdy(here["begin"]),
        "INX_MAIN_SIGNAL_MAST_ARM_IND": ml["sm"],
        "INX_MAIN_LEFT_CHANNEL_CODE": ml["lc"],
        "INX_MAIN_RIGHT_CHANNEL_CODE": ml["rc"],
        "INX_MAIN_FLOW_CODE": ml["tf"],
        "INX_CROSS_BEGIN_DATE": acb.mdy(there["begin"]),
        "INX_CROSS_SIGNAL_MAST_ARM_IND": cs["sm"],
        "INX_CROSS_LEFT_CHANNEL_CODE": cs["lc"],
        "INX_CROSS_RIGHT_CHANNEL_CODE": cs["rc"],
        "INX_CROSS_FLOW_CODE": cs["tf"],
        "INX_BEGIN_DATE": acb.mdy(r["InventoryItemStartDate"]),
        "INX_RECORD_DATE": acb.mdy(r["Int_Date_Of_Record"]),
        "INX_INTERSECTION_NAME": acb.s(r["Intersection_Name"]),
        "INX_DESIGN_DATE": acb.mdy(r["Int_Geometry_Begin_Date"]),
        "INX_CONTROL_DATE": acb.mdy(r["Int_Control_Begin_Date"]),
        "INX_CONTROL_CODE": acb.code(r["Intersection_Control"]),
        "INX_MAIN_LANES_AMT": ml["nl"],
        "INX_MAIN_OVERRIDE_LENGTH_AMT": ml["len"],
        "INX_CROSS_LANES_AMT": cs["nl"],
        "INX_CROSS_OVERRIDE_LENGTH_AMT": cs["len"],
        "INX_ROUTE_NAME": crl.norm_route(here["num"]),
        "INX_BEGIN_PM_AMT": pm_text,
        "INX_END_PM_AMT": pm_text,
        "INX_DISTRICT_CODE": crl.norm_district(r["District_Code"]),
        "INX_COUNTY_CODE": county,
        "INX_PM_PREFIX_CODE": prefix,
        "INX_PM_SUFFIX_CODE": acb.pfx(here["pmsfx"]),
        "INX_ROUTE_SUFFIX_CODE": acb.pfx(here["sfx"]),
        "INX_SEG_ORDER_ID": acb.integer(net.value_at(route, county, prefix, pm)),
        "INX_HIGHWAY_GROUP": crl.code_of(at("HG")),
        "INX_CITY_CODE": city_codes.norm_city(at("CITY")),
        "INX_POPULATION_GROUP": crl.code_of(at("POP")),
        "INX_X_ROUTE_NAME": crl.norm_route(x_route) if x_route else "",
        "INX_X_PM_PREFIX_CODE": acb.pfx(there["pfx"]),
        "INX_X_POSTMILE_AMT": acb.num(there["pm"]) if acb.s(there["pm"]) else "",
        "INX_X_PM_SUFFIX_CODE": acb.pfx(there["pmsfx"]),
        "INX_X_ROUTE_SUFFIX_CODE": acb.pfx(there["sfx"]),
        "INX_X_SEG_ORDER_ID": acb.integer(x_seg) if x_seg else "",
        "INX_EXTRACT_DATE": extract,
    }
    return [route] + [v.get(c, "") for c in SPEC.header]


def _live(r, asof):
    """The site's universe filter: not retired, the geometry live at `asof`."""
    if acb.s(r["LRS_DATE_RETIRE"]):
        return False
    begin = crl.to_serial(r["Int_Geometry_Begin_Date"])
    end = crl.to_serial(r["Int_Geometry_End_Date"])
    if begin is None or begin > asof:
        return False
    return end is None or end > asof


def build_rows(lib, index, *, asof, asof_day, events):
    """(rows, stats): one row per intersection per state route it is on."""
    events.on_log("  reading the span layers the file looks intersections up in...")

    def read(layer, attr, **dates):
        entry = index.get(layer) or {}
        ov = acb.PostmileOverlay.read(lib["present"][layer], layer, attr,
                                      asof=asof, expected_rows=entry.get("rows"),
                                      **dates)
        events.on_log(f"  {layer}: read")
        return ov

    overlays = {tag: read(layer, attr) for tag, (layer, attr) in RANGE_LAYERS.items()}
    net = read(NET_LAYER, "SegOrderId", start="Network_Start_Date",
               end="Network_End_Date")
    legs, detail = _read_legs(lib, index, events)
    extract = f"{asof_day.month:02d}/{asof_day.day:02d}/{asof_day.year:04d}"

    entry = index.get(INX_LAYER) or {}
    rows, not_live, crossed = [], 0, 0
    for r in crl.stream_layer(lib["present"][INX_LAYER], _INX_COLS,
                              layer_name=INX_LAYER, expected_rows=entry.get("rows")):
        if events.is_cancelled():
            return [], {}
        if not acb.s(r["Main_RouteNum"]) or not acb.s(r["Main_PMMeasure"]) \
                or not _live(r, asof):
            not_live += 1
            continue
        mine = legs.get(acb.s(r["INTERSECTION_ID"]), ())
        main, cross = _side(r, "Main"), _side(r, "Cross")
        rows.append(_project(r, main, cross, mine, detail, overlays, net, extract))
        if acb.s(cross["num"]) and acb.s(cross["pm"]) \
                and acb.route_token(cross["num"], cross["sfx"]) \
                != acb.route_token(main["num"], main["sfx"]):
            crossed += 1
            rows.append(_project(r, cross, main, mine, detail, overlays, net, extract))

    pm_at = 1 + SPEC.header.index("INX_BEGIN_PM_AMT")
    county_at = 1 + SPEC.header.index("INX_COUNTY_CODE")
    rows.sort(key=lambda row: (row[0], row[county_at],
                               crl.pm_units(row[pm_at]) or 0))
    stats = {"rows": len(rows), "routes": len({row[0] for row in rows}),
             "not_live": not_live, "crossing_rows": crossed}
    events.on_log(f"  built {len(rows):,} intersection records "
                  f"({stats['routes']} routes; {crossed:,} listed again under "
                  f"their crossing state route); {not_live:,} layer rows are not "
                  "a live intersection on a route")
    return rows, stats


def marker_rows(stats):
    return [
        ("Layer rows not live on a route", stats.get("not_live", 0)),
        ("Rows listed again under the crossing state route",
         stats.get("crossing_rows", 0)),
        ("Row universe", "IM Intersection Detail rows with a main route and "
                         "postmile, LRS_DATE_RETIRE empty and the geometry live at the "
                         "as-of date; a state x state intersection is listed in "
                         "both routes' files, as the site lists it"),
        ("Leg rule", "the mainline block from the leg on the file's route, the "
                     "cross-street block from a leg on another route; within "
                     "each, Increasing, then the _P roadbed, then the lowest "
                     "APPROACH_ID (the site's rule)"),
        ("Columns with no source", ", ".join(SPEC.unsourced)),
    ]


def is_arcgis_report(path):
    return acb.is_build(path, SPEC)


def report_facts(path):
    return acb.report_facts(path)


def consolidate(events=None, confirm_overwrite=None, day=None, *,
                out_path=None, asof=None, lib_root=None):
    """Build the Clean Road Intersection file FROM THE ARCGIS LAYERS as of
    `asof` (set it to the export day being compared)."""
    del day                       # ConsolidateWorker passes it; no run days here
    return acb.consolidate(sys.modules[__name__], events, confirm_overwrite,
                           out_path=out_path, asof=asof, lib_root=lib_root)
