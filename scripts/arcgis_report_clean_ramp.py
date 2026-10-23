"""Build OUR Clean Road Ramp file (the TASAS "CA RAMPS" layout) from the ArcGIS
layer library (v0.50.0).

The site's `clean_ramp.js` re-renders its Ramp Detail record set flat in the
34-column RAM_* layout; this build projects the same record set from the
layers:

  * the ramp universe and its location are `SHS Ramp Pt` (the site's layer
    132): the as-of point rows, one per Ramp_Name, with the district / county /
    route / suffix / postmile prefix, measure and suffix the row carries, and
    its inventory start / end dates as RAM_BEGIN_DATE / RAM_END_DATE;
  * the ramp attributes are `SHS Ramp` (layer 131), joined on Ramp_Name: the
    design code and its domain name (RAM_DESIGN_DESC, written whole as the
    site writes it: 'D- Diamond Type Ramp'), the description, the
    on/off indicator (0 OFF, 1 ON, 2 Z) and Area 4 (Y/N);
  * the highway the ramp sits on gives Highway Group, City, Population and
    `SHS Inv Network Date.SegOrderId`, looked up the way the site looks them up
    (`queryRangeLayer` / `hl_queryLayer122Field`): at the ramp point's own AR
    measure on its own RouteID — a secondary-alignment ramp tries its `_S` route
    first — with the current calibration preferred (`MeasureOverlay`). Unlike an
    intersection, a ramp point carries its AR measure, so no postmile sampling is
    needed;
  * RAM_PRIMARY_DIRECTION_CODE is the `Route Direction` table's leading letter
    — the route's TO direction for a ramp on the right (primary) alignment, its
    FROM direction on the left — and RAM_NETWORK_ID the site's constant 2.

The 13 columns the site gives no source stay blank, except
RAM_BEGIN_OFFSET_AMT: the business requirement asks for the odometer and the
Ramp Pt row carries it (ODMeasure), so the build writes it — the comparison
shows it as context, never counted, because the site's file leaves it blank.
RAM_EXTRACT_DATE is the build's as-of date. Console-free; returns a
ConsolidateResult.
"""
from __future__ import annotations

import logging
import sys

import arcgis_clean_road_build as acb
import city_codes
import clean_road_columns as crc
import clean_road_layers as crl
import paths

log = logging.getLogger(__name__)

REPORT_NAME = "Clean Road Ramp (ArcGIS)"
SPEC = crc.RAMP
BUILD_VERSION = 1
FILENAME = "clean_ramp_from_layers.xlsx"
OUT_DIR = paths.OUTPUT_ROOT / "arcgis_reports"
OUT_PATH = OUT_DIR / FILENAME
SIDECAR_KEY = "arcgis_report_build"
# The lane's surface (`arcgis_reports`): the built sheet's header, and the
# columns the comparison shows but never counts — the ones the site has no
# source for.
HEADER = [acb.ROUTE_COL] + list(SPEC.header)
CONTEXT_COLUMNS = tuple(SPEC.unsourced)
TEXT_COLUMNS = ("RAM_DESCRIPTION", "RAM_DESIGN_DESC")

PT_LAYER = "SHS Ramp Pt"
RAMP_LAYER = "SHS Ramp"
DIR_LAYER = "Route Direction"
NET_LAYER = "SHS Inv Network Date"
# tag -> (layer, attribute) for the values the ramp's highway gives it
RANGE_LAYERS = {
    "HG": ("SHS Highway Group", "Highway_Group"),
    "CITY": ("City", "City_Code"),
    "POP": ("SHS Population", "Population_Code"),
}
REQUIRED_LAYERS = tuple(sorted(
    {PT_LAYER, RAMP_LAYER, DIR_LAYER, NET_LAYER}
    | {layer for layer, _attr in RANGE_LAYERS.values()}))

_PT_COLS = ["District", "RouteNum", "RouteSuffix", "Alignment", "County",
            "PMPrefix", "PMMeasure", "PMSuffix", "ODMeasure", "Ramp_Name",
            "InventoryItemStartDate", "InventoryItemEndDate", "RouteID",
            "ARMeasure", "LRSFromDate", "LRSToDate"]
_RAMP_COLS = ["Ramp_Name", "Ramp_Description", "Ramp_Design", "Area4_Ind",
              "Ramp_On_Off_Ind", "InventoryItemStartDate",
              "InventoryItemEndDate", "LRSFromDate", "LRSToDate"]
_DIR_COLS = ["ROUTE", "FROM_", "TO_"]

_ON_OFF = {"0": "OFF", "1": "ON", "2": "Z"}

PROVENANCE = {
    "RAM_NETWORK_ID": ("(constant)", "2, as the site writes it"),
    "RAM_PRIMARY_DIRECTION_CODE": (f"{DIR_LAYER}.TO_ / FROM_",
                                   "leading letter: TO on the right alignment, "
                                   "FROM on the left"),
    "RAM_DESIGN_CODE": (f"{RAMP_LAYER}.Ramp_Design", "code"),
    "RAM_DESIGN_DESC": (f"{RAMP_LAYER}.Ramp_Design",
                        "the coded value's domain name, whole"),
    "RAM_BEGIN_OFFSET_AMT": (f"{PT_LAYER}.ODMeasure",
                             "natural decimal (the site leaves it blank)"),
    "RAM_BEGIN_DATE": (f"{PT_LAYER}.InventoryItemStartDate", "MM/DD/YYYY"),
    "RAM_DISTRICT_CODE": (f"{PT_LAYER}.District", "2 digits"),
    "RAM_COUNTY_CODE": (f"{PT_LAYER}.County", "TASAS county code"),
    "RAM_ROUTE_NAME": (f"{PT_LAYER}.RouteNum", "3 digits"),
    "RAM_ROUTE_SUFFIX_CODE": (f"{PT_LAYER}.RouteSuffix", "code ('.' = none)"),
    "RAM_PM_PREFIX_CODE": (f"{PT_LAYER}.PMPrefix", "code ('.' = none)"),
    "RAM_PM_LOC_AMT": (f"{PT_LAYER}.PMMeasure", "natural decimal"),
    "RAM_PM_SUFFIX_CODE": (f"{PT_LAYER}.PMSuffix", "code ('.' = none)"),
    "RAM_SEG_ORDER_ID": (f"{NET_LAYER}.SegOrderId",
                         "span at the ramp's AR measure on its _P route"),
    "RAM_ON_OFF_CODE": (f"{RAMP_LAYER}.Ramp_On_Off_Ind", "0 OFF, 1 ON, 2 Z"),
    "RAM_AREA_4_IND": (f"{RAMP_LAYER}.Area4_Ind", "Yes/No -> Y/N"),
    "RAM_DESCRIPTION": (f"{RAMP_LAYER}.Ramp_Description", "verbatim"),
    "RAM_CITY_CODE": ("City.City_Code",
                      "span at the ramp's AR measure, TASAS code"),
    "RAM_END_DATE": (f"{PT_LAYER}.InventoryItemEndDate", "MM/DD/YYYY"),
    "RAM_POP_GROUP": ("SHS Population.Population_Code",
                      "span at the ramp's AR measure, code"),
    "RAM_HIGHWAY_GROUP": ("SHS Highway Group.Highway_Group",
                          "span at the ramp's AR measure, code"),
    "RAM_EXTRACT_DATE": ("(the build's as-of date)", "MM/DD/YYYY"),
}
assert set(PROVENANCE) | set(SPEC.unsourced) == set(SPEC.header), \
    "every Clean Road Ramp column is either sourced or unsourced"


def _live(r, asof):
    """The site's two filters on one row, as of the build date: the LRS window
    (`LRSFromDate <= D < LRSToDate`) and the inventory window
    (`InventoryItemStartDate <= D` and no end, or an end after D)."""
    if not crl.is_asof(crl.to_serial(r["LRSFromDate"]),
                       crl.to_serial(r["LRSToDate"]), asof):
        return False
    start = crl.to_serial(r["InventoryItemStartDate"])
    end = crl.to_serial(r["InventoryItemEndDate"])
    if start is not None and start > asof:
        return False
    return end is None or end > asof


def _inventory_live(r, asof):
    start = crl.to_serial(r["InventoryItemStartDate"])
    end = crl.to_serial(r["InventoryItemEndDate"])
    return (start is None or start <= asof) and (end is None or end > asof)


def _read_attrs(lib, index, asof, events):
    """{Ramp_Name: SHS Ramp row} — the site's `clr_fetchRampAttrs` pick: the
    first row in layer order whose inventory window holds the as-of date. Its
    query has no calibration filter, so a ramp whose only attribute row sits on
    a retired calibration still has attributes."""
    entry = index.get(RAMP_LAYER) or {}
    out = {}
    for r in crl.stream_layer(lib["present"][RAMP_LAYER], _RAMP_COLS,
                              layer_name=RAMP_LAYER,
                              expected_rows=entry.get("rows")):
        name = acb.s(r["Ramp_Name"])
        if name and name not in out and _inventory_live(r, asof):
            out[name] = r
    events.on_log(f"  {RAMP_LAYER}: {len(out):,} as-of ramps")
    return out


def _read_directions(lib, index):
    """{route number: (FROM_, TO_)} from the Route Direction table."""
    entry = index.get(DIR_LAYER) or {}
    out = {}
    for r in crl.stream_layer(lib["present"][DIR_LAYER], _DIR_COLS,
                              layer_name=DIR_LAYER,
                              expected_rows=entry.get("rows")):
        n = acb.integer(r["ROUTE"])
        if n and n not in out:
            out[n] = (acb.s(r["FROM_"]), acb.s(r["TO_"]))
    return out


def _direction(directions, route_num, alignment):
    frm, to = directions.get(acb.integer(route_num), ("", ""))
    d = frm if crl.norm_alignment(alignment) == "L" else to
    return d[:1].upper()


def _range_value(overlay, route_id, ar):
    """The site's `queryRangeLayer` pick for one point: a secondary-alignment
    ramp ('…_S') reads its own route first, then the primary one."""
    if ar is None or not route_id:
        return ""
    if route_id.endswith("_S"):
        hit = overlay.value_at(route_id, ar)
        if hit:
            return hit
        route_id = route_id[:-2] + "_P"
    return overlay.value_at(route_id, ar)


def _design_name(v):
    """The design code's domain name ('D- Diamond Type Ramp'), or "" for a code
    outside the domain (the layer holds a bare 'V' / 'P' on 25 ramps) — the
    site looks the name up in the domain and finds none."""
    text = acb.s(v)
    return text if acb.code(text) != text else ""


def _project(pt, attrs, overlays, net, directions, extract):
    route = acb.route_token(pt["RouteNum"], pt["RouteSuffix"])
    county = crl.norm_county(pt["County"])
    prefix = acb.pfx(pt["PMPrefix"])
    pm = crl.pm_units(pt["PMMeasure"])
    a = attrs or {}
    rid = acb.s(pt["RouteID"])
    try:
        ar = float(pt["ARMeasure"])
    except (TypeError, ValueError):  # silent-ok: a point with no AR measure has nothing to look up, and every lookup renders that blank, as the site does
        ar = None

    def at(layer_tag):
        return _range_value(overlays[layer_tag], rid, ar)

    design = acb.code(a.get("Ramp_Design")).upper() if a else ""
    v = {
        "RAM_NETWORK_ID": "2",
        "RAM_PRIMARY_DIRECTION_CODE": _direction(directions, pt["RouteNum"],
                                                 pt["Alignment"]),
        "RAM_DESIGN_CODE": design,
        "RAM_DESIGN_DESC": _design_name(a.get("Ramp_Design")) if a else "",
        "RAM_BEGIN_OFFSET_AMT": acb.num(pt["ODMeasure"]) if acb.s(pt["ODMeasure"]) else "",
        "RAM_BEGIN_DATE": acb.mdy(pt["InventoryItemStartDate"]),
        "RAM_DISTRICT_CODE": crl.norm_district(pt["District"]),
        "RAM_COUNTY_CODE": county,
        "RAM_ROUTE_NAME": crl.norm_route(pt["RouteNum"]),
        "RAM_ROUTE_SUFFIX_CODE": acb.pfx(pt["RouteSuffix"]),
        "RAM_PM_PREFIX_CODE": prefix,
        "RAM_PM_LOC_AMT": acb.num(pt["PMMeasure"]) if pm is not None else "",
        "RAM_PM_SUFFIX_CODE": acb.pfx(pt["PMSuffix"]),
        "RAM_SEG_ORDER_ID": acb.integer(net.value_at(acb.primary_route_id(rid), ar))
                            if ar is not None else "",
        "RAM_ON_OFF_CODE": _ON_OFF.get(acb.code(a.get("Ramp_On_Off_Ind")), "")
                           if a else "",
        "RAM_AREA_4_IND": acb.yn(a.get("Area4_Ind")) if a else "",
        "RAM_DESCRIPTION": acb.s(a.get("Ramp_Description")) if a else "",
        "RAM_CITY_CODE": city_codes.norm_city(at("CITY")),
        "RAM_END_DATE": acb.mdy(pt["InventoryItemEndDate"]),
        "RAM_POP_GROUP": crl.code_of(at("POP")),
        "RAM_HIGHWAY_GROUP": crl.code_of(at("HG")),
        "RAM_EXTRACT_DATE": extract,
    }
    return [route] + [v.get(c, "") for c in SPEC.header]


def build_rows(lib, index, *, asof, asof_day, events):
    """(rows, stats): one row per as-of ramp point, ordered by odometer within
    each route as the site orders its file."""
    events.on_log("  reading the span layers the file looks ramps up in...")

    def read(layer, attr, **dates):
        entry = index.get(layer) or {}
        return acb.MeasureOverlay.read(lib["present"][layer], layer, attr,
                                       asof=asof, expected_rows=entry.get("rows"),
                                       **dates)

    overlays = {tag: read(layer, attr) for tag, (layer, attr) in RANGE_LAYERS.items()}
    net = read(NET_LAYER, "SegOrderId", start="Network_Start_Date",
               end="Network_End_Date")
    attrs = _read_attrs(lib, index, asof, events)
    directions = _read_directions(lib, index)
    extract = f"{asof_day.month:02d}/{asof_day.day:02d}/{asof_day.year:04d}"

    entry = index.get(PT_LAYER) or {}
    seen, rows, no_attrs, history = set(), [], 0, 0
    odo = {}
    for pt in crl.stream_layer(lib["present"][PT_LAYER], _PT_COLS,
                               layer_name=PT_LAYER,
                               expected_rows=entry.get("rows")):
        if events.is_cancelled():
            return [], {}
        name = acb.s(pt["Ramp_Name"])
        if not name or not _live(pt, asof):
            history += 1
            continue
        if name in seen:
            continue
        seen.add(name)
        a = attrs.get(name)
        if a is None:
            no_attrs += 1
        row = _project(pt, a, overlays, net, directions, extract)
        rows.append(row)
        try:
            odo[id(row)] = float(pt["ODMeasure"])
        except (TypeError, ValueError):  # silent-ok: an unreadable odometer only places the ramp after its route's measured ones, as the site's own sort does
            odo[id(row)] = float("inf")

    rows.sort(key=lambda row: (row[0], odo[id(row)]))
    stats = {"rows": len(rows), "routes": len({row[0] for row in rows}),
             "not_current": history, "no_attributes": no_attrs}
    events.on_log(f"  built {len(rows):,} ramp records ({stats['routes']} "
                  f"routes); {no_attrs:,} ramp points have no as-of SHS Ramp "
                  "attribute row")
    return rows, stats


def marker_rows(stats):
    return [
        ("Point rows not live at the as-of date", stats.get("not_current", 0)),
        ("Ramp points with no SHS Ramp attributes", stats.get("no_attributes", 0)),
        ("Row universe", "SHS Ramp Pt rows live at the as-of date (LRS window and "
                         "inventory window), one per Ramp_Name"),
        ("Columns with no source on the site",
         ", ".join(SPEC.unsourced) + " — all blank here too except "
         "RAM_BEGIN_OFFSET_AMT, written from SHS Ramp Pt.ODMeasure"),
    ]


def is_arcgis_report(path):
    return acb.is_build(path, SPEC)


def report_facts(path):
    return acb.report_facts(path)


def consolidate(events=None, confirm_overwrite=None, day=None, *,
                out_path=None, asof=None, lib_root=None):
    """Build the Clean Road Ramp file FROM THE ARCGIS LAYERS as of `asof` (set
    it to the export day being compared)."""
    del day                       # ConsolidateWorker passes it; no run days here
    return acb.consolidate(sys.modules[__name__], events, confirm_overwrite,
                           out_path=out_path, asof=asof, lib_root=lib_root)
