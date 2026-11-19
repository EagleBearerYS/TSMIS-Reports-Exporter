"""Clean Road files — the vendor's source list, held against the app's builds.

On 2026-11-16 the TSMIS vendor documented where each Clean Road attribute
comes from: the feature-service layer (by name and number) for 34 Highway and
12 Ramp attributes, a source for 32 Intersection attributes, and the service's
whole layer list. This module is that list as data, so three things can be
checked instead of asserted:

  * every attribute column a build writes comes from the layer AND field the
    list names for it (`mismatches`, run by `check_clean_road_sources`);
  * every layer the library refreshes carries the number the list gives it
    (`SERVICE_LAYERS` against `arcgis_refresh.LAYER_IDS`);
  * the comparison workbooks say so, column by column, and say which service
    the layers were pulled from next to which source the export came from
    (`notes_lines`, `source_note`).

What the list does NOT settle is recorded here too, because a reader of a
comparison will ask:

  * it names attributes the site's file has no column for (`unplaced`);
  * it names SHS Tolls for a toll type the site's code does not read, and SHS
    Ramp's inventory start date for a create date the site leaves blank;
  * for Intersection it names one flat "Master Intersection" layer on another
    service. The site's own code (the 9.1 and the 11.17 dev captures alike)
    reads the IM tables of the main service instead, and so do the builds —
    the library holds no Master Intersection layer, so that source is neither
    read nor checked here (`INTERSECTION_SOURCE`).

The list numbers its layers on a development copy of the service; the numbers
are the same on the production service the library is refreshed from (the
library's own index records each layer's number, and the check holds the two
together). Import-light: no openpyxl, no build modules at import time.
"""
import re
from collections import namedtuple

LIST_DATE = "2026-11-16"
LIST_NAME = f"the TSMIS vendor's source list of {LIST_DATE}"

# --------------------------------------------------------------------------- #
# the service's layers and tables, by number (the list's "Featureservice" tab)
# --------------------------------------------------------------------------- #
SERVICE_LAYERS = {
    0: "IM Intersection Point", 1: "Calibration Point", 2: "Centerline",
    3: "Postmile", 4: "AllRoads", 5: "Odometer", 11: "Terrain Type",
    41: "HPMS Urban Code", 42: "HPMS Turn Lanes R", 43: "HPMS Turn Lanes L",
    44: "HPMS Through Lanes", 46: "Redline", 48: "HPMS Shldr Width L",
    50: "HPMS Sample Panel", 51: "HPMS Route Signing", 74: "City",
    85: "County Code", 86: "HPMS AADT", 87: "HPMS AADT Combination",
    88: "HPMS AADT Single Unit", 89: "HPMS Access Control",
    90: "HPMS Dir Factor", 91: "HPMS F System", 92: "HPMS Facility Type",
    93: "HPMS K Factor", 94: "HPMS Lane Width", 95: "HPMS Median Type",
    96: "HPMS Median Width", 97: "HPMS National Network", 98: "HPMS NHFN",
    99: "HPMS NHS", 100: "HPMS Ownership", 101: "HPMS Pct Peak Combo",
    102: "HPMS Pct Peak Single", 103: "HPMS Shoulder Type", 104: "HPMS Tolls",
    107: "Posted Speed Limit", 108: "Private Roads", 109: "SHS Access Control",
    110: "SHS Barrier", 111: "SHS Centerline RumblStr",
    112: "SHS Curb Landscape", 113: "SHS Design Speed", 114: "SHS District",
    115: "SHS Forest HWY", 116: "SHS Highway Group",
    117: "SHS Horizontal Curve", 118: "SHS I Shld RumblStr L",
    119: "SHS I Shld RumblStr R", 120: "SHS I Shld Width L",
    121: "SHS I Shld Width R", 122: "SHS Inv Network Date", 123: "SHS Landmark",
    124: "SHS Median", 125: "SHS Non Add Mileage", 126: "SHS O Shld RumblStr L",
    127: "SHS O Shld RumblStr R", 128: "SHS O Shld Width L",
    129: "SHS O Shld Width R", 130: "SHS Population", 131: "SHS Ramp",
    132: "SHS Ramp Pt", 133: "SHS Route Break", 134: "SHS Special Feature L",
    135: "SHS Special Feature R", 136: "SHS Surface Type L",
    137: "SHS Surface Type R", 138: "SHS Tolls", 139: "SHS Travel Way L",
    140: "SHS Travel Way R", 141: "SHS Vertical Grade",
    143: "HPMS Structure Type", 144: "HPMS Shldr Width R",
    148: "IM Roadway Segment", 149: "Master Intersection AOI",
    153: "Traffic Volume Segments", 156: "MS2 Traffic Station",
    157: "Traffic Volume Ramps", 200: "SHS Bikeway Inventory",
    203: "IM Filter Points", 213: "HSL Equation Points", 214: "HSL Supporting",
    215: "Highway Sequence Listing", 217: "HighwayLog Left Postmile AllRoads",
    218: "Translate Postmile to AllRoads", 303: "SHS Ped Inventory",
    305: "Equation Points",
    # tables
    47: "Centerline Sequence", 146: "IM Intersection Approach Segments",
    147: "IM Intersection Route Table", 150: "IM Intersection Approach Detail",
    151: "IM Intersection Detail",
    211: "IM Complex Intersection Influence Segments",
    212: "IM Complex Intersection Cross Reference", 216: "Highway Log",
    302: "IM Intersection Manager Metadata", 304: "Route Direction",
}
LAYER_NUMBERS = {name: number for number, name in SERVICE_LAYERS.items()}
assert len(LAYER_NUMBERS) == len(SERVICE_LAYERS) == 97, "a layer name repeats"


def layer_tag(layer):
    """'SHS Median #124' — a layer with the number the list gives it."""
    number = LAYER_NUMBERS.get(layer)
    return f"{layer} #{number}" if number is not None else layer


# --------------------------------------------------------------------------- #
# the attribute lists
# --------------------------------------------------------------------------- #
# field: the layer's own field name (checked against the library's layer
# files 2026-11-17 — the list writes two with a space, 'Barrier Type' and
# 'Toll Facility_Name'); columns: the file column(s) it fills, () when the
# site's file has no column for it; note: what a reader needs to know.
Source = namedtuple("Source", "field layer columns note")

HIGHWAY = (
    Source("City_Code", "City", ("THY_CITY_CODE",), ""),
    Source("Posted_Speed_Limit", "Posted Speed Limit", (),
           "the file has no posted-speed column"),
    Source("SHS_Access_Control", "SHS Access Control",
           ("THY_HIGHWAY_ACCESS_CODE",), ""),
    Source("Barrier_Type", "SHS Barrier", ("THY_MEDIAN_BARRIER_CODE",), ""),
    Source("Curb_Landscape", "SHS Curb Landscape",
           ("THY_CURB_LANDSCAPE_CODE",), ""),
    Source("Design_Speed", "SHS Design Speed", ("THY_DESIGN_SPEED_AMT",), ""),
    Source("Forest_Hwy", "SHS Forest HWY", ("THY_TOLL_FOREST_CODE",), ""),
    Source("Highway_Group", "SHS Highway Group",
           ("THY_HIGHWAY_GROUP_CODE",), ""),
    Source("Median_Type", "SHS Median", ("THY_MEDIAN_TYPE_CODE",), ""),
    Source("Median_Width", "SHS Median", ("THY_MEDIAN_WIDTH_AMT",), ""),
    Source("Median_Variance", "SHS Median",
           ("THY_MEDIAN_WIDTH_VAR_CODE",), ""),
    Source("Non_Add_Mileage", "SHS Non Add Mileage", ("THY_NON_ADD_CODE",), ""),
    Source("Population_Code", "SHS Population", ("THY_POPULATION_CODE",), ""),
    Source("Terrain_Type", "Terrain Type", ("THY_TERRAIN_CODE",), ""),
    Source("Toll_Type", "SHS Tolls", ("THY_TOLL_FOREST_CODE",),
           "the list says the site's code does not read it; the build writes "
           "a toll span as TASAS code 1"),
    Source("Facility_Name", "SHS Tolls", (),
           "the file has no toll-facility column"),
    Source("Travel_Way_Width_L", "SHS Travel Way L",
           ("THY_LT_TRAV_WAY_WIDTH_AMT",), ""),
    Source("Total_Num_Lanes_L", "SHS Travel Way L", ("THY_LT_LANES_AMT",), ""),
    Source("Thru_Num_Lanes_L", "SHS Travel Way L", (),
           "the file's one lanes column is the total"),
    Source("Travel_Way_Width_R", "SHS Travel Way R",
           ("THY_RT_TRAV_WAY_WIDTH_AMT",), ""),
    Source("Total_Num_Lanes_R", "SHS Travel Way R", ("THY_RT_LANES_AMT",), ""),
    Source("Thru_Num_Lanes_R", "SHS Travel Way R", (),
           "the file's one lanes column is the total"),
    Source("Special_Feature_Type_L", "SHS Special Feature L",
           ("THY_LT_SPEC_FEATURES_CODE",), ""),
    Source("Special_Feature_Type_R", "SHS Special Feature R",
           ("THY_RT_SPEC_FEATURES_CODE",), ""),
    Source("Surface_Type_L", "SHS Surface Type L",
           ("THY_LT_SURF_TYPE_CODE",), ""),
    Source("Surface_Type_R", "SHS Surface Type R",
           ("THY_RT_SURF_TYPE_CODE",), ""),
    Source("Shld_Width_Treated_In_L", "SHS I Shld Width L",
           ("THY_LT_I_SHD_TRT_WIDTH_AMT",), ""),
    Source("Shld_Width_Total_In_L", "SHS I Shld Width L",
           ("THY_LT_I_SHD_TOT_WIDTH_AMT",), ""),
    Source("Shld_Width_Treated_In_R", "SHS I Shld Width R",
           ("THY_RT_I_SHD_TRT_WIDTH_AMT",), ""),
    Source("Shld_Width_Total_In_R", "SHS I Shld Width R",
           ("THY_RT_I_SHD_TOT_WIDTH_AMT",), ""),
    Source("Shld_Width_Treated_Out_L", "SHS O Shld Width L",
           ("THY_LT_O_SHD_TRT_WIDTH_AMT",), ""),
    Source("Shld_Width_Total_Out_L", "SHS O Shld Width L",
           ("THY_LT_O_SHD_TOT_WIDTH_AMT",), ""),
    Source("Shld_Width_Treated_Out_R", "SHS O Shld Width R",
           ("THY_RT_O_SHD_TRT_WIDTH_AMT",), ""),
    Source("Shld_Width_Total_Out_R", "SHS O Shld Width R",
           ("THY_RT_O_SHD_TOT_WIDTH_AMT",), ""),
)

RAMP = (
    Source("Ramp_Description", "SHS Ramp", ("RAM_DESCRIPTION",), ""),
    Source("Ramp_Name", "SHS Ramp", (),
           "the key the ramp point and its attributes join on; the file "
           "prints no ramp name"),
    Source("Ramp_Design", "SHS Ramp", ("RAM_DESIGN_CODE", "RAM_DESIGN_DESC"), ""),
    Source("Area4_Ind", "SHS Ramp", ("RAM_AREA_4_IND",), ""),
    Source("Ramp_On_Off_Ind", "SHS Ramp", ("RAM_ON_OFF_CODE",), ""),
    Source("Side_of_Hwy", "SHS Ramp", (), "the file has no such column"),
    Source("Ramp_Length", "SHS Ramp", (), "the file has no such column"),
    Source("Ramp_Num_Lanes", "SHS Ramp", (), "the file has no such column"),
    Source("Highway_Group", "SHS Highway Group", ("RAM_HIGHWAY_GROUP",), ""),
    Source("City_Code", "City", ("RAM_CITY_CODE",), ""),
    Source("Population_Code", "SHS Population", ("RAM_POP_GROUP",), ""),
    Source("InventoryItemStartDate", "SHS Ramp", ("RAM_CREATE_DATE",),
           "the site's file leaves the column blank; the build writes it, "
           "shown and never counted"),
)

# The Intersection tab names ONE source for all 32 attributes: a flat layer on
# another service, with a Major_* / Minor_* field per leg. Field names as the
# list gives them; the last three are the ones it could not find.
INTERSECTION_SOURCE = "Master Intersection"
INTERSECTION_FIELDS = (
    "Description", "Intersection_Geometry", "Intersection_Geometry_Eff_Date",
    "Intersection_Control", "Intersection_Control_Eff_Date",
    "Intersection_Lighted_Ind", "Intersection_Lighted_Eff_Date",
    "Major_Population_Code", "Major_Highway_Group", "Major_Pri_Num_Lanes",
    "Major_Sec_Num_Lanes", "Major_Total_Num_Lanes", "Major_Signal_Arm_Ind",
    "Major_Left_Channel", "Major_Right_Channel_Ind", "Major_N_Distance",
    "Major_FlowCode", "Major_Effective_Date", "Minor_Population_Code",
    "Minor_Highway_Group", "Minor_Pri_Num_Lanes", "Minor_Sec_Num_Lanes",
    "Minor_Total_Num_Lanes", "Minor_Signal_Arm_Ind", "Minor_Left_Channel",
    "Minor_Right_Channel_Ind", "Minor_N_Distance", "Minor_FlowCode",
    "Minor_Effective_Date",
)
INTERSECTION_NOT_FOUND = ("Major_Left_Channel_Ind", "Minor_Left_Channel_Ind",
                          "Record Create Date")
assert len(INTERSECTION_FIELDS) + len(INTERSECTION_NOT_FOUND) == 32

_LISTS = {"clean_highway": HIGHWAY, "clean_ramp": RAMP}
assert len(HIGHWAY) == 34 and len(RAMP) == 12


def sources(report_key):
    """The list's attribute rows for a report ('clean_highway' / 'clean_ramp');
    () for Intersection, whose one source the library does not hold."""
    return _LISTS.get(report_key, ())


def by_column(report_key):
    """{file column: [Source, ...]} — the list rows that fill each column."""
    out = {}
    for src in sources(report_key):
        for col in src.columns:
            out.setdefault(col, []).append(src)
    return out


def unplaced(report_key):
    """The attributes the list names that the site's file has no column for."""
    return tuple(src for src in sources(report_key) if not src.columns)


# --------------------------------------------------------------------------- #
# the list against a build
# --------------------------------------------------------------------------- #
def _highway_claim(column):
    """(layer text, field text) the Highway build declares for a column."""
    import clean_highway_columns as chc
    _tier, layer, field, _note = chc.PROVENANCE[column]
    return layer, field


def _point_claim(provenance, column):
    """The 'Layer.Field' text an Intersection / Ramp build declares."""
    return provenance.get(column, ("", ""))[0]


def mismatches(report_key, provenance=None):
    """Every place a build's declared source differs from the list — [] when
    each column the list covers is built from the layer and field it names.
    `provenance` is the Ramp build's PROVENANCE (Highway's lives in
    `clean_highway_columns`)."""
    out = []
    for column, rows in by_column(report_key).items():
        for src in rows:
            if report_key == "clean_highway":
                layer, field = _highway_claim(column)
                # THY_TOLL_FOREST_CODE is one column fed by two layers; its
                # declared field is the coverage rule, not one field name.
                ok = src.layer in layer and (len(rows) > 1 or src.field == field)
                claim = f"{layer} ({field})"
            else:
                claim = _point_claim(provenance or {}, column)
                ok = claim.startswith(f"{src.layer}.{src.field}")
            if not ok:
                out.append(f"{column}: the list says {src.layer}.{src.field}, "
                           f"the build declares {claim or '(nothing)'}")
    return out


# --------------------------------------------------------------------------- #
# which service
# --------------------------------------------------------------------------- #
_SERVICE_RE = re.compile(r"^(.*?/FeatureServer)", re.IGNORECASE)
_RUN_TAG_RE = re.compile(r"(\d{4}-\d{2}-\d{2}(?: \w+-\w+(?: \w+-site)?)?)")


def service_of(source):
    """The feature service of one INDEX `Data Source` cell (everything up to
    'FeatureServer'), or "" when the cell names none."""
    m = _SERVICE_RE.match(str(source or "").strip())
    return m.group(1) if m else ""


def library_service(index):
    """The one service a library's layers were pulled from, "" for an empty or
    unrecorded index, and 'several services' when its layers disagree."""
    found = {service_of(row.get("source")) for row in (index or {}).values()}
    found.discard("")
    if not found:
        return ""
    return found.pop() if len(found) == 1 else "several services"


def export_source(path):
    """(day, src, env) the export's file name records — the run-folder tag the
    app stamps into every consolidated workbook — or None."""
    import paths
    from pathlib import Path

    for tag in reversed(_RUN_TAG_RE.findall(Path(str(path)).stem)):
        parsed = paths.parse_run_folder(tag)
        if parsed and " " in tag:
            return parsed
    return None


def source_note(service, export_path):
    """One Notes line: the service the layers came from beside the source the
    export was taken from, and whether the two are the same data source. The
    site reads one service per data source (SSOR / ARS) and environment; the
    library is refreshed from the production SSOR service."""
    import arcgis_refresh

    head = (f"DATA SOURCE: the layers were pulled from {service}."
            if service and service != "several services" else
            "DATA SOURCE: this build does not record which feature service its "
            "layers were pulled from — rebuild it with this version to have it "
            "stated here." if not service else
            "DATA SOURCE: the layer library mixes layers from several services.")
    tagged = export_source(export_path)
    if tagged is None:
        return head + (" The export's file name does not say which data source "
                       "it was taken from.")
    _day, src, env = tagged
    label = f"{src.upper()} {env}"
    if not service or service == "several services":
        return head + f" The export was taken from the site's {label} source."
    is_default = service.rstrip("/").lower() == arcgis_refresh.SERVICE_DEFAULT.lower()
    if src == "ssor" and env == "prod" and is_default:
        return head + (" The export was taken from the site's SSOR prod "
                       "source, which reads this same service: both sides "
                       "come from one data source.")
    return head + (f" ⚠ The export was taken from the site's {label} source, "
                   "which reads a DIFFERENT service from the one these layers "
                   "came from. The two are meant to carry the same data, but "
                   "this comparison cannot show that they do — compare an "
                   "export taken from the source the layers were pulled from.")


# --------------------------------------------------------------------------- #
# the Notes lines
# --------------------------------------------------------------------------- #
def _unplaced_text(report_key):
    rows = unplaced(report_key)
    return "; ".join(f"{src.field} ({layer_tag(src.layer)}: {src.note})"
                     for src in rows)


def notes_lines(report_key, provenance=None):
    """The Notes paragraphs that hold a build against the list."""
    if report_key == "clean_intersection":
        return (
            f"SOURCES: {LIST_NAME} names one source for the Intersection file, "
            f"a flat '{INTERSECTION_SOURCE}' layer on another feature service. "
            "The site's own code does not read it: the 9.1 and the 11.17 dev "
            "captures both read the IM tables of the main service "
            f"({layer_tag('IM Intersection Detail')}, "
            f"{layer_tag('IM Intersection Approach Segments')}, "
            f"{layer_tag('IM Intersection Approach Detail')}) and the three "
            "span layers, and so does this build. The layer library holds no "
            f"{INTERSECTION_SOURCE} layer, so this comparison says nothing "
            "about it.",
        )
    covered = by_column(report_key)
    bad = mismatches(report_key, provenance)
    verdict = (f"all {len(covered)} agree" if not bad else
               f"{len(bad)} DISAGREE — " + "; ".join(bad))
    lines = [
        f"SOURCES: {LIST_NAME} gives the layer and field of "
        f"{len(sources(report_key))} attributes of this file. "
        f"{len(covered)} file columns are built from them here, each from the "
        f"layer and field the list names: {verdict}. The table below shows each "
        "layer with the number the list gives it (#).",
    ]
    extra = _unplaced_text(report_key)
    if extra:
        lines.append("The list also names attributes the site's file has no "
                     f"column for, so nothing here compares them: {extra}.")
    notes = [f"{' / '.join(src.columns)} — {src.field}: {src.note}"
             for src in sources(report_key) if src.columns and src.note]
    if notes:
        lines.append("Where the list and the site's code part: " + "; ".join(notes) + ".")
    return tuple(lines)


def highway_table_lines(counted):
    """The Highway column table for the vs-TSMIS Notes: each column · the layer
    (with its number) and field the build paints it from · whether the list
    covers it · counted or context."""
    import clean_highway_columns as chc

    listed = by_column("clean_highway")
    out = []
    for name in chc.ARC_HEADER:
        _tier, layer, field, _note = chc.PROVENANCE[name]
        if layer:
            shown = " + ".join(layer_tag(part.strip()) for part in layer.split("+"))
            src = f"{shown} ({field})" if field else shown
        else:
            src = "no ArcGIS source"
        on_list = "on the list" if name in listed else "not an attribute on the list"
        role = "counted" if name in counted else "context"
        out.append(f"    {name}  ·  built from {src}  ·  {on_list}  ·  {role}")
    return tuple(out)


def point_table_lines(spec, provenance, report_key, context=None):
    """The Intersection / Ramp column table: each column the build fills · its
    layer field(s) · whether the list covers it · counted or context
    (`context` = the comparison's own context columns)."""
    listed = by_column(report_key)
    names = sorted(LAYER_NUMBERS, key=len, reverse=True)
    shown_only = set(spec.unsourced if context is None else context)
    out = []
    for col in spec.header:
        if col not in provenance:
            continue
        src, how = provenance[col]
        layer = next((n for n in names if src.startswith(n)), "")
        tagged = layer_tag(layer) + src[len(layer):] if layer else src
        role = "context" if col in shown_only else "counted"
        tail = "  ·  on the list" if col in listed else ""
        out.append(f"    {col}  ·  {tagged} ({how}){tail}  ·  {role}")
    return tuple(out)
