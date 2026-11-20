"""The SITE's Clean Road exports vs the TSN clean-road extracts (v0.48.0).

Nine "files"-kind comparisons over the regression-locked compare_core engine —
three per report (Highway / Intersection / Ramp):

  * <KIND>_VS_TSN        — the consolidated Excel export vs the TSN extract;
  * <KIND>_PDF_VS_TSN    — the consolidated PRINT vs the TSN extract;
  * <KIND>_PDF_VS_EXCEL  — the print vs the Excel export of the same run (the
                           site checked against itself).

Side A is always a workbook from `clean_road_consolidate` (sheet = the export's
own sheet, "Route" + the export's exact header). The TSN side is the raw
`CA HIGHWAYS` / `CA INTERSECTIONS` / `CA RAMPS` extract (`Sheet 1`) or the TSN
library's normalized copy (marker-gated). Role gates keep the sides honest: the
print flavors REQUIRE the PDF-conversion marker on side A, the Excel flavors
and the self-check's Excel side REFUSE it, and no side may be the app's own
ArcGIS build (`ArcGIS Build` marker) — that is a different comparison.

Row identity — the physical location both systems carry natively:
    route (name + suffix) · county · PM prefix · postmile (decimal-canonical)
    · roadbed (the R/L/X postmile suffix)
For Highway that is the segment's BEGIN postmile (the END is a compared field,
so a stretch the systems cut differently pairs instead of splitting); for
Intersection the intersection's postmile; for Ramp the ramp's.

Format-only reconciliations, all measured on the 2026-10-02 statewide pull
against the 2025-09 extracts and named on the Notes sheet:
  * dates compare as dates (the site prints MM/DD/YYYY, Intersection dates
    pass the layer's own YY-MM-DD through; TSN stores dates);
  * amounts compare as numbers ('02' = '2', '9.60' = '9.6');
  * county codes drop the TASAS trailing period the site writes for the
    two-letter counties ('SD.' = 'SD');
  * Highway: the site writes the equation marker 'E' into the postmile-suffix
    column on equation rows (THY_EQUATE_CODE = E on every one), so it is not a
    roadbed for the KEY — the column itself is still compared; a blank
    THY_TOLL_FOREST_CODE is TSN's code 0 (neither toll nor forest);
  * Intersection: TSN prefixes its design codes with the domain letter
    ('IT' = the site's 'T'); the control type applies the TSNR/MIRE crosswalk
    the Intersection Detail comparison already uses (TSN's signal sub-types
    J–P fold to the Signalized code S on both sides);
  * Ramp: TSN prefixes its design codes ('RD' = 'D'); the site's On/Off 'Z'
    is TSN's 'OTH'.
Runs of spaces compare as one space in every flavor (the engine's Excel-TRIM
rule) — which is what lets the print, whose HTML collapses them, compare
against the Excel file and the TSN extract, which keep them; the print flavors
also fold tabs and line breaks the same way.

Then every value of both sides is written back the way the site's Excel file
writes it (`rows_as_written`, v0.54.1): MM/DD/YYYY dates (YY-MM-DD for the
Intersection layer's), Highway's three-decimal postmiles and two-digit lanes
and widths, 'SD.', a blank toll / forest code, 'Z'. It is one to one with the
compared form, so it moves no count; it makes the sheets read like the file.

CONTEXT columns (the TSMIS file's own value shown, never counted): the columns
the site has no source for (blank in its file by design), TSN's bookkeeping
(ids, lifecycle, create/update), the TASAS change-tracking columns, the
extract date, Highway's two offset columns (each system's own cumulative), the
Intersection connection id (a different id space) and the Ramp design
description (each system's own wording of the compared design code).

TSN is FROZEN at the September-2025 cutover (roadmap D5): every difference is
TSMIS-vs-the-2025-snapshot drift, never by itself a defect. Console-free.
"""
import re
from dataclasses import dataclass, replace
from datetime import date, datetime
from functools import partial
from pathlib import Path

try:
    from openpyxl import load_workbook
    _DEPS_OK = True
except ImportError:
    _DEPS_OK = False

import arcgis_clean_road_build as acb
import clean_highway_columns as chc
import clean_road_columns as crc
import compare_clean_highway_tsn as cht
import compare_tsn_common as ctc
import comparison_contract as cc
from compare_core import CompareSchema
from compare_intersection_detail_tsn import _norm_control_type
from paths import today_str
from pdf_table_lib import pdf_source_marker_state

# CMP-AUD-037: the Intersection / Ramp library slots are stamped with this
# version (tsn_load_clean_road mirrors it; the catalog's normalization_version
# mirrors both). Highway shares the ArcGIS comparison's slot and marker.
NORMALIZATION_VERSION = 1
_HIGHWAY_TSN_MARKER = cht.NORMALIZATION_VERSION

_WS = re.compile(r"\s+")

# The marker sheets the app's own ArcGIS builds carry: the CA HIGHWAYS build's
# and the Intersection / Ramp builds'. Either one means the workbook is OUR
# reconstruction from the layers — never the site's export, never TSN.
_BUILD_MARKERS = (chc.ARC_MARKER_SHEET, acb.MARKER_SHEET)


def _is_arcgis_build(wb):
    return any(m in wb.sheetnames for m in _BUILD_MARKERS)


def _s(v):
    return "" if v is None else str(v).strip()


def _date(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return ctc.iso_date(v)


def _county(v):
    return _s(v).upper().rstrip(".")


@dataclass(frozen=True)
class _Profile:
    """One report's comparison rules (see the module docstring)."""
    spec: object
    context: tuple
    dates: frozenset
    amounts: frozenset
    text: tuple                 # free-text columns (representation disclosure)
    tsn_design_prefix: str      # TSN's domain letter on the design code, or ""
    design_col: str
    notes: tuple
    # How the site's file WRITES these columns (`as_written`): amounts with
    # three decimals ('0.780'), whole numbers with at least two digits ('03'),
    # dates as YY-MM-DD ('73-10-19'; every other date is MM/DD/YYYY).
    three_places: frozenset = frozenset()
    two_digits: frozenset = frozenset()
    short_dates: frozenset = frozenset()


def _is_numeric_col(name):
    return name.endswith(("_AMT", "_ID", "_ADT")) or name in (
        "THY_TOLL_FOREST_CODE", "THY_CURB_LANDSCAPE_CODE", "THY_MAINT_SVC_LVL_CODE",
        "THY_NATIONAL_LANDS_CODE", "THY_SCENIC_FREEWAY_CODE", "RAM_ADT",
        "INX_MAINLINE_ADT", "INX_XSTREET_ADT")


def _profile(spec, *, extra_context, design_col, tsn_design_prefix, text, notes,
             three_places=(), two_digits=(), short_dates=()):
    header = spec.header
    dates = frozenset(c for c in header if c.endswith("_DATE"))
    amounts = frozenset(c for c in header if _is_numeric_col(c) and c not in dates)
    context = tuple(c for c in header
                    if c in set(spec.unsourced) | set(extra_context)
                    or c in spec.site_only or c.endswith("_EXTRACT_DATE"))
    for name in (*three_places, *two_digits, *short_dates):
        if name not in header:
            raise ValueError(f"{spec.label} has no column {name!r}")
    return _Profile(spec=spec, context=context, dates=dates, amounts=amounts,
                    text=text, tsn_design_prefix=tsn_design_prefix,
                    design_col=design_col, notes=notes,
                    three_places=frozenset(three_places),
                    two_digits=frozenset(two_digits),
                    short_dates=frozenset(short_dates))


_COMMON_NOTES = (
    "The TSN side is FROZEN at the September-2025 cutover (TSMIS replaced "
    "TSN), so every difference here is TSMIS-vs-that-snapshot drift — read the "
    "counts as change since then, never by themselves as defects.",
    "Every value is shown the way the TSMIS file writes it, and TSN's value "
    "is written the same way (its dates, numbers and county codes), so a row "
    "reads like the file. Format never counts: a date is the same date however "
    "it is written, an amount the same number ('02' = '2', '9.60' = '9.6'), "
    "and a two-letter county the same with or without its period ('SD.' = "
    "'SD').",
)

_HIGHWAY = _profile(
    crc.HIGHWAY,
    extra_context=tuple(c for c in chc.CONTEXT_COLUMNS if c in chc.HEADER),
    design_col="", tsn_design_prefix="",
    text=("THY_LANDMARK_SHORT_DESC",),
    notes=_COMMON_NOTES + (
        "Rows pair on route + county + PM prefix + BEGIN postmile + roadbed "
        "(the R/L/X postmile suffix). The END postmile and length are compared "
        "fields, so a stretch the two systems cut differently pairs on its "
        "begin and shows its end as a difference instead of splitting into two "
        "one-sided rows. The site segments the road at the Highway Log's event "
        "points, TSN at its own record breaks, so one-sided rows are mostly "
        "segmentation; TSN also carries the unconstructed routes the site's "
        "file doesn't.",
        "The site writes the equation marker 'E' into THY_PM_SUFFIX_CODE on "
        "equation rows (every one also carries THY_EQUATE_CODE = E); it is not a "
        "roadbed, so the row key treats it as none. The column is still "
        "compared, so the marker shows against TSN's blank.",
        "THY_TOLL_FOREST_CODE: the site leaves the cell blank where no toll or "
        "forest span exists; TSN writes its code 0 for that, which shows as a "
        "blank here.",
        "CONTEXT (the TSMIS file's own value is shown, never counted; TSN's is "
        "on the TSN sheet): the 22 columns the site has no source "
        "for (blank in its file by design — ids, lifecycle dates, the four "
        "change-tracking flags, the ADT profile trio, maintenance service, the "
        "federal-aid trio, national lands, scenic freeway), TSN's bookkeeping, "
        "THY_LAST_SIG_CHG_DATE and THY_FUNCTIONAL_CLASS_CODE (TSN's own "
        "change-tracking date; TSN leaves the class blank), the two offset "
        "columns (each system's own cumulative) and the extract date.",
    ),
    three_places=("THY_BEGIN_OFFSET_AMT", "THY_END_OFFSET_AMT", "THY_BEGIN_PM_AMT",
                  "THY_END_PM_AMT", "THY_LENGTH_MILES_AMT"),
    two_digits=("THY_LT_LANES_AMT", "THY_RT_LANES_AMT",
                "THY_LT_O_SHD_TOT_WIDTH_AMT", "THY_LT_O_SHD_TRT_WIDTH_AMT",
                "THY_LT_TRAV_WAY_WIDTH_AMT", "THY_LT_I_SHD_TOT_WIDTH_AMT",
                "THY_LT_I_SHD_TRT_WIDTH_AMT", "THY_RT_I_SHD_TOT_WIDTH_AMT",
                "THY_RT_I_SHD_TRT_WIDTH_AMT", "THY_RT_TRAV_WAY_WIDTH_AMT",
                "THY_RT_O_SHD_TOT_WIDTH_AMT", "THY_RT_O_SHD_TRT_WIDTH_AMT"))

_INTERSECTION = _profile(
    crc.INTERSECTION,
    extra_context=("INX_CONNECTION_ID",),
    design_col="INX_DESIGN_CODE", tsn_design_prefix="I",
    text=("INX_INTERSECTION_NAME",),
    notes=_COMMON_NOTES + (
        "Rows pair on route + county + PM prefix + the intersection's postmile "
        "+ roadbed (the R/L postmile suffix).",
        "INX_DESIGN_CODE: TSN prefixes its intersection design codes with the "
        "domain letter 'I' ('IT' = the site's 'T'); the letter is dropped from "
        "the TSN value before comparing.",
        "INX_CONTROL_CODE applies the TSNR/MIRE crosswalk the Intersection "
        "Detail comparison uses: TSN's signal sub-types J, K, L, M, N and P and "
        "TSMIS's S all fold to S (Signalized) on both sides.",
        "The site passes the intersection layer's own YY-MM-DD dates through "
        "('73-10-19' is 1973-10-19; a 2-digit year below 30 is 20xx), so they "
        "are shown that way and TSN's dates are written the same way (a date "
        "outside 1930–2029 in full). INX_EXTRACT_DATE is MM/DD/YYYY, as the "
        "file writes it.",
        "CONTEXT (the TSMIS file's own value is shown, never counted; TSN's is "
        "on the TSN sheet): the 10 columns the site has no source "
        "for (placement ids, create/update, end date, both ADTs, the LSC date), "
        "INX_CONNECTION_ID (the site numbers intersections in a different id "
        "space) and the extract date.",
    ),
    short_dates=tuple(c for c in crc.INTERSECTION.header
                      if c.endswith("_DATE") and c != "INX_EXTRACT_DATE"))

_RAMP = _profile(
    crc.RAMP,
    extra_context=("RAM_DESIGN_DESC",),
    design_col="RAM_DESIGN_CODE", tsn_design_prefix="R",
    text=("RAM_DESCRIPTION",),
    notes=_COMMON_NOTES + (
        "Rows pair on route + county + PM prefix + the ramp's postmile + "
        "roadbed (the R/L postmile suffix) — unique on both sides statewide.",
        "RAM_DESIGN_CODE: TSN prefixes its ramp design codes with the domain "
        "letter 'R' ('RD' = the site's 'D'); the letter is dropped from the TSN "
        "value before comparing. RAM_DESIGN_DESC is each system's own wording "
        "of that code (the site's ArcGIS domain vs TSN's code table) and is "
        "shown as context.",
        "RAM_ON_OFF_CODE: the site writes 'Z' for the 'other' class, TSN 'OTH'; "
        "they compare equal, and TSN's 'OTH' shows as 'Z'.",
        "RAM_BEGIN_DATE: the site fills it from the ramp layer's inventory start "
        "date, which is the date TSN keeps in RAM_CHANGE_DATE (TSN's "
        "RAM_BEGIN_DATE is its own record-version date). The column is compared "
        "as named, so this mapping difference shows on most ramps.",
        "CONTEXT (the TSMIS file's own value is shown, never counted; TSN's is "
        "on the TSN sheet): the two site-only columns (RAM_HPMS_ID, "
        "RAM_RAMP_ROUTE_NAME — TSN has neither), the 13 columns the site has no "
        "source for (ids, offset, create/update, ADT, change date, connection "
        "id), the design description and the extract date.",
    ))

_PROFILES = {p.spec.key: p for p in (_HIGHWAY, _INTERSECTION, _RAMP)}


# --------------------------------------------------------------------------- #
# row projection
# --------------------------------------------------------------------------- #
def _value(profile, name, raw, side, collapse):
    """One cell as compared. `side` is "site" or "tsn"; `collapse` folds
    whitespace runs (the print flavors)."""
    if name in profile.dates:
        return _date(raw)
    if name.endswith("_COUNTY_CODE"):
        return _county(raw)
    if name == "THY_TOLL_FOREST_CODE":
        return cht._norm_amount(raw) or "0"
    if name == "THY_CHANGE_PER_MILE_AMT":
        return cht._norm_cell(name, raw)
    if name in profile.amounts:
        return cht._norm_amount(raw)
    # The installed-Excel reading (CMP-AUD-197): the export serializes control
    # characters as OOXML `_xHHHH_` escapes; decoded, any such character stays
    # compared content (the Ramp Detail vs-TSN rule).
    text = ctc.decode_ooxml_escapes(_s(raw))
    if collapse:
        text = _WS.sub(" ", text)
    if name == profile.design_col and side == "tsn" and profile.tsn_design_prefix:
        if len(text) == 2 and text[0].upper() == profile.tsn_design_prefix:
            text = text[1:]
    elif name == "INX_CONTROL_CODE":
        text = _norm_control_type(text)
    elif name == "RAM_ON_OFF_CODE" and text.upper() == "Z":
        text = "OTH"
    return text


def _self_value(profile, name, raw):
    """The PDF-vs-Excel projection: both sides are the SITE's own renders, so
    every value is verbatim except what one render structurally cannot carry —
    the whitespace runs the HTML print collapses, the Excel export's OOXML
    escapes and edge tab padding (the CMP-AUD-197 same-source rule) — and a
    date the workbook may have typed."""
    if isinstance(raw, (datetime, date)):
        return _date(raw)
    return _WS.sub(" ", ctc.same_source_render_text(_s(raw)))


# --------------------------------------------------------------------------- #
# the site's own spelling (v0.54.1)
# --------------------------------------------------------------------------- #
_ISO_DAY = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
_PLAIN_NUMBER = re.compile(r"(-?\d+)(?:\.(\d+))?")


def _written_date(text, short):
    """An ISO date the way the site's file writes it: MM/DD/YYYY, or YY-MM-DD
    for the Intersection layer's dates. `iso_date` reads a 2-digit year below
    30 as 20xx, so only 1930–2029 goes back to two digits; any other date, and
    anything that is not a real ISO date, stays as it is."""
    m = _ISO_DAY.fullmatch(text)
    if not m:
        return text
    yyyy, mm, dd = m.groups()
    try:
        date(int(yyyy), int(mm), int(dd))
    except ValueError:  # silent-ok: an impossible date is kept verbatim on purpose (CMP-AUD-038), so it shows as written
        return text
    if short:
        return f"{yyyy[2:]}-{mm}-{dd}" if 1930 <= int(yyyy) <= 2029 else text
    return f"{mm}/{dd}/{yyyy}"


def _written_three_places(text):
    """A canonical number with three decimals ('0.78' -> '0.780'); one with
    more decimals than three, or no number, stays as it is (never rounded)."""
    m = _PLAIN_NUMBER.fullmatch(text)
    if not m or len(m.group(2) or "") > 3:
        return text
    return f"{m.group(1)}.{(m.group(2) or '').ljust(3, '0')}"


def as_written(profile, name, value):
    """One compared value (after `_value`) the way the site's Clean Road file
    writes that column, so a sheet reads like the file: MM/DD/YYYY dates (the
    Intersection layer's YY-MM-DD), Highway's three-decimal postmiles and
    two-digit lanes and widths, the period after a two-letter county, a blank
    toll / forest code for none, 'Z' for the Ramp's 'other' class.

    One to one with the compared form: two values are equal after this exactly
    when they were equal before it, so no count can move. A typed key cell is
    left to `rows_as_written`."""
    if not isinstance(value, str) or isinstance(value, cc.PhysicalKey) or not value:
        return value
    if name in profile.dates:
        return _written_date(value, name in profile.short_dates)
    if name in profile.three_places:
        return _written_three_places(value)
    if name in profile.two_digits:
        return "0" + value if len(value) == 1 and "0" <= value <= "9" else value
    if name.endswith("_COUNTY_CODE"):
        return value + "." if len(value) == 2 else value
    if name == "THY_TOLL_FOREST_CODE":
        return "" if value == "0" else value
    if name == "RAM_ON_OFF_CODE":
        return "Z" if value == "OTH" else value
    return value


_US_DAY = re.compile(r"(\d{2})/(\d{2})/(\d{4})")
_SHORT_DAY = re.compile(r"(\d{2})-(\d{2})-(\d{2})")
_THREE_PLACES = re.compile(r"(-?\d+)\.(\d{3})")


def as_compared(profile, name, text):
    """`as_written` undone: a shown value back in the compared spelling, exactly
    (ISO dates, canonical numbers, no county period, toll 0, 'OTH'). The
    pairing's character likeness reads this (`CompareSchema.likeness_text`),
    so the file's spelling never decides which rows of one location pair."""
    if not isinstance(text, str) or not text:
        if name == "THY_TOLL_FOREST_CODE" and text == "":
            return "0"
        return text
    if name in profile.dates:
        m = (_SHORT_DAY if name in profile.short_dates else _US_DAY).fullmatch(text)
        if not m:
            return text
        if name in profile.short_dates:
            yy, mm, dd = m.groups()
            yyyy = str((1900 if int(yy) >= 30 else 2000) + int(yy))
        else:
            mm, dd, yyyy = m.groups()
        try:
            date(int(yyyy), int(mm), int(dd))
        except ValueError:  # silent-ok: not a date as_written made; it was kept verbatim
            return text
        return f"{yyyy}-{mm}-{dd}"
    if name in profile.three_places:
        m = _THREE_PLACES.fullmatch(text)
        if not m:
            return text
        frac = m.group(2).rstrip("0")
        return m.group(1) + ("." + frac if frac else "")
    if name in profile.two_digits:
        return text[1:] if len(text) == 2 and text[0] == "0" and "0" <= text[1] <= "9" else text
    if name.endswith("_COUNTY_CODE"):
        return text[:-1] if len(text) == 3 and text.endswith(".") else text
    if name == "RAM_ON_OFF_CODE":
        return "OTH" if text == "Z" else text
    return text


def _written_key(profile, key):
    """The key cell with the postmile as the file writes it ('0.780', the
    prefix and roadbed in their own columns). Only the visible text changes:
    the typed identity, which pairs the rows, is the same object."""
    if not isinstance(key, cc.PhysicalKey) or is_unplaced(key):
        return key
    raw = next((c.value for c in key.physical_identity.raw_claims
                if c.name == "postmile"), None)
    if raw in (None, ""):
        return key
    text = as_written(profile, profile.spec.pm_col, ctc.decimal_pm(raw))
    return cc.physical_key(text, key.physical_identity) if text else key


def rows_as_written(profile, rows):
    """Loaded [route, *header] rows with every value as the site's file
    writes it (`as_written`) — the last step before the engine, so every rule
    that reads the compared form has already run."""
    header = profile.spec.header
    pm_col = profile.spec.pm_col
    return [[row[0]] + [_written_key(profile, v) if name == pm_col
                        else as_written(profile, name, v)
                        for name, v in zip(header, row[1:])]
            for row in rows]


_ROUTE_RE = re.compile(r"(\d+)([A-Z]*)")


def route_identity(name, suffix):
    """One route token from a row's route-name + route-suffix cells, both sides
    alike. TSN writes the bare number ('005') with the suffix in its own column;
    the site's Clean Road Highway file writes the selected route LABEL ('005S')
    as the name and fills the suffix column on some rows only — so a suffix
    already in the name is never appended twice ('005S' + 'S' = '005S')."""
    text = _s(name).upper()
    if text.endswith(".0"):
        text = text[:-2]
    sfx = _s(suffix).upper()
    m = _ROUTE_RE.fullmatch(text)
    if not m:
        return text + sfx
    digits, own = m.groups()
    base = f"{int(digits):03d}"
    if own:
        return base + own + (sfx if sfx and sfx != own else "")
    return base + sfx


UNPLACED = "no postmile"


def _unplaced_key(spec, g, route, county, prefix, roadbed, source_hint):
    """The key of a row whose postmile cell is blank. With no postmile the row
    cannot be placed against the other side, so it keys on what it does carry —
    its odometer, where the file has one — under a label no real postmile can
    take, and pairs only with a row the same file shape leaves unplaced the same
    way (the site's other render, or another day's export)."""
    if not county:
        raise ValueError(
            f"{spec.label} row (route {route}) has neither a postmile nor a "
            f"county in {source_hint} — cannot key it to a location")
    odometer = cht._norm_amount(g.get(spec.odometer_col)) if spec.odometer_col else ""
    component = f"{UNPLACED} @ odometer {odometer}" if odometer else UNPLACED
    identity = cc.make_physical_identity(
        route, county, component,
        (cc.RawIdentityClaim("route", route),
         cc.RawIdentityClaim("county", county),
         cc.RawIdentityClaim("postmile_prefix", prefix),
         cc.RawIdentityClaim("postmile", ""),
         cc.RawIdentityClaim("roadbed", roadbed),
         cc.RawIdentityClaim("odometer", odometer)),
        f"{route} / {county} / {component}")
    return cc.physical_key(component, identity)


def is_unplaced(key):
    """Was this row keyed without a postmile (see `_unplaced_key`)?"""
    return str(key).startswith(UNPLACED)


def _key(profile, vals, source_hint, route=None):
    """(route, the physical key) for one row. `route` overrides the row's own
    route cells (the cross-environment path keys on the file's route token,
    which the engine requires the identity to equal)."""
    spec = profile.spec
    g = dict(zip(spec.header, vals))
    if route is None:
        route = route_identity(g.get(spec.route_col), g.get(spec.suffix_col))
    roadbed = _s(g.get(spec.pm_suffix_col)).upper()
    if roadbed not in ("R", "L", "X"):
        roadbed = ""                     # the site's equation marker 'E' is no roadbed
    county = _county(g.get(spec.county_col))
    prefix = _s(g.get(spec.pm_prefix_col)).upper()
    if not _s(g.get(spec.pm_col)):
        return route, _unplaced_key(spec, g, route, county, prefix, roadbed,
                                    source_hint)
    return route, cht._physical_span_key(
        route, county, prefix, g.get(spec.pm_col), roadbed, source_hint)


def _row(profile, vals, side, source_hint, *, same_source=False, collapse=False):
    """[route, *header values] with the postmile cell as the physical key."""
    spec = profile.spec
    route, key = _key(profile, vals, source_hint)
    out = [route]
    for name, raw in zip(spec.header, vals):
        if name == spec.pm_col:
            out.append(key)
        elif same_source:
            out.append(_self_value(profile, name, raw))
        else:
            out.append(_value(profile, name, raw, side, collapse))
    return out


# --------------------------------------------------------------------------- #
# loaders + role gates
# --------------------------------------------------------------------------- #
def _open(path):
    name = Path(path).name
    try:
        return name, load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise ValueError(f"Could not open {name}: {type(e).__name__}: {e}")


def load_site(path, profile, *, pdf, same_source=False, collapse=False):
    """Side A (and the self-check's side B): a consolidated SITE export —
    `clean_road_consolidate`'s sheet with "Route" + the export's exact header.
    `pdf` True = the print consolidation (marker required), False = the Excel
    one (marker refused)."""
    spec = profile.spec
    label = f"{spec.label} ({'PDF' if pdf else 'Excel'})"
    if pdf:
        ctc.require_pdf_source(path, f"TSMIS {label}", spec.label)
    else:
        ctc.reject_pdf_source(path, f"TSMIS {label}", spec.label)
    name, wb = _open(path)
    try:
        if _is_arcgis_build(wb):
            raise ValueError(
                f"{name} is the app's own ArcGIS-built Clean Road workbook, not "
                "the site's export — pick the consolidated Clean Road export "
                "(the Consolidate tab, or the matrix builds it).")
        if spec.sheet not in wb.sheetnames:
            raise ValueError(f"{name} has no '{spec.sheet}' sheet — pick the "
                             f"consolidated {label} workbook.")
        it = wb[spec.sheet].iter_rows(values_only=True)
        header = [_s(c) for c in (next(it, None) or ())]
        while header and not header[-1]:
            header.pop()
        if header != ["Route"] + list(spec.header):
            raise ValueError(
                f"{name} is not a consolidated {label} workbook with the "
                f"{len(spec.header)}-column export header — re-consolidate it "
                "with this version.")
        rows = []
        for r in it:
            if not ctc.row_has_data(r):
                continue
            vals = list(r[1:1 + len(spec.header)])
            vals += [None] * (len(spec.header) - len(vals))
            rows.append(_row(profile, vals, "site", f"{name} ({spec.sheet})",
                             same_source=same_source, collapse=collapse))
        return rows
    finally:
        wb.close()


def _tsn_values(profile, raw):
    """A TSN row (in its own header order) onto the site header order: the
    site-only columns are simply absent on the TSN side."""
    spec = profile.spec
    by_name = dict(zip(spec.tsn_header, raw))
    return [by_name.get(c) for c in spec.header]


def load_tsn(path, profile, *, collapse=False):
    """Side B: the raw TSN extract (`Sheet 1`, exact header) or the TSN
    library's normalized copy (marker-gated). A PDF conversion, the ArcGIS build
    or a consolidated site export is refused — none of them is TSN."""
    spec = profile.spec
    if pdf_source_marker_state(path) != 0:
        raise ValueError(
            f"{Path(path).name} is one of this app's Clean Road PDF conversions, "
            f"so it cannot stand as the TSN side — pick the TSN {spec.tsn_label} "
            "extract or the TSN library's normalized copy.")
    name, wb = _open(path)
    try:
        if _is_arcgis_build(wb):
            raise ValueError(
                f"{name} is an ArcGIS-built Clean Road workbook, so it cannot "
                f"stand as the TSN side — pick the TSN {spec.tsn_label} extract or "
                "the TSN library's normalized copy.")
        if spec.normalized_sheet in wb.sheetnames:
            it = wb[spec.normalized_sheet].iter_rows(values_only=True)
            header = [_s(c) for c in (next(it, None) or ())]
            ctc.require_shared_header_prefix(header, spec.tsn_header, (), name,
                                             spec.label)
            version = (_HIGHWAY_TSN_MARKER if spec is crc.HIGHWAY
                       else NORMALIZATION_VERSION)
            ctc.require_current_normalization(
                wb, name, version, "pre-v1: no in-workbook normalization marker")
            return [_row(profile, _tsn_values(profile, list(r)), "tsn",
                         f"{name} ({spec.normalized_sheet})", collapse=collapse)
                    for r in it if ctc.row_has_data(r)]
    finally:
        wb.close()
    return tsn_rows_from_raw(path, profile, collapse=collapse)


def required_identity(spec):
    """The TSN identity claims a raw row must carry."""
    return (spec.county_col, spec.route_col, spec.pm_col)


def tsn_rows_from_raw(path, profile, *, collapse=False):
    """Every row of the exact raw TSN statewide extract."""
    spec = profile.spec
    with ctc.exact_raw_rows(path, crc.TSN_RAW_SHEET, spec.tsn_header,
                            f"Clean Road {spec.tsn_label}",
                            required_nonblank=required_identity(spec)) as (_h, rows_in):
        return [_row(profile, _tsn_values(profile, list(r)), "tsn",
                     f"{Path(path).name} ({crc.TSN_RAW_SHEET})", collapse=collapse)
                for r in rows_in]


# --------------------------------------------------------------------------- #
# schemas + adapters
# --------------------------------------------------------------------------- #
def _schema(profile, side_a, side_b, *, notes_title, notes, context, representation,
            context_side="", written=False):
    """`context_side` names the TSMIS file's side: its own value is what a
    context column shows (`CompareSchema.context_display_side`). `written`:
    the loader writes values the file's way (`rows_as_written`), so the
    pairing's likeness reads them back in the compared spelling."""
    spec = profile.spec
    header = list(spec.header)
    widths = {c: 26 for c in profile.text}
    return CompareSchema(
        report_name=spec.label,
        header=header,
        side_a=side_a,
        side_b=side_b,
        id_noun=spec.noun,
        id_noun_plural=spec.noun_plural,
        pair_noun="postmile",
        sides_noun="systems" if side_b == "TSN" else "renders",
        date_fields=tuple(c for c in header if c in profile.dates),
        data_widths=widths,
        cmp_widths={c: 30 for c in profile.text},
        one_sided_note_extra=(f" ({spec.noun_plural} one side carries at a "
                              "physical location the other doesn't)"),
        key_field=header.index(spec.pm_col),
        context_fields=context,
        context_header_fill="808080" if context else "",
        context_display_side=context_side if context else "",
        likeness_text=partial(as_compared, profile) if written else None,
        legend_writer=ctc.make_notes_writer(notes_title, notes),
        representation_fields=representation,
    )


def _unplaced_routes(rows):
    """(count, sorted routes) of the loaded rows keyed without a postmile."""
    routes = set()
    count = 0
    for r in rows:
        if any(is_unplaced(v) for v in r[1:] if isinstance(v, cc.PhysicalKey)):
            count += 1
            routes.add(r[0])
    return count, sorted(routes)


def unplaced_note_lines(spec, rows_a, rows_b, label_a, label_b):
    """The Notes lines for rows a side's file lists with a BLANK postmile, or
    () when every row has one (so a normal comparison's Notes are unchanged)."""
    lines = []
    for label, rows in ((label_a, rows_a), (label_b, rows_b)):
        count, routes = _unplaced_routes(rows)
        if not count:
            continue
        shown = ", ".join(routes[:12]) + (f" and {len(routes) - 12} more"
                                          if len(routes) > 12 else "")
        many = count != 1
        lines.append(
            f"{count:,} {spec.noun_plural if many else spec.noun} in the "
            f"{label} file {'have' if many else 'has'} NO postmile (the postmile "
            f"cells are blank) — route{'s' if len(routes) != 1 else ''} {shown}. "
            "The blank is in that file itself, not something this comparison "
            "did. Without a postmile a row cannot be placed against the other "
            f"side, so each one is keyed '{UNPLACED} @ odometer …' on its own "
            "odometer and pairs only with a row the other side leaves unplaced "
            "at the same odometer; otherwise it lists as one-sided, and whatever "
            "the other side has over that stretch lists as one-sided too.")
    return tuple(lines)


_SELF_NOTES = (
    "Both sides are the SITE's own renders of the same run: the print (read "
    "back from its PDF) and the Excel export. Every value compares verbatim; "
    "a run of spaces counts as one space (the HTML print collapses them, the "
    "Excel file keeps them).",
    "The print leaves out the columns the site has no source for; they are "
    "blank in the Excel file too, so they compare equal here.",
    f"The Excel file writes a few control characters as escapes "
    f"('{ctc.CR_ESCAPE_TEXT}' is a carriage return) and can pad a cell's edge "
    "with tabs; the print cannot "
    "carry either, so they are decoded away and never count (the owner's "
    "PDF-vs-Excel ruling, CMP-AUD-197).",
)


class _CleanRoadCompare:
    """One Clean Road file-vs-file comparison: compare(path_a, path_b, …) +
    suggest_name(path_a), with the side labels carried to the workbook and
    the Compare tab's file pickers."""

    def __init__(self, profile, flavor):
        spec = profile.spec
        self._profile = profile
        self._flavor = flavor
        if flavor == "excel_tsn":
            self.file_a_label, self.file_b_label = "TSMIS", "TSN"
            self.REPORT_NAME = f"{spec.label} — TSMIS vs TSN"
            tag, collapse = "TSMIS_vs_TSN", False
        elif flavor == "pdf_tsn":
            self.file_a_label, self.file_b_label = "TSMIS (PDF)", "TSN"
            self.REPORT_NAME = f"{spec.label} — TSMIS (PDF) vs TSN"
            tag, collapse = "TSMIS_PDF_vs_TSN", True
        elif flavor == "pdf_excel":
            self.file_a_label, self.file_b_label = "TSMIS (PDF)", "TSMIS (Excel)"
            self.REPORT_NAME = f"{spec.label} — TSMIS PDF vs Excel"
            tag, collapse = "TSMIS_PDF_vs_Excel", True
        else:
            raise ValueError(f"unknown Clean Road comparison flavor {flavor!r}")
        self._tag = f"{tag}_{spec.label.replace(' ', '')}"
        self._collapse = collapse
        if flavor == "pdf_excel":
            self._notes_title = f"{spec.label} — TSMIS PDF vs Excel: comparison notes"
            self._notes = _SELF_NOTES
            self._schema = _schema(
                profile, self.file_a_label, self.file_b_label,
                notes_title=self._notes_title, notes=self._notes, context=(),
                representation=())
        else:
            lines = profile.notes + ((
                "This side was read back from the site's PRINT, whose HTML "
                "collapses runs of spaces; a run of spaces compares as one space "
                "in every flavor, and this flavor folds tabs and line breaks the "
                "same way on both sides.",) if collapse else (
                "The Excel export writes a few control characters as escapes "
                f"('{ctc.CR_ESCAPE_TEXT}' is a carriage return); they are read "
                "decoded, the way "
                "installed Excel shows the cell, and the character itself still "
                "compares.",))
            self._notes_title = f"{self.REPORT_NAME}: comparison notes"
            self._notes = lines
            self._schema = _schema(
                profile, self.file_a_label, self.file_b_label,
                notes_title=self._notes_title, notes=lines,
                context=profile.context, representation=profile.text,
                context_side=self.file_a_label, written=True)

    def suggest_name(self, path_a=None):
        return f"{self._tag}_Comparison {today_str()}.xlsx"

    def _load_pair(self, path_a, path_b):
        p = self._profile
        if self._flavor == "pdf_excel":
            rows_a = load_site(path_a, p, pdf=True, same_source=True)
            rows_b = load_site(path_b, p, pdf=False, same_source=True)
        else:
            rows_a = rows_as_written(p, load_site(
                path_a, p, pdf=(self._flavor == "pdf_tsn"), collapse=self._collapse))
            rows_b = rows_as_written(p, load_tsn(path_b, p, collapse=self._collapse))
        return rows_a, rows_b, None

    def compare(self, path_a, path_b, out_path, events=None, confirm_overwrite=None,
                mode="formulas", commit_guard=None):
        # Per call, never on the shared instance: the matrix runs cells in
        # parallel, and each one's Notes must describe its own two files.
        found = {}

        def loader(a, b):
            rows_a, rows_b, extra = self._load_pair(a, b)
            found["lines"] = unplaced_note_lines(
                self._profile.spec, rows_a, rows_b,
                self.file_a_label, self.file_b_label)
            return rows_a, rows_b, extra

        def notes(wb):
            return ctc.make_notes_writer(
                self._notes_title,
                tuple(self._notes) + tuple(found.get("lines", ())))(wb)

        return ctc.run_files_compare(
            replace(self._schema, legend_writer=notes), path_a, path_b, out_path,
            banner=(f"{self._profile.spec.label} Comparison — {self.file_a_label} "
                    f"vs {self.file_b_label}"),
            has_route=True, loader=loader, deps_ok=_DEPS_OK,
            deps_msg="Required components are missing (openpyxl).",
            side_a=self.file_a_label, side_b=self.file_b_label,
            events=events, confirm_overwrite=confirm_overwrite, mode=mode,
            commit_guard=commit_guard)


HIGHWAY_VS_TSN = _CleanRoadCompare(_HIGHWAY, "excel_tsn")
HIGHWAY_PDF_VS_TSN = _CleanRoadCompare(_HIGHWAY, "pdf_tsn")
HIGHWAY_PDF_VS_EXCEL = _CleanRoadCompare(_HIGHWAY, "pdf_excel")
INTERSECTION_VS_TSN = _CleanRoadCompare(_INTERSECTION, "excel_tsn")
INTERSECTION_PDF_VS_TSN = _CleanRoadCompare(_INTERSECTION, "pdf_tsn")
INTERSECTION_PDF_VS_EXCEL = _CleanRoadCompare(_INTERSECTION, "pdf_excel")
RAMP_VS_TSN = _CleanRoadCompare(_RAMP, "excel_tsn")
RAMP_PDF_VS_TSN = _CleanRoadCompare(_RAMP, "pdf_tsn")
RAMP_PDF_VS_EXCEL = _CleanRoadCompare(_RAMP, "pdf_excel")


def profile_for(key):
    """The comparison profile for a Clean Road export key (Excel or print)."""
    spec = crc.BY_KEY.get(key) or crc.BY_PDF_KEY.get(key)
    return _PROFILES[spec.key] if spec is not None else None


__all__ = ["NORMALIZATION_VERSION", "HIGHWAY_VS_TSN", "HIGHWAY_PDF_VS_TSN",
           "HIGHWAY_PDF_VS_EXCEL", "INTERSECTION_VS_TSN", "INTERSECTION_PDF_VS_TSN",
           "INTERSECTION_PDF_VS_EXCEL", "RAMP_VS_TSN", "RAMP_PDF_VS_TSN",
           "RAMP_PDF_VS_EXCEL", "load_site", "load_tsn", "tsn_rows_from_raw",
           "profile_for"]
