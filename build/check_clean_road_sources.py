"""build/check_clean_road_sources.py — the Clean Road builds against the vendor's
source list (`clean_road_sources`).

The list says which feature-service layer and field each Clean Road attribute
comes from. This check holds the builds to it, so a build can never drift onto
another layer or field without the suite saying so:

  * every Highway and Ramp column the list covers is built from the layer and
    field the list names — and a build that declares another source is caught;
  * every layer the library refreshes carries the number the list gives it, and
    every layer a Clean Road build requires is on the list;
  * the attributes the file has no column for, and the Intersection source the
    site's code does not read (the library refreshes it from its own service
    since 2026-11-22, and every field the list found is one of its fields) are
    exactly the ones the module documents — a change in any of them is a
    decision, not a drift;
  * the data-source line a comparison prints: the layers' feature service beside
    the source the export was taken from.

Run with the build venv:
    build\\.venv\\Scripts\\python.exe build\\check_clean_road_sources.py
"""
import sys

from _checklib import Checker, scripts_path

scripts_path()

import arcgis_refresh  # noqa: E402
import arcgis_report_clean_intersection as acri  # noqa: E402
import arcgis_report_clean_ramp as acrr  # noqa: E402
import clean_highway_columns as chc  # noqa: E402
import clean_road_columns as crc  # noqa: E402
import clean_road_layers as crl  # noqa: E402
import clean_road_sources as crs  # noqa: E402
import compare_clean_highway_arcgis as cmp_hwy  # noqa: E402
import compare_clean_ramp_arcgis as cmp_ram  # noqa: E402
import consolidate_clean_highway as cch  # noqa: E402

c = Checker()


def test_agreement():
    print("the builds against the list:")
    bad = crs.mismatches("clean_highway")
    c.check("every Highway column the list covers is built from the layer and field "
            "it names", bad == [], bad)
    bad = crs.mismatches("clean_ramp", acrr.PROVENANCE)
    c.check("every Ramp column the list covers is built from the layer and field it "
            "names", bad == [], bad)
    wrong = dict(acrr.PROVENANCE, RAM_CITY_CODE=("SHS Population.City_Code", "x"))
    c.check("a build that declares another layer for a column is caught",
            crs.mismatches("clean_ramp", wrong)
            == ["RAM_CITY_CODE: the list says City.City_Code, the build declares "
                "SHS Population.City_Code"], crs.mismatches("clean_ramp", wrong))
    old = chc.PROVENANCE["THY_MEDIAN_WIDTH_AMT"]
    chc.PROVENANCE["THY_MEDIAN_WIDTH_AMT"] = (old[0], old[1], "Median_Type", old[3])
    try:
        bad = crs.mismatches("clean_highway")
    finally:
        chc.PROVENANCE["THY_MEDIAN_WIDTH_AMT"] = old
    c.check("...and so is another field of the right layer",
            len(bad) == 1 and bad[0].startswith("THY_MEDIAN_WIDTH_AMT"), bad)

    hwy, ram = crs.by_column("clean_highway"), crs.by_column("clean_ramp")
    c.check("the list covers 29 Highway columns from 30 attributes (toll and forest "
            "share one) and 9 Ramp columns from 8",
            len(hwy) == 29 and sum(len(v) for v in hwy.values()) == 30
            and len(ram) == 9 and sum(len(v) for v in ram.values()) == 8 + 1,
            (len(hwy), len(ram)))
    c.check("every covered column is a real column of the site's file",
            set(hwy) <= set(crc.HIGHWAY.header) and set(ram) <= set(crc.RAMP.header))
    counted_h = set(cmp_hwy.SHARED_HEADER) - set(cmp_hwy.CONTEXT_FIELDS)
    counted_r = set(crc.RAMP.header) - set(cmp_ram._COMPARE.context)
    c.check("the comparisons COUNT every covered column — the one exception is "
            "RAM_CREATE_DATE, which the site's file leaves blank",
            set(hwy) <= counted_h and set(ram) - counted_r == {"RAM_CREATE_DATE"},
            (set(hwy) - counted_h, set(ram) - counted_r))


def test_layers():
    print("the layers:")
    off = [(name, number, crs.LAYER_NUMBERS.get(name))
           for name, number in arcgis_refresh.LIBRARY_LAYERS
           if name not in arcgis_refresh.LAYER_SERVICES
           and crs.LAYER_NUMBERS.get(name) != number]
    c.check("every layer the library refreshes from the main service carries the number "
            "the list gives it", off == [], off)
    c.check("Master Intersection is refreshed from its own service, where production "
            "numbers it 204 (the list's development copy numbers it 0)",
            arcgis_refresh.LAYER_SERVICES.get(crs.INTERSECTION_SOURCE)
            == arcgis_refresh.CSMS_SERVICE
            and arcgis_refresh.LAYER_IDS[crs.INTERSECTION_SOURCE] == 204
            and crs.INTERSECTION_LIST_NUMBER == 0
            and crs.INTERSECTION_SOURCE not in crs.LAYER_NUMBERS)
    required = (set(cch.REQUIRED_LAYERS) | set(acri.REQUIRED_LAYERS)
                | set(acrr.REQUIRED_LAYERS))
    c.check("every layer a Clean Road build requires is on the list",
            required <= set(crs.LAYER_NUMBERS), required - set(crs.LAYER_NUMBERS))
    named = {src.layer for key in ("clean_highway", "clean_ramp")
             for src in crs.sources(key)}
    c.check("every layer the list names is a library layer (Posted Speed Limit "
            "joined the refresh 2026-11-18)",
            named <= set(crl.EXPECTED_LAYERS), named - set(crl.EXPECTED_LAYERS))
    c.check("the layers the site's code reads that the list does not name as a source "
            "are library layers too (Calibration Point, HPMS F System)",
            {"Calibration Point", "HPMS F System"} <= set(crl.EXPECTED_LAYERS)
            and crs.LAYER_NUMBERS["Calibration Point"] == 1
            and crs.LAYER_NUMBERS["HPMS F System"] == 91)
    c.check("every list layer that fills a column is one its build requires",
            {src.layer for src in crs.sources("clean_highway") if src.columns}
            <= set(cch.REQUIRED_LAYERS)
            and {src.layer for src in crs.sources("clean_ramp") if src.columns}
            <= set(acrr.REQUIRED_LAYERS))
    c.check("a layer shows with its number",
            crs.layer_tag("SHS Median") == "SHS Median #124"
            and crs.layer_tag("not a layer") == "not a layer")


def test_documented_gaps():
    print("what the list names that the file or the site's code does not carry:")
    c.check("Highway: four attributes with no column",
            [s.field for s in crs.unplaced("clean_highway")]
            == ["Posted_Speed_Limit", "Facility_Name", "Thru_Num_Lanes_L",
                "Thru_Num_Lanes_R"])
    c.check("Ramp: four attributes with no column",
            [s.field for s in crs.unplaced("clean_ramp")]
            == ["Ramp_Name", "Side_of_Hwy", "Ramp_Length", "Ramp_Num_Lanes"])
    c.check("Intersection: one source, now a library layer, 29 fields found and 3 the "
            "list could not find",
            crs.INTERSECTION_SOURCE in crl.EXPECTED_LAYERS
            and crl.EXPECTED_LAYERS[crs.INTERSECTION_SOURCE] == "I"
            and len(crs.INTERSECTION_FIELDS) == 29
            and len(crs.INTERSECTION_NOT_FOUND) == 3)
    header = set(crs.MASTER_INTERSECTION_HEADER)
    c.check("every field the list found is a field of the production layer, and none "
            "of the three it could not find is",
            set(crs.INTERSECTION_FIELDS) <= header
            and not header & set(crs.INTERSECTION_NOT_FOUND),
            (set(crs.INTERSECTION_FIELDS) - header, header & set(crs.INTERSECTION_NOT_FOUND)))
    c.check("the production layer: 72 fields, both legs with their own route and "
            "postmile, the inventory dates, no GlobalID",
            len(crs.MASTER_INTERSECTION_HEADER) == 72
            and {"Major_PMMeasure", "Minor_PMMeasure", "Major_ARRouteID",
                 "InventoryItemStartDate", "InventoryItemEndDate"} <= header
            and "GlobalID" not in header and "Minor_District" not in header)
    text = " ".join(crs.notes_lines("clean_intersection"))
    c.check("the Intersection notes say which tables the build reads instead",
            "Master Intersection" in text and "IM Intersection Detail #151" in text
            and "with the IM tables only" in text, text)
    text = " ".join(crs.notes_lines("clean_highway"))
    c.check("the Highway notes give the verdict, the attributes with no column and "
            "where the list and the site's code part",
            "all 29 agree" in text and "Posted_Speed_Limit (Posted Speed Limit #107"
            in text and "Toll_Type" in text, text)
    text = " ".join(crs.notes_lines("clean_ramp", acrr.PROVENANCE))
    c.check("the Ramp notes do the same", "all 9 agree" in text
            and "Side_of_Hwy" in text and "RAM_CREATE_DATE" in text, text)
    rows = crs.highway_table_lines({"THY_MEDIAN_TYPE_CODE"})
    c.check("the Highway column table: every build column, its layer with the list's "
            "number, on the list or not, counted or context",
            len(rows) == len(chc.ARC_HEADER)
            and any("THY_MEDIAN_TYPE_CODE  ·  built from SHS Median #124 (Median_Type)"
                    "  ·  on the list  ·  counted" in r for r in rows)
            and any("THY_TOLL_FOREST_CODE  ·  built from SHS Tolls #138 + SHS Forest "
                    "HWY #115 ((coverage))  ·  on the list  ·  context" in r for r in rows)
            and any("THY_LANDMARK_SHORT_DESC  ·  built from SHS Landmark #123" in r
                    and "not an attribute on the list" in r for r in rows), rows[:3])
    rows = crs.point_table_lines(crc.RAMP, acrr.PROVENANCE, "clean_ramp",
                                 cmp_ram._COMPARE.context)
    c.check("the Ramp column table tags each layer, marks the list's columns, and "
            "says counted or context as the comparison itself has it (the extract "
            "date and the two columns the site leaves blank are context)",
            any("RAM_DESCRIPTION  ·  SHS Ramp #131.Ramp_Description (verbatim)"
                "  ·  on the list  ·  counted" in r for r in rows)
            and any("RAM_PM_LOC_AMT  ·  SHS Ramp Pt #132.PMMeasure" in r
                    and "on the list" not in r for r in rows)
            and all(r.endswith("context") for r in rows
                    if r.split()[0] in ("RAM_EXTRACT_DATE", "RAM_CREATE_DATE",
                                        "RAM_BEGIN_OFFSET_AMT"))
            and sum(r.endswith("context") for r in rows) == 3, rows[:3])


def test_service():
    print("which service:")
    default = arcgis_refresh.SERVICE_DEFAULT
    index = {"SHS Tolls": {"source": default + "/138"},
             "City": {"source": default + "/74"}}
    c.check("a library's service is its layers' common feature service",
            crs.service_of(default + "/138") == default
            and crs.library_service(index) == default
            and crs.service_of("https://example.invalid/FeatureServer;VERSION=x/138")
            == "https://example.invalid/FeatureServer")
    c.check("an index that records none says none; mixed sources say so",
            crs.library_service({"A": {"source": "synthetic"}}) == ""
            and crs.library_service({}) == ""
            and crs.library_service(dict(index, X={"source":
                                                   "https://other/FeatureServer/1"}))
            == "several services")
    csms = arcgis_refresh.CSMS_SERVICE
    with_mi = dict(index, **{"Master Intersection": {"source": f"{csms}/204"}})
    c.check("Master Intersection's own service never makes the library read as mixed; "
            "asked for it alone, its service is its own",
            crs.library_service(with_mi) == default
            and crs.library_service(with_mi, ["Master Intersection"]) == csms
            and crs.library_service(with_mi, ["SHS Tolls", "City"]) == default
            and crs.library_service(with_mi, ["SHS Tolls", "Master Intersection"])
            == "several services")
    name = "clean_road_highway_consolidated 2026-11-06 ssor-prod dev-site.xlsx"
    c.check("an export's file name says the data source it was taken from",
            crs.export_source(name) == ("2026-11-06", "ssor", "prod")
            and crs.export_source("C:/x/tsmis_clean_ramp_pdf_consolidated 2026-10-02 "
                                  "ars-prod.xlsx") == ("2026-10-02", "ars", "prod")
            and crs.export_source("clean_road_highway_consolidated.xlsx") is None
            and crs.export_source("report 2026-11-06.xlsx") is None)
    same = crs.source_note(default, name)
    c.check("SSOR prod export + the default service: one data source",
            "both sides come from one data source" in same and "⚠" not in same, same)
    ars = crs.source_note(
        default, "clean_road_highway_consolidated 2026-11-13 ars-prod dev-site.xlsx")
    c.check("ARS prod export + the default service: the same data, up to a day behind "
            "the layers (the owner, 2026-11-19: ARS takes 24 hours to update)",
            "ARS prod source" in ars and "up to 24 hours" in ars
            and "up to a day behind the layers" in ars and "⚠" not in ars, ars)
    for other in ("clean_road_highway_consolidated 2026-11-13 ars-test dev-site.xlsx",
                  "clean_road_highway_consolidated 2026-11-13 ssor-test dev-site.xlsx"):
        note = crs.source_note(default, other)
        c.check(f"{other.split()[2]}: flagged as a different service",
                "⚠" in note and "DIFFERENT service" in note, note)
    note = crs.source_note("https://other/FeatureServer",
                           "clean_road_highway_consolidated 2026-11-13 ars-prod dev-site.xlsx")
    c.check("an ARS prod export against layers from another service is flagged",
            "⚠" in note and "DIFFERENT service" in note, note)
    note = crs.source_note("https://other/FeatureServer", name)
    c.check("an SSOR prod export against layers from another service is flagged too",
            "⚠" in note, note)
    c.check("a build that records no service says so instead of guessing",
            "does not record which feature service" in crs.source_note("", name)
            and "⚠" not in crs.source_note("", name))
    c.check("an export with no source in its name says so",
            "does not say which data source" in crs.source_note(default, "x.xlsx"))


def main():
    c.check("the list holds 97 layers and tables, 34 Highway and 12 Ramp attributes",
            len(crs.SERVICE_LAYERS) == 97 and len(crs.HIGHWAY) == 34
            and len(crs.RAMP) == 12)
    test_agreement()
    test_layers()
    test_documented_gaps()
    test_service()
    return c.summary()


if __name__ == "__main__":
    sys.exit(main())
