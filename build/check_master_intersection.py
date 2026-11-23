"""build/check_master_intersection.py — Clean Road Intersection built from the
Master Intersection layer (2026-11-22), the source the vendor's list names,
beside the build from the IM tables the site's code reads.

Drives the SHIPPED path on synthetic data (no corpus files):

  * the registry: the Master Intersection rows come first and are labeled, and
    they compare against the Clean Road Intersection export's own editions;
    the IM rows are labeled by their source too;
  * the build through its real `consolidate()`: the row universe (current at
    the as-of date, a major route, the geometry in effect; a row a rebuild
    replaced reads as of its day), the second row under a minor state route
    with the legs swapped, lanes from the leg's total, highway group and
    population from the file route's leg, county and district from the major
    leg, the columns the layer has no field for left blank. Its marker names
    its own report and the layer's own service, and it never stands as an
    IM-tables build;
  * the comparison: a clean match on the same records; a changed value the
    list attributes to the layer counts; the city, the begin date and the
    segment orders never do; the Notes say which source each side is;
  * the vendor's list: its 29 found fields fill 23 file columns, each from the
    field it names, and a build that declares another field is caught.

Run with the build venv:
    build\\.venv\\Scripts\\python.exe build\\check_master_intersection.py
"""
import datetime
import sys
from pathlib import Path

from _checklib import Checker, scripts_path, temp_dir

scripts_path()

from openpyxl import Workbook, load_workbook  # noqa: E402

import arcgis_clean_road_build as acb  # noqa: E402
import arcgis_layers  # noqa: E402
import arcgis_refresh  # noqa: E402
import arcgis_report_clean_intersection as acri  # noqa: E402
import arcgis_report_clean_intersection_mi as mi  # noqa: E402
import arcgis_reports  # noqa: E402
import clean_road_columns as crc  # noqa: E402
import clean_road_layers as crl  # noqa: E402
import clean_road_sources as crs  # noqa: E402
import compare_clean_intersection_arcgis as cmp_inx  # noqa: E402
import compare_clean_intersection_mi_arcgis as cmp_mi  # noqa: E402
import consolidate_clean_road_intersection as con_inx  # noqa: E402
from events import Events  # noqa: E402

c = Checker()
D = datetime.datetime
ASOF = "2026-10-05"
DAY = "2026-10-22 ssor-prod dev-site"


class _Quiet(Events):
    def on_log(self, msg):
        pass


def _sheet(path, title, header, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.append(list(header))
    for r in rows:
        ws.append([r.get(col) for col in header])
    wb.save(path)


def _mi(iid, pm, **over):
    """A Master Intersection row: a Tee on route 001 with a local cross street."""
    r = {f: None for f in crs.MASTER_INTERSECTION_HEADER}
    r.update({
        "Master_Intersection_ID": str(iid), "Date_Of_Record": D(1973, 10, 19),
        "Intersection_Geometry": "T", "Intersection_Geometry_Eff_Date": D(1964, 1, 1),
        "Intersection_Control": "B", "Intersection_Control_Eff_Date": D(1973, 10, 19),
        "Intersection_Lighted_Ind": "Yes", "Intersection_Lighted_Eff_Date": D(1980, 2, 3),
        "Description": "MAIN ST LT", "Major_District": "8", "Major_County": "SBD",
        "Major_RouteNum": "001", "Major_RouteSuffix": None, "Major_PMPrefix": None,
        "Major_PMSuffix": None, "Major_PMMeasure": pm,
        "Major_Effective_Date": D(1990, 5, 15), "Major_Signal_Arm_Ind": "No",
        "Major_Left_Channel": "P", "Major_Right_Channel_Ind": "No",
        "Major_FlowCode": "P", "Major_Pri_Num_Lanes": 1, "Major_Sec_Num_Lanes": 1,
        "Major_Total_Num_Lanes": 2, "Major_N_Distance": 250,
        "Major_Highway_Group": "D", "Major_Population_Code": "U", "Major_City": "SBD",
        "Minor_County": "SBD", "Minor_Effective_Date": D(1964, 1, 1),
        "Minor_Signal_Arm_Ind": "Yes", "Minor_Left_Channel": "N",
        "Minor_Right_Channel_Ind": "Yes", "Minor_FlowCode": "N",
        "Minor_Total_Num_Lanes": 4, "Minor_N_Distance": 150,
        "InventoryItemStartDate": D(2025, 9, 11), "InventoryItemEndDate": None,
        "OBJECTID": iid})
    r.update(over)
    return r


def _library(root, rows):
    root.mkdir(parents=True, exist_ok=True)
    name = f"44_{mi.MI_LAYER}.xlsx"
    _sheet(root / name, "204", crs.MASTER_INTERSECTION_HEADER, rows)
    _sheet(root / crl.INDEX_NAME, "INDEX", list(crl.INDEX_HEADER),
           [dict(zip(crl.INDEX_HEADER,
                     (name, mi.MI_LAYER, len(rows), len(crs.MASTER_INTERSECTION_HEADER),
                      mi.MI_LAYER, f"{arcgis_refresh.CSMS_SERVICE}/204")))])


def _read_build(path):
    wb = load_workbook(path, read_only=True)
    it = wb[mi.SPEC.sheet].iter_rows(values_only=True)
    header = [str(x) for x in next(it)]
    rows = [dict(zip(header, r)) for r in it]
    wb.close()
    return rows


def _site_export(folder, rows):
    run = folder / crc.INTERSECTION.key
    run.mkdir(parents=True, exist_ok=True)
    for route in sorted({r["Route"] for r in rows}):
        wb = Workbook()
        ws = wb.active
        ws.title = crc.INTERSECTION.sheet
        ws.append(list(crc.INTERSECTION.header))
        for r in rows:
            if r["Route"] == route:
                ws.append([r.get(h) or None for h in crc.INTERSECTION.header])
        wb.save(run / f"{DAY} {crc.INTERSECTION.key}_route_{route}.xlsx")
    out = folder / f"{crc.INTERSECTION.key}_consolidated {DAY}.xlsx"
    res = con_inx.consolidate(events=_Quiet(), confirm_overwrite=lambda p: True,
                              input_dir=run, out_path=out)
    c.check("the site export consolidates", res.status == "ok", res.message)
    return out


def _counts(result):
    co = getattr(result, "comparison_outcome", None)
    k = co.counts if co else None
    return {f: getattr(k, f, None) for f in ("paired_rows", "side_a_only_rows",
                                             "side_b_only_rows", "differing_cells")}


def _q():
    return dict(events=_Quiet(), confirm_overwrite=lambda p: True, mode="preview")


# --------------------------------------------------------------------------- #
def test_registry():
    print("the registry: the Master Intersection rows first, both builds labeled:")
    keys = list(arcgis_reports.KEYS)
    c.check("Master Intersection comes right before the IM-tables build",
            keys.index("clean_intersection_mi") + 1 == keys.index("clean_intersection"))
    eds = {e.key: e for e in arcgis_reports.editions()}
    m, p = eds["clean_intersection_mi"], eds["clean_intersection_mi_pdf"]
    i, ip = eds["clean_intersection"], eds["clean_intersection_pdf"]
    c.check("its rows are the Intersection export's editions under its own keys",
            (m.family, p.family) == ("clean_intersection_mi",) * 2
            and (m.subdir, m.consolidator) == (i.subdir, i.consolidator)
            and (p.subdir, p.consolidator) == (ip.subdir, ip.consolidator)
            and m.comparable and p.comparable)
    c.check("labeled by source: Master Intersection and IM tables",
            (m.label, p.label, i.label, ip.label)
            == ("Clean Road: Intersection (Master Intersection)",
                "Clean Road: Intersection (PDF, Master Intersection)",
                "Clean Road: Intersection (IM tables)",
                "Clean Road: Intersection (PDF, IM tables)"),
            (m.label, p.label, i.label, ip.label))
    c.check("its codes are its own", (m.code, p.code) == ("CR-INT-MI", "CR-INT-PDF-MI")
            and (i.code, ip.code) == ("CR-INT", "CR-INT-PDF"))
    s = arcgis_reports.spec("clean_intersection_mi")
    c.check("its build and comparator are the Master Intersection modules",
            s.build is mi and s.compare is cmp_mi and not s.why
            and arcgis_reports.label_of("clean_intersection_mi_pdf") == p.label)
    c.check("the build reads Master Intersection alone",
            mi.REQUIRED_LAYERS == ("Master Intersection",)
            and mi.MI_LAYER in crl.EXPECTED_LAYERS)


def test_build(tmp):
    print("the build from the layer, through its real consolidate():")
    lib = tmp / "mi_layers"
    rows = [
        _mi(10, 0.2),                                               # local cross street
        _mi(11, 1.5, Minor_RouteNum="133", Minor_PMMeasure=2.5,      # state x state
            Minor_Highway_Group="U", Minor_Population_Code="R", Minor_City="RIA",
            Minor_PMPrefix="R"),
        _mi(12, 3.0, Intersection_Control="S",                       # replaced by a rebuild
            InventoryItemEndDate=D(2026, 1, 1)),
        _mi(12, 3.0, InventoryItemStartDate=D(2026, 1, 1)),          # ...its current row
        _mi(13, 4.0, InventoryItemStartDate=D(2027, 1, 1)),          # not yet loaded
        _mi(14, 5.0, Intersection_Geometry_Eff_Date=D(2027, 1, 1)),  # not yet in effect
        _mi(15, None, Major_RouteNum=None),                         # a shell row
    ]
    _library(lib, rows)
    out = tmp / mi.FILENAME
    res = mi.consolidate(events=_Quiet(), confirm_overwrite=lambda p: True,
                         out_path=out, asof=ASOF, lib_root=lib)
    c.check("the build runs from the layer", res.status == "ok", res.message)
    got = _read_build(out)
    c.check("the universe: current at the as-of date, a major route, the geometry in "
            "effect; the state x state intersection is listed under BOTH routes",
            [(r["Route"], r["INX_CONNECTION_ID"]) for r in got]
            == [("001", "10"), ("001", "11"), ("001", "12"), ("133", "11")],
            [(r["Route"], r["INX_CONNECTION_ID"]) for r in got])
    a, b, r12, x = got
    c.check("the replaced row reads as of its day: the current row's control code",
            r12["INX_CONTROL_CODE"] == "B", r12["INX_CONTROL_CODE"])
    c.check("the site's forms: codes, Y/N, MM/DD/YYYY, natural postmile, begin = end, "
            "district 08",
            (a["INX_DESIGN_CODE"], a["INX_LIGHTED_IND"], a["INX_LIGHTED_BEGIN_DATE"],
             a["INX_RECORD_DATE"], a["INX_BEGIN_PM_AMT"], a["INX_END_PM_AMT"],
             a["INX_COUNTY_CODE"], a["INX_DISTRICT_CODE"], a["INX_INTERSECTION_NAME"])
            == ("T", "Y", "02/03/1980", "10/19/1973", "0.2", "0.2", "SBD", "08",
                "MAIN ST LT"), a)
    c.check("the major leg is the mainline block, the minor leg the cross street; lanes "
            "are each leg's total",
            (b["INX_MAIN_LANES_AMT"], b["INX_MAIN_OVERRIDE_LENGTH_AMT"],
             b["INX_MAIN_SIGNAL_MAST_ARM_IND"], b["INX_MAIN_BEGIN_DATE"],
             b["INX_CROSS_LANES_AMT"], b["INX_CROSS_SIGNAL_MAST_ARM_IND"],
             b["INX_CROSS_RIGHT_CHANNEL_CODE"], b["INX_CROSS_BEGIN_DATE"],
             b["INX_HIGHWAY_GROUP"], b["INX_POPULATION_GROUP"])
            == ("2", "250", "N", "05/15/1990", "4", "Y", "Y", "01/01/1964", "D", "U"), b)
    c.check("the minor route's row: on 133 at its postmile, crossing 001, the legs, "
            "their dates and the highway group and population swapped; county and "
            "district stay the major leg's",
            (x["INX_ROUTE_NAME"], x["INX_BEGIN_PM_AMT"], x["INX_PM_PREFIX_CODE"],
             x["INX_X_ROUTE_NAME"], x["INX_X_POSTMILE_AMT"], x["INX_MAIN_LANES_AMT"],
             x["INX_CROSS_LANES_AMT"], x["INX_MAIN_BEGIN_DATE"],
             x["INX_CROSS_BEGIN_DATE"], x["INX_HIGHWAY_GROUP"],
             x["INX_POPULATION_GROUP"], x["INX_COUNTY_CODE"], x["INX_DISTRICT_CODE"])
            == ("133", "2.5", "R", "001", "1.5", "4", "2", "01/01/1964", "05/15/1990",
                "U", "R", "SBD", "08"), x)
    c.check("no field in the layer: begin date and both segment orders stay blank; the "
            "city is the file route's leg's own; the unsourced columns stay blank",
            all(not r[col] for r in got for col in
                ("INX_BEGIN_DATE", "INX_SEG_ORDER_ID", "INX_X_SEG_ORDER_ID")
                + crc.INTERSECTION.unsourced)
            and (b["INX_CITY_CODE"], x["INX_CITY_CODE"]) == ("SBD", "RIA"))
    facts = acb.report_facts(out)
    c.check("the marker names its own report and the layer's own service",
            facts.get("report") == "clean_intersection_mi"
            and facts.get("layer_service") == arcgis_refresh.CSMS_SERVICE, facts)
    c.check("it never stands as an IM-tables build",
            mi.is_arcgis_report(out) and not acri.is_arcgis_report(out))
    return out, got


def test_compare(tmp, build, got):
    print("the comparison with the site's file:")
    site_rows = [dict(r) for r in got]
    for r in site_rows:
        r["INX_EXTRACT_DATE"] = "10/22/2026"
    same = _site_export(tmp / "site_same", site_rows)
    counts = _counts(cmp_mi.compare(str(build), str(same), str(tmp / "c1.xlsx"), **_q()))
    c.check("the same records: a clean match",
            counts == {"paired_rows": 4, "side_a_only_rows": 0, "side_b_only_rows": 0,
                       "differing_cells": 0}, counts)
    site_rows[0]["INX_HIGHWAY_GROUP"] = "U"           # a listed attribute: counts
    site_rows[0]["INX_CITY_CODE"] = "XX"              # not on the list: context
    site_rows[1]["INX_SEG_ORDER_ID"] = "150"          # no field in the layer: context
    site_rows[1]["INX_BEGIN_DATE"] = "01/01/2022"     # no field in the layer: context
    diff = _site_export(tmp / "site_diff", site_rows)
    counts = _counts(cmp_mi.compare(str(build), str(diff), str(tmp / "c2.xlsx"), **_q()))
    c.check("a listed attribute counts; the city, the begin date and a segment order "
            "never do",
            counts["differing_cells"] == 1 and counts["paired_rows"] == 4, counts)
    c.check("the comparison's context is the unsourced columns, the extract date and "
            "the four the layer does not give",
            set(cmp_mi.CONTEXT_FIELDS) == set(crc.INTERSECTION.unsourced)
            | {"INX_EXTRACT_DATE"} | set(mi.EXTRA_CONTEXT), cmp_mi.CONTEXT_FIELDS)
    res = cmp_mi.compare(str(same), str(build), str(tmp / "c3.xlsx"), **_q())
    c.check("swapped sides are refused", res.status == "error", res.status)
    res = cmp_inx.compare(str(build), str(same), str(tmp / "c4.xlsx"), **_q())
    c.check("the IM-tables comparison refuses a Master Intersection build", res.status
            == "error" and "not an ArcGIS-built" in (res.message or ""), res.message)
    written = cmp_mi.compare(str(build), str(same), str(tmp / "c5.xlsx"),
                             events=_Quiet(), confirm_overwrite=lambda p: True,
                             mode="values")
    wb = load_workbook(tmp / "c5.xlsx", read_only=True)
    notes = " ".join(str(v) for row in wb["Notes"].iter_rows(values_only=True)
                     for v in row if v)
    title = next((str(v) for r in wb["Summary"].iter_rows(max_row=4, values_only=True)
                  for v in r if v not in (None, "")), "")
    wb.close()
    c.check("the workbook is titled by its source",
            written.status == "ok" and "Master Intersection" in title, title)
    c.check("the Notes say which source each side is: the vendor's stated layer here, "
            "the IM tables beside it, and the layer's service against the export's",
            "Master Intersection #204" in notes and "where it is layer 0" in notes
            and "beside this one" in notes and "all 23 agree" in notes
            and "the Master Intersection layer was pulled from "
                f"{arcgis_refresh.CSMS_SERVICE}" in notes
            and "INX_CITY_CODE" in notes and "Master_Intersection_ID" in notes, notes[:300])


_DAY0 = datetime.time(0, 0)   # Excel's day 0, 12/30/1899, as openpyxl reads it
_DAY0_FIELDS = ("Date_Of_Record", "Intersection_Geometry_Eff_Date",
                "Intersection_Control_Eff_Date", "Intersection_Lighted_Eff_Date",
                "Major_Effective_Date", "Minor_Effective_Date")
_DAY0_COLS = ("INX_RECORD_DATE", "INX_DESIGN_DATE", "INX_CONTROL_DATE",
              "INX_LIGHTED_BEGIN_DATE", "INX_MAIN_BEGIN_DATE", "INX_CROSS_BEGIN_DATE")


def test_day0_and_export_date(tmp):
    print("Excel's day 0 and the build's own export date (2026-11-22 audit):")
    c.check("a bare time is Excel's day 0, written 12/30/1899",
            crl.to_serial(_DAY0) == 0.0 and acb.mdy(_DAY0) == "12/30/1899")
    lib = tmp / "day0_layers"
    lib.mkdir(parents=True, exist_ok=True)
    rows = [_mi(20, 0.5), _mi(21, 1.0, **{f: _DAY0 for f in _DAY0_FIELDS})]
    name, other = f"44_{mi.MI_LAYER}.xlsx", "16_SHS Ramp.xlsx"
    _sheet(lib / name, "204", crs.MASTER_INTERSECTION_HEADER, rows)
    _sheet(lib / other, "SHS Ramp", ["OBJECTID"], [{"OBJECTID": 1}])
    # A library refreshed in parts: Master Intersection on 11-22, the rest on 11-19.
    arcgis_layers.write_index(lib, {
        mi.MI_LAYER: {"file": name, "rows": len(rows),
                      "fields": len(crs.MASTER_INTERSECTION_HEADER), "path": mi.MI_LAYER,
                      "source": f"{arcgis_refresh.CSMS_SERVICE}/204",
                      "exported_at": "2026-11-22T13:58:10"},
        "SHS Ramp": {"file": other, "rows": 1, "fields": 1, "path": "SHS Ramp",
                     "source": "synthetic", "exported_at": "2026-11-19T09:00:00"}})
    out = tmp / f"day0_{mi.FILENAME}"
    res = mi.consolidate(events=_Quiet(), confirm_overwrite=lambda p: True,
                         out_path=out, asof=ASOF, lib_root=lib)
    c.check("the build runs", res.status == "ok", res.message)
    got = _read_build(out)
    c.check("a row whose every date is day 0 is current, as the site lists it (it was "
            "read as having no geometry date and left out)",
            [r["INX_CONNECTION_ID"] for r in got] == ["20", "21"],
            [r["INX_CONNECTION_ID"] for r in got])
    c.check("... and its dates are written 12/30/1899",
            len(got) == 2 and all(got[1][col] == "12/30/1899" for col in _DAY0_COLS),
            [got[1].get(col) for col in _DAY0_COLS] if len(got) == 2 else got)

    drop = arcgis_layers.drop_info(lib)
    own = arcgis_layers.for_layers(drop, mi.REQUIRED_LAYERS)
    c.check("the library dates itself by its oldest layer and says it is mixed",
            drop["exported"] == "2026-11-19" and drop["mixed"], drop)
    c.check("a build's date is its own layers' oldest, and mixed only over them; the "
            "fingerprint stays the library's",
            own["exported"] == "2026-11-22" and not own["mixed"]
            and own["fingerprint"] == drop["fingerprint"]
            and arcgis_layers.consistent_asof(mi.REQUIRED_LAYERS, drop) == "2026-11-22",
            own)
    c.check("a layer the manifest does not date leaves the library's date",
            arcgis_layers.for_layers(drop, ("No Such Layer",)) is drop)
    c.check("the build's marker states the date of the layer it read, not 11-19",
            acb.report_facts(out).get("drop_exported") == "2026-11-22",
            acb.report_facts(out))

    site_rows = [dict(r) for r in got]
    for r in site_rows:
        r["INX_EXTRACT_DATE"] = "10/22/2026"
    for col in _DAY0_COLS:              # the site prints day 0 with its 2-digit year
        site_rows[1][col] = "99-12-30"
    site = _site_export(tmp / "site_day0", site_rows)
    counts = _counts(cmp_mi.compare(str(out), str(site), str(tmp / "cz.xlsx"), **_q()))
    c.check("the site's '99-12-30' is the same day 0: a clean match",
            counts == {"paired_rows": 2, "side_a_only_rows": 0, "side_b_only_rows": 0,
                       "differing_cells": 0}, counts)


def test_sources():
    print("the vendor's list against the build:")
    found = crs.sources("clean_intersection_mi")
    covered = crs.by_column("clean_intersection_mi")
    c.check("29 found fields fill 23 file columns, each from the field the list names",
            len(found) == 29 and len(covered) == 23
            and crs.mismatches("clean_intersection_mi", mi.PROVENANCE) == [])
    c.check("the six with no column of their own: the legs' Pri and Sec lanes, and the "
            "minor leg's highway group and population",
            sorted(s.field for s in crs.unplaced("clean_intersection_mi"))
            == sorted(["Major_Pri_Num_Lanes", "Major_Sec_Num_Lanes",
                       "Minor_Pri_Num_Lanes", "Minor_Sec_Num_Lanes",
                       "Minor_Highway_Group", "Minor_Population_Code"]))
    wrong = dict(mi.PROVENANCE, INX_MAIN_LANES_AMT=(
        "Master Intersection.Major_Pri_Num_Lanes", "integer"))
    bad = crs.mismatches("clean_intersection_mi", wrong)
    c.check("a build that declares another field is caught",
            len(bad) == 1 and bad[0].startswith("INX_MAIN_LANES_AMT"), bad)
    c.check("every column the build fills is sourced or unsourced, and what it does not "
            "take from the layer is a real column",
            set(mi.PROVENANCE) | set(mi.SPEC.unsourced) == set(mi.SPEC.header)
            and set(mi.EXTRA_CONTEXT) <= set(mi.SPEC.header))
    note = crs.source_note(arcgis_refresh.CSMS_SERVICE,
                           f"clean_road_intersection_consolidated {DAY}.xlsx")
    c.check("the data-source line names the layer's service and what the site reads",
            "Master Intersection layer was pulled from" in note
            and "IM tables" in note and "⚠" not in note, note)


def main():
    test_registry()
    with temp_dir("mi_check_") as tmp:
        build, got = test_build(Path(tmp))
        test_compare(Path(tmp), build, got)
        test_day0_and_export_date(Path(tmp))
    test_sources()
    return c.summary()


if __name__ == "__main__":
    sys.exit(main())
