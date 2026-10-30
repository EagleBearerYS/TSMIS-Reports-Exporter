"""build/check_three_way.py — the three-way lane (v0.51.0): the site's export vs
the TSN extract vs our ArcGIS build, cell by cell.

Drives the SHIPPED path on synthetic data (no corpus files):

  * the classes partition every compared cell — report differs (TSN = ArcGIS ≠
    TSMIS), TSN differs, ArcGIS differs, all differ, agree — and an unplaceable
    ArcGIS anchor is undecided;
  * Clean Road Ramp end to end: the real layer build, the real consolidation of a
    per-route site file, a raw TSN extract, `compare_clean_road_three_way` — the
    published counts, every sheet, the finding list naming the ArcGIS layer and
    field and linking to its grid row, the grid's cell texts, a record TSN and the
    layers have that the report leaves out, the role gates;
  * the published counts are bound to the workbook (a rewrite reads as unknown);
  * every column of all three Clean Road reports traces to library layers (or
    says it has none), and the corrected Highway provenance;
  * the pairing equals the two two-sided comparisons' own;
  * the matrix: rows, the missing-side states, build_cell through the real
    consolidation and comparison, a fresh cell, and staleness when the export,
    the TSN dataset or the ArcGIS build moves.

Run with the build venv:
    build\\.venv\\Scripts\\python.exe build\\check_three_way.py
"""
import datetime
import os
import sys
import time
from pathlib import Path

from _checklib import Checker, patch, scripts_path, temp_dir

scripts_path()

from openpyxl import Workbook, load_workbook  # noqa: E402

import arcgis_layers  # noqa: E402
import arcgis_report_clean_ramp as acrr  # noqa: E402
import arcgis_report_intersection_detail as ari  # noqa: E402
import clean_highway_columns as chc  # noqa: E402
import clean_road_columns as crc  # noqa: E402
import clean_road_layers as crl  # noqa: E402
import compare_clean_ramp_arcgis as cmp_ram  # noqa: E402
import compare_clean_road_three_way as t3  # noqa: E402
import compare_clean_road_tsn as crt  # noqa: E402
import consolidate_clean_road_ramp as con_ram  # noqa: E402
import consolidation_meta  # noqa: E402
import matrix  # noqa: E402
import outcome  # noqa: E402
import paths  # noqa: E402
import three_way  # noqa: E402
import three_way_matrix as twm  # noqa: E402
from consolidate_clean_highway import UNAVAILABLE_TOKEN  # noqa: E402
from events import ConsolidateResult, Events  # noqa: E402

c = Checker()
D = datetime.datetime
ASOF = "2026-10-05"
DAY = "2026-10-22"
SRC = "ssor-prod dev-site"
RUN = f"{DAY} {SRC}"


def _sheet(path, title, header, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.append(list(header))
    for r in rows:
        ws.append([r.get(col) for col in header])
    wb.save(path)


def _span(begin, end, attr, value, ar=None):
    frm, to = ar or (begin, end)
    return {"District": "District 12", "RouteNum": "001", "RouteSuffix": ".",
            "Alignment": "Right", "BeginCounty": "Orange", "EndCounty": "Orange",
            "BeginPMPrefix": ".", "EndPMPrefix": ".", "BeginPMMeasure": begin,
            "EndPMMeasure": end, "LRSFromDate": D(2020, 1, 1), "LRSToDate": None,
            "RouteID": "SHS_001._P", "FromARMeasure": frm, "ToARMeasure": to,
            "InventoryItemStartDate": D(1964, 1, 1), "InventoryItemEndDate": None,
            "Network_Start_Date": D(1964, 1, 1), "Network_End_Date": None,
            attr: value}


_AR = ["FromARMeasure", "ToARMeasure", "InventoryItemStartDate", "InventoryItemEndDate"]


def _pt(name, pm, od):
    return {"District": "District 12", "RouteNum": "001", "RouteSuffix": ".",
            "Alignment": "Right", "County": "Orange", "PMPrefix": ".", "PMMeasure": pm,
            "PMSuffix": ".", "ODMeasure": od, "Ramp_Name": name,
            "InventoryItemStartDate": D(1976, 2, 25), "InventoryItemEndDate": None,
            "RouteID": "SHS_001._P", "ARMeasure": od, "LRSFromDate": D(2020, 1, 1),
            "LRSToDate": None}


def _attrs(name, desc, on_off):
    return {"Ramp_Name": name, "Ramp_Description": desc,
            "Ramp_Design": "D- Diamond Type Ramp", "Area4_Ind": "Yes",
            "Ramp_On_Off_Ind": on_off, "InventoryItemStartDate": D(1976, 2, 25),
            "InventoryItemEndDate": None, "LRSFromDate": D(2020, 1, 1), "LRSToDate": None}


def _ramp_build(tmp):
    """Three live ramps on route 001 built from a synthetic layer library."""
    lib = tmp / "layers"
    lib.mkdir(parents=True, exist_ok=True)
    hg, city, pop = (ari.OVERLAY_LAYERS[k] for k in ("HG", "CITY", "POP"))
    layers = [
        (acrr.PT_LAYER, acrr._PT_COLS, [_pt("R1", 0.606, 1.6), _pt("R2", 0.9, 1.9),
                                         _pt("R3", 1.2, 2.2)]),
        (acrr.RAMP_LAYER, acrr._RAMP_COLS,
         [_attrs("R1", "001/NB OFF TO A ST", "0- OFF"),
          _attrs("R2", "001/NB ON FR B ST", "1- ON"),
          _attrs("R3", "001/NB OFF TO C ST", "0- OFF")]),
        (acrr.DIR_LAYER, acrr._DIR_COLS, [{"ROUTE": 1, "FROM_": "S", "TO_": "N"}]),
        (hg[0], ari._SPAN_ID + _AR + [hg[1]], [_span(0, 50, hg[1], "D- Divided")]),
        (city[0], ari._SPAN_ID + _AR + [city[1]], [_span(0, 50, city[1], "")]),
        (pop[0], ari._SPAN_ID + _AR + [pop[1]], [_span(0, 50, pop[1], "U- Urban")]),
        ("SHS Inv Network Date", ari._SPAN_ID + ["FromARMeasure", "ToARMeasure",
                                                 "Network_Start_Date", "Network_End_Date",
                                                 "SegOrderId"],
         [_span(0, 9, "SegOrderId", 150)]),
    ]
    made = []
    for n, (layer, cols, rows) in enumerate(layers, start=1):
        name = f"{n:02d}_{layer}.xlsx"
        _sheet(lib / name, layer, cols, rows)
        made.append((name, layer, len(rows), len(cols)))
    _sheet(lib / crl.INDEX_NAME, "INDEX", list(crl.INDEX_HEADER),
           [dict(zip(crl.INDEX_HEADER, (f, layer, n, k, str(lib),
                                        f"https://example.test/{layer}")))
            for f, layer, n, k in made])
    out = tmp / "builds" / acrr.FILENAME
    res = acrr.consolidate(events=Events(), confirm_overwrite=lambda p: True,
                           out_path=out, asof=ASOF, lib_root=lib)
    assert res.status == "ok", res.message
    wb = load_workbook(out, read_only=True)
    it = wb[crc.RAMP.sheet].iter_rows(values_only=True)
    header = [str(x) for x in next(it)]
    rows = [dict(zip(header, r)) for r in it]
    wb.close()
    return out, rows, lib


def _site(folder, rows):
    """The site's per-route file, consolidated by the app — what a cell reads."""
    run = folder / "clean_ramp"
    run.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = crc.RAMP.sheet
    ws.append(list(crc.RAMP.header))
    for r in rows:
        ws.append([r.get(h) or None for h in crc.RAMP.header])
    wb.save(run / f"{RUN} clean_ramp_route_001.xlsx")
    out = folder / "consolidated.xlsx"
    res = con_ram.consolidate(events=Events(), confirm_overwrite=lambda p: True,
                              input_dir=run, out_path=out)
    assert res.status == "ok", res.message
    return out, run


def _tsn(path, rows):
    """A raw TSN CA RAMPS extract (`Sheet 1`, the exact header)."""
    _sheet(path, crc.TSN_RAW_SHEET, crc.RAMP_TSN_HEADER, rows)
    return path


def _fixture(tmp):
    """(build, site export, TSN extract): R1 and R2 on all three, R3 left out of
    the report, one ramp only the report has, one only TSN has."""
    build, rows, lib = _ramp_build(tmp)
    by = {r["RAM_PM_LOC_AMT"]: r for r in rows}
    r1, r2, r3 = by["0.606"], by["0.9"], by["1.2"]
    site = [dict(r1, RAM_DESCRIPTION="001/NB OFF TO A STREET"),          # report differs
            dict(r2, RAM_AREA_4_IND="N", RAM_ON_OFF_CODE="OFF"),          # ArcGIS / all
            dict(r2, RAM_PM_LOC_AMT="5.5", RAM_DESCRIPTION="ONLY IN TSMIS")]
    for r in site:
        r["RAM_BEGIN_OFFSET_AMT"] = ""
        r["RAM_EXTRACT_DATE"] = "10/22/2026"
    export, run = _site(tmp / "site", site)
    tsn_rows = [dict(r1, RAM_BEGIN_DATE="2015-01-01"),                     # TSN differs
                dict(r2, RAM_AREA_4_IND="N", RAM_ON_OFF_CODE="OTH"),
                dict(r3),
                dict(r3, RAM_PM_LOC_AMT="7.7", RAM_DESCRIPTION="ONLY IN TSN")]
    for r in tsn_rows:
        r["RAM_EXTRACT_DATE"] = "2025-09-08"
        r["RAM_DESIGN_CODE"] = "RD"                                      # TSN's domain letter
    tsn = _tsn(tmp / "CA RAMPS.xlsx", tsn_rows)
    return build, export, tsn, run, lib


# --------------------------------------------------------------------------- #
def test_classes():
    print("the classes partition every compared cell:")
    cls = three_way.classify
    c.check("all three agree", cls("2", "2", "2") == three_way.MATCH)
    c.check("TSN = ArcGIS ≠ TSMIS is 'report differs'", cls("3", "2", "2") == three_way.REPORT)
    c.check("TSMIS = ArcGIS ≠ TSN is 'TSN differs' (expected)", cls("3", "2", "3") == three_way.TSN)
    c.check("TSMIS = TSN ≠ ArcGIS is 'ArcGIS differs'", cls("3", "3", "1") == three_way.ARCGIS)
    c.check("three different values", cls("3", "2", "1") == three_way.ALL)
    c.check("a blank is a value like any other", cls("", "2", "2") == three_way.REPORT)
    c.check("an unplaceable ArcGIS anchor is undecided",
            cls("3", "3", UNAVAILABLE_TOKEN, UNAVAILABLE_TOKEN) == three_way.UNDECIDED)
    c.check("the class list leads with the one to read first",
            three_way.CLASSES[0] == three_way.REPORT)


def test_ramp_end_to_end(tmp):
    print("Clean Road Ramp three ways, end to end on the shipped path:")
    build, export, tsn, _run, lib = _fixture(tmp)
    out = tmp / "ramp_three_way.xlsx"
    res = t3.RAMP.compare(export, tsn, build, out, events=Events())
    c.check("the comparison runs", res.status == "ok", res.message)
    got = three_way.read_counts(out)
    c.check("the counts are published beside the workbook", got is not None)
    if got is None:
        return
    c.check("records: two on all three, one TSN + ArcGIS (missing from the report), "
            "one only in the report, one only in TSN",
            {k: v for k, v in got.records.items() if v} ==
            {"all": 2, "tsn_arcgis": 1, "tsmis": 1, "tsn": 1}, got.records)
    c.check("cells: one report differs, one TSN differs, one ArcGIS differs, one all "
            "differ, the rest agree",
            (got.cells["report"], got.cells["tsn"], got.cells["arcgis"],
             got.cells["all"], got.cells["undecided"]) == (1, 1, 1, 1, 0)
            and got.cells["match"] == 2 * got.counted_columns - 4, got.cells)
    c.check("18 Ramp columns are compared three ways (RAM_DESIGN_DESC counts only "
            "against ArcGIS, so it is shown here)",
            got.counted_columns == 18
            and "RAM_DESIGN_DESC" not in t3.RAMP.counted(False), got.counted_columns)
    c.check("per column: the report-differs cell is RAM_DESCRIPTION",
            got.columns["RAM_DESCRIPTION"]["report"] == 1
            and got.columns["RAM_BEGIN_DATE"]["tsn"] == 1
            and got.columns["RAM_AREA_4_IND"]["arcgis"] == 1
            and got.columns["RAM_ON_OFF_CODE"]["all"] == 1)

    wb = load_workbook(out)
    c.check("sheets in reading order",
            wb.sheetnames == ["Summary", "Report differs", "Three-way", "ArcGIS differs",
                              "All differ", "Layer trace", "Notes"], wb.sheetnames)
    rd = list(wb["Report differs"].iter_rows(min_row=3, values_only=True))
    c.check("the report-differs list has the one cell", len(rd) == 1, rd)
    row = rd[0] if rd else (None,) * 11
    c.check("...naming the column, the three values, the ArcGIS layer and field",
            row[2] == "RAM_DESCRIPTION" and row[3] == "001/NB OFF TO A STREET"
            and row[4] == row[5] == "001/NB OFF TO A ST"
            and row[6] == "SHS Ramp" and row[7] == "Ramp_Description", row)
    c.check("...where to look it up", "route 001" in str(row[9]) and "0.606" in str(row[9]),
            row[9])
    link = str(row[10] or "")
    c.check("...and a link to its grid row", link.startswith("=HYPERLINK(\"#'Three-way'!"), link)
    kind = wb["Report differs"].cell(row=3, column=11).data_type
    c.check("...a live link, not its formula shown as text", kind == "f", kind)
    grid = wb["Three-way"]
    target = int(link.split("!")[1].split(":")[0]) if "!" in link else 0
    head = [x.value for x in grid[2]]
    gr = {h: x for h, x in zip(head, grid[target])} if target else {}
    c.check("the linked grid row is that record, the cell leading with the TSMIS value",
            str(gr.get("Location").value if gr else "").endswith("0.606")
            and gr["RAM_DESCRIPTION"].value
            == "001/NB OFF TO A STREET ≠ 001/NB OFF TO A ST (TSN & ArcGIS)",
            gr and gr["RAM_DESCRIPTION"].value)
    c.check("...filled the report-differs red",
            gr and gr["RAM_DESCRIPTION"].fill.start_color.rgb.endswith("FFC7CE"))
    rows = {str(r[1]): r for r in grid.iter_rows(min_row=3, values_only=True)}
    r2 = next((r for k, r in rows.items() if k.endswith("/ 0.9")), None)
    ci = {h: i for i, h in enumerate(head)}
    c.check("the TSN-differs and ArcGIS-differs cells read 'TSMIS (other: …)'",
            r2 and r2[ci["RAM_AREA_4_IND"]] == "N (ArcGIS: Y)"
            and r2[ci["RAM_ON_OFF_CODE"]] == "OFF ≠ TSN OTH ≠ ArcGIS ON", r2)
    r3 = next((r for k, r in rows.items() if k.endswith("/ 1.2")), None)
    c.check("the ramp TSN and the layers have but the report doesn't is its own record",
            r3 and r3[ci["Presence"]] == "TSN + ArcGIS (not in TSMIS)", r3)
    trace = {r[0]: r for r in wb["Layer trace"].iter_rows(min_row=3, values_only=True)}
    c.check("the layer trace lists every column", set(trace) == set(crc.RAMP.header))
    c.check("...a sourced column with its layer, field, library file and service",
            trace["RAM_DESCRIPTION"][2] == "SHS Ramp"
            and trace["RAM_DESCRIPTION"][3] == "Ramp_Description"
            and trace["RAM_DESCRIPTION"][5].endswith("_SHS Ramp.xlsx")
            and trace["RAM_DESCRIPTION"][6] == "https://example.test/SHS Ramp",
            trace["RAM_DESCRIPTION"])
    c.check("...and says which columns are shown only",
            trace["RAM_ADT"][1] == "shown only" and trace["RAM_DESCRIPTION"][1] == "three ways")
    summary = " ".join(str(v) for r in wb["Summary"].iter_rows(values_only=True)
                       for v in r if v is not None)
    c.check("the Summary states the vintage gap between the build and the export",
            "VINTAGE" in summary and ASOF in summary, summary[:300])
    wb.close()

    print("role gates:")
    bad = t3.RAMP.compare(build, tsn, build, tmp / "x1.xlsx", events=Events())
    c.check("our ArcGIS build cannot stand as the site's export", bad.status == "error",
            bad.message)
    bad = t3.RAMP.compare(export, export, build, tmp / "x2.xlsx", events=Events())
    c.check("a site export cannot stand as TSN", bad.status == "error", bad.message)
    bad = t3.RAMP.compare(export, tsn, export, tmp / "x3.xlsx", events=Events())
    c.check("a site export cannot stand as the ArcGIS build", bad.status == "error",
            bad.message)

    print("a partial input is named as the reason:")
    part = tmp / "export_partial.xlsx"
    part.write_bytes(export.read_bytes())
    consolidation_meta.write_outcome(part, ConsolidateResult(
        status="ok", completion=outcome.PARTIAL, skipped_inputs=2))
    out2 = tmp / "ramp_three_way_partial.xlsx"
    res = t3.RAMP.compare(part, tsn, build, out2, events=Events())
    got2 = three_way.read_counts(out2)
    c.check("an export consolidated with skipped routes makes the comparison partial",
            res.status == "ok" and got2 is not None
            and got2.completion == outcome.PARTIAL, (res.status, got2))
    wb2 = load_workbook(out2, read_only=True)
    said = " ".join(str(v) for r in wb2["Summary"].iter_rows(values_only=True)
                    for v in r if v is not None)
    wb2.close()
    c.check("...and the Summary names that input, not a capped pairing it didn't have",
            "export_partial.xlsx" in said and "capped" not in said, said[:400])

    print("the counts are bound to the workbook:")
    time.sleep(0.05)
    os.utime(out, (time.time() + 5, time.time() + 5))
    c.check("a workbook rewritten after publication reads as unknown",
            three_way.read_counts(out) is None)

    print("pairing equals the two two-sided comparisons':")
    rows_m, rows_n, rows_a, _pdf = t3.RAMP.load(export, tsn, build)
    records, _states, counts, _trim = three_way.compute(
        t3.RAMP.schema(False), rows_m, rows_n, rows_a)
    vs_tsn = crt.RAMP_VS_TSN.compare(export, tsn, tmp / "p1.xlsx", events=Events(),
                                     mode="preview").comparison_outcome.counts
    vs_arc = cmp_ram.compare(build, export, tmp / "p2.xlsx", events=Events(),
                             mode="preview").comparison_outcome.counts
    c.check("TSMIS rows paired with TSN = the TSMIS-vs-TSN comparison's paired rows",
            counts.records["all"] + counts.records["tsmis_tsn"] == vs_tsn.paired_rows,
            (counts.records, vs_tsn.paired_rows))
    c.check("TSMIS rows paired with ArcGIS = the Reports-vs-ArcGIS comparison's",
            counts.records["all"] + counts.records["tsmis_arcgis"] == vs_arc.paired_rows,
            (counts.records, vs_arc.paired_rows))
    mn = counts.cells["report"] + counts.cells["tsn"] + counts.cells["all"]
    c.check("TSMIS-vs-TSN differing cells agree with that comparison's",
            mn + counts.two_side["tsmis_tsn"] == vs_tsn.differing_cells,
            (mn, vs_tsn.differing_cells))
    ma = counts.cells["report"] + counts.cells["arcgis"] + counts.cells["all"]
    extra = sum(n for key, n in vs_arc.per_field_counts.items()
                if key.split(":", 1)[1] not in t3.RAMP.counted(False))
    c.check("TSMIS-vs-ArcGIS differing cells agree with that comparison's (less the "
            "columns TSN can't compare)",
            ma + counts.two_side["tsmis_arcgis"] == vs_arc.differing_cells - extra,
            (ma, vs_arc.differing_cells, extra))


def test_traces():
    print("every column traces to the ArcGIS layers:")
    for report in t3.REPORTS.values():
        spec = report.spec
        missing = [col for col in spec.header if col not in report.trace]
        c.check(f"{spec.label}: a trace for every column", not missing, missing)
        unknown = sorted({layer for tr in report.trace.values() for layer in tr.layers
                          if layer not in crl.EXPECTED_LAYERS})
        c.check(f"{spec.label}: every traced layer is a library layer", not unknown, unknown)
        silent = [col for col in report.counted(False)
                  if not report.trace[col].layers and not report.trace[col].rule]
        c.check(f"{spec.label}: a compared column with no layer says why", not silent, silent)
    hwy = t3.HIGHWAY.trace
    c.check("Highway's district is the covering span's own District field (the build "
            "never reads SHS District)",
            hwy["THY_DISTRICT_CODE"].layers == ("SHS Highway Group",)
            and hwy["THY_DISTRICT_CODE"].fields == "District"
            and chc.PROVENANCE["THY_DISTRICT_CODE"][1] == "SHS Highway Group")
    c.check("a block effective date names its primary layer first, and the build's "
            "own provenance says the primary-layer rule",
            hwy["THY_LEFT_ROAD_EFF_DATE"].layers[0] == "SHS Travel Way L"
            and hwy["THY_MEDIAN_EFF_DATE"].layers[0] == "SHS Median"
            and "SHS Median's own" in chc.PROVENANCE["THY_MEDIAN_EFF_DATE"][3])
    inx = t3.INTERSECTION.trace
    c.check("an intersection leg column names the approach layers and the leg rule",
            inx["INX_MAIN_LANES_AMT"].layers[0] == "IM Intersection Approach Detail"
            and "IM Intersection Approach Segments" in inx["INX_MAIN_LANES_AMT"].layers
            and "ROUTE_ID" in inx["INX_MAIN_LANES_AMT"].rule)
    c.check("the provenance parser keeps the longest layer name (SHS Ramp Pt, not SHS Ramp)",
            t3.trace_from_provenance("SHS Ramp Pt.ODMeasure", "x").layers == ("SHS Ramp Pt",))


def test_matrix(tmp):
    print("the Three-way matrix:")
    saved_root = paths.OUTPUT_ROOT
    try:
        paths.OUTPUT_ROOT = tmp / "output"
        build, export, tsn, run, lib = _fixture(tmp / "fx")
        day = paths.OUTPUT_ROOT / RUN / "clean_ramp"
        day.mkdir(parents=True)
        for f in run.iterdir():
            (day / f.name).write_bytes(f.read_bytes())
        token = {"v": "tok-1"}
        tsn_src = lambda *_a, **_k: {"kind": "consolidated", "path": str(tsn),  # noqa: E731
                                     "mtime": 1.0, "identity_token": token["v"]}
        with patch(acrr, "OUT_PATH", build), patch(matrix, "tsn_source", tsn_src), \
                patch(twm, "_resolve_tsn",
                      lambda *_a, **_k: (tsn, token["v"], lambda: True)):
            c.check("rows: the six Clean Road editions, Excel then PDF per report",
                    [r.key for r in twm._rows()] ==
                    ["clean_highway", "clean_highway_pdf", "clean_intersection",
                     "clean_intersection_pdf", "clean_ramp", "clean_ramp_pdf"])
            snap = twm.tw_matrix_snapshot(SRC, [DAY], today="2099-01-01")
            cell = snap["cells"]["clean_ramp"][DAY]["cmp"]
            c.check("a cell with its three sides present is buildable, not built",
                    matrix.cell_buildable(cell) and not cell["built"], cell)
            c.check("the PDF edition with no print exported needs its export",
                    snap["cells"]["clean_ramp_pdf"][DAY]["cmp"]["missing_side"] == "export")
            hwy = snap["cells"]["clean_highway"][DAY]["cmp"]
            c.check("a report with no export that day names every missing side",
                    "export" in hwy["missing_sides"], hwy)
            c.check("the add-day picker offers the exported day",
                    DAY in twm.available_days(SRC))
            hid = twm.tw_matrix_snapshot(SRC, [DAY], hidden=["clean_ramp", "clean_ramp_pdf"],
                                         today="2099-01-01")
            c.check("hiding both Ramp rows keeps Ramp's TSN state for the three-sides card",
                    "clean_ramp" not in hid["rows"]
                    and hid["tsn_meta"].get("clean_ramp", {}).get("ready") is True,
                    hid["tsn_meta"].get("clean_ramp"))

            res = twm.build_cell(SRC, DAY, "clean_ramp", None, Events())
            c.check("build_cell consolidates the day and compares the three", res.status == "ok",
                    res.message)
            out = twm.day_out_path(DAY, SRC, "clean_ramp")
            c.check("...into the three-way store under a self-identifying name",
                    out.exists() and out.name == f"clean_ramp_three_way {RUN}.xlsx"
                    and out.parent.parent.name == "three-way-by-day")
            snap = twm.tw_matrix_snapshot(SRC, [DAY], today="2099-01-01")
            cell = snap["cells"]["clean_ramp"][DAY]["cmp"]
            c.check("the cell reads fresh with the published counts",
                    cell["built"] and not cell["stale"]
                    and cell["three_way"]["report"] == 1
                    and cell["three_way"]["missing_from_report"] == 1, cell)
            token["v"] = "tok-2"
            cell = twm.tw_matrix_snapshot(SRC, [DAY], today="2099-01-01")[
                "cells"]["clean_ramp"][DAY]["cmp"]
            c.check("a moved TSN dataset reads the cell stale",
                    cell["stale"] and cell["reason"] == "source_identity_changed", cell)
            token["v"] = "tok-1"
            refreshed = acrr.consolidate(events=Events(), confirm_overwrite=lambda p: True,
                                         out_path=build, asof="2026-10-06", lib_root=lib)
            cell = twm.tw_matrix_snapshot(SRC, [DAY], today="2099-01-01")[
                "cells"]["clean_ramp"][DAY]["cmp"]
            c.check("a refreshed ArcGIS build reads the cell stale",
                    refreshed.status == "ok" and cell["stale"]
                    and cell["reason"] in ("source_identity_changed", "layers_newer"), cell)
            res = twm.build_cell(SRC, DAY, "clean_ramp", None, Events())
            cell = twm.tw_matrix_snapshot(SRC, [DAY], today="2099-01-01")[
                "cells"]["clean_ramp"][DAY]["cmp"]
            c.check("rebuilding it reads fresh again", res.status == "ok" and not cell["stale"],
                    cell)
            added = day / f"{RUN} clean_ramp_route_002.xlsx"
            added.write_bytes((day / f"{RUN} clean_ramp_route_001.xlsx").read_bytes())
            os.utime(added, (time.time() + 5, time.time() + 5))
            cell = twm.tw_matrix_snapshot(SRC, [DAY], today="2099-01-01")[
                "cells"]["clean_ramp"][DAY]["cmp"]
            c.check("a new route in the export folder reads the cell stale",
                    cell["stale"] and cell["reason"] in ("export_newer", "inputs_changed"),
                    cell)
            c.check("the scoped rebuild list picks the stale cell",
                    twm.cells_to_rebuild(twm.tw_matrix_snapshot(
                        SRC, [DAY], today="2099-01-01")) == [(DAY, "clean_ramp")])
            try:
                twm.build_cell(SRC, DAY, "clean_ramp_pdf", None, Events())
                refused = False
            except ValueError:
                refused = True
            c.check("build_cell refuses an edition not exported that day", refused)
    finally:
        paths.OUTPUT_ROOT = saved_root


def main():
    test_classes()
    test_traces()
    with patch(arcgis_layers, "drop_info",
               lambda lib_root=None: {"fingerprint": "v2:test", "exported": ASOF}):
        with temp_dir("tsmis_3way_") as tmp:
            test_ramp_end_to_end(Path(tmp))
        with temp_dir("tsmis_3way_mx_") as tmp:
            test_matrix(Path(tmp))
    return c.summary()


if __name__ == "__main__":
    sys.exit(main())
