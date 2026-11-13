"""Shared comprehensive runtime self-test for the bundled app.

Exercises EVERY real code path the frozen app depends on -- launch the system /
bundled browser + page.pdf() + a download, pdfplumber text/word/table extraction
(exactly what the Ramp Summary consolidator uses), an openpyxl write/read
round-trip, the dynamically-imported matrix modules (the F6 trio), the GUI bridge
(build the js_api + initial state, confirm the bundled ui/ assets), and --
best-effort -- a hidden WebView2 window cycle through the real JS bridge.

ONE body, two callers, so the exact shipped exe and the dev venv prove the same
thing (DRY):
  * scripts/gui_main.py  `--self-test`  -- the EXACT windowed release exe runs
    this before any real window is created; this is the build.ps1 -SelfTest
    release gate, exercising the precise artifact that ships (not a different
    console build -- the prior gate built a separate `full_smoke` exe).
  * build/full_smoke.py                 -- the dev/venv tool (same body; the
    render-stack step proves PIL/pypdfium2 WORK — they ship since v0.21.0).

`run(emit)` returns 0 on success and RAISES on any MANDATORY failure (import /
asset / registry / browser / pdf / openpyxl / dynamic-module). Only the hidden
WebView WINDOW probe may skip -- an environment that can't show a window is
tolerated; the import/registry sub-checks never skip. `emit` is a line sink
(defaults to print); gui_main passes a windowed-safe sink because the release exe
has no console/stdout.

Console-free-core note: this is a DIAGNOSTIC DRIVER (only gui_main and the build
tool import it), never imported by an engine module, and it writes only through
the injected `emit` -- so the no-console-in-core convention is preserved. The
heavy third-party imports live INSIDE run(), so merely importing this module
(e.g. the packaging reachability check, or normal startup) stays cheap.
"""
import sys
import tempfile
import threading
import time
from pathlib import Path

# A small page in the Ramp Summary shape (a heading + a table) -- enough to drive
# page.pdf() and exercise the exact pdfplumber calls the consolidator makes.
HTML = """
<h1>Route 005 Ramp Summary</h1>
<table border=1>
  <tr><th>Ramp</th><th>Count</th></tr>
  <tr><td>NB On</td><td>1234</td></tr>
  <tr><td>SB Off</td><td>5678</td></tr>
</table>
"""

# A miniature Clean Road print (v0.48.0) in the site's own print structure — a
# cover with its ROUTE line, then the header row + data rows as table cells — so
# the frozen bundle proves the pypdfium2 TEXT + path calls clean_road_print makes
# (the evidence step below only rasterizes with pdfium).
CLEAN_ROAD_HTML = """
<style>html,body{margin:0;background:#fff;font-family:Arial,sans-serif}
.cover{page-break-after:always}
table.t{border-collapse:collapse;font-size:8px;white-space:nowrap}
table.t th{background:#fff;border-bottom:1px solid #ccc;padding:2px 5px}
table.t td{padding:2px 5px;border-bottom:1px solid #eee}
table.t tbody tr:nth-child(even){background:#eef0f2;-webkit-print-color-adjust:exact}</style>
<div class="cover"><div>LOCATION CRITERIA:</div><div>ROUTE 005S</div></div>
<table class="t"><thead><tr><th>NETWORK<br>ID</th><th>ON_OFF<br>CODE</th>
<th>DESCRIPTION</th></tr></thead><tbody>
<tr><td>2</td><td>ON</td><td>005/NB ON FR MAIN ST</td></tr>
<tr><td>2</td><td>OFF</td><td>005/SB OFF TO EL CAMINO REAL</td></tr>
</tbody></table>
"""

# Flat modules the app imports DYNAMICALLY (the matrix tab loads them lazily) that
# a frozen bundle MUST carry -- the F6 trio. Importing them here is the runtime
# half of build/check_app_modules.py's offline packaging contract: if the bundle
# ever drops one, the exact shipped exe fails its own gate with a precise error.
# v0.21.0 adds the lazily-imported visual-evidence modules to the contract
# (v0.22.0: the Intersection Detail adapter joins; v0.24.0: Highway Log;
# v0.25.0: Highway Sequence + its lazily-resolved PDF consolidator/comparisons;
# v0.26.0: the Ramp Detail trio; RB-4: the env-only Ramp Summary adapter the
# HF-10 camera resolves lazily).
_DYNAMIC_REPORT_MODULES = ("matrix", "day_matrix", "pdf_excel_matrix",
                           "report_library",
                           "visual_evidence", "evidence_highway_detail",
                           "evidence_intersection_detail",
                           "evidence_highway_log",
                           "evidence_highway_sequence",
                           "evidence_ramp_summary",
                           "consolidate_tsmis_highway_sequence_pdf",
                           "compare_highway_sequence_pdf",
                           "evidence_ramp_detail",
                           "consolidate_tsmis_ramp_detail_pdf",
                           "compare_ramp_detail_pdf",
                           # v0.48.0: the second editions' PDF consolidators the
                           # matrix resolves lazily, the pdfium-based Clean Road
                           # print reader behind three of them, and the new
                           # comparison modules.
                           "consolidate_tsmis_intersection_summary_pdf",
                           "consolidate_tsmis_highway_summary_pdf",
                           "consolidate_tsmis_clean_highway_pdf",
                           "consolidate_tsmis_clean_intersection_pdf",
                           "consolidate_tsmis_clean_ramp_pdf",
                           "clean_road_print",
                           "compare_clean_road_tsn", "compare_summary_editions",
                           "compare_env_editions",
                           # The ArcGIS reports + "Reports vs ArcGIS" lane: every
                           # module here is imported INSIDE a GUI endpoint, so
                           # nothing else in the frozen bundle proves it resolves.
                           # The registry pulls each report's build + comparator
                           # behind it, and those pull the clean-road substrate.
                           "arcgis_reports",
                           "arcgis_matrix",
                           "arcgis_report_highway_detail",
                           "compare_highway_detail_arcgis",
                           "arcgis_report_intersection_detail",
                           "compare_intersection_detail_arcgis",
                           # v0.46.0: the in-app layer refresh (the Layers tab).
                           "arcgis_pro", "arcgis_refresh", "arcgis_selftest",
                           "gui_arcgis_layers_api", "gui_worker_arcgis",
                           # v0.51.0: the Three-way matrix and its Clean Road
                           # adapters, imported inside its endpoints and worker.
                           "three_way", "three_way_matrix",
                           "compare_clean_road_three_way",
                           # v0.52.0: the Three Days matrix, its engine and
                           # the workbook writers the engine imports on use
                           # (the live-formulas copy since v0.52.1).
                           "three_day", "three_day_workbook",
                           "three_day_formulas", "three_day_matrix")

# What a hosted job needs importable in its worker process (step 5c): the two
# modules that define the hosted workers, and the builders their run() imports
# on use.
_WORKER_PROCESS_MODULES = ("gui_worker_matrix", "gui_worker_arcgis",
                           "tsn_library", "arcgis_reports", "arcgis_matrix",
                           "three_way_matrix", "three_day_matrix")

# What the hidden window is asked when its page never finishes booting: the
# files that did not load, the page's own failure banner, and whether the bridge
# arrived. Plain ES5 so it runs however little of the page did. (The browser
# always asks for a favicon and the app has none, so that one is left out.)
_PAGE_DIAGNOSIS_JS = """(function () {
  var out = {};
  try {
    out.not_loaded = performance.getEntriesByType('resource')
      .filter(function (e) {
        return e.responseStatus !== 200 && !/favicon\\.ico$/.test(e.name);
      })
      .map(function (e) { return e.name.split('/').pop() + ' (status ' + e.responseStatus + ')'; });
  } catch (e) { out.not_loaded = 'unavailable: ' + e; }
  var banner = document.getElementById('fatalBanner');
  out.banner = banner ? banner.innerText.slice(0, 200) : null;
  out.bridge = typeof (window.pywebview && window.pywebview.api
                       && window.pywebview.api.get_initial_state);
  return JSON.stringify(out);
})()"""


def _writer_probe(path, classic):
    """Write one sheet of the cell shapes a comparison holds through
    `compare_core._SheetRows` — with openpyxl's own row path forced when
    `classic` — and return ``(package members, direct path used)``.
    ``docProps/core.xml`` is left out: it holds only the creation time."""
    import os
    from zipfile import ZipFile

    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    import compare_core as core

    saved = os.environ.get("TSMIS_CLASSIC_XLSX")
    if classic:
        os.environ["TSMIS_CLASSIC_XLSX"] = "1"
    else:
        os.environ.pop("TSMIS_CLASSIC_XLSX", None)
    try:
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet("Probe")
        ws.column_dimensions["A"].width = 18
        body = Font(name="Arial", size=10)
        bold = Font(name="Arial", size=10, bold=True)
        tint = PatternFill("solid", start_color="FFE699")
        wrap = Alignment(wrap_text=True, vertical="top")
        rows = core._SheetRows(ws)
        ws.append(core._header_row(ws, ["Key", "Exact", "Tinted", "Count", "Link", "Bare"]))
        values = ["plain", " lead", "trail ", "a<b&c>\"d'", "new\nline", "tab\there",
                  "é ≠ \U0001f642", "", "=SUM(A1)", "+1", "-x", "@y", "#N/A",
                  12, 1.5, 0.1 + 0.2, 12345678901234567890, True, False, None]
        for i, value in enumerate(values):

            def fill(put, value=value, i=i):
                put(f"K{i:03d}", bold, guard=True)
                put(value, body, guard=True, exact_source_numeric=True)
                put(value, body, fill=tint, align=wrap, guard=True)
                put(i, body)
                put(f'=HYPERLINK("#Probe!{i + 2}:{i + 2}",{i + 2})', body)
                put(value if isinstance(value, str) else None, body)

            rows.row(fill)
        wb.save(str(path))
        active = rows.fast
    finally:
        if saved is None:
            os.environ.pop("TSMIS_CLASSIC_XLSX", None)
        else:
            os.environ["TSMIS_CLASSIC_XLSX"] = saved
    with ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist() if n != "docProps/core.xml"}, active


def run(emit=None):
    """Run the comprehensive self-test. Returns 0 on success; raises on any
    mandatory failure. `emit(line)` is the output sink (default print). The temp
    workspace is ALWAYS removed -- on success and on a mandatory-failure raise
    (so a failing CI run doesn't leave diagnostics behind)."""
    import shutil
    emit = emit or print
    tmp = Path(tempfile.mkdtemp())
    try:
        return _exercise(tmp, emit)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _exercise(tmp, emit):
    """The comprehensive exercise body. run() owns the temp dir + its cleanup."""
    emit("=" * 60)
    emit("TSMIS Exporter -- full bundle self-test")
    emit("=" * 60)

    import openpyxl
    import pdfplumber
    from playwright.sync_api import sync_playwright
    from common import launch_browser                  # drives system Edge/Chrome

    emit(f"frozen={getattr(sys, 'frozen', False)}  "
         f"openpyxl={openpyxl.__version__}  pdfplumber={pdfplumber.__version__}")

    # 1. Chromium: launch + render to PDF (Ramp Summary path) + a download.
    pdf_path = tmp / "page.pdf"
    with sync_playwright() as p:
        browser = launch_browser(p, headless=True)    # system Edge/Chrome (or bundled)
        ctx = browser.new_context(accept_downloads=True)
        page = ctx.new_page()
        page.set_content(HTML)
        page.pdf(path=str(pdf_path), format="Letter", print_background=True)
        page.set_content('<a id="d" href="data:text/plain,hello" download="x.txt">d</a>')
        with page.expect_download() as dl:
            page.click("#d")
        dl.value.save_as(str(tmp / "x.txt"))
        cr_pdf = tmp / "clean_road.pdf"
        page.set_content(CLEAN_ROAD_HTML)
        page.pdf(path=str(cr_pdf), format="Letter", landscape=True,
                 print_background=True,
                 margin={"top": "0.35in", "bottom": "0.35in",
                         "left": "0.35in", "right": "0.35in"})
        browser.close()
    assert pdf_path.stat().st_size > 0, "page.pdf produced nothing"
    emit(f"chromium: PDF {pdf_path.stat().st_size} bytes, download ok")

    # 2. pdfplumber: exactly the calls consolidate_ramp_summary makes.
    with pdfplumber.open(str(pdf_path)) as pdf:
        text = pdf.pages[0].extract_text() or ""
        words = pdf.pages[0].extract_words()
        tables = pdf.pages[0].extract_tables()
    assert "Route 005" in text, f"extract_text failed: {text!r}"
    assert any(w.get("text") == "1234" for w in words), "extract_words failed"
    emit(f"pdfplumber: text={len(text)} chars, words={len(words)}, tables={len(tables)}")

    # 3. openpyxl: write + read round-trip (consolidator output path).
    xlsx = tmp / "wb.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["Route", "Ramp", "Count"]); ws.append(["005", "NB On", 1234])
    wb.save(str(xlsx))
    wb2 = openpyxl.load_workbook(str(xlsx))
    assert wb2.active["C2"].value == 1234, "openpyxl round-trip failed"
    emit("openpyxl: write/read round-trip ok")

    # 3b. The comparison writer. A statewide comparison writes its rows straight
    #     into the sheet's stream instead of through one openpyxl cell per value
    #     (xlsx_fast_rows), and that is only allowed to be a SPEED change. Only
    #     this exact bundle can show that its openpyxl lets that path in and that
    #     the bytes it writes are openpyxl's own, so the same sheet is written
    #     both ways and every package member compared.
    direct, active = _writer_probe(tmp / "writer_direct.xlsx", classic=False)
    classic, _ = _writer_probe(tmp / "writer_classic.xlsx", classic=True)
    differing = sorted(n for n in set(direct) | set(classic)
                       if direct.get(n) != classic.get(n))
    assert not differing, \
        f"the comparison writer's direct rows differ from openpyxl's own in {differing}"
    emit("comparison writer: direct rows "
         + ("active" if active else "NOT active (openpyxl's own row path: slower, same output)")
         + f", {len(direct)} package members identical to openpyxl's")

    # 4. The visual-evidence render stack (SHIPS since v0.21.0 -- the excludes
    #    proof this step used to be is inverted): rasterize a PDF page
    #    (pdfplumber.to_image -> pypdfium2), draw the highlight box (PIL), and
    #    embed the PNG in a workbook (openpyxl's image insert needs PIL). These
    #    are exactly the calls visual_evidence makes, so a bundle missing any
    #    piece fails ITS OWN gate here, not the user's first evidence run.
    from PIL import ImageDraw
    from openpyxl.drawing.image import Image as XLImage
    with pdfplumber.open(str(pdf_path)) as pdf:
        img = pdf.pages[0].to_image(resolution=72).original.convert("RGB")
    ImageDraw.Draw(img).rectangle([10, 10, 60, 30], outline=(220, 20, 20), width=3)
    png = tmp / "evidence.png"
    img.save(str(png))
    assert png.stat().st_size > 0, "PDF page render produced nothing"
    wbe = openpyxl.Workbook()
    wbe.active.add_image(XLImage(str(png)), "A1")
    ev_xlsx = tmp / "evidence.xlsx"
    wbe.save(str(ev_xlsx))
    assert ev_xlsx.stat().st_size > 0, "evidence-workbook image embed failed"
    emit(f"evidence render stack: page->PNG {png.stat().st_size} bytes, "
         "highlight + workbook embed ok")
    emit(f"cryptography loaded (required by pdfminer): {'cryptography' in sys.modules}")

    # 4b. The Clean Road print reader (v0.48.0): read the miniature print back
    #     through pypdfium2's text + path APIs exactly as the consolidators do.
    import clean_road_columns as crc
    import clean_road_print as crp
    read = crp.read_print(cr_pdf, crc.RAMP)
    col = {k: i for i, k in enumerate(crc.RAMP.header)}
    got = [(r[col["RAM_NETWORK_ID"]], r[col["RAM_ON_OFF_CODE"]], r[col["RAM_DESCRIPTION"]])
           for r in read.rows]
    assert got == [("2", "ON", "005/NB ON FR MAIN ST"),
                   ("2", "OFF", "005/SB OFF TO EL CAMINO REAL")], \
        f"clean road print read back wrong: {got}"
    assert read.route_claim == "005S", f"clean road print route claim: {read.route_claim!r}"
    emit(f"clean road print reader: {len(read.rows)} rows, {len(read.printed)} columns, "
         f"route {read.route_claim} ok")

    # 5. The F6 trio: prove the frozen bundle carries the dynamically-imported
    #    matrix modules (precise failure here beats a confusing gui_api ImportError).
    import importlib
    for m in _DYNAMIC_REPORT_MODULES:
        importlib.import_module(m)
    emit(f"dynamic report modules import: {', '.join(_DYNAMIC_REPORT_MODULES)} ok")

    # 5b. The ArcGIS layer refresh (v0.46.0): the worker ships as a plain .py data
    #     file that ArcGIS Pro's own Python runs, so nothing above proves it is in
    #     the bundle. Run the SHIPPED script against the stand-in ArcPy into a temp
    #     library and read the result back with the builds' own reader.
    import arcgis_pro
    import arcgis_selftest
    assert arcgis_pro.worker_script().is_file(), \
        f"layer export worker missing: {arcgis_pro.worker_script()}"
    arcgis_selftest.run(tmp / "arcgis", emit)

    # 5c. Worker processes: comparison / build jobs run in a child process, and a
    #     frozen exe starts one by re-launching ITSELF — so only this exact
    #     artifact can prove it works. A child that starts and then fails is a
    #     packaging defect and fails the gate. An operating system that refuses
    #     to create the process at all is an environment restriction: the app
    #     keeps those jobs in-process there, so that is reported, not failed.
    #     The child also imports the modules a real job runs (the worker classes
    #     and the builders they import on use), so a job that could only ever
    #     fall back to the window's process fails here instead of going unseen.
    import job_host
    probed = job_host.probe(modules=_WORKER_PROCESS_MODULES)
    if probed["ok"]:
        emit(f"worker process: started (pid {probed['pid']}), imported "
             f"{len(_WORKER_PROCESS_MODULES)} job modules and answered in "
             f"{probed['seconds']:.1f}s")
    elif probed["refused"]:
        emit("worker process: skipped, this environment refuses child processes "
             f"({probed['reason']}) -- jobs will run in-process")
    else:
        raise AssertionError(
            "worker process probe failed: "
            f"{probed['detail'] or probed['reason'] or 'no answer from a child process'} "
            f"(answer pid {probed['pid']}, {probed['seconds']:.1f}s)")

    # 6. GUI bridge: js_api + initial state + bundled ui/ assets MUST work (these
    #    catch a prune/exclude that broke pywebview/pythonnet or lost the ui/ assets).
    import webview
    import gui_api

    class _NoWorker:                 # no background browser probe / GitHub update
        def __init__(self, *a, **k): pass     # the gate must be deterministic and
        def start(self): pass                 # offline-safe

    gui_api.CheckWorker = _NoWorker
    gui_api.UpdateWorker = _NoWorker

    api = gui_api.GuiApi()
    state = api.get_initial_state()
    assert state["reports"] and state["routes"], "GUI initial state incomplete"
    ui_index = gui_api._ui_index_path()
    assert ui_index.exists(), f"UI assets missing: {ui_index}"
    emit(f"gui: bridge api ok ({len(state['reports'])} reports, "
         f"{len(state['routes'])} routes, ui={ui_index})")

    # 6b. The ArcGIS reports tab + the Reports-vs-ArcGIS rows (v0.47.0): every
    #     report and every export edition resolve through the frozen registry
    #     (the editions come from the catalog), and the tab's payload reads.
    import arcgis_reports
    info = api.arcgis_reports_info()
    assert not info.get("error"), f"arcgis_reports_info: {info.get('error')}"
    assert [r["key"] for r in info["reports"]] == list(arcgis_reports.KEYS), \
        "the ArcGIS reports tab lost a report"
    editions = arcgis_reports.editions()
    comparable = [e.key for e in editions if e.comparable]
    assert comparable, "no export edition can be compared against an ArcGIS build"
    emit(f"arcgis reports: {len(info['reports'])} reports, {len(editions)} export "
         f"editions ({len(comparable)} comparable)")

    # 7. Hidden WebView window cycle -- the ONLY skippable sub-check. The
    #    import/asset/registry checks above already passed and never skip; an
    #    environment that simply can't show a window (a headless CI box) is
    #    tolerated, since that is a display capability, not a packaging defect.
    res = {}

    def _drive(w):
        try:
            deadline = time.time() + 30
            booted = False
            while time.time() < deadline:                # wait for app.js to boot
                if w.evaluate_js("typeof window.__tsmis !== 'undefined'"):
                    booted = True
                    break
                time.sleep(0.25)
            if not booted:
                # Name the cause: which files did not load, and what the page
                # itself reported. (A bare "test_state of undefined" took a
                # browser net log to explain.)
                res["err"] = ("the page never finished booting: "
                              f"{w.evaluate_js(_PAGE_DIAGNOSIS_JS)}")
                return
            res["state"] = w.evaluate_js("window.__tsmis.test_state()")
        except Exception as e:
            res["err"] = f"{type(e).__name__}: {e}"
        finally:
            w.destroy()

    window = webview.create_window("smoke", str(ui_index), js_api=api, hidden=True)
    window.events.loaded += lambda: threading.Thread(
        target=_drive, args=(window,), daemon=True).start()
    watchdog = threading.Timer(60, lambda: (res.setdefault("err", "watchdog timeout"),
                                            window.destroy()))
    watchdog.daemon = True
    watchdog.start()
    try:
        webview.start(gui="edgechromium")
    except Exception as e:
        emit(f"gui: window skipped, environment can't start WebView2 "
             f"({type(e).__name__}: {e})")
    else:
        watchdog.cancel()
        if not res.get("state"):
            raise AssertionError(f"gui window cycle failed: {res.get('err', 'no JS state')}")
        emit(f"gui: WebView window + JS bridge ok ({res['state']})")

    emit("")
    emit("SMOKE OK -- every app-required code path works.")
    return 0
