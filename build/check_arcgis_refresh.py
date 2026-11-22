"""Golden check for the in-app ArcGIS layer refresh (v0.46.0): the worker that
runs in ArcGIS Pro's Python (scripts/arcgis_worker/export_layers.py), the
session runner (arcgis_pro), the refresh/probe orchestration (arcgis_refresh),
the library manifest + drop identity (arcgis_layers), and the GUI wiring.

ArcPy exists only inside ArcGIS Pro, so every run here drives the SHIPPED worker
script against the stand-in ArcPy in arcgis_selftest — in-process for the
orchestration cases, and once as a REAL subprocess (this venv's python.exe with
a stand-in `arcpy.py` beside a copy of the worker) through the real launcher,
environment and event-file protocol. Offline; openpyxl only.

    build\\.venv\\Scripts\\python.exe build\\check_arcgis_refresh.py
"""
import ast
import contextlib
import functools
import io
import json
import logging
import os
import queue
import shutil
import sys
import tempfile
import threading
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT / "scripts"), str(ROOT)]

import arcgis_layers  # noqa: E402
import arcgis_pro  # noqa: E402
import arcgis_refresh as ar  # noqa: E402
import arcgis_selftest as ast_  # noqa: E402
import clean_road_layers as crl  # noqa: E402
import paths  # noqa: E402
from events import Events  # noqa: E402

_fail = []


def check(name, cond, detail=""):
    print(f"  [{'OK ' if cond else 'FAIL'}] {name}" + (f" — {detail}" if detail and not cond else ""))
    if not cond:
        _fail.append(name)


@contextlib.contextmanager
def _patch(obj, name, value):
    old = getattr(obj, name)
    setattr(obj, name, value)
    try:
        yield
    finally:
        setattr(obj, name, old)


@contextlib.contextmanager
def _library():
    """A temp library seeded with a manual drop; ARCGIS_LAYERS_ROOT points at it."""
    tmp = Path(tempfile.mkdtemp(prefix="agrefresh_"))
    lib = tmp / "arcgis_layers"
    ast_.seed_manual_library(lib)
    with _patch(paths, "ARCGIS_LAYERS_ROOT", lib):
        try:
            yield lib
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


def _refresh(spec, layers, **kw):
    fake = ast_.make_fake_arcpy(spec)
    logs = []
    res = ar.refresh(layers, events=Events(on_log=logs.append), python="fake-python.exe",
                     launcher=ast_.inprocess_launcher(fake), **kw)
    return res, {r["name"]: r for r in res.layers}, logs


# --------------------------------------------------------------------------- #
def test_contract():
    print("the library contract — 44 layers, the worker script, its Python floor:")
    names = [n for n, _id in ar.LIBRARY_LAYERS]
    check("LIBRARY_LAYERS is exactly clean_road_layers.EXPECTED_LAYERS",
          sorted(names) == sorted(crl.EXPECTED_LAYERS) and len(names) == 44)
    check("the layers added 2026-11-18 come after the first 40, so those keep their "
          "file numbers",
          names[40:43] == ["Calibration Point", "HPMS F System", "Posted Speed Limit"]
          and ar.LAYER_ORDER["Route Direction"] == 40
          and [ar.LAYER_IDS[n] for n in names[40:43]] == [1, 91, 107])
    check("Master Intersection (2026-11-22) comes last, layer 204 of its own service",
          names[43] == "Master Intersection" and ar.LAYER_ORDER["Master Intersection"] == 44
          and ar.LAYER_IDS["Master Intersection"] == 204
          and ar.LAYER_SERVICES == {"Master Intersection": ar.CSMS_SERVICE})
    check("the second service is on the same host as the first (one ArcGIS Pro sign-in)",
          ar.CSMS_SERVICE.split("/")[2] == ar.SERVICE_DEFAULT.split("/")[2]
          and ar.CSMS_SERVICE.endswith("/csms_supporting_prod/FeatureServer"))
    check("every partly exported layer is a library layer",
          set(ar.LAYER_FILTERS) <= set(names) and all(ar.LAYER_FILTERS.values()))
    check("every layer read from another service is a library layer",
          set(ar.LAYER_SERVICES) <= set(names))
    ids = [(ar.LAYER_SERVICES.get(n, ar.SERVICE_DEFAULT), i) for n, i in ar.LIBRARY_LAYERS]
    check("every fallback layer id is unique on its service", len(set(ids)) == len(ids))
    check("the probe layer is a library layer", ar.PROBE_LAYER in ar.LAYER_IDS)
    script = arcgis_pro.worker_script()
    check("the worker script ships beside the modules", script.is_file(), str(script))
    src = script.read_text(encoding="utf-8")
    try:
        tree = ast.parse(src, feature_version=(3, 9))
        parsed = True
    except SyntaxError as e:
        tree, parsed = None, False
        print(f"     {e}")
    check("the worker parses as Python 3.9 (ArcGIS Pro 3.0's floor)", parsed)
    allowed = {"argparse", "json", "os", "re", "sys", "threading", "time", "traceback",
               "urllib", "urllib.parse", "urllib.request", "arcpy"}
    used = set()
    for node in ast.walk(tree) if tree else ():
        if isinstance(node, ast.Import):
            used |= {a.name for a in node.names}
        elif isinstance(node, ast.ImportFrom) and node.module:
            used.add(node.module)
    check("the worker imports only the standard library + ArcPy",
          used <= allowed, f"unexpected: {sorted(used - allowed)}")
    check("the worker never writes the sign-in token into an event",
          "token=token" not in src and '"token":' not in src.replace('"token": token["token"]', ""))


def test_scenario():
    print("the shared scenario (also the frozen exe's self-test):")
    tmp = Path(tempfile.mkdtemp(prefix="agscenario_"))
    try:
        ast_.run(tmp, emit=lambda line: print(f"     {line}"))
        check("arcgis_selftest.run passes", True)
    except AssertionError as e:
        check("arcgis_selftest.run passes", False, str(e))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_truncated_export_is_refused():
    print("a truncated export never reaches the library:")
    with _library() as lib:
        before = (lib / "11_SHS Tolls.xlsx").read_bytes()
        spec = ast_.sample_spec()
        spec["layers"]["138"]["truncate"] = 2
        res, by, _logs = _refresh(spec, ["SHS Tolls"])
        check("the layer fails as truncated",
              by["SHS Tolls"]["status"] == "failed"
              and "truncated" in (by["SHS Tolls"].get("message") or ""),
              str(by["SHS Tolls"]))
        check("the run reads failed", res.status == "failed", res.status)
        check("the library's file is untouched",
              (lib / "11_SHS Tolls.xlsx").read_bytes() == before)
        check("the manifest still says 3 rows", crl.read_index(lib)["SHS Tolls"]["rows"] == 3)
        check("nothing was backed up", not (lib / "_previous").exists())


def test_cancel():
    print("Cancel stops the worker; finished layers stay, the rest are untouched:")
    with _library() as lib:
        cancel = threading.Event()
        spec = ast_.sample_spec()
        spec["layers"]["74"]["delay"] = 30          # City stalls in Table To Excel
        fake = ast_.make_fake_arcpy(spec)

        def on_progress(p):
            if p.get("done") == 1:                  # SHS Tolls is in: cancel
                cancel.set()

        res = ar.refresh(["SHS Tolls", "City"], events=Events(), cancel_event=cancel,
                         on_progress=on_progress, python="fake-python.exe",
                         launcher=ast_.inprocess_launcher(fake))
        by = {r["name"]: r for r in res.layers}
        check("the run reads cancelled", res.status == "cancelled", res.status)
        check("the finished layer was kept", by["SHS Tolls"]["status"] == "exported")
        check("the stopped layer reads cancelled", by["City"]["status"] == "cancelled",
              str(by["City"]))
        check("the stopped layer never reached the library", not (lib / "33_City.xlsx").exists())
        check("staging is empty after a cancel", not any(ar.staging_dir().glob("*")))


def test_catalog_by_name():
    print("signed in, layers are matched by NAME from the service's own list:")

    class _Resp(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    service_json = {"layers": [{"id": 999, "name": "SHS Tolls"},
                               {"id": 74, "name": "City"},
                               {"id": 85, "name": "County Codes LRS"}],
                    "tables": []}
    seen = {}

    def fake_urlopen(req, timeout=None):
        seen["url"] = req.full_url
        seen["body"] = req.data.decode()
        return _Resp(json.dumps(service_json).encode())

    with _library() as lib, _patch(urllib.request, "urlopen", fake_urlopen):
        spec = ast_.sample_spec(signed_in=True)
        spec["layers"]["999"] = spec["layers"].pop("138")     # the service moved it
        res, by, logs = _refresh(spec, ["SHS Tolls", "City", "Route Direction",
                                        "County Code"])
        check("a moved layer is found by name", by["SHS Tolls"]["status"] == "exported",
              str(by["SHS Tolls"]))
        missed = by["Route Direction"].get("message") or ""
        check("a layer the service doesn't list fails by name",
              by["Route Direction"]["status"] == "failed"
              and "no layer named" in missed and "is called" not in missed, missed)
        renamed = by["County Code"].get("message") or ""
        check("a renamed layer fails and the message says what the service calls it",
              by["County Code"]["status"] == "failed"
              and "no layer named 'County Code'" in renamed
              and "The service's layer 85 is called 'County Codes LRS'." in renamed,
              renamed)
        check("the token goes in the POST body, never the URL",
              "token=" not in seen.get("url", "") and "token=" in seen.get("body", ""))
        check("no log line carries the token", not any("fake-token" in m for m in logs))
        check("the manifest records the URL the layer was read from",
              crl.read_index(lib)["SHS Tolls"]["source"].endswith("/FeatureServer/999"))


def test_other_service():
    print("a layer of another service is read from that service, matched by its own list:")

    class _Resp(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    main_json = {"layers": [{"id": 138, "name": "SHS Tolls"}], "tables": []}
    # As the list's development copy numbers it: the name finds it anyway.
    csms_json = {"layers": [{"id": 0, "name": "Master Intersection"},
                            {"id": 7, "name": "Something Else"}], "tables": []}
    seen = []

    def fake_urlopen(req, timeout=None):
        seen.append((req.full_url, req.data.decode()))
        body = csms_json if "csms_supporting_prod" in req.full_url else main_json
        return _Resp(json.dumps(body).encode())

    with _library() as lib, _patch(urllib.request, "urlopen", fake_urlopen):
        spec = ast_.sample_spec(signed_in=True)
        spec["layers"]["0"] = spec["layers"].pop("204")       # moved, as on dev
        res, by, logs = _refresh(spec, ["SHS Tolls", "Master Intersection"])
        mi = by["Master Intersection"]
        check("both layers export", res.status == "ok" and mi["status"] == "exported"
              and mi.get("rows") == 12, f"{res.status}: {mi}")
        urls = [u for u, _b in seen]
        check("each service's own layer list is read, once",
              urls.count(ar.SERVICE_DEFAULT) == 1 and urls.count(ar.CSMS_SERVICE) == 1, urls)
        check("the token goes in the POST body for both, never a URL",
              all("token=" not in u and "token=" in b for u, b in seen))
        check("the manifest records the service and number the layer was read from",
              crl.read_index(lib)["Master Intersection"]["source"]
              == f"{ar.CSMS_SERVICE}/0", crl.read_index(lib)["Master Intersection"]["source"])
        check("the log names the second service",
              any(f"and Master Intersection from {ar.CSMS_SERVICE}" in m for m in logs)
              and any(f"Read {ar.CSMS_SERVICE}'s layer list" in m for m in logs), logs[:6])
        import clean_road_sources as crs
        index = crl.read_index(lib)
        check("the builds still read one service; Master Intersection has its own",
              crs.library_service(index) == ar.SERVICE_DEFAULT
              and crs.library_service(index, ["Master Intersection"]) == ar.CSMS_SERVICE,
              (crs.library_service(index), crs.library_service(index, ["Master Intersection"])))
    with _library():
        spec = ast_.sample_spec()
        spec["layers"]["204"]["service"] = ar.SERVICE_DEFAULT  # not where the app looks
        _res, by, _logs = _refresh(spec, ["Master Intersection"])
        check("a layer asked of the wrong service is not found there",
              by["Master Intersection"]["status"] == "failed"
              and "000732" in (by["Master Intersection"].get("message") or ""),
              str(by["Master Intersection"]))
    status = None
    with _library():
        status = ar.library_status()
    check("the Layers tab is told which layers come from which other service",
          status and status["other_services"] == [{"url": ar.CSMS_SERVICE,
                                                   "layers": ["Master Intersection"]}]
          and status["expected"] == 44, str(status and status.get("other_services")))


def test_partial_layer():
    print("a layer the library holds in part is exported with its where clause:")
    header = ["OBJECTID", "NetworkId", "RouteId", "Measure"]
    rows = [[i, 2 if i % 3 else 1, f"LA.001..{i:02d}", i / 10] for i in range(1, 31)]
    kept = [r for r in rows if r[1] == 2]

    def spec():
        s = ast_.sample_spec()
        s["layers"]["1"] = {"name": "Calibration Point", "header": header, "rows": rows,
                            "where": {"NetworkId = 2": ["NetworkId", 2]}}
        return s

    with _library() as lib:
        res, by, logs = _refresh(spec(), ["Calibration Point", "SHS Tolls"])
        got = by["Calibration Point"]
        check("the layer exports", got["status"] == "exported", str(got))
        check("only the rows the clause keeps reach the library",
              got.get("rows") == len(kept) == 20 and got.get("server_rows") == 20, str(got))
        check("a new layer takes the next file number",
              got.get("file") == "41_Calibration Point.xlsx"
              and (lib / "41_Calibration Point.xlsx").is_file(), str(got.get("file")))
        index = crl.read_index(lib)
        check("the manifest's Data Source states the clause",
              index["Calibration Point"]["source"].endswith(
                  "/FeatureServer/1 (rows where NetworkId = 2)"),
              index["Calibration Point"]["source"])
        check("a whole layer's Data Source is its URL alone",
              index["SHS Tolls"]["source"].endswith("/FeatureServer/138"))
        check("the log says which rows were counted",
              any("Calibration Point" in m and "where NetworkId = 2" in m for m in logs))
        import clean_road_sources as crs
        check("the library still reads as one service",
              crs.library_service(index) == ar.SERVICE_DEFAULT, crs.library_service(index))
    with _library(), _patch(ar, "LAYER_FILTERS", {"Calibration Point": "NetworkID IS 2"}):
        _res, by, _logs = _refresh(spec(), ["Calibration Point"])
        check("a clause the service refuses fails that layer with ArcGIS's own message",
              by["Calibration Point"]["status"] == "failed"
              and "Invalid expression" in (by["Calibration Point"].get("message") or ""),
              str(by["Calibration Point"]))


def test_index_write_failure_rolls_back():
    print("a manifest that can't be written undoes that layer's swap:")
    with _library() as lib:
        before = (lib / "11_SHS Tolls.xlsx").read_bytes()

        def locked(*_a, **_k):
            raise PermissionError("[WinError 32] 00_INDEX.xlsx is open in Excel")

        logging.disable(logging.CRITICAL)        # the expected failure logs a traceback
        try:
            with _patch(ar, "write_index", locked):
                res, by, _logs = _refresh(ast_.sample_spec(), ["SHS Tolls"])
        finally:
            logging.disable(logging.NOTSET)
        check("the layer fails with the Excel hint",
              by["SHS Tolls"]["status"] == "failed"
              and "close it in Excel" in (by["SHS Tolls"].get("message") or ""),
              str(by["SHS Tolls"]))
        check("the library file is the one from before",
              (lib / "11_SHS Tolls.xlsx").read_bytes() == before)
        check("the manifest is the one from before", crl.read_index(lib)["SHS Tolls"]["rows"] == 3)


def test_not_signed_in_hint():
    print("a refresh that reads nothing while signed out says so plainly:")
    with _library():
        spec = ast_.sample_spec()
        spec["layers"] = {}                                    # every read refused
        res, _by, _logs = _refresh(spec, ["SHS Tolls"])
        check("the run reads failed", res.status == "failed", res.status)
        check("the message says ArcGIS Pro is not signed in",
              "not signed in" in res.message and "Sign me in automatically" in res.message,
              res.message)
    hint = ar.signin_hint({"signed_in": True, "portal": "https://www.arcgis.com/"})
    check("a different active portal is named", "not the TSMIS portal" in hint, hint)
    check("a signed-in TSMIS portal needs no hint",
          ar.signin_hint({"signed_in": True,
                          "portal": "https://rhapps-prod.dot.ca.gov/portal/"}) == "")


def test_real_subprocess():
    print("the REAL launcher: a separate python.exe, the event file, exit + crash + stall:")
    tmp = Path(tempfile.mkdtemp(prefix="agproc_"))
    worker_dir = tmp / "worker"
    worker_dir.mkdir()
    shutil.copy2(arcgis_pro.worker_script(), worker_dir / arcgis_pro.WORKER_SCRIPT)
    (worker_dir / "arcpy.py").write_text(ast_.FAKE_ARCPY_SOURCE, encoding="utf-8")
    spec_path = tmp / "fake_arcpy.json"
    lib = tmp / "arcgis_layers"
    ast_.seed_manual_library(lib)
    script = worker_dir / arcgis_pro.WORKER_SCRIPT
    os.environ["TSMIS_FAKE_ARCPY"] = str(spec_path)     # passes the clean env through
    try:
        with _patch(paths, "ARCGIS_LAYERS_ROOT", lib), \
                _patch(arcgis_pro, "worker_script", lambda: script):
            spec_path.write_text(json.dumps(ast_.sample_spec()), encoding="utf-8")
            res = ar.refresh(["SHS Tolls", "City"], events=Events(), python=sys.executable)
            by = {r["name"]: r for r in res.layers}
            check("a real worker process exports both layers", res.status == "ok",
                  f"{res.status}: {res.message}")
            check("its rows reach the library",
                  by.get("City", {}).get("rows") == 40 and (lib / "33_City.xlsx").is_file())
            out = (ar.session_dir() / "worker-output.txt")
            check("the worker's console output is kept for diagnosis", out.is_file())

            spec_path.write_text(json.dumps(ast_.sample_spec(crash=True)), encoding="utf-8")
            res = ar.refresh(["SHS Tolls"], events=Events(), python=sys.executable)
            check("a worker that dies is a failed run, not a hang",
                  res.status == "failed", f"{res.status}: {res.message}")

            spec_path.write_text(json.dumps(ast_.sample_spec(import_delay=30)), encoding="utf-8")
            outcome = arcgis_pro.run_worker(
                ar.session_dir(), {"protocol": 1, "mode": "probe", "service": "x",
                                   "probe_layer": {"name": "SHS Tolls", "id": 138,
                                                   "out": str(tmp / "p.xlsx")}},
                lambda _ev: None, None, sys.executable, start_limit=2)
            check("a worker that never starts the run is stopped at the start limit",
                  "did not start" in outcome.timed_out, outcome.timed_out)
    finally:
        os.environ.pop("TSMIS_FAKE_ARCPY", None)
        shutil.rmtree(tmp, ignore_errors=True)


def test_layer_limit_scales():
    print("the per-layer time limit grows with the layer's own row count:")
    flat = arcgis_pro.LAYER_LIMIT_S
    check("a small layer keeps the flat limit", arcgis_pro.layer_limit_for(27) == flat)
    check("an unknown or unreadable count keeps the flat limit",
          arcgis_pro.layer_limit_for(None) == flat and arcgis_pro.layer_limit_for("x") == flat
          and arcgis_pro.layer_limit_for(0) == flat)
    big = arcgis_pro.layer_limit_for(824914)
    check("County Code (824,914 rows, 54 min on 2026-10-05) gets 3x its measured time",
          big >= 3 * 3234 and big > flat, str(big))
    with _library() as lib:
        spec = ast_.sample_spec()
        spec["layers"]["74"]["delay"] = 2.5                  # City: 40 rows, 2.5 s
        request = {"protocol": 1, "mode": "export",
                   "service": "https://example.invalid/FeatureServer",
                   "layers": [{"name": "City", "id": 74, "out": str(lib / "City.xlsx")}]}

        def _run(per_row):
            return arcgis_pro.run_worker(
                ar.session_dir(), request, lambda _ev: None, None, "fake-python.exe",
                launcher=ast_.inprocess_launcher(ast_.make_fake_arcpy(spec)),
                layer_limit=1, layer_seconds_per_row=per_row)

        roomy = _run(0.1)                                    # 40 rows -> a 4 s limit
        check("once counted, a layer gets its row-scaled limit over the floor",
              not roomy.timed_out, roomy.timed_out)
        tight = _run(0)                                      # no allowance -> the 1 s floor
        check("with no per-row allowance the floor still stops a stuck layer",
              "took longer than" in tight.timed_out, tight.timed_out)


def test_drop_dates():
    print("the drop's dates — every layer's own export time, the oldest wins:")
    from openpyxl import Workbook
    tmp = Path(tempfile.mkdtemp(prefix="agdates_"))
    try:
        for name in ("01_SHS Tolls.xlsx", "02_City.xlsx"):
            wb = Workbook()
            wb.active.append(["OBJECTID"])
            wb.save(tmp / name)
        rows = {"SHS Tolls": {"file": "01_SHS Tolls.xlsx", "rows": 0, "fields": 1,
                              "exported_at": "2026-09-30T08:00:00"},
                "City": {"file": "02_City.xlsx", "rows": 0, "fields": 1,
                         "exported_at": "2026-10-02T14:05:00"}}
        arcgis_layers.write_index(tmp, rows)
        drop = arcgis_layers.drop_info(tmp)
        check("the drop dates to its OLDEST layer",
              drop["exported"] == "2026-09-30" and drop["exported_source"] == "index", str(drop))
        check("a drop exported over several days says so",
              drop["mixed"] is True and drop["newest_at"] == "2026-10-02T14:05:00")
        check("a build reading only City may go as of City's own date",
              arcgis_layers.consistent_asof(["City"], drop) == "2026-10-02")
        check("a build reading both is held to the oldest",
              arcgis_layers.consistent_asof(["City", "SHS Tolls"], drop) == "2026-09-30")
        check("a layer with no recorded time falls back to the drop's date",
              arcgis_layers.consistent_asof(["Route Direction"], drop) == "2026-09-30")
        full, created = arcgis_layers.read_index_full(tmp)
        check("the manifest's own stamp is the oldest layer's time",
              created is not None and created.isoformat() == "2026-09-30T08:00:00",
              str(created))
        check("the six classic columns still read with the builds' reader",
              crl.read_index(tmp)["City"]["file"] == "02_City.xlsx")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_row_count():
    print("row counting reads the FIRST sheet's XML, exactly:")
    from openpyxl import Workbook
    tmp = Path(tempfile.mkdtemp(prefix="agrows_"))
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["a", "b"])
        for i in range(2500):
            ws.append([i, f"text {i}"])
        other = wb.create_sheet("Other")
        for i in range(7):
            other.append([i])
        wb.save(tmp / "w.xlsx")
        check("2,501 rows in the first sheet (header + 2,500)",
              arcgis_layers.count_sheet_rows(tmp / "w.xlsx") == 2501)
        info = arcgis_layers.inspect_export(tmp / "w.xlsx", server_rows=2500)
        check("inspect_export: 2,500 data rows, 2 columns",
              info["data_rows"] == 2500 and info["fields"] == 2)
        try:
            arcgis_layers.inspect_export(tmp / "w.xlsx", server_rows=2501)
            refused = False
        except ValueError:
            refused = True
        check("fewer rows than the server counted is refused", refused)
        (tmp / "bad.xlsx").write_bytes(b"not a workbook")
        try:
            arcgis_layers.inspect_export(tmp / "bad.xlsx")
            bad = False
        except ValueError:
            bad = True
        check("an unreadable file is refused as ValueError", bad)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_gui_wiring():
    print("the GUI wiring — endpoints, the queue job, the worker's messages:")
    import gui_api
    import gui_arcgis_layers_api as gal
    import gui_matrix
    import gui_worker_arcgis as gwa

    for name in ("arcgis_layers_info", "refresh_arcgis_layers", "check_arcgis_pro",
                 "choose_arcgis_python", "clear_arcgis_python"):
        check(f"GuiApi exposes {name}", callable(getattr(gui_api.GuiApi, name, None)))
    check("GuiApi carries the Layers mixin", gal.GuiArcgisLayersMixin in gui_api.GuiApi.__mro__)

    started = []

    class _StubWorker:
        def __init__(self, *args):
            started.append(args)

        def start(self):
            pass

    class _Host(gal.GuiArcgisLayersMixin):
        def __init__(self):
            self._lock = threading.RLock()
            self._matrix = None
            self.cancel_event = threading.Event()
            self.logs = []

        def _emit_log(self, m):
            self.logs.append(m)

        def _set_dot(self, *_a):
            pass

        def _emit(self, _e):
            pass

        def _gated_queue(self):
            return "Q"

    host = _Host()
    job = gui_matrix.GuiMatrixMixin._make_job.__get__(_JobFactory())(
        "arcgis_refresh", "cell", "Refresh 2 ArcGIS layers", which="arcgis",
        layers=["SHS Tolls", "City"])
    check("a refresh job carries its layers", job["layers"] == ["SHS Tolls", "City"])
    with _patch(gal, "ArcgisLayerRefreshWorker", _StubWorker):
        ok = host._dispatch_arcgis_refresh_job(job)
    check("dispatch starts the refresh worker with the layers",
          ok and started and started[-1][0] == ["SHS Tolls", "City"])
    check("dispatch sets the layers phase and total",
          host._matrix == {"phase": "layers", "row": None, "cell": None, "done": 0, "total": 2})

    # The real worker thread over the in-process stand-in: its message stream.
    with _library():
        fake = ast_.make_fake_arcpy(ast_.sample_spec())
        wrapped = functools.partial(ar.refresh, python="fake-python.exe",
                                    launcher=ast_.inprocess_launcher(fake))
        q = queue.Queue()
        with _patch(ar, "refresh", wrapped):
            w = gwa.ArcgisLayerRefreshWorker(["SHS Tolls", "Route Direction"], q,
                                             threading.Event())
            w.start()
            w.join(timeout=60)
        msgs = []
        while not q.empty():
            msgs.append(q.get())
        kinds = [k for k, _p in msgs]
        done = [p for k, p in msgs if k == "matrix_done"]
        check("exactly one matrix_done ends the job", kinds.count("matrix_done") == 1
              and kinds[-1] == "matrix_done")
        check("matrix_done names the job and counts the layers",
              done and done[0]["label"] == "Layer refresh" and done[0]["total"] == 2
              and done[0]["succeeded"] == 1 and done[0]["failed"] == 1, str(done))
        check("progress goes out as matrix_cell with the layer as the row",
              any(k == "matrix_cell" and p.get("row") == "SHS Tolls" for k, p in msgs))


class _JobFactory:
    """Just enough of GuiApi for _make_job (the shared id sequence)."""

    class _Coord:
        n = 0

        def next_seq(self):
            self.n += 1
            return self.n

    _coord = _Coord()


def main():
    print("=== ArcGIS layer refresh (in-app export) ===")
    test_contract()
    test_scenario()
    test_truncated_export_is_refused()
    test_cancel()
    test_catalog_by_name()
    test_other_service()
    test_partial_layer()
    test_index_write_failure_rolls_back()
    test_not_signed_in_hint()
    test_real_subprocess()
    test_layer_limit_scales()
    test_drop_dates()
    test_row_count()
    test_gui_wiring()
    print()
    if _fail:
        print(f"FAILED: {len(_fail)} check(s): {_fail}")
        return 1
    print("ALL ARCGIS-REFRESH CHECKS PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
