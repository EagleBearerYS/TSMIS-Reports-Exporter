#!/usr/bin/env python3
"""Photograph every screen of the GUI and list what overflows.

Walks the real UI (scripts/ui) in its built-in ``#mock`` preview with a headless
browser: every tab and sub-tab, each matrix with day columns added, the options
corners scrolled, Remove all / Add all, and one confirm dialog. For each view it
writes ``<out>/<view>.png`` and records an overflow audit in ``<out>/report.json``:

    HSCROLL    a container that scrolls sideways (something inside is too wide)
    CLIP       text cut off by its box with no ellipsis
    ELLIPSIS   text cut with an ellipsis (fine when the full text is in a tooltip)
    OFFSCREEN  text that starts on screen and ends past the window's right edge

Usage (from the repo root; needs Playwright's Chromium, as tools/screenshots.py):

    python tools/ui_walk.py <out_dir> [width height [dark|light [view-prefix ...]]]

The two sizes worth running are 1164x741 (the default window's page area) and
910x512 (a 1366x768 work PC at 150% DPI), in both themes. A view prefix limits
the screenshots to matching views; every view's setup still runs, because later
views continue from earlier ones.

The mock shows the layout with plausible data, not the real app. To look at the
REAL window the same way, start the app with
``WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9333`` and attach
with ``playwright.chromium.connect_over_cdp("http://127.0.0.1:9333")`` -- the
page is then the live one, bridge and data included (docs/gui.md).
"""
import json
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

if len(sys.argv) < 2:
    raise SystemExit(__doc__)
ui_dir = Path(__file__).resolve().parent.parent / "scripts" / "ui"
out_dir = Path(sys.argv[1])
width = int(sys.argv[2]) if len(sys.argv) > 3 else 1164
height = int(sys.argv[3]) if len(sys.argv) > 3 else 741
theme = sys.argv[4] if len(sys.argv) > 4 else "dark"
only = set(sys.argv[5:])
out_dir.mkdir(parents=True, exist_ok=True)
URL = (ui_dir / "index.html").resolve().as_uri() + "#mock"

HELPERS = r"""
window.__sleep = (ms) => new Promise((r) => setTimeout(r, ms));
window.__tab = async (id) => { document.getElementById(id).click(); await __sleep(350); };
window.__group = async (g) => {
  document.querySelector('[role=tab][data-group="' + g + '"]').click(); await __sleep(700);
};
window.__addDay = async (selId, btnId, n) => {
  for (let i = 0; i < n; i++) {
    const sel = document.getElementById(selId);
    if (!sel || !sel.options.length) break;
    // prefer a day that holds exports (skip the export-less "today")
    let pick = [...sel.options].find((o) => !/nothing exported/.test(o.textContent)) || sel.options[0];
    sel.value = pick.value; sel.dispatchEvent(new Event('change', { bubbles: true }));
    document.getElementById(btnId).click(); await __sleep(900);
  }
};
window.__bulk = async (hostId, text) => {
  const b = [...document.querySelectorAll('#' + hostId + ' .mx-toggle-all')].find((x) => x.textContent.trim() === text);
  if (!b) throw new Error('no bulk button ' + text + ' in ' + hostId);
  b.click(); await __sleep(1000);
};
window.__chip = async (hostId, text) => {
  const b = [...document.querySelectorAll('#' + hostId + ' .mx-toggle:not(.mx-toggle-all)')].find((x) => x.textContent.trim() === text);
  if (!b) throw new Error('no chip ' + text + ' in ' + hostId);
  b.click(); await __sleep(900);
};
window.__desc = (e) => ((e.id ? '#' + e.id : '') + (e.className && typeof e.className === 'string'
  ? '.' + e.className.trim().split(/\s+/).slice(0, 3).join('.') : '')) || e.tagName.toLowerCase();
window.__audit = () => {
  const out = []; const vw = innerWidth;
  for (const e of document.querySelectorAll('body *')) {
    const r = e.getBoundingClientRect();
    if (r.width <= 0 || r.height <= 0) continue;
    const cs = getComputedStyle(e);
    if (cs.visibility === 'hidden') continue;
    const d = __desc(e);
    if (e.scrollWidth > e.clientWidth + 2 && (cs.overflowX === 'auto' || cs.overflowX === 'scroll'))
      out.push('HSCROLL ' + d + ' sw=' + e.scrollWidth + ' cw=' + e.clientWidth);
    else if (cs.textOverflow === 'ellipsis' && e.scrollWidth > e.clientWidth + 1)
      out.push('ELLIPSIS ' + d + ' "' + e.textContent.trim().slice(0, 60) + '"');
    else if (cs.overflowX === 'hidden' && e.scrollWidth > e.clientWidth + 2 && e.children.length === 0)
      out.push('CLIP ' + d + ' "' + e.textContent.trim().slice(0, 60) + '"');
    if (r.right > vw + 1 && r.left < vw && e.children.length === 0 && e.textContent.trim())
      out.push('OFFSCREEN ' + d + ' right=' + Math.round(r.right));
  }
  return [...new Set(out)];
};
window.__scrollCfg = (frac) => {
  const col = document.querySelector('.col.col-config');
  const sc = [col, ...col.querySelectorAll('*')].find((e) => e.scrollHeight > e.clientHeight + 30
    && ['auto', 'scroll'].includes(getComputedStyle(e).overflowY) && e.getBoundingClientRect().height > 0);
  if (sc) sc.scrollTop = (sc.scrollHeight - sc.clientHeight) * frac;
  return sc ? __desc(sc) + ' ' + sc.scrollHeight + '/' + sc.clientHeight : 'none';
};
window.__scrollEl = (sel, frac) => {
  const sc = document.querySelector(sel);
  if (sc) sc.scrollTop = (sc.scrollHeight - sc.clientHeight) * frac;
  return sc ? sc.scrollHeight + '/' + sc.clientHeight : 'none';
};
"""

VIEWS = [
    ("01-export", "await __tab('tabExport');"),
    ("02-export-mid", "await __tab('tabExport'); __scrollCfg(0.5);"),
    ("03-export-bottom", "await __tab('tabExport'); __scrollCfg(1);"),
    ("04-consolidate", "await __tab('tabConsolidate');"),
    ("05-consolidate-bottom", "await __tab('tabConsolidate'); __scrollCfg(1);"),
    ("06-compare-env", "await __tab('tabCompare'); await __group('env');"),
    ("07-compare-env-bottom", "await __tab('tabCompare'); await __group('env'); __scrollCfg(1);"),
    ("08-compare-tsn", "await __tab('tabCompare'); await __group('tsn');"),
    ("09-compare-self", "await __tab('tabCompare'); await __group('self');"),
    ("10-day-matrix-empty", "await __tab('tabCompare'); await __group('tsn_by_day');"),
    ("11-day-matrix", "await __tab('tabCompare'); await __group('tsn_by_day'); await __addDay('dayMatrixAddDay','btnDayAddDay',3);"),
    ("12-day-matrix-config-mid", "__scrollEl('#dayMatrixConfig', 0.45);"),
    ("13-day-matrix-config-bottom", "__scrollEl('#dayMatrixConfig', 1);"),
    ("13b-day-remove-all", "await __bulk('dayMatrixReportToggles', 'Remove all'); __scrollEl('#dayMatrixConfig', 1);"),
    ("13c-day-pick-two", "await __chip('dayMatrixReportToggles', 'Highway Log'); await __chip('dayMatrixReportToggles', 'Clean Road: Ramp'); __scrollEl('#dayMatrixConfig', 0.75);"),
    ("13d-day-add-all", "await __bulk('dayMatrixReportToggles', 'Add all'); __scrollEl('#dayMatrixConfig', 0);"),
    ("13e-day-export-confirm", "document.getElementById('btnDayExportToday').click(); await __sleep(500);"),
    ("13f-day-export-confirm-closed", "[...document.querySelectorAll('#modalOverlay button')].find((b) => /Cancel/.test(b.textContent)).click(); await __sleep(300);"),
    ("14-baseline-matrix", "await __tab('tabCompare'); await __group('baseline_by_day'); await __addDay('baselineMatrixAddDay','btnBaselineAddDay',3);"),
    ("15-baseline-config-bottom", "__scrollEl('#baselineMatrixConfig', 1);"),
    ("16-pve-matrix", "await __tab('tabCompare'); await __group('pdf_vs_excel_by_day'); await __addDay('pveMatrixAddDay','btnPveAddDay',3);"),
    ("17-three-days", "await __tab('tabCompare'); await __group('three_days'); document.getElementById('btnThreeDayAddSet').click(); await __sleep(900);"),
    ("18-three-days-config-bottom", "__scrollEl('#threeDayMatrixConfig', 1);"),
    ("19-everything-export", "await __tab('tabEverything'); await __tab('subEveryExport');"),
    ("20-everything-export-bottom", "__scrollCfg(1);"),
    ("21-everything-matrix", "await __tab('tabEverything'); await __tab('subEveryMatrix'); await __sleep(700);"),
    ("22-everything-matrix-config-bottom", "__scrollEl('#matrixConfig', 1);"),
    ("23-arcgis-layers", "await __tab('tabArcgis'); await __tab('subAgLayers'); await __sleep(500);"),
    ("24-arcgis-layers-bottom", "__scrollCfg(1);"),
    ("25-arcgis-builds", "await __tab('tabArcgis'); await __tab('subAgBuilds'); await __sleep(500);"),
    ("26-arcgis-matrix", "await __tab('tabArcgis'); await __tab('subAgReports'); await __sleep(500); await __addDay('agMatrixAddDay','btnAgAddDay',3);"),
    ("27-arcgis-threeway", "await __tab('tabArcgis'); await __tab('subAgThreeWay'); await __sleep(500); await __addDay('twMatrixAddDay','btnTwAddDay',3);"),
    ("28-arcgis-cleanroad", "await __tab('tabArcgis'); await __tab('subAgCleanRoad'); await __sleep(500);"),
    ("29-arcgis-cleanroad-bottom", "__scrollCfg(1);"),
    ("30-settings", "await __tab('tabSettings'); await __sleep(600);"),
    ("31-settings-25", "__scrollCfg(0.25);"),
    ("32-settings-50", "__scrollCfg(0.5);"),
    ("33-settings-75", "__scrollCfg(0.75);"),
    ("34-settings-bottom", "__scrollCfg(1);"),
]

report = {}
with sync_playwright() as p:
    browser = p.chromium.launch()
    ctx = browser.new_context(viewport={"width": width, "height": height},
                              color_scheme=theme, device_scale_factor=1)
    page = ctx.new_page()
    errors = []
    page.on("console", lambda m: errors.append(f"{m.type}: {m.text}") if m.type in ("error", "warning") else None)
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
    page.goto(URL)
    page.wait_for_function("typeof S !== 'undefined' && !!S.st", timeout=20000)
    page.evaluate(HELPERS)
    page.wait_for_timeout(600)
    for name, js in VIEWS:
        if only and name not in only and not any(name.startswith(o) for o in only):
            # still run the setup so later views (which continue from it) work
            pass
        try:
            page.evaluate("async () => {" + js + "}")
        except Exception as e:  # noqa: BLE001 - recorded per view
            report[name] = {"error": str(e)[:400]}
            continue
        page.wait_for_timeout(350)
        if only and not any(name.startswith(o) for o in only):
            continue
        page.screenshot(path=str(out_dir / f"{name}.png"))
        info = page.evaluate("""() => ({
            audit: __audit(),
            docW: document.documentElement.scrollWidth, vw: innerWidth,
            docH: document.documentElement.scrollHeight, vh: innerHeight,
            body: document.body.className,
        })""")
        report[name] = info
    report["_console"] = errors
    browser.close()

(out_dir / "report.json").write_text(json.dumps(report, indent=1), encoding="utf-8")
for name, info in report.items():
    if name == "_console":
        print("console:", len(info), "messages")
        for line in info[:20]:
            print("   ", line[:200])
        continue
    if "error" in info:
        print(f"{name}: ERROR {info['error']}")
        continue
    flags = []
    if info["docW"] > info["vw"]:
        flags.append(f"PAGE-HSCROLL {info['docW']}>{info['vw']}")
    if info["docH"] > info["vh"]:
        flags.append(f"page-vscroll {info['docH']}>{info['vh']}")
    print(f"{name}: {len(info['audit'])} audit lines {' '.join(flags)}")
    for line in info["audit"]:
        print("    ", line[:170])
