# TSMIS Reports Exporter

> Bulk-export Caltrans TSMIS reports for every California state route — from a single click.

[![Version](https://img.shields.io/github/v/release/EagleBearerYS/TSMIS-Reports-Exporter?label=version&color=blue)](https://github.com/EagleBearerYS/TSMIS-Reports-Exporter/releases/latest)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-0078D6?logo=windows)](#)
[![Python](https://img.shields.io/badge/python-3.11-3776AB?logo=python&logoColor=white)](#)
[![Automation](https://img.shields.io/badge/automation-Playwright-2EAD33?logo=microsoftedge&logoColor=white)](#)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

A portable Windows desktop tool that bulk-exports reports from the Caltrans
**Transportation System Management Information System (TSMIS)** for every
California state route. Pick the report types you want, sign in once, and the
tool walks every route for you — no copy-pasting route numbers, no babysitting
the browser. It also combines the per-route files into one workbook per report
and compares the results: against TSN (the older system TSMIS replaced), between
environments, between export days, PDF against Excel, and against the ArcGIS
layers.

It is distributed as a single zip — unzip, double-click, done — in two flavors:
the standard build drives the **Google Chrome or Microsoft Edge already on the
machine** (nothing extra to install), and a *with-browser* build ships its own
**Built-in Chromium** and uses it by default, for PCs where managed browsers get
in the way. The same engine also runs from a set of `.bat` scripts for
development and as a fallback.

<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="docs/screenshot-dark.png">
    <img alt="TSMIS Reports Exporter — Export tab" src="docs/screenshot-light.png" width="820">
  </picture>
</p>

---

## Table of Contents

- [Features](#features)
- [Supported reports](#supported-reports)
- [Getting started (end users)](#getting-started-end-users)
- [Usage](#usage)
- [Developer setup](#developer-setup)
- [Building the app](#building-the-app)
- [Project structure](#project-structure)
- [Tech stack](#tech-stack)
- [Known limitations](#known-limitations)
- [Contributing](#contributing)
- [License & disclaimer](#license--disclaimer)

---

## Features

- **Eleven reports, twenty-two editions, any combination.** Every report the site
  can export, each as an Excel file and as a print-layout PDF — export one, several,
  or all at once. Tick both editions of a report and it's generated once, saved both ways.
- **One login for everything.** A single SSO + MFA sign-in covers every report. On
  a Caltrans PC, an export with no saved login first tries to sign in by itself
  through Microsoft Edge.
- **Pick your routes.** Run all routes by default, or narrow to a subset
  (`5, 99, 101` — any casing/zero-padding, suffixes like `101U` accepted).
- **Resumable.** Re-running the same day skips routes already on disk, so an
  interrupted run picks up where it left off.
- **Resilient.** Per-route timeouts, automatic retry of stragglers under lighter
  load, failure screenshots, and rotating logs.
- **Skip, pause & cancel.** Skip a slow route mid-run, pause between routes, or
  cancel the whole run — the current route stops promptly, not after a long wait.
- **Consolidation.** Combine a report's per-route files into a single workbook —
  every report, both editions. The PDF editions are read back into Excel, so the
  two editions of a report can be checked against each other.
- **Comparisons for every report.** Each report, in both editions, compares
  against TSN, between environments, PDF against Excel, and against an earlier
  export day; the Clean Road reports also compare across three export days at
  once. The result is a discrepancy workbook — matching values shown plainly,
  differences in red, per-route coverage — saved as a values workbook (opens
  instantly), a live-formulas workbook (recalculates in Excel), or both.
- **Comparison matrices.** Four grids on the Compare tab — vs TSN, vs Baseline
  (an earlier day, say), PDF vs Excel and Three Days — where each row is a report,
  each column an export day, and each cell one comparison: build it, see how many
  cells differ (green = identical), open the workbook. *Export today* on the
  vs TSN Matrix exports every report on the grid for today, then compares.
- **Everything tab.** *Refresh & export* pulls every ticked report for every
  ticked environment into one folder that is overwritten each time, so it always
  holds the latest copy — pause it and pick it up later, even after closing the
  app. Its *Comparison matrix* is a grid of reports by environments built from
  that folder.
- **ArcGIS tab** *(needs ArcGIS Pro)*. Pulls the TSMIS map layers with the ArcGIS
  Pro installed on the PC, builds reports from them (Highway Detail, Intersection
  Detail and the Clean Road files so far), and compares your exports against those
  builds; the Clean Road files can also be compared against TSN and the ArcGIS
  build at once. Ignore the tab if you don't have ArcGIS Pro.
- **TSN library.** The app doesn't download TSN data: import each report's TSN
  files once under Settings ▸ *TSN reports*, and every vs-TSN comparison reads them
  from the `tsn_library\` folder.
- **Run reports.** Every run records a per-route outcome CSV (saved / empty /
  skipped / failed).
- **One-click validation.** Settings ▸ *Validate & package results* runs every
  report already on this PC through the real comparison pipeline and saves the
  outcomes, TSN freshness, and logs into one credential-safe file for support.
- **Optional fast mode.** Run several browsers in parallel for a 2.5–3×+ speedup.
- **Browser, your way.** The standard download uses the machine's installed
  Chrome/Edge, keeping it small (~170 MB unzipped); the with-browser download adds a
  Built-in Chromium that works even where org-managed browsers interfere. The
  title bar shows which browser is exporting (*Export via*); when both Google
  Chrome and the Built-in Chromium are present, choose between them under
  Settings ▸ *Export browser*. Microsoft Edge handles automatic sign-in and is
  the fallback.
- **Updates itself.** When a newer version is out, an *Update to v…* button
  appears in the title bar — click it, then *Restart to update*; your reports,
  login and settings are kept.

## Supported reports

| Report | Output format | Output folder |
|---|---|---|
| TSAR: Ramp Summary | PDF (Letter) | `output\<run>\ramp_summary\` |
| TSAR: Ramp Summary (Excel) | XLSX | `output\<run>\ramp_summary_excel\` |
| TSAR: Ramp Detail | XLSX | `output\<run>\ramp_detail\` |
| TSAR: Ramp Detail (PDF) | PDF (Letter, landscape) | `output\<run>\ramp_detail_pdf\` |
| Highway Sequence Listing | XLSX | `output\<run>\highway_sequence\` |
| Highway Sequence Listing (PDF) | PDF (Letter, portrait) | `output\<run>\highway_sequence_pdf\` |
| Highway Log | XLSX | `output\<run>\highway_log\` |
| Highway Log (PDF) | PDF (Letter, landscape) | `output\<run>\highway_log_pdf\` |
| Intersection Summary | XLSX | `output\<run>\intersection_summary\` |
| Intersection Summary (PDF) | PDF (Letter, portrait) | `output\<run>\intersection_summary_pdf\` |
| Intersection Detail | XLSX | `output\<run>\intersection_detail\` |
| Intersection Detail (PDF) | PDF (Letter, landscape) | `output\<run>\intersection_detail_pdf\` |
| Highway Detail | XLSX | `output\<run>\highway_detail\` |
| Highway Detail (PDF) | PDF (Letter, landscape) | `output\<run>\highway_detail_pdf\` |
| Highway Summary | XLSX | `output\<run>\highway_summary\` |
| Highway Summary (PDF) | PDF (Letter, portrait) | `output\<run>\highway_summary_pdf\` |
| Clean Road: Highway / Intersection / Ramp | XLSX | `output\<run>\clean_<kind>\` |
| Clean Road: Highway / Intersection / Ramp (PDF) | PDF (Letter, landscape) | `output\<run>\clean_<kind>_pdf\` |

`<run>` is the run folder, named for the day, the data source and environment, and
the TSMIS site the export came from — `<date> <source>-<environment> <site>-site`,
e.g. `output\2026-11-13 ssor-prod main-site\highway_log\`. A main-site export and a
development-site export never share a folder.

All eleven reports export in both editions — the regular Excel export and a
print-layout **PDF** — and ticking both editions of a report generates it once and
saves both. Every edition consolidates and compares: vs TSN, between environments,
PDF vs Excel, and against an earlier export day. The app also lists the site's
**Route History Table**, greyed out as *not yet available* — the site has no export
for it yet.

New reports are released first on the **development** TSMIS site and move to the
main site once they are approved, so a report can be greyed out or fail on the main
site and still export from the development one. A fresh copy uses the main site;
switch under Settings ▸ *TSMIS site addresses* ▸ **Use development site** (**Use
main site** switches back). This is separate from the *Environment* picker.

**TSN files** are not downloaded by the app. Import them once under Settings ▸
*TSN reports* (*Import raw…*, then *Rebuild*); they are kept in `tsn_library\`, one
folder per report, and every vs-TSN comparison reads them from there. TSN's data
stopped changing in September 2025, so most differences against TSN are things that
changed in TSMIS since then, not errors.

---

## Getting started (end users)

> **Requirements:** Windows 10/11 and TSMIS credentials. No Python or other
> setup needed. The standard download also needs **Google Chrome or Microsoft
> Edge** installed; the with-browser download brings its own.

1. **Download** the variant that fits from the
   [Releases page](https://github.com/EagleBearerYS/TSMIS-Reports-Exporter/releases):
   - `…-win64.zip` — standard; uses the Chrome/Edge already on the PC (smallest).
   - `…-win64-with-browser.zip` — ships its own Built-in Chromium and uses it by
     default (best on managed PCs where Edge sign-in misbehaves and Chrome
     isn't installed).
   - `…-batch-source.zip` — the console/`.bat` flow for developers (see
     [Developer setup](#developer-setup)).
2. **Unblock it** before extracting: right-click the zip → **Properties** →
   tick **Unblock** → **OK**. (It's an unsigned app, so Windows marks downloads
   as untrusted — see [Known limitations](#known-limitations).)
3. **Extract** the folder somewhere you can write, and keep its path **short** —
   `C:\TSMIS` is ideal. Windows and Excel limit how long a file's full path can
   be, and a deep folder (a OneDrive Desktop, for example) can make exports and
   comparisons fail to save or open.
4. **Run** `TSMIS Exporter.exe`. On first launch, choose
   **More info → Run anyway** if SmartScreen appears. Keep the folder together —
   the `.exe` needs the `_internal` folder next to it.

Exports, logs, and your saved session are written **next to the `.exe`** (or under
`%LOCALAPPDATA%\TSMIS Exporter` if that folder is read-only), so the tool is fully
portable and leaves nothing behind on the system. The `Start Here.txt` in the zip
is the full quick start; `IT-README.txt` answers IT and security questions.

## Usage

1. **Log in** — click **Log in**. A browser window opens: sign in with your
   Caltrans account and complete MFA until the TSMIS report page loads, then
   click **I've finished logging in** (or just close the browser window). The
   session is saved and reused until it expires; your password goes into the
   browser, never into the app. On a Caltrans PC you can also just start an
   export: with no saved login, the app first tries to sign in by itself with
   Microsoft Edge and your Windows account, and asks you to log in only if that
   doesn't work.
2. **Export** — tick the report types, optionally enter a subset of routes
   (blank = all), and click **Start export**. The **Data source** (SSOR / ARS)
   and **Environment** (Prod / Test / Dev) pickers under the title choose what
   you export — defaults SSOR + Prod. Files land in a run folder named for the
   day, the source and environment, and the TSMIS site, e.g.
   `output\2026-11-13 ssor-prod main-site\highway_log\`. Progress, counts, and a
   live log are shown; use **Skip route** for a slow route, **Pause** to hold
   the run between routes, or **Cancel** to stop.
3. **Consolidate** *(optional)* — on the Consolidate tab, pick a report (and
   which export day — newest by default) and click **Start consolidation** to
   combine that day's per-route files into one workbook under the run folder's
   `consolidated\`. Or tick *Consolidate each report when its export finishes*
   on the Export tab (the PDF editions, Ramp Summary aside, still go through the
   Consolidate tab).
4. **Compare** *(optional)* — the Compare tab has seven sub-tabs. Three compare
   two things you pick and ask where to save the result: **Cross-environment**
   (one report from two export folders), **vs TSN** (a report's consolidated
   workbook against its TSN copy) and **Self-consistency** (a report's PDF
   edition against its Excel edition). Tick *Values workbook*, *Live-formulas
   workbook*, or both; differences show in red as `value ≠ value`. A big
   live-formulas workbook opens with blank cells — press **F9** in Excel and
   wait. The other four are matrices (**vs TSN**, **vs Baseline**, **PDF vs
   Excel**, **Three Days**): grids of reports by export day in which each cell
   is one comparison to build and open.
5. **Save run report** *(optional)* — export a CSV of every route's outcome.

Generated folders keep the files you use (XLSX, PDF, images, and run reports) at
the top level. Internal outcome records, provenance, comparison payloads, locks,
and Matrix caches are grouped under one `_state` subfolder in the artifact
folder they describe. Keep `_state` with its artifacts if you move or archive
them. An app-owned destination may also contain `.tsmis-owned.json` at its root;
that small marker is intentionally separate because Settings ▸ *Delete all
reports…* uses it to prove the folder is safe to delete.

**Fast mode** (experimental): tick **⚡ Fast mode** and choose how many browsers
to run in parallel. Each one uses ~0.5 GB RAM — 3–4 is a good number; higher
counts give a bigger speedup on a well-resourced PC. Fast mode needs a saved
login, and **Skip route** is off while it runs.

**If something goes wrong:** Settings ▸ Debugging ▸ **Save support bundle…** saves
one zip with the logs, run reports and settings — it includes this PC's name and
folder paths, but never your login. Send it to whoever gave you the app, with a
note of what you clicked and what happened.

---

## Developer setup

The packaged app is the product, but the same console-free engine runs from
`.bat` scripts (and plain Python) for development.

```bash
git clone https://github.com/EagleBearerYS/TSMIS-Reports-Exporter.git
cd TSMIS-Reports-Exporter
```

Then either double-click the numbered `.bat` files in order, or run the
equivalents directly (Python 3.11):

| Step | `.bat` | Equivalent |
|---|---|---|
| Install deps | `1. setup (one time).bat` | `pip install -r requirements.txt` + `playwright install chromium --no-shell` |
| Sign in | `2. login (update login).bat` | `python scripts/login.py` |
| Export | `3. run_export (main script).bat` | `python scripts/export_multi.py` |
| Consolidate | `4. consolidate (combine reports).bat` | `python scripts/consolidate_*.py` |
| Fast export | `5. fast export (experimental).bat` | set `TSMIS_FAST_WORKERS`, then export |
| GUI (dev) | `run app (GUI preview).bat` | `python scripts/gui_main.py` |

Setup downloads a Built-in Chromium that becomes the default browser; if that
download is skipped or fails, the tool falls back to the system Chrome/Edge.
Offline verification is a suite of golden checks under `build/check_*.py` (run with
the build venv; CI runs them all) plus the frozen-bundle self-test; the live-export
path against TSMIS still requires a login to exercise end to end.

## Building the app

From the repo root (PowerShell):

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1
```

This creates an isolated build venv, runs PyInstaller (onefolder, windowed), and
prunes the bundle to runtime-only files. The result is `dist\TSMIS Exporter\`
(~170 MB) — zip it to distribute.

Add `-BundleChromium` to build the with-browser variant (downloads Playwright's
Chromium into the bundle, which then defaults to it). Add `-SelfTest` to build
and run a headless self-test that verifies the frozen bundle (browser PDF +
download, PDF parsing, Excel, GUI) — a real release gate, not just "it
compiled":

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1 -SelfTest
```

**Releases** are built and published by CI: push a `v*` tag (or run the
`release` workflow manually) and `.github/workflows/release.yml` gates both
variants with the frozen self-test, then publishes all three zips.

See [`docs/build-and-release.md`](docs/build-and-release.md) for build internals,
bundle hygiene, and the DLP guard.

## Project structure

```
scripts/        Core engine (console-free) + console & GUI drivers
  common.py       Compatibility shim re-exporting the engine leaves (auth_nav,
                  report_nav, session, site_target, routes, timeouts, …)
  exporter.py     The shared per-route export loop (+ parallel variant)
  report_catalog.py  The report-metadata source of truth; reports.py derives from it
  cli.py          Console adapters backing the .bat flow
  gui_main.py     Desktop app entry (pywebview / Edge WebView2 window)
  gui_api.py      GUI bridge; task_coordinator.py owns task state; gui_worker*.py = background workers
  ui/             The interface itself (plain HTML/CSS/JS, no build step)
  export_*.py     One thin file per report type (a ReportSpec)
  consolidate_*.py  Combine per-route exports into one workbook
  compare_core.py   The correctness-locked comparison-workbook engine (compare_*.py = its report families)
  matrix*.py, *_matrix.py  The comparison matrices (Everything, by day, vs baseline, PDF vs Excel, three days, ArcGIS)
  tsn_library.py  The TSN library; arcgis_*.py = the ArcGIS layer library and the reports built from it
  updater.py      The in-app updater
build/          Reproducible PyInstaller build (build.ps1, app.spec, prune, self-test) + the check_*.py suite
output/         Generated: dated run folders (per-report folders + consolidated/), comparisons/, run_reports/
tsn_library/    Imported TSN files (local only, git-ignored)
arcgis_layers/  The ArcGIS layer library (local only, git-ignored)
*.bat           Numbered launchers for the console workflow
docs/project-guide.md       Router → conventions + an index into the docs/ library
docs/           In-depth knowledge library (start at docs/INDEX.md)
```

A single report catalog (`scripts/report_catalog.py`) is the metadata source of
truth; `scripts/reports.py` derives the GUI and the multi-report selector from it,
so the two never drift. For the full design — the `Events`/`ReportSpec` seam, retry
logic, fast mode, and packaging — see [`docs/project-guide.md`](docs/project-guide.md) and the
[`docs/`](docs/INDEX.md) knowledge library.

## Tech stack

- **Python 3.11** (standard library + a handful of pinned runtime deps)
- **[Playwright](https://playwright.dev/python/)** — browser automation, driving
  the system Chrome/Edge or the optional Built-in Chromium
- **[pdfplumber](https://github.com/jsvine/pdfplumber)** — PDF parsing
  (consolidation), plus pypdfium2 and Pillow for reading and rendering PDF pages
- **[openpyxl](https://openpyxl.readthedocs.io/)** — Excel reading and writing
  (consolidation and comparison workbooks)
- **[pywebview](https://pywebview.flowrl.com/)** — desktop GUI shell (Edge
  WebView2 rendering a vanilla HTML/CSS/JS interface)
- **[PyInstaller](https://pyinstaller.org/)** — portable onefolder packaging

## Known limitations

- **The app is unsigned.** Windows SmartScreen / Defender and corporate DLP may
  flag it on first run. Unblock the zip before extracting (Properties → Unblock)
  and choose **More info → Run anyway**. Code-signing is the planned fix.
- **The standard build needs a browser installed.** It drives the machine's Chrome
  or Edge and does not bundle one — use the *with-browser* download for
  machines where that's a problem.
- **Long folder paths break saves.** Windows and Excel limit how long a file's
  full path can be, so keep the app in a short folder such as `C:\TSMIS`.
- **The ArcGIS tab needs ArcGIS Pro** installed and signed in on the PC; the rest
  of the app works without it.
- **One job at a time.** An export, consolidation or comparison runs on its own —
  wait for the running one, or cancel it.

## Contributing

Issues and pull requests are welcome. Start with
[`CONTRIBUTING.md`](CONTRIBUTING.md) (setup, the one-command check suite, the
three rules people trip on), then [`docs/project-guide.md`](docs/project-guide.md) — it documents the
architecture and the conventions that keep the codebase maintainable, notably:

- The core engine is **console-free** — report progress via the `Events` sink and
  raise exceptions; never `print`/`input`/`sys.exit` in core modules.
- User-facing messages must be **UI-neutral** (shown in both the console and GUI).
- Adding a report type or consolidator follows a documented recipe.

Keep commit messages short and imperative (e.g. `add route 395`).

## License & disclaimer

Released under the [MIT License](LICENSE). Bundled third-party components retain
their own licenses.

This is an **unofficial** utility that automates exporting reports a user is
already authorized to access. It is not affiliated with, endorsed by, or
supported by Caltrans or the State of California.
