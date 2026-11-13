TSMIS Reports Exporter
======================

WHAT IT DOES
  Downloads TSMIS reports for every California state route in one run,
  combines the per-route files into one workbook, and compares the results:
  against TSN, between environments, between export days, PDF against Excel,
  and against the ArcGIS layers.

  Reports (each one as an Excel file and as a print-layout PDF):
    - Highway Log, Highway Sequence Listing
    - Ramp Summary, Ramp Detail
    - Intersection Summary, Intersection Detail
    - Highway Summary, Highway Detail
    - Clean Road: Highway, Intersection, Ramp

HOW TO RUN
  Double-click  "TSMIS Exporter.exe"  in this folder.
  Keep this whole folder together -- the app needs the "_internal" folder next
  to the .exe. Don't move the .exe out on its own.
  Nothing to install; you don't need Python. Windows 10 or 11.
  * If you received this as a .zip: right-click the zip -> Properties -> tick
    "Unblock" -> OK, BEFORE extracting it. Extract into a folder you can write
    to, and keep its path SHORT -- C:\TSMIS is ideal. Windows and Excel limit
    how long a file's full path can be, and a deep folder (a OneDrive Desktop,
    for example) can make exports and comparisons fail to save or open. The
    app says so in its Activity log at start-up if the path is long.
  * The first time, Windows may say the publisher is unknown. That is expected
    -- the app is not code-signed. Choose "More info" -> "Run anyway".
  * The standard download uses the Google Chrome or Microsoft Edge already on
    this PC; the "with-browser" download also brings its own browser.
  * IT or security questions: see "IT-README.txt" in this folder.

FIRST TIME: SIGN IN
  The status at the top reads "No saved login -- click Log in".
  1. Click  "Log in".
  2. A browser window opens. Sign in with your Caltrans account and complete
     MFA, until the TSMIS report page loads.
  3. Click  "I've finished logging in"  (the same button) or just close the
     browser window -- your session is saved either way.
  The status dot turns green and reads "Session ready". If the saved login
  later stops working, the app tells you to click "Log in" again. Your
  password goes into the browser, never into the app; the app keeps only the
  signed-in session, in its "data" folder.
  On a Caltrans PC you can also just start an export: the app first tries to
  sign in by itself with Microsoft Edge and your Windows account, and asks
  you to log in only if that does not work. When it works, the status reads
  "Automatic sign-in ready (no saved login needed)". Click the status to see
  both ways ("Saved login" and "Edge one-click"). Fast mode needs a saved
  login.

EXPORT REPORTS
  Under the title, "Data source" (SSOR / ARS) and "Environment" (Prod / Test /
  Dev) choose what you export. They start on SSOR and Prod.
  On the Export tab:
    - Tick ONE OR MORE reports (they run one after another); "All" and "None"
      tick or clear the whole list. Tick both the Excel and the PDF version
      of a report and it is generated once per route and saved both ways.
    - Routes: leave blank for ALL routes, or type a few (e.g. 5, 99, 101) or
      use "Choose...".
    - Click  "Start export".
  Files go into a folder named for the day, the data source and environment,
  and the TSMIS site, with one folder per report inside:
      output\<date> <source>-<environment> <site>-site\<report>\
      e.g.  output\2026-11-13 ssor-prod main-site\highway_log\
  The card on the right shows the exact folder before you start.
  Already-downloaded routes are skipped, so you can stop and re-run the same
  day to resume. "Skip route" moves past one slow route, "Pause" holds the
  run between routes, "Cancel" stops it. When a run ends, the summary has
  "Open run folder", "Retry failed" and "Save run report..." (a CSV of how
  each route turned out).
  Fast mode (optional): runs several browsers at once -- faster, but heavier
  on this PC; 3-4 is a good number. "Skip route" is off in fast mode.
  A route can fail with a "TSMIS site error". That comes from the TSMIS
  website, not this app; it is recorded as failed and the run continues.

WHICH TSMIS SITE (MAIN OR DEVELOPMENT)
  A fresh copy uses the main site, tsmis.dot.ca.gov ("main-site" in folder
  names). New reports are released first on the development site,
  tsmis-dev.dot.ca.gov ("dev-site"), and move to the main site once they are
  approved -- so a report can be greyed out or fail on the main site and
  still export from the development site. To switch: Settings -> "TSMIS site
  addresses" -> "Use development site"; "Use main site" switches back.
  This is separate from the "Environment" picker. If you are not sure which
  site to use, ask the person who gave you this copy.

COMBINE FILES (CONSOLIDATE)
  The Consolidate tab combines one report's per-route files into a single
  workbook. Pick the report and the "Export day" (newest first), then click
  "Start consolidation". The workbook is saved under "consolidated" in that
  day's folder, e.g.  output\2026-11-13 ssor-prod main-site\consolidated\
  Or tick "Consolidate each report when its export finishes" on the Export
  tab (PDF versions, except Ramp Summary, still need the Consolidate tab).

COMPARE
  The Compare tab has seven sub-tabs. The first three compare two things you
  pick, and ask where to save the result:
    - Cross-environment: one report from two export folders -- for example
      two environments, or the development site against the main site.
    - vs TSN: a report's consolidated workbook against its TSN copy. Pick an
      "Export day" (only days already consolidated are listed).
    - Self-consistency: a report's PDF version against its Excel version.
  Tick "Values workbook" (opens instantly), "Live-formulas workbook"
  (recalculates in Excel), or both. Differences are in red. A big
  live-formulas workbook opens with blank cells: press F9 in Excel and wait.

  The other four are matrices. A matrix is a grid: each row is a report,
  each column is an export day, each cell is one comparison. Under "Matrix
  options" you add the days and choose the reports: click a report to take
  it off the grid or put it back, or use "Add all" / "Remove all". Point at
  a cell to see its buttons: build the comparison, then open the workbook
  ("Build all" builds every cell). A built cell shows how many cells differ
  (green = identical). A cell that cannot be built says why ("needs export",
  "needs TSN").
  TSN is the older system TSMIS replaced, and its data stopped changing in
  September 2025 -- so most differences against TSN are things that changed
  in TSMIS since then, not errors.
    - vs TSN Matrix: each day's export against TSN. "Export today" exports
      every report on the matrix for today, then compares.
    - vs Baseline Matrix: each day against a baseline, such as an earlier day.
    - PDF vs Excel Matrix: each day's PDF version against its Excel version.
    - Three Days Matrix: the Clean Road reports on three export days at once.

EVERYTHING TAB
  "Refresh & export" exports every ticked report for every ticked
  environment into ONE folder that is overwritten each time, so it always
  holds the latest copy (default: output\All Reports (current)). You can
  pause it and pick it up later, even after closing the app. "Comparison
  matrix" is a grid of reports by environments, built from that folder.

ARCGIS TAB (ONLY IF YOU HAVE ARCGIS PRO)
  Checks reports against the TSMIS map layers: it pulls the layers with
  ArcGIS Pro, builds reports from them, and compares your exports against
  those. It needs ArcGIS Pro installed and signed in on this PC. If you do
  not have ArcGIS Pro, ignore this tab.

TSN FILES (NEEDED FOR EVERY "vs TSN" COMPARISON)
  The app does not download TSN data -- you add the TSN files once. Go to
  Settings -> "TSN reports", click "Import raw..." next to a report, pick
  its TSN file(s), then click "Rebuild". A green dot means that report is
  ready. After an app update some may show as out of date: click "Rebuild
  all out of date". The files are kept in the "tsn_library" folder, one
  folder per report. Ask the person who gave you this copy for the TSN
  files.

SETTINGS WORTH KNOWING (THE GEAR TAB)
  * Reliability: time limits per route. Raise them if the TSMIS server is
    slow and routes keep timing out.
  * TSMIS site addresses: "Use development site" -- see WHICH TSMIS SITE.
  * Data & storage: shows where everything is saved. "Delete all reports..."
    lists what it will remove and asks first; it deletes every exported,
    consolidated and comparison file. Logs, your login, your settings and
    the TSN files stay.
  * Updates: when a newer version is out, an "Update to v..." button appears
    in the title bar. Click it, then "Restart to update" -- your reports,
    login and settings are kept. Settings also has "Check for updates".

WHERE THINGS ARE (IN THIS FOLDER)
  output\         exports, consolidated workbooks, comparisons, run reports
  tsn_library\    your TSN files
  arcgis_layers\  the ArcGIS layers (ArcGIS tab only)
  data\           logs (data\logs), settings, your saved login
  _internal\      the program itself -- leave it alone
  The "Output folder" button at the bottom of the window opens "output". If
  this folder is read-only, the app saves under %LOCALAPPDATA%\TSMIS Exporter
  instead (the bottom of the window shows the path) and cannot update itself.

IF SOMETHING GOES WRONG
  The app does one job at a time: wait for the running one, or cancel it.
  * Best: Settings -> Debugging -> "Save support bundle...". It saves one zip
    with the logs, run reports and settings. It includes this PC's name and
    folder paths, but never your login. Send that zip to the person who gave
    you this copy.
  * Or click "Logs" at the bottom of the window and send those files.
  * Say what you clicked and what happened. For a failed route, Settings ->
    "Failure screenshots" opens pictures of the page when it failed.
