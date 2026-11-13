TSMIS Reports Exporter -- Notes for IT / Security
=================================================

If you have come across this app on a managed PC and want to know what it is,
what it touches on disk, and what it does on the network, this page answers
that. It is written for the person reviewing or approving the tool. The
separate "Start Here.txt" in this same folder covers day-to-day use.


IN ONE PARAGRAPH
  This is an in-house productivity tool used by Caltrans staff to bulk-download
  TSMIS reports -- the same reports a person can already export by hand from
  the TSMIS website, one route at a time -- for all California state routes in
  one pass. It is a portable Windows app: no installer, no admin rights. It is
  just this folder of files, and it runs as the signed-in user. It signs in to
  TSMIS with the user's own Caltrans SSO account and saves the reports into
  this folder. The source is open and inspectable (see VERIFY IT YOURSELF).

  App:     TSMIS Exporter
  Version: shown in the app's title bar and under Settings > About
  Source:  https://github.com/EagleBearerYS/TSMIS-Reports-Exporter  (public)


WHAT IT DOES NOT DO
  * No administrator / elevation. The .exe ships an "asInvoker" manifest -- it
    runs as the current user and never prompts for UAC.
  * No installer, no system changes. It does not write to the registry, to
    C:\Windows, to Program Files, or to any other user's profile. "Uninstall"
    is deleting this folder.
  * No background presence. It installs no service, no scheduled task, no
    startup / Run-key entry, and no drivers. It runs only while its window is
    open.
  * No telemetry. No analytics, crash-reporting service, license server, or
    "phone home" of any kind. (All outbound traffic is listed below -- it is
    only TSMIS and, for update checks, GitHub.)
  * It does not collect the user's password. Sign-in happens in a normal
    browser window with full MFA; the app only keeps the resulting session,
    the same way a browser keeps you signed in. (See SAVED LOGIN.)
  * No inbound network service. Nothing listens for off-box connections. Two
    loopback-only (127.0.0.1) ports exist: the app's own window loads its
    interface from a small page server inside the app (port 42001, open while
    the app is open), and a short-lived debug port is used during interactive
    Edge sign-in -- see BROWSER AUTOMATION.
  * It does not scan, map, or probe the local network.


NETWORK CONNECTIONS IT MAKES
  Outbound only, and only these:

  1. The TSMIS site (tsmis.dot.ca.gov, or tsmis-dev.dot.ca.gov when the user
     selects the development site in Settings) -- on every sign-in and export.
     This is the whole point of the tool: load the report page and pull report
     data using the user's Caltrans SSO session. (The page itself fetches its
     data from a Caltrans intranet host.)
  2. The Caltrans sign-in provider (the ArcGIS portal / Azure AD identity
     provider) -- during sign-in only. Standard OAuth/SAML: the same hosts a
     browser hits to log in to TSMIS by hand.
  3. GitHub -- api.github.com and GitHub's release-download CDN, only to check
     for and download an app update (on launch, and when the version is
     clicked). Read-only, public repo, no credentials. See UPDATES.
  4. (Only on request) Playwright's Chromium download -- only if someone uses
     the optional "Download Built-in Chromium" action or the developer .bat
     setup. The normal app download never fetches a browser.
  5. (Only on request, and only on a PC with ArcGIS Pro) The Caltrans ArcGIS
     server (rhapps-prod.dot.ca.gov) -- when someone uses the ArcGIS tab's
     layer refresh. The app starts ArcGIS Pro's own Python, which reads the
     TSMIS map layers from that server with ArcGIS Pro's own sign-in. See
     OTHER PROGRAMS IT STARTS.

  There are no other destinations. The tool honors the system proxy and works
  through corporate TLS inspection, because its update check validates TLS
  against the Windows certificate store -- not a private bundled CA list.


DATA AND FILES
  Everything the tool keeps lives in THIS folder (or, if this folder is
  read-only, under %LOCALAPPDATA%\TSMIS Exporter). The only writes outside it
  are short-lived scratch files in the user's %TEMP% while a comparison runs,
  and files the user saves somewhere else on purpose through a Save dialog (a
  comparison workbook, a support bundle, a different "Export Everything"
  folder).

  * output\        -- the PDF / Excel reports it produces, the combined and
    comparison workbooks, and run-report CSVs. This is the tool's product.
  * tsn_library\   -- TSN reference files the user imports for comparisons.
  * arcgis_layers\ -- map layers exported through ArcGIS Pro (ArcGIS tab only).
  * data\          -- the app's own working data: rotating logs (plain text),
    the settings file, and the Edge WebView2 / sign-in browser profiles.
  * _internal\     -- the bundled runtime that lets the app run without
    installing anything: the Python interpreter, support DLLs, and the
    browser-automation driver (node.exe). Treat it as program files; it is
    not user data.

  The reports never leave the PC on their own. The only outbound traffic is the
  list in NETWORK CONNECTIONS above.


SAVED LOGIN (the one sensitive file)
  After sign-in the tool stores the session -- browser cookies, a Playwright
  "storage_state" file -- so the user need not log in every run.
  * It is the SESSION, not the password. The password and MFA are entered in a
    real browser window the tool never reads.
  * It lives only in this user's app folder, protected by normal NTFS file
    permissions. It is treated as a credential and is deliberately EXCLUDED
    from the app's diagnostic "support bundle" (that zip carries logs and
    settings only -- never the login or the browser profiles).
  * At-rest encryption of this file (Windows DPAPI) is a planned hardening
    item.


BROWSER AUTOMATION (two options that can look unusual, explained)
  The tool drives a browser to do what a person would do on the TSMIS site.
  Two non-default options are set, both scoped to the tool's OWN automated
  browser pages:
  * Local Network Access permission. The TSMIS page loads its data from a
    Caltrans *intranet* host; current Chromium would otherwise block that
    behind an "allow local network?" prompt that an automated run cannot
    click. This grants that one permission to the tool's own pages only. It
    does not scan, expose, or open the local network.
  * A loopback debug port (127.0.0.1, a random free port) -- used ONLY during
    the interactive Microsoft Edge sign-in, so the tool can re-attach if
    managed Edge relaunches itself into a work profile mid-login. It is bound
    to loopback, open only while that one sign-in window is up, and accepts no
    off-box connection.


UPDATES (why a self-updating .exe is safe here)
  The tool can update itself from its GitHub releases. The mechanics are
  deliberately conservative:
  * It only ever looks at this one public repo's published releases -- never
    prereleases or drafts.
  * It downloads the update and verifies its SHA-256 checksum against the
    published value BEFORE doing anything; a mismatch refuses the update.
  * It replaces only its own known program files (the .exe, the _internal
    runtime and the two text files in this folder). User data (output\,
    tsn_library\, arcgis_layers\, data\) is never part of an update and is
    never touched. A failed swap rolls back to the previous version.
  * If the folder is read-only (locked down), it cannot self-update at all --
    it just opens the release page in the browser.


KNOWN: THE .EXE IS NOT CODE-SIGNED YET
  This is in-house tooling and the executable is not yet Authenticode-signed,
  so on first run SmartScreen may say the publisher is unknown ("More info" ->
  "Run anyway"), and EDR / AV may look twice at a self-updating exe.
  Code-signing is the planned fix and will settle both. Until then, the
  controls above -- checksum-verified updates, known-file-only swap, no admin,
  local-only data -- stand in for it. The build is also deliberately NOT
  UPX-packed and carries proper version / icon / manifest metadata, precisely
  to avoid antivirus false positives.


OTHER PROGRAMS IT STARTS
  * A web browser already on the PC (Google Chrome or Microsoft Edge), or the
    app's own Chromium where that variant is used -- to sign in and to export.
    It is driven by the bundled automation driver (node.exe under _internal\).
  * A second "TSMIS Exporter.exe" under the first, while a comparison or build
    runs -- see VERIFY IT YOURSELF.
  * ArcGIS Pro's own Python (python.exe under the ArcGIS Pro install), only
    when the ArcGIS tab's layer refresh or "Check ArcGIS Pro" is used, and
    only if ArcGIS Pro is installed. It is started directly, with no shell
    and no window, runs one script shipped in _internal\, and exits when the
    refresh ends.
  Nothing is started with a command shell, PowerShell, or elevation.


VERIFY IT YOURSELF
  You do not have to take this page's word for any of it:
  * The source is open. The full code is in the GitHub repo listed at the top;
    every network call, file path, and browser flag described here is in it.
  * It is just a folder. Inspect _internal\, read the logs under data\logs\
    (plain text), and open the files in output\ directly.
  * Watch it live. Resource Monitor / netstat will show the outbound
    connections above and nothing else; Process Explorer will show it running
    as the user, with no service and no driver.
  * A second "TSMIS Exporter.exe" under the first is expected while a
    comparison or build runs. The app starts a copy of itself to do that work
    so its window stays responsive. It is the same program, running as the
    same user with the same files, and it exits when the job ends. If policy
    blocks an app from starting itself, nothing breaks: the job runs inside
    the main process instead and the log says so.
  * The diagnostic "support bundle" the app can produce (logs + settings, no
    credentials) is a quick way to see exactly what it has been doing.


QUESTIONS
  Contact the tool's maintainer -- the person who gave you this copy. Deeper
  technical detail ships with the source, in the repo's docs.
