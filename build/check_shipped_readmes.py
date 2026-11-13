"""The two text files that ship beside the exe must describe the app that ships.

`build/dist_readme.txt` (shipped as "Start Here.txt") and `build/it_readme.txt`
("IT-README.txt") are the first things a new user and a reviewing IT person
read, and nothing held them to the code: by v0.52.1 the quick start still
described a four-report app with a browser dropdown that had been gone for
thirty-five releases, and the IT notes named a retired folder, a version from
v0.14 and the repository's old address.

This check cannot judge prose. It holds the parts that rot:

  * every button or label the quick start QUOTES still exists in the UI source,
    so renaming a button without updating the guide fails here;
  * every tab is named, and nothing retired is;
  * its example folder is the name the app really gives a run folder;
  * the IT notes name the real update source, every outbound host the code
    knows, and every data folder -- and carry no hard-coded version;
  * both files stay plain ASCII and within 80 columns (they open in Notepad).

Run with the build venv:
    build\\.venv\\Scripts\\python.exe build\\check_shipped_readmes.py
"""
import re
import sys
from pathlib import Path
from urllib.parse import urlsplit

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT / "scripts"), str(ROOT)]

import arcgis_refresh  # noqa: E402
import paths  # noqa: E402
import site_target  # noqa: E402
import updater  # noqa: E402

_fail = []


def check(name, cond):
    print(f"  [{'OK ' if cond else 'FAIL'}] {name}")
    if not cond:
        _fail.append(name)


START = (ROOT / "build" / "dist_readme.txt").read_text(encoding="utf-8")
IT = (ROOT / "build" / "it_readme.txt").read_text(encoding="utf-8")
UI_DIR = ROOT / "scripts" / "ui"
UI = "\n".join(p.read_text(encoding="utf-8")
               for p in sorted(UI_DIR.glob("*"))
               if p.suffix in (".html", ".js") and p.name != "mock.js")
# The window also shows strings the bridge sends (status texts, dialogs).
BRIDGE = "\n".join(p.read_text(encoding="utf-8")
                   for p in sorted((ROOT / "scripts").glob("gui_*.py")))

# What the quick start tells a reader to click or look for, exactly as written
# there. Each must exist in what the window can show.
QUOTED = [
    "Log in", "I've finished logging in", "Session ready",
    "Automatic sign-in ready (no saved login needed)",
    "Start export", "Skip route", "Pause", "Cancel", "Open run folder",
    "Retry failed", "Start consolidation",
    "Consolidate each report when its export finishes",
    "Values workbook", "Live-formulas workbook", "Matrix options",
    "Add all", "Remove all", "Build all", "Use development site",
    "Use main site", "TSMIS site addresses", "TSN reports", "Rebuild",
    "Rebuild all out of date", "Delete all reports", "Check for updates",
    "Restart to update", "Failure screenshots", "Output folder", "Logs",
    "Refresh & export", "Comparison matrix",
]


def _normal(text):
    """Typographic punctuation and HTML entities folded to what a plain-text
    file can say, so the guide's ASCII quote matches the window's real label."""
    for a, b in (("…", "..."), ("—", "--"), ("’", "'"),
                 ("“", '"'), ("”", '"'), ("&amp;", "&")):
        text = text.replace(a, b)
    # A quoted label may wrap across two lines of the guide (or of the HTML).
    return re.sub(r"\s+", " ", text)


def test_plain_text():
    print("both files are plain ASCII, 80 columns at most:")
    for name, text in (("Start Here", START), ("IT-README", IT)):
        bad = sorted({c for c in text if ord(c) > 126 or (ord(c) < 32 and c not in "\n\r")})
        check(f"{name}: ASCII only" + (f" (found {bad[:5]!r})" if bad else ""), not bad)
        wide = [i for i, line in enumerate(text.splitlines(), 1) if len(line) > 80]
        check(f"{name}: no line over 80 columns" + (f" (lines {wide[:5]})" if wide else ""),
              not wide)


def test_quick_start_matches_the_window():
    print("the quick start names what the window really shows:")
    shown = _normal(UI + "\n" + BRIDGE)
    guide = _normal(START)
    for label in QUOTED:
        check(f"it quotes {label!r}, and the window still has it",
              f'"{label}' in guide and label in shown)
    tabs = re.findall(r'id="tab(?:Export|Consolidate|Compare|Everything|Arcgis)"[^>]*>'
                      r'(?:\s*<svg.*?</svg>)?\s*([A-Za-z]+)', UI, re.S)
    check(f"the five named tabs are read from the page ({', '.join(tabs)})", len(tabs) == 5)
    for tab in tabs:
        check(f"it has a section for the {tab} tab", tab.upper() in START.upper())
    subtabs = re.findall(r'"(vs TSN Matrix|vs Baseline Matrix|PDF vs Excel Matrix|'
                         r'Three Days Matrix)"', UI)
    for sub in sorted(set(subtabs)):
        check(f"it names the {sub}", sub in START)
    check("all four Compare matrices are found in the page", len(set(subtabs)) == 4)


def test_quick_start_has_nothing_retired():
    print("the quick start names nothing that is gone:")
    for gone, why in (("Browser dropdown", "the title bar only shows what will export"),
                      ("input\\", "the input folder was retired in v0.30.0"),
                      ("output\\2026-06-10\\", "run folders carry source, environment and site"),
                      ("Reset to production", "the button is Use main site")):
        check(f"no {gone!r} ({why})", gone not in START)
    example = paths.run_folder_name("ssor", "prod", "2026-11-13", site_target.HOST_MAIN)
    check(f"its example folder is a real run-folder name ({example})",
          f"output\\{example}\\highway_log\\" in START)
    check("it names both TSMIS sites by their real hosts",
          site_target.TSMIS_HOST in START and site_target.TSMIS_DEV_HOST in START)


def test_it_notes_match_the_code():
    print("the IT notes match the code:")
    check(f"the source is the repository the updater asks ({updater.GITHUB_REPO})",
          f"github.com/{updater.GITHUB_REPO}" in IT)
    check("no hard-coded version (it went stale at 0.14.2)",
          not re.search(r"^\s*Version:\s*\d", IT, re.M))
    arcgis_host = urlsplit(arcgis_refresh.SERVICE_DEFAULT).hostname
    for host in (site_target.TSMIS_HOST, site_target.TSMIS_DEV_HOST,
                 "api.github.com", arcgis_host):
        check(f"outbound host {host} is listed", host in IT)
    for folder in ("output\\", "tsn_library\\", "arcgis_layers\\", "data\\", "_internal\\"):
        check(f"data folder {folder} is listed", folder in IT)
    check("the retired input folder is not listed", "input\\" not in IT)
    for item in updater._BUNDLE_ITEMS:
        if item.endswith(".txt"):
            check("an update replacing the two text files is disclosed",
                  "two text files" in IT)
            break
    import webview
    port = str(webview.settings.get("DEFAULT_HTTP_PORT", ""))
    check(f"the window's own loopback page server is disclosed (port {port})",
          bool(port) and port in IT and "127.0.0.1" in IT)


def main():
    test_plain_text()
    test_quick_start_matches_the_window()
    test_quick_start_has_nothing_retired()
    test_it_notes_match_the_code()
    if _fail:
        print(f"\nFAILED: {len(_fail)} check(s): {_fail}")
        return 1
    print("\nAll shipped-readme checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
