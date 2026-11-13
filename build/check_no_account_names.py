"""Guard: no person's account name in anything git would commit.

Every report prints the account that asked for it. The TSMIS site prints the
signed-in e-mail address on each PDF cover page, and a TSN print carries a
mainframe user ID -- both on a "SUBMITTOR : ..." line. An audit script, a
fixture or a note written from a real export can therefore carry a person's
account into this public repository, and the release's source zip publishes
everything git tracks.

This check FAILS, in the files git tracks or would add, on:

  * any e-mail-shaped string;
  * a value printed after a SUBMITTOR label ("SUBMITTOR : ...");
  * a quoted submitter value ('"submitter": "..."');
  * a user ID named in prose ("submitted by ...", "submitter ..."),

unless the value is one of the made-up placeholders in ALLOWED. Local data git
ignores (output\\, the saved sign-in state) is never read. A failure prints the
file and line with the value masked, so the check's own log cannot leak what it
found.

Where a pin needs the account, match its shape or read it from the bound source
at run time -- see SUBMITTER_ADDRESS and TSN_SUBMITTER_ID in
build\\phase8_intersection_summary_comparison.py.

Without a git checkout (a source zip) there is nothing to ask, and the check
passes with a note.

Run with the build venv:
    build\\.venv\\Scripts\\python.exe build\\check_no_account_names.py
"""
import os
import re
import subprocess
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

ADDRESS = re.compile(
    r"(?<![A-Za-z0-9._%+-])[A-Za-z0-9._%+-]+@"
    r"[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}(?![A-Za-z0-9-])")
PRINTED_LINE = re.compile(r"SUBMITTOR\s*:\s*([A-Za-z0-9][A-Za-z0-9._%+@-]*)")
QUOTED_VALUE = re.compile(r"""["']submitter["']\s*:\s*["']([^"'\n]+)["']""", re.I)
IN_PROSE = re.compile(r"(?i:submitted\s+by|submitter)\s+`?([A-Z][A-Z0-9]{5,7})\b")

# (rule, what a failure calls it). ADDRESS masks the name part only.
RULES = (
    (ADDRESS, "an e-mail address"),
    (PRINTED_LINE, "a value on a SUBMITTOR line"),
    (QUOTED_VALUE, "a quoted submitter value"),
    (IN_PROSE, "a user ID named in prose"),
)

# Made-up values only, each with the reason it is here. Never a real account.
ALLOWED = {
    # Cover-page fixtures in the Stage-8 audit gates.
    "first.last@dot.ca.gov",
    # TSN print fixtures: a mainframe ID's shape with every letter masked, and the
    # plain word three TSN checks use for the field.
    "trxxxxxx",
    "test",
    # check_updater.py: a URL that hides its real host behind a user-info part.
    "github.com@evil.test",
}

_SNIFF = 8192
# This file's own self-test lines are shaped like account names on purpose.
_SELF = "build/" + os.path.basename(__file__)


def _git_files():
    """Paths git tracks or would add, or None where git cannot say."""
    try:
        got = subprocess.run(
            ["git", "-C", REPO, "ls-files", "-z", "--cached", "--others",
             "--exclude-standard"],
            capture_output=True, timeout=120)
    except (OSError, subprocess.SubprocessError):
        return None
    if got.returncode != 0:
        return None
    return [p for p in got.stdout.decode("utf-8", "replace").split("\0") if p]


def _masked(rule, value):
    if rule is ADDRESS:
        return "****@" + value.split("@", 1)[1]
    return "****"


def line_offenders(line):
    """(what, masked value) for each account name one line of text carries."""
    found = []
    for rule, what in RULES:
        for match in rule.finditer(line):
            value = match.group(1) if rule.groups else match.group(0)
            if value.lower() not in ALLOWED:
                found.append((what, _masked(rule, value)))
    return found


def find_offenders(paths):
    offenders = []
    for rel in paths:
        if rel.replace("\\", "/") == _SELF:
            continue
        try:
            with open(os.path.join(REPO, rel), "rb") as fh:
                raw = fh.read()
        except OSError:
            continue                       # silent-ok: deleted in the worktree, still in the index
        if b"\0" in raw[:_SNIFF]:
            continue                       # a binary file (png, ico)
        for lineno, line in enumerate(raw.decode("utf-8", "replace").splitlines(), 1):
            for what, shown in line_offenders(line):
                offenders.append((rel, lineno, what, shown))
    return offenders


def self_test():
    """The rules catch what they are for and leave ordinary text alone."""
    caught = (
        "by someone.real@dot.ca.gov today",
        "SUBMITTOR : AB12345",
        "SUBMITTOR:someone.real@dot.ca.gov",
        '"submitter": "AB12345",',
        "reference 09/15/2025 · submitted by AB12345 · generated",
        "submitter `AB12345`, report title",
    )
    passed = (
        "SUBMITTOR : first.last@dot.ca.gov",
        "SUBMITTOR : TRXXXXXX",
        '"submitter": "TRXXXXXX",',
        '"submitter": submitter.group(1), "route": route,',
        r'"submitter": r"SUBMITTOR\s*:\s*(\S+)",',
        'rf"SUBMITTOR : ({SUBMITTER_ADDRESS.pattern})"',
        "the SUBMITTOR line, a submitter user ID, STATEWIDE",
        "· submitted by {ident.get('submitter')} · generated",
        "uses: actions/checkout@v4 and pkg@1.2.3 on tsmis.dot.ca.gov",
        "https://github.com@evil.test/x",
    )
    missed = [line for line in caught if not line_offenders(line)]
    wrong = [line for line in passed if line_offenders(line)]
    return missed, wrong


def main():
    missed, wrong = self_test()
    if missed or wrong:
        print(f"FAIL  account-name guard: its own rules are off "
              f"({len(missed)} not caught, {len(wrong)} wrongly caught).")
        sys.exit(1)
    paths = _git_files()
    if paths is None:
        print("OK  account-name guard: [SKIP] not a git checkout, so there is no "
              "list of files to commit.")
        return
    offenders = find_offenders(paths)
    if offenders:
        print(f"FAIL  account-name guard: {len(offenders)} account name(s) in files "
              f"git would commit (values masked):")
        for rel, lineno, what, shown in offenders:
            print(f"  {rel}:{lineno}: {what} ({shown})")
        print("  Remove it. If a check must pin an account, match its shape or read "
              "it from the source at run time; a made-up placeholder goes in ALLOWED "
              "with its reason.")
        sys.exit(1)
    print(f"OK  account-name guard: no account name in {len(paths)} files git would "
          f"commit (placeholders allowed: {len(ALLOWED)}).")


if __name__ == "__main__":
    main()
