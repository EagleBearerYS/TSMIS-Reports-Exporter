"""The lite copy of a comparison workbook (v0.55.0) — the version the owner
delivers, made from a finished values workbook by `lite_copy`.

Locks, on REAL values workbooks built by compare_core.run_compare (with a Notes
legend and a Provenance sheet, as the Clean Road comparisons carry):

  * the sheets: the full list minus Spot Check, Notes and Provenance, in order,
    every state kept (the very-hidden snapshot sheets stay very hidden);
  * the Summary: the title without its 'Discrepancy Report (…)' tail, then the
    ROW COUNTS / MATCH STATUS / ROUTE COVERAGE rows exactly as the full Summary
    has them, one blank row, the DIFFERENCES BY FIELD table, nothing after, and
    no conditional format left (it coloured the verdict line);
  * every other part is the values workbook's own compressed bytes;
  * the filter names still point at their own sheets, every CRC reads back and
    openpyxl opens the copy; a flat (no-route) comparison works the same;
  * refusals — no Summary, a Summary without its blocks, a lite path that is
    the workbook itself, a guard that says no — leave no file behind;
  * the zip64-capable fallback writes the same parts;
  * the matrix step (`matrix._settle_lite_copy`) writes the copy beside a
    published comparison, removes a stale one when the option is off, writes
    none for a partial comparison (the notes that say what is missing would be
    lost) and never writes over, or deletes, a comparison input.

    build\\.venv\\Scripts\\python.exe build\\check_lite_copy.py
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.dom import minidom

_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_ROOT / "scripts"))

import artifact_store  # noqa: E402
import lite_copy  # noqa: E402
import matrix  # noqa: E402
from compare_core import CompareSchema, run_compare  # noqa: E402
from events import Events  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402

FAILS = []
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REWRITTEN = {"xl/worksheets/sheet1.xml", "xl/workbook.xml",
             "xl/_rels/workbook.xml.rels", "[Content_Types].xml"}


def check(label, ok, detail=""):
    print(f"  {'ok' if ok else 'FAIL'}: {label}" + (f"\n      {detail}" if detail and not ok else ""))
    if not ok:
        FAILS.append(label)


def _notes(wb):
    ws = wb.create_sheet("Notes")
    ws.append(["How to read this workbook"])


ROUTED = CompareSchema(report_name="Clean Road Check", header=["PM", "County", "Note", "Desc"],
                       side_a="TSMIS", side_b="TSN", id_noun="segment",
                       id_noun_plural="segments", context_fields=("Note",),
                       legend_writer=_notes)
ROWS_A = [["001", "1.000", "ALA", "n", "A"], ["001", "2.000", "ALA", "n", "B"],
          ["001", "3.000", "ALA", "x", "C"], ["002", "1.000", "SF", "n", "D"]]
ROWS_B = [["001", "1.000", "ALA", "n", "A"], ["001", "2.000", "ALA", "n", "Z"],
          ["001", "3.000", "ALA", "y", "C"], ["003", "1.000", "LA", "n", "E"]]
FLAT = CompareSchema(report_name="Flat Check", header=["Key", "F1"], side_a="A",
                     side_b="B", id_noun="row", id_noun_plural="rows")
PROVENANCE = {"recipe": {"report": "Clean Road Check", "banner": "TSMIS vs TSN"},
              "inputs": [{"role": "TSMIS", "kind": "file", "selection": "C:/x/a.xlsx",
                          "sha256": "0" * 64}]}


# --------------------------------------------------------------------------- #
# reading a workbook without openpyxl's help
# --------------------------------------------------------------------------- #
def _parts(path):
    with zipfile.ZipFile(path) as z:
        wb = minidom.parseString(z.read("xl/workbook.xml"))
        rels = minidom.parseString(z.read("xl/_rels/workbook.xml.rels"))
        target = {r.getAttribute("Id"): r.getAttribute("Target").lstrip("/")
                  for r in rels.getElementsByTagName("Relationship")}
        sheets = [(s.getAttribute("name"), s.getAttribute("state") or "visible",
                   target[s.getAttributeNS(REL, "id")])
                  for s in wb.getElementsByTagNameNS(MAIN, "sheet")]
        names = [(d.getAttribute("name"), d.getAttribute("localSheetId"),
                  "".join(n.data for n in d.childNodes))
                 for d in wb.getElementsByTagNameNS(MAIN, "definedName")]
        summary = next(p for n, _s, p in sheets if n == "Summary")
        doc = minidom.parseString(z.read(summary))
    rows = {}
    for row in doc.getElementsByTagNameNS(MAIN, "row"):
        vals = {}
        for c in row.getElementsByTagNameNS(MAIN, "c"):
            col = re.match(r"[A-Z]+", c.getAttribute("r")).group(0)
            if c.getAttribute("t") == "inlineStr":
                vals[col] = "".join(t.firstChild.data if t.firstChild else ""
                                    for t in c.getElementsByTagNameNS(MAIN, "t"))
            else:
                f = c.getElementsByTagNameNS(MAIN, "f")
                v = c.getElementsByTagNameNS(MAIN, "v")
                vals[col] = ("=" + f[0].firstChild.data if f and f[0].firstChild
                             else v[0].firstChild.data if v and v[0].firstChild else "")
        if any(vals.values()):
            rows[int(row.getAttribute("r"))] = (vals.get("B", ""), vals.get("C", ""),
                                                vals.get("D", ""))
    cfs = doc.getElementsByTagNameNS(MAIN, "conditionalFormatting").length
    return sheets, names, rows, cfs


def _raw(path):
    """{member: its compressed bytes} of a zip."""
    out = {}
    with open(path, "rb") as f, zipfile.ZipFile(f) as z:
        for info in z.infolist():
            f.seek(info.header_offset)
            head = f.read(30)
            n = int.from_bytes(head[26:28], "little")
            e = int.from_bytes(head[28:30], "little")
            f.seek(info.header_offset + 30 + n + e)
            out[info.filename] = f.read(info.compress_size)
    return out


def _expected_lite_rows(full_rows):
    order = sorted(full_rows)
    start = next(r for r in order if full_rows[r][0] == "ROW COUNTS")
    until = next(r for r in order if full_rows[r][0].startswith("FIELD-LEVEL DISCREPANCIES"))
    table = next(r for r in order if full_rows[r][0] == "DIFFERENCES BY FIELD")
    end = table
    while end + 1 in full_rows:
        end += 1
    kept = [full_rows[r] for r in range(start, until)]
    out = {2 + i: row for i, row in enumerate(kept)}
    first = 2 + len(kept) + 1
    out.update({first + i: full_rows[r] for i, r in enumerate(range(table, end + 1))})
    return out


def _filter_names(sheets, names):
    """{(sheet the filter belongs to, its range)} of a workbook's filter names."""
    return {(sheets[int(local)][0], text) for name, local, text in names
            if name == "_xlnm._FilterDatabase"}


# --------------------------------------------------------------------------- #
def check_routed(tmp):
    print("a routed values workbook (Notes + Provenance sheets, as Clean Road has):")
    full = tmp / "routed.xlsx"
    res = run_compare(ROUTED, ROWS_A, ROWS_B, True, full, mode="values",
                      confirm_overwrite=lambda _p: True, provenance=PROVENANCE)
    assert res.status == "ok", res.message
    lite = lite_copy.write_lite_copy(full)
    check("the lite copy is '<name> (lite).xlsx' beside the workbook",
          lite == tmp / "routed (lite).xlsx" and lite.exists())
    f_sheets, f_names, f_rows, f_cfs = _parts(full)
    l_sheets, l_names, l_rows, l_cfs = _parts(lite)
    check("the fixture carries the three sheets a lite copy leaves out",
          {"Spot Check", "Notes", "Provenance"} <= {n for n, _s, _p in f_sheets})
    want = [(n, s) for n, s, _p in f_sheets if n not in lite_copy.DROPPED_SHEETS]
    check("the lite sheets are the full list minus Spot Check, Notes and Provenance, "
          "in order, states kept",
          [(n, s) for n, s, _p in l_sheets] == want, f"{l_sheets}")
    check("...the snapshot sheets stay very hidden",
          [s for n, s, _p in l_sheets if n.startswith("__CMP_E2_SNAPSHOT")]
          == ["veryHidden", "veryHidden"])
    check("the title loses its 'Discrepancy Report (…)' tail",
          l_rows.get(1, ("",))[0] == "TSMIS vs TSN — Clean Road Check"
          and f_rows[2][0].startswith("TSMIS vs TSN — Clean Road Check — Discrepancy Report"),
          f"{l_rows.get(1)}")
    expected = _expected_lite_rows(f_rows)
    got = {r: v for r, v in l_rows.items() if r != 1}
    check("the Summary keeps the counts blocks and the per-field table exactly, "
          "one blank row between, nothing after", got == expected,
          f"got {got}\n      want {expected}")
    check("...the counts blocks start at row 2 and the per-field table "
          "follows one blank row",
          l_rows.get(2, ("",))[0] == "ROW COUNTS"
          and min(r for r, v in l_rows.items() if v[0] == "DIFFERENCES BY FIELD") - 2
          == max(r for r in l_rows if r < min(r2 for r2, v in l_rows.items()
                                              if v[0] == "DIFFERENCES BY FIELD")))
    check("...and has no formula left", not any(
        str(c).startswith("=") for v in l_rows.values() for c in v))
    check("the verdict line's conditional format goes with it",
          f_cfs == 1 and l_cfs == 0, f"full {f_cfs}, lite {l_cfs}")
    full_raw, lite_raw = _raw(full), _raw(lite)
    kept = [m for m in lite_raw if m not in REWRITTEN]
    check("every other part is the values workbook's own compressed bytes",
          kept and all(lite_raw[m] == full_raw[m] for m in kept)
          and set(lite_raw) - set(full_raw) == set())
    check("only the Summary, the workbook part, its rels and the content types "
          "are rewritten", REWRITTEN <= set(lite_raw)
          and all(lite_raw[m] != full_raw[m] for m in REWRITTEN))
    gone = set(full_raw) - set(lite_raw)
    dropped_parts = {p for n, _s, p in f_sheets if n in lite_copy.DROPPED_SHEETS}
    check("the three sheets' parts are gone and nothing else",
          gone == dropped_parts, f"{sorted(gone)} vs {sorted(dropped_parts)}")
    with zipfile.ZipFile(lite) as z:
        ct = z.read("[Content_Types].xml").decode()
        check("the content types no longer name them",
              not any(f'"/{p}"' in ct for p in dropped_parts))
        check("every CRC reads back", z.testzip() is None)
    check("the filter names still point at their own sheets",
          _filter_names(l_sheets, l_names)
          == {fn for fn in _filter_names(f_sheets, f_names)
              if fn[0] not in lite_copy.DROPPED_SHEETS}
          and _filter_names(l_sheets, l_names))
    wb = load_workbook(lite, read_only=True, data_only=True)
    try:
        check("openpyxl opens it and reads the title",
              wb.sheetnames == [n for n, _s in want]
              and wb["Summary"]["B1"].value == "TSMIS vs TSN — Clean Road Check")
    finally:
        wb.close()
    again = tmp / "again (lite).xlsx"
    lite_copy.write_lite_copy(full, again)
    check("writing it twice gives the same bytes", again.read_bytes() == lite.read_bytes())
    return full, lite


def check_flat(tmp):
    print("a flat (no-route) values workbook:")
    full = tmp / "flat.xlsx"
    res = run_compare(FLAT, [["k1", "x"], ["k2", "y"]], [["k1", "x"], ["k2", "z"]],
                      False, full, mode="values", confirm_overwrite=lambda _p: True)
    assert res.status == "ok", res.message
    lite = lite_copy.write_lite_copy(full)
    _fs, _fn, f_rows, _fc = _parts(full)
    _ls, _ln, l_rows, _lc = _parts(lite)
    check("no ROUTE COVERAGE block, and the rest is kept exactly",
          {r: v for r, v in l_rows.items() if r != 1} == _expected_lite_rows(f_rows)
          and not any(v[0].startswith("ROUTE COVERAGE") for v in l_rows.values()))
    check("its title is 'A vs B — Flat Check'", l_rows[1][0] == "A vs B — Flat Check")


def _refused(fn, fragment):
    try:
        fn()
    except ValueError as e:
        return fragment in str(e)
    return False


def check_refusals(tmp, full):
    print("refusals leave nothing behind:")
    plain = tmp / "plain.xlsx"
    wb = Workbook()
    wb.active.title = "Comparison"
    wb.save(plain)
    check("a workbook with no Summary is refused",
          _refused(lambda: lite_copy.write_lite_copy(plain), "no Summary sheet"))
    odd = tmp / "odd.xlsx"
    wb = Workbook()
    wb.active.title = "Summary"
    wb.active["B2"] = "Some other summary"
    wb.save(odd)
    check("a Summary without the comparison's blocks is refused",
          _refused(lambda: lite_copy.write_lite_copy(odd), "has no ROW COUNTS block"))
    check("a lite path that is the workbook itself is refused",
          _refused(lambda: lite_copy.write_lite_copy(full, full), "overwrite its own"))
    target = tmp / "guarded (lite).xlsx"
    check("a destination guard that says no is obeyed",
          _refused(lambda: lite_copy.write_lite_copy(full, target, allow=lambda p: False),
                   "destination changed"))
    leftovers = [p.name for p in tmp.iterdir()
                 if p.name.endswith(".tmp") or p.name in (
                     "plain (lite).xlsx", "odd (lite).xlsx", target.name)]
    check("no lite copy or temporary file was left", not leftovers, f"{leftovers}")
    check("lite_title keeps a title without the tail",
          lite_copy.lite_title("A vs B — R") == "A vs B — R"
          and lite_copy.lite_title("A vs B — R — Discrepancy Report (x)") == "A vs B — R")


def check_fallback(tmp, full, lite):
    print("the zip64-capable fallback:")
    saved = lite_copy._MAX32
    try:
        lite_copy._MAX32 = 64                     # every offset now "needs" zip64
        out = lite_copy.write_lite_copy(full, tmp / "fallback (lite).xlsx")
    finally:
        lite_copy._MAX32 = saved
    with zipfile.ZipFile(out) as z, zipfile.ZipFile(lite) as zl:
        same_parts = z.namelist() == zl.namelist()
        same_bytes = all(z.read(n) == zl.read(n) for n in zl.namelist())
        crc_ok = z.testzip() is None
    check("it writes the same parts with the same contents, every CRC good",
          same_parts and same_bytes and crc_ok)


def _published(out, warnings=()):
    def produce(tmp_path):
        return run_compare(ROUTED, ROWS_A, ROWS_B, True, Path(tmp_path), mode="values",
                           confirm_overwrite=lambda _p: True, warnings=warnings)
    return artifact_store.commit_workbook(Path(out), produce, expect_sheet="Comparison",
                                          requested_mode="values")


def check_matrix_step(tmp):
    print("the matrix step (matrix._settle_lite_copy):")
    day = tmp / "day"
    day.mkdir()
    out = day / "clean_check_vs_tsn 2026-11-19 ars-prod dev-site.xlsx"
    res = _published(out)
    assert res.status == "ok", res.message
    lite = matrix._lite_sibling(out)
    logs = []
    events = Events(on_log=logs.append)
    matrix._settle_lite_copy(out, res, True, events)
    check("with the option on, the lite copy is written beside the comparison",
          lite.exists() and lite.name ==
          "clean_check_vs_tsn 2026-11-19 ars-prod dev-site (lite).xlsx")
    check("...and the run log names it", any(lite.name in m for m in logs))
    matrix._settle_lite_copy(out, res, False, events)
    check("with the option off, a values refresh removes the prior lite copy",
          not lite.exists() and any("Removed the prior lite copy" in m for m in logs))

    partial_out = day / "partial.xlsx"
    partial = _published(partial_out, warnings=("route 005.xlsx could not be read",))
    assert partial.status == "ok", partial.message
    logs.clear()
    matrix._settle_lite_copy(partial_out, partial, True, events)
    check("a partial comparison gets no lite copy, and the log says why",
          not matrix._lite_sibling(partial_out).exists()
          and any("partial" in m for m in logs), f"{logs}")

    lite.write_bytes(b"an input that happens to have the lite name")
    logs.clear()
    matrix._settle_lite_copy(out, res, True, events, source_paths=(lite,))
    check("a lite copy that would overwrite a comparison input is not written, "
          "and the input is kept", lite.read_bytes().startswith(b"an input"))
    matrix._settle_lite_copy(out, res, False, events, source_paths=(lite,))
    check("...nor deleted when the option is off", lite.exists())
    lite.unlink()

    logs.clear()
    matrix._settle_lite_copy(out, res, True, events, commit_guard=lambda path=None: False)
    check("a destination that is no longer the matrix's gets nothing",
          not lite.exists() and any("not written" in m for m in logs))


def main():
    tmp = Path(tempfile.mkdtemp(prefix="check_lite_"))
    try:
        full, lite = check_routed(tmp)
        check_flat(tmp)
        check_refusals(tmp, full)
        check_fallback(tmp, full, lite)
        check_matrix_step(tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("\nall good" if not FAILS else f"\n{len(FAILS)} FAILED")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
