"""CompareSchema.context_display_side (v0.54.1): a context column can show one
side's OWN value on a matched row instead of the coalescing default.

Locks: the default stays the coalescing display (side A's value, else side
B's) in the Python truth, the values workbook and the formula twin; naming a
side shows exactly that side's value, blank where it is blank, in all three;
a context cell never asserts either way; a one-sided row still shows its own
side's value; a name that is neither side is refused.

Also CompareSchema.likeness_text (v0.54.1): by default the pairing of a
repeated key reads the shown text; with the hook it reads the spelling the
hook gives back, and the counts are the same either way.

    python build/check_context_display_side.py
"""
import sys
import tempfile
from dataclasses import replace
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_ROOT / "scripts"))

import compare_core as core  # noqa: E402
from compare_core import CompareSchema, run_compare  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

FAILS = []


def check(label, ok):
    print(f"  {'ok' if ok else 'FAIL'}: {label}")
    if not ok:
        FAILS.append(label)


BASE = CompareSchema(
    report_name="Context side",
    header=["Key", "F1", "Context"],
    side_a="BUILD",
    side_b="SITE",
    id_noun="row",
    id_noun_plural="rows",
    context_fields=("Context",),
)

A = [["k1", "same", "build value"],
     ["k2", "same", ""],
     ["k3", "same", "a only"],
     ["ka", "x", "build one-sided"]]
B = [["k1", "same", "site value"],
     ["k2", "same", "site value"],
     ["k3", "same", ""],
     ["kb", "y", "site one-sided"]]


def _context_cells(sc, tmp, name):
    """{key: the Comparison sheet's Context cell} of a values workbook."""
    out = tmp / name
    result = run_compare(sc, A, B, False, out, mode="values")
    assert result.status == "ok", result.message
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        rows = list(wb["Comparison"].iter_rows(values_only=True))
    finally:
        wb.close()
    head = [str(h) for h in rows[0]]
    k, c = head.index("Key"), head.index("Context")
    return {r[k]: ("" if r[c] is None else r[c]) for r in rows[1:] if r[k]}, result


def main():
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        cells, res = _context_cells(BASE, tmp, "default.xlsx")
        check("the default shows side A's value, else side B's",
              cells == {"k1": "build value", "k2": "site value", "k3": "a only",
                        "ka": "build one-sided", "kb": "site one-sided"})
        default_counts = res.comparison_outcome.counts

        sc_b = replace(BASE, context_display_side="SITE")
        cells, res = _context_cells(sc_b, tmp, "site.xlsx")
        check("naming side B shows only its own value, blank where it is blank",
              cells == {"k1": "site value", "k2": "site value", "k3": "",
                        "ka": "build one-sided", "kb": "site one-sided"})
        check("... and moves no count",
              res.comparison_outcome.counts == default_counts)

        sc_a = replace(BASE, context_display_side="BUILD")
        cells, res = _context_cells(sc_a, tmp, "build.xlsx")
        check("naming side A shows only its own value",
              cells == {"k1": "build value", "k2": "", "k3": "a only",
                        "ka": "build one-sided", "kb": "site one-sided"})
        check("... and moves no count",
              res.comparison_outcome.counts == default_counts)

        live = run_compare(sc_b, A, B, False, tmp / "live.xlsx", mode="formulas")
        check("a live-formulas workbook builds with a named side",
              live.status == "ok"
              and live.comparison_outcome.counts == default_counts)

    rt, rn = ["k2", "same", ""], ["k2", "same", "site value"]
    for sc, want in ((BASE, "site value"), (sc_b, "site value"), (sc_a, "")):
        cell = core.compared_cell(sc, 2, rt, rn, 0)
        check(f"compared_cell ({sc.context_display_side or 'default'}): "
              f"{want!r}, never asserting",
              cell.display == want and not cell.asserting)
        check(f"_field_value agrees ({sc.context_display_side or 'default'})",
              core._field_value(sc, rt, rn, 0, 2, "N") == want)

    lay_default = type("L", (), {"sc": BASE, "only_a": "BUILD only",
                                 "only_b": "SITE only"})()
    lay_b = type("L", (), {"sc": sc_b, "only_a": "BUILD only",
                           "only_b": "SITE only"})()
    lay_a = type("L", (), {"sc": sc_a, "only_a": "BUILD only",
                           "only_b": "SITE only"})()
    default_expr = core._field_display_expr(lay_default, 2, "S", "T", "N", "X")
    check("the formula twin's default is unchanged",
          default_expr == 'IF(S="BUILD only",T,IF(S="SITE only",N,IF(T="",N,T)))')
    check("the formula twin shows side B's own value when named",
          core._field_display_expr(lay_b, 2, "S", "T", "N", "X")
          == 'IF(S="BUILD only",T,IF(S="SITE only",N,N))')
    check("the formula twin shows side A's own value when named",
          core._field_display_expr(lay_a, 2, "S", "T", "N", "X")
          == 'IF(S="BUILD only",T,IF(S="SITE only",N,T))')

    bad = replace(BASE, context_display_side="TSN")
    try:
        core.compared_cell(bad, 2, rt, rn, 0)
        refused = False
    except ValueError as e:
        refused = "names neither side" in str(e)
    check("a side name that is neither side is refused", refused)

    check_likeness()


# CompareSchema.likeness_text (v0.54.1): the pairing's character likeness can
# read a cell in another spelling than the one shown. Two rows of one key on
# each side, every value differing, so both pairings cost the same number of
# differing cells and only the likeness decides: as shown, AAAA is nearest
# AAAB; in the spelling the hook gives back, AAAA is BBBA's twin.
LIKE = CompareSchema(
    report_name="Likeness", header=["Key", "F"], side_a="A", side_b="B",
    id_noun="row", id_noun_plural="rows")
LA = [["k", "AAAA"], ["k", "BBBB"]]
LB = [["k", "AAAB"], ["k", "BBBA"]]
SPELLING = {"AAAA": "p", "BBBB": "q", "AAAB": "q", "BBBA": "p"}


def _pairs(sc, tmp, name):
    out = tmp / name
    result = run_compare(sc, LA, LB, False, out, mode="values")
    assert result.status == "ok", result.message
    wb = load_workbook(out, read_only=True, data_only=True)
    try:
        rows = list(wb["Comparison"].iter_rows(values_only=True))
    finally:
        wb.close()
    head = [str(h) for h in rows[0]]
    return sorted(r[head.index("F")] for r in rows[1:] if r[0]), result


def check_likeness():
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        plain, res_plain = _pairs(LIKE, tmp, "plain.xlsx")
        check("by default the pairing reads the shown text (AAAA with AAAB)",
              plain == ["AAAA ≠ AAAB", "BBBB ≠ BBBA"])
        hooked, res_hooked = _pairs(
            replace(LIKE, likeness_text=lambda name, text: SPELLING.get(text, text)),
            tmp, "hooked.xlsx")
        check("with likeness_text it reads the spelling the hook gives back "
              "(AAAA with BBBA)", hooked == ["AAAA ≠ BBBA", "BBBB ≠ AAAB"])
        check("... and the counts are the same either way",
              res_plain.comparison_outcome.counts == res_hooked.comparison_outcome.counts)

    print("\nall good" if not FAILS else f"\n{len(FAILS)} FAILED")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
