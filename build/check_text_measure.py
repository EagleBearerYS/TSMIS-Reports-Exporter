"""Text widths are assembled from character and pair lengths — exactly.

Column widths are stored in every comparison workbook, and they come from
measuring each distinct value in the real font. `compare_core._font_length`
replaces one ``font.getlength(text)`` per value with a sum over per-character
and per-pair lengths that fill a table once. Under Pillow's BASIC layout engine
that sum is the same number ``getlength`` returns (a string's length is its
glyph advances plus the kerning of each adjacent pair, all whole multiples of
1/64 px), so no stored width can move.

This holds the identity on the fonts the workbooks are measured in, over text of
every kind the reports carry and over random strings, and holds that any OTHER
layout engine is measured directly rather than assumed additive.

Run from the repo root:  build\\.venv\\Scripts\\python.exe build\\check_text_measure.py
"""
from __future__ import annotations

import random
import string
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT / "scripts"), str(ROOT / "build"), str(ROOT)]

import compare_core as core  # noqa: E402

_failures: list[str] = []


def check(name, cond, detail=""):
    print(f"  [{'OK ' if cond else 'FAIL'}] {name}")
    if not cond:
        if detail:
            print(f"       {detail}")
        _failures.append(name)


_REAL = [
    "A", " ", "AV", "To", "Ty", "fi", "ffl", "L'", "T.", "Y,", "r.", "P.", "1.", "11", "W.A.",
    "005/NB ON FR MAIN ST", "RTE 101 SB OFF TO EL CAMINO REAL", "R012.345", "L000.000R",
    "INTERSECTION DETAIL — MAJOR / MINOR", "Comparison row", "Key (helper)", "(blank)",
    "12345678901234567890", "0.30000000000000004", "1.50 ≠ 1.5", "AVALON BLVD ≠ AVALON BL",
    "++", "+", "  two  spaces  ", "tab\there", "line\nbreak", "é ≠ ü — ½ ° Ñ ñ", "🙂 𝔘 ￾",
    "WAWAWAWAWAWAWAWA", "iiiiiiiiiiiiiiii", "T" * 200, "x",
]


def main():
    try:
        from PIL import ImageFont
    except ImportError:
        check("Pillow is not installed: widths fall back to the character estimate "
              "(nothing to hold here)", True)
        return
    fonts = []
    for bold in (False, True):
        for size in (10.0, 11.0):
            fonts.extend(core._fonts_for(bold, size))
    if not fonts:
        check("no measuring font on this machine: widths use the character estimate "
              "(nothing to hold here)", True)
        return
    basic = [f for f in fonts if getattr(f, "layout_engine", None) == ImageFont.Layout.BASIC]
    print(f"{len(fonts)} measuring fonts, {len(basic)} with the BASIC layout engine")

    rng = random.Random(11)
    alphabet = (string.ascii_letters + string.digits + " .,-/()&'#:;+_\"%*=<>@!?[]"
                + "é≠üÑñ—°½“”’\t")
    randoms = ["".join(rng.choice(alphabet) for _ in range(rng.randint(1, 60)))
               for _ in range(4000)]
    texts = _REAL + randoms

    mismatches = []
    for font in fonts:
        core._GLYPH_TABLES.pop(id(font), None)
        for text in texts:
            direct = font.getlength(text)
            assembled = core._font_length(font, text)
            if direct != assembled:
                mismatches.append((text[:40], direct, assembled))
    check(f"{len(texts) * len(fonts):,} measurements equal getlength exactly",
          not mismatches, repr(mismatches[:3]))
    if basic:
        tables = core._GLYPH_TABLES[id(basic[0])]
        check("a BASIC-engine font is measured from its tables",
              tables[1] is not None and len(tables[1]) > 50 and len(tables[2]) > 500,
              f"{len(tables[1] or ())} characters, {len(tables[2] or ())} pairs")

    class _Shaping:
        """A stand-in for a font whose engine shapes across characters."""
        layout_engine = "raqm"

        def __init__(self):
            self.calls = []

        def getlength(self, text):
            self.calls.append(text)
            return float(len(text) ** 2)             # deliberately NOT additive

    shaping = _Shaping()
    check("any other layout engine is measured directly, never assembled",
          core._font_length(shaping, "abcd") == 16.0 and shaping.calls == ["abcd"],
          repr(shaping.calls))

    print("the public measure is unchanged by it:")
    core._PX_CACHE.clear()
    fonts10 = core._fonts_for(False, 10.0)
    for text in ("005/NB ON FR MAIN ST", "AVALON BLVD ≠ AVALON BL", "T" * 80):
        expected = (max(f.getlength(text) for f in fonts10)
                    / core._MEASURE_SCALE * core._MEASURE_MARGIN)
        check(f"_text_px({text[:18]!r}…) is the widest font's own length",
              core._text_px(text, size_pt=10.0) == expected)
    saved = dict(core._GLYPH_TABLES)
    limit = core._GLYPH_TABLE_MAX
    core._GLYPH_TABLE_MAX = 5
    try:
        font = fonts10[0]
        check("the tables are bounded, and a reset changes no answer",
              all(core._font_length(font, t) == font.getlength(t) for t in texts[:400]))
    finally:
        core._GLYPH_TABLE_MAX = limit
        core._GLYPH_TABLES.update(saved)


if __name__ == "__main__":
    print("assembled text widths equal the font's own:")
    main()
    if _failures:
        print(f"\n{len(_failures)} check(s) FAILED")
        raise SystemExit(1)
    print("all good")
