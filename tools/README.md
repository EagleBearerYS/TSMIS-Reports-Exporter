# tools/

Developer utilities (not shipped with the app).

## clean_road_layer_check.py

Checks a Clean Road Highway file (an export or the app's build), or a Reports vs
ArcGIS Highway workbook, against the raw ArcGIS layer files directly, the way a
person checks one cell by hand, with none of the build's placement logic. Run it
before a comparison workbook is delivered:

```bash
python tools/clean_road_layer_check.py index arcgis_layers 2026-10-30 layers.pkl
python tools/clean_road_layer_check.py workbook "output/comparisons/arcgis-by-day/<day>/clean_highway_vs_layers <day>.xlsx" layers.pkl
python tools/clean_road_layer_check.py spot "output/comparisons/arcgis-by-day/<day>/clean_highway_vs_layers <day>.xlsx" arcgis_layers 2026-10-30 "output/<day>/clean_highway" --out spot.txt
```

`workbook` says what share of the workbook's differences are false (the layer
says what TSMIS says, so the build was wrong); `file` says how often a file
differs from its own layers; `spot` prints random cells beside the raw rows for
a person to read. The as-of is the library's export date
(`arcgis_layers/_refresh/last_run.json`). Detail and the 2026-11-03 numbers:
[docs/comparison-engine.md](../docs/comparison-engine.md) §9o.

## screenshots.py

Regenerates the marketing screenshots and the social/OG card by rendering the
real GUI (`scripts/ui`) in its built-in `#mock` preview mode with a headless
browser, in both light and dark themes.

```bash
python -m playwright install chromium   # once, if not already present
python tools/screenshots.py
```

Outputs (overwritten in place):

- `docs/screenshot-light.png`, `docs/screenshot-dark.png` — used by the README.
- `docs/og-image.png` — the 1200×630 social preview card. The app screenshot is
  embedded as a **data: URI** (the card renders via `set_content`, whose
  about:blank origin makes Chromium silently block `file://` images — the bug
  that shipped an empty app window through v0.20.x), and the script **fails if
  the image didn't decode** rather than writing a broken card.

The landing page lives on the **`gh-pages`** branch and keeps its own copies at
the branch root. After regenerating, copy the three files onto `gh-pages` to
refresh the site:

```bash
git checkout gh-pages
git checkout main -- docs/screenshot-light.png docs/screenshot-dark.png docs/og-image.png
git mv -f docs/screenshot-light.png screenshot-light.png
git mv -f docs/screenshot-dark.png  screenshot-dark.png
git mv -f docs/og-image.png         og-image.png
git commit -m "Refresh site screenshots" && git push
```

Then **bump the `?v=` query on the two `og-image.png` meta URLs in the
gh-pages `index.html`** (`og:image` + `twitter:image`, currently `?v=2`) —
social scrapers cache the card by URL, so a regenerated image behind the same
URL keeps showing the old card on new shares.
