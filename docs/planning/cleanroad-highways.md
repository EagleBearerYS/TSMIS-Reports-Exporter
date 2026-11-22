# Clean Road — CA HIGHWAYS from the ArcGIS layers (coverage assessment)

**Status: CA HIGHWAYS SHIPPED (v0.29.0) and refined through v0.39.3. CA
INTERSECTIONS and CA RAMPS are the next large build — roadmap G1.**

This is the design/coverage record for building our own clean-road files from the
owner's TSMIS ArcGIS layer exports and comparing them against the TSN extract
(`CA HIGHWAYS <date>.xlsx`, the `THY_*` table). The highways build is live; the
other two are censused to schema level and not started.

**Read in this order if you're starting G1** (the CA INTERSECTIONS / CA RAMPS
builds):

1. **[The consolidation model](#the-consolidation-model-proven)** and
   **[the build-time rules](#app-integration--shipped-v0290-the-arcgis-tab)** —
   the measured rules the highways build settled over six probe rounds
   (`CODE- Label` domains, cross-county chain-walk, the ADT anchor, PM-continued
   offsets, primary-layer effective dates). **Do not re-derive these.** They cost
   days and several of them overturned a plausible first guess.
2. **[The other two clean-road files](#the-other-two-clean-road-files-censused-2026-07-22--schema-level-mapping)**
   — the layer→column mapping for both, already censused, with the two
   derivations left deliberately open (`INX_LSC_DATE`'s eff-date class and
   `RAM_PRIMARY_DIRECTION_CODE`).
3. `docs/comparison-engine.md` §9j for how the highways comparison is wired, and
   `scripts/tsn_load_clean_road.py`, whose Intersection and Ramp slots are
   deliberately normalizer-less until these builds land.

Everything below is **measured** — no column mapping is assumed. The original
census ran against the 2026-07-20 layer export and the 2025-09-08 TSN extract;
the library was refreshed on 2026-08-19 (40 layers, `ground-truth/Clean Road and
ArcGIS layers/All Layers 8.19/`) and the rules re-measured against the real
2026-08-17 export. Census scripts and exact numbers live locally beside the data in
`Downloads\TSMIS\ground-truth\Clean Road and ArcGIS layers\_analysis\` (see its
README; local-only, like all real data; filed there from `_inbox\Cleanroad\` on
2026-10-05).

> **The as-of does NOT follow the layer library.** `resolve_default_asof()` takes
> it from the staged TSN extract, so a default build off fresh layers still
> reconstructs the extract's date. A reconstruction is only meaningful when its
> as-of matches whatever it is being compared against — set it explicitly.

---

## The shape of the problem (measured)

- **TSN side:** `CA HIGHWAYS` = 60,083 rows × 74 `THY_*` columns, **current-state
  only** (`THY_END_DATE` is the constant sentinel `3000-01-01`), extract date
  stamped in every row (`THY_EXTRACT_DATE`, here 2025-09-08). Same 60,083-record
  universe as the TSAR Highway Detail TSN extract (56 cols) — the clean-road
  table is the SOURCE table the HD report projects from.
- **TSMIS side:** single-attribute **LRS event layers** (one worksheet per layer
  in the export workbook). Shared span shape: `District, RouteNum('001'),
  RouteSuffix('.'=none, S, U), Alignment, Begin/EndCounty, Begin/EndPMPrefix
  ('.'=none), Begin/EndPMMeasure, Begin/EndPMSuffix, Begin/EndODMeasure,
  <attribute column(s)>, InventoryItemStartDate, InventoryItemEndDate,
  LRSFromDate, LRSToDate, RouteID('SHS_001._P'), From/ToARMeasure, LocError`.
- **History:** every layer carries its full slice history. A row is **current
  iff `LRSToDate` is empty** (the owner's rule — confirmed), and the slices
  support **as-of-date reconstruction**: `LRSFromDate <= D < LRSToDate`. The
  layer export (2026-07-20) is 10 months newer than the TSN extract
  (2025-09-08); a fair comparison must either select slices **as-of the TSN
  extract date** or use a same-dated TSN extract.
- **`InventoryItemStartDate` is the DOMAIN effective date** (Excel serial;
  26956 = 1973-10-19 — matches THY eff dates), distinct from the LRS lifecycle
  dates. Do not confuse the two when filtering.
- **City / County Code / Posted Speed cover ALL public roads** (that's why they
  are 481k/825k rows) — the highways build keeps only `RouteID` beginning
  `SHS_`.

## The consolidation model (proven)

Build, per route: select each needed layer's slices as-of the chosen date, keep
single-county spans keyed **county + PM interval** (split the rare cross-county
spans, ≤5 per layer on route 001), union all breakpoints, and emit one output
row per homogeneous stretch with each layer's value painted across it.

**Proof (route 001, ORA+MEN, 502 base THY rows, as-of 2025-09-08):**

| layer column → THY column | agreement |
|---|---|
| Design_Speed → THY_DESIGN_SPEED_AMT | **502/502** |
| Terrain_Type → THY_TERRAIN_CODE | **502/502** |
| Median_Type / Width / Variance → THY_MEDIAN_* | **502/502** ×3 |
| Barrier_Type → THY_MEDIAN_BARRIER_CODE | **502/502** |
| Highway_Group → THY_HIGHWAY_GROUP_CODE | **502/502** |
| SHS_Access_Control → THY_HIGHWAY_ACCESS_CODE | **502/502** |
| Travel_Way_Width_R/L → THY_RT/LT_TRAV_WAY_WIDTH_AMT | **502/502** ×2 |
| Total_Num_Lanes_R/L → THY_RT/LT_LANES_AMT | **502/502** ×2 |
| Shld_Width_Total/Treated_Out_R → THY_RT_O_SHD_TOT/TRT | **502/502** ×2 |
| Shld_Width_Total/Treated_In_R → THY_RT_I_SHD_TOT/TRT | **502/502** ×2 |
| Special_Feature_Type_R → THY_RT_SPEC_FEATURES_CODE | **502/502** (absent span ≡ THY `Z`) |
| Surface_Type_R/L → THY_RT/LT_SURF_TYPE_CODE | **501/502** ×2 (one 0.001-mi boundary sliver at MEN 82.011) |

Traps the proof flushed out (each cost a wrong first result):
- **Join on county+PM, never on the odometers.** THY's `BEGIN/END_OFFSET_AMT`
  is a PM-continued cumulative (route 001 begins at offset 0.129 = its begin
  PM), while the layers' `ODMeasure` is a separately calibrated odometer
  (route-001 end: THY 645.121 vs layer 645.377) and `ARMeasure` is a third
  family (645.377 vs 639.308). Midpoint-matching in OD space silently scrambles
  every fine-grained attribute (28–39% agreement) while coarse ones stay ~85% —
  a false "mostly works" signal.
- **`Total_Num_Lanes`, not `Thru_Num_Lanes`,** is THY's lanes column (100% vs 0%).
- **Layer L/R == THY LT/RT** — no side swap (tested both ways).
- **Boundary slivers:** the two systems can disagree by 0.001 PM on a breakpoint;
  the eventual comparison needs a sliver/tolerance policy, not exact-edge faith.

## Column coverage (74 THY columns)

**Value-proven pulls (on 001)** — the 24 attribute columns in the table above,
plus `THY_TOLL_FOREST_CODE` (mux: **1 = toll** — SHS Tolls spans; verified
geographically: SF 080 Bay Bridge, 160 Antioch, ORA 073/241/261; **2 = forest**
— SHS Forest HWY spans: MPA 140, TRI 003/036, SBD 038, TUO 108/120…; 0 = neither)
and `THY_LANDMARK_SHORT_DESC` (SHS Landmark `Landmarks_Short`, byte-equal spot
check at route-001 begin).

**Mapped, same pattern, not yet value-run** (no reason to doubt, still unproven):
left-side shoulder/special-feature layers (only the R side was run),
`THY_CURB_LANDSCAPE_CODE` ← SHS Curb Landscape, `THY_POPULATION_CODE` ← SHS
Population, `THY_NON_ADD_CODE` ← SHS Non Add Mileage (flag → `N`, default `A`),
`THY_CITY_CODE` ← City (`SHS_` rows only), `THY_RECORD_DATE` + `THY_SEG_ORDER_ID`
← SHS Inv Network Date (`Network_Start_Date`, `SegOrderId`).

**Block effective dates — two rules, deliberately (v0.39.2, DA1).**
`_block_eff_date` computes `THY_LEFT_ROAD_EFF_DATE` / `THY_MEDIAN_EFF_DATE` /
`THY_RIGHT_ROAD_EFF_DATE`. The clean-road build (this doc's subject) keeps the ORIGINAL
rule — the OLDEST `InventoryItemStartDate` among the block's member layers, measured on
route 001. A REPORT build passes `primary_eff=True` and gets the rule measured against
the real TSMIS export over 44-46k rows: **the block's PRIMARY layer's own date**
(`BLOCK_PRIMARY` — travel way for a roadbed, the median layer for the median), falling
back to the NEWEST of the block's other layers only where the primary has no covering
span. It scored 79-80% where the oldest-member rule scored 56-60%.

They differ only because the clean-road side has never been measured: Highway Detail
PRINTS `THY_LEFT_ROAD_EFF_DATE`, so the report's rule is very likely right here too, but
these columns are COUNTED in the vs-TSN comparison and its blessed canary and the dev box
has no staged TSN extract. Roadmap **DA5** owns that measurement.

**The vendor's own mapping, cross-checked (2026-09-02).** The dev site 9.1 capture
ships `clean_highway.js` — the site's Clean Road File (Highway) export, the same 74-column
THY table rendered by the vendor from the same layers (52 columns sourced, 22 left blank).
Its column registry is a second, independent mapping table, read against `PROVENANCE`:

| Verdict | Columns |
|---|---|
| **Agrees layer-for-layer** | every attribute column — surface type (136/137), lanes + travel-way width (139/140), special features (134/135), outer/inner shoulder widths (128/129, 120/121), median type/width/variance (124), curb-landscape (112), barrier (110), access control + its date (109), terrain (11), design speed (113), non-add (125, `N` else `A`), landmark (123), population (130), highway group (116), network start date + `SegOrderId` (122), equate `E` from equation points, route breaks from 133 |
| **Same rule** | the block eff dates are the PRIMARY layer's own `InventoryItemStartDate` (139 / 124 / 140) — the v0.39.2 `primary_eff` rule, independently chosen by the vendor (DA5 support) |
| **Sources we mark "no TSMIS source"** | `THY_FUNCTIONAL_CLASS_CODE` ← layer 91 `F_System`; `THY_LAST_SIG_CHG_DATE` ← the newest `InventoryItemStartDate` among the layers whose change made the row significant (lanes, access, treated-shoulder >4, travel-way >5, median-width to/from 0 — the Highway Log's own significance test). TSN's FUNCTIONAL_CLASS is all-null, so the first is context anyway; the second is a real candidate for our build |
| **Differs by design** | `THY_TOLL_FOREST_CODE` = layer 115's `Forest_Hwy` value alone (no toll mux from 138); `THY_BREAK_DESC` = BEG/END only (no CNTY/DIST/BEGL/ENDR); the offsets are OD measures (AR→OD translate) and length = OD end − begin; district comes off the HL point stream (85/116) rather than SHS District (114); `THY_CITY_CODE` reads layer 74's `City_Code` FIELD directly — our layer export carries city NAMES and translates them through `city_codes` |
| **Left blank by the vendor** | `THY_PROFILE_CODE` / `THY_ADT_AMT` / `THY_CHANGE_PER_MILE_AMT` (we derive them from Traffic Volume Segments), `THY_POPULATION_GROUP_CODE` ("awaiting group field name"), and the same bookkeeping / SIG_CHG_IND / federal-aid / lands / scenic set we leave empty |

One consequence worth acting on when the site's Highway export is available statewide: it is
a third THY-shaped side, so "site export vs our build" measures the two reconstructions of
the same layers against each other with no TSN vintage gap at all.

**Build-only columns (v0.39.1)** — ours, with no TSN counterpart, so CONTEXT in the
vs-TSN comparison and never counted there. `chc.HEADER` stays exactly TSN's 74 (the
library loads the raw extract through it as an exact gate); `chc.ARC_HEADER` =
`HEADER + BUILD_ONLY_COLUMNS` is what the build writes, a strict superset so every
existing index holds.

| Column | Source | Why it exists |
|---|---|---|
| `THY_POPULATION_EFF_DATE` | SHS Population (`InventoryItemStartDate`) | The TSN table gives four blocks an eff date but population only a CODE — yet **Highway Detail PRINTS the Rural/Urban effective date on 99.3% of its rows** (`RU Eff`), so a projection of this table could not fill a column the report always shows (DA2). Same single-layer passthrough as `THY_ACCESS_EFF_DATE`; the layer's value distribution matches the printed column in the same rank order (1964-01-01 / 2010-12-31 lead both). |

**Identity/structural (synthesized by the overlay):** district, county, route,
route suffix, PM prefix/begin/end/suffix, length (= end−begin), the offsets
(cumulate county PM spans along the route — matches THY's convention), and
`THY_PM_SUFFIX_CODE` (mirrors HG for R/L/X rows — counts identical: 1027/965/340).
`THY_BREAK_DESC` is composed: BEG/END/CNTY/DIST are structural (route termini,
county/district lines), R-BR/U-BR ride SHS Route Break, plus equates from
Equation Points (`THY_EQUATE_CODE`).

**Owner decision (2026-07-22): the built file keeps the FULL 74-column THY
header.** A column we cannot source is still PRESENT in the built workbook and
in the comparison sheet — left empty, and explicitly noted as "no TSMIS ArcGIS
source found" (per-column provenance record + the comparison Notes sheet; the
comparison shows it one-sided rather than silently dropping it). Discrepancies
in a correctly-sourced column are the comparison's job to surface, never a
reason to re-source it. Three tiers:

1. **Sourced** — filled from the mapped layer; provenance = layer + column
   (+ FeatureServer URL/layer id from the INDEX sheet).
2. **No TSMIS source** — present + empty + noted (the table below).
3. **TSN-internal bookkeeping** (`THY_ID`, `THY_ELEMENT_ID`, create/update
   user+date, lifecycle `BEGIN/END_DATE`) — present + empty + noted "TSN-internal";
   never fabricated. Exception: `THY_EXTRACT_DATE` is stamped with the build's
   as-of date (that is literally what the column means on our side — flagged in
   the notes so it reads as ours, not TSN's).

**Gaps — cannot be built from the current export:**

| THY columns | why | action |
|---|---|---|
| `THY_ADT_AMT`, `THY_CHANGE_PER_MILE_AMT`, `THY_PROFILE_CODE` | **DELIVERED + source-confirmed (2026-07-22): Traffic Volume Segments.** Not a verbatim pull — a DERIVED family: TVS carries the profile anchors (`AADT`/`AADT_AHEAD`/`AADT_BACK`/`AADT_YEAR` spans) and THY interpolates ADT along them (`CHANGE_PER_MILE` = the slope — `(AHEAD−BACK)/len` matched THY to 4+ decimals on 001, sign convention pending; all 28 THY `P` rows co-locate with TVS spans; `S` = interpolated rows). TVS keeps 3 year-vintages LRS-live concurrently (2022/23/24 as-of Sep-2025), so the consolidator selects latest-year-per-location. `AADT_CODE` is a binary flag, NOT the P/S source. | pin the interpolation + vintage rule at build time (statewide fit, like the block eff dates) |
| `THY_MAINT_SVC_LVL_CODE` | no maintenance layer anywhere in the ~90-layer list | ask the owner whether TSMIS carries it at all |
| `THY_FEDERAL_AID_CODE`, `THY_FA_ROUTE_PREFIX_CODE`, `THY_FA_ROUTE_NAME` | legacy federal-aid system; no layer | likely TSN-legacy-only — owner confirm |
| `THY_NATIONAL_LANDS_CODE`, `THY_SCENIC_FREEWAY_CODE` | no layer | owner confirm |
| the 4 `*_SIG_CHG_IND` + `THY_LAST_SIG_CHG_DATE` | TASAS-side change tracking | likely not derivable — exclude from comparison |
| block eff dates (`THY_LEFT/RIGHT_ROAD_EFF_DATE`, `THY_MEDIAN_EFF_DATE`, `THY_ACCESS_EFF_DATE`) | single-layer item dates agree only 82–96% — the block date is a composite (candidate: max over the block's member layers) | resolve while building; PARTIAL until then |
| bookkeeping (`THY_ID`, `THY_ELEMENT_ID`, create/update user+date, lifecycle dates, `THY_EXTRACT_DATE`) | TSN-internal | synthesize/ignore — never compare |
| `THY_POPULATION_GROUP_CODE`, `THY_FUNCTIONAL_CLASS_CODE` | **all-null in the TSN extract** | nothing to build |

## The layer export list = the IN-APP INPUT CONTRACT (CA HIGHWAYS)

> **v0.46.0: the app produces this library itself** (ArcGIS ▸ Layers ▸ Refresh):
> ArcGIS Pro's own Python reads each of the 40 layers below straight from the
> `lrs_tsmis` FeatureServer (default version) and writes it with Table To Excel —
> field names as headers, coded domains as descriptions, `NN_<Layer Name>.xlsx` plus
> `00_INDEX.xlsx` (now with per-layer Exported At / Exported By). The contract below
> is unchanged and a manual export in the same shape still reads; the owner no longer
> exports by hand. One difference from the manual drop: Route Direction comes from
> its own service layer, without the map join's extra columns (no build reads them).

**Owner direction (2026-07-22): the app input is ONE .xlsx PER LAYER**, dropped
into `arcgis_layers/`, re-exported by the owner from this list. **The COMPLETE
library manifest covers ALL THREE clean-road files — 40 layers, in the ArcGIS
contents (screenshot) order** (H = highways, I = intersections, R = ramps;
✅ = already delivered 2026-07-22):

City (H I R) · County Code (H) · Equation Points (H) · IM Complex Intersection
Cross Reference (I) · IM Complex Intersection Influence Segments (I) · IM
Intersection Approach Detail (I) · IM Intersection Approach Segments (I) · IM
Intersection Detail (I) · IM Intersection Point (I) · IM Intersection Route
Table (I) · Route Direction (R) · SHS Access Control (H) · SHS Barrier (H) ·
SHS Curb Landscape (H) · SHS Design Speed (H) · SHS District (H) · SHS Forest
HWY (H) · SHS Highway Group (H I R) · SHS I Shld Width L (H) · SHS I Shld
Width R (H) · SHS Inv Network Date (H I R) · SHS Landmark (H) · SHS Median (H)
· SHS Non Add Mileage (H) · SHS O Shld Width L (H) · SHS O Shld Width R (H) ·
SHS Population (H I R) · SHS Ramp (R) · SHS Ramp Pt (R) · SHS Route Break (H) ·
SHS Special Feature L (H) · SHS Special Feature R (H) · SHS Surface Type L (H)
· SHS Surface Type R (H) · SHS Tolls (H) · SHS Travel Way L (H) · SHS Travel
Way R (H) · Terrain Type (H) · Traffic Volume Ramps (R) ✅ · Traffic Volume
Segments (H I) ✅

The HIGHWAYS consolidator reads its 29; the intersections/ramps sets are staged
for their phases (mapping censused, value proofs pending those phases).

**The delivered per-layer format (2026-07-22 preview — the app input contract):**
a folder of numbered files `NN_<Layer Name>.xlsx` (one sheet, named after the
layer) plus **`00_INDEX.xlsx`** — `Excel File | ArcGIS Layer or Table | Rows
Exported | Fields Exported | ArcGIS Contents Path | Data Source`. The INDEX is
the verification manifest AND the audit-provenance record: the consolidator
verifies each file's row count against it and stamps each output column with
the layer's FeatureServer URL + layer id. One dialect wrinkle: this export
style resolves coded domains to LABELS (`District 7` / `Los Angeles` / `Right`)
where the earlier bundle exports carried CODES (`7` / `LA` / `R`) — the
consolidator normalizes BOTH dialects (fixed 58-county + district + alignment
tables); if the export dialog offers "use domain codes", codes are mildly
preferred but not required.

Export rules: **all rows, all columns** (the history slices are load-bearing —
they enable as-of-date builds); ¹ = all-roads layers, filter `RouteID LIKE
'SHS_%'` at export if ArcGIS allows (≈90% smaller), else the app filters on
load; **filename = layer name** (sheet names truncate at 31 chars, so the
FILENAME is the reliable identity; the `NN_` order prefix is fine). The
consolidator rejects files that match no expected layer name, so the drop-zone
can't accumulate dead weight.

Not needed for highways: Posted Speed Limit (no THY column), SHS Ramp / Ramp Pt
(→ the RAMPS file later), everything in IMLayers.xlsx + Master Intersection AOI
+ IM Roadway Segment (→ the INTERSECTIONS file later), and the 8 empty-in-export
layers (Vertical Grade, Horizontal Curve, 4× RumblStr, Ped Inventory, Bikeway
Inventory — header-only AND unused by THY).

## The other two clean-road files (censused 2026-07-22 — schema-level mapping)

**CA INTERSECTIONS (16,626 rows × 55 `INX_*` cols; current-only — `INX_END_DATE`
all-null)** maps onto the **IM family**: IM Intersection Detail carries the
intersection-level core (name, `Intersection_Geometry`→DESIGN, `Intersection_
Control`→CONTROL, lighted ind, all their begin dates, `Int_Date_Of_Record`→
RECORD_DATE, main+cross route/PM identity, `Cross_AADT`→XSTREET_ADT, Main/Cross
begin dates, the 285-row `INX_X_*` state-crossing block); IM Intersection
Approach Detail + Approach Segments carry the per-leg MAIN_/CROSS_ attributes
(signal mast arm, left/right channelization, flow, lanes, override lengths);
IM Intersection Route Table ties ON/AT routes; the Complex pair covers the
crossing references + influence lengths. Shared with highways: SHS Highway
Group (`INX_HIGHWAY_GROUP`), City (`INX_CITY_CODE`), SHS Population
(`INX_POPULATION_GROUP`), SHS Inv Network Date (`INX_SEG_ORDER_ID` — same
3..361700 id space as THY), Traffic Volume Segments (`INX_MAINLINE_ADT` — the
same interpolation family; note TSN carries NEGATIVE ADTs here, a data quirk
for the comparison to surface). Eff-date-class open item: `INX_LSC_DATE`.
All-null besides END_DATE: `INX_UPDATE_USER/DATE`, `INX_X_ROUTE_SUFFIX_CODE`.

**CA RAMPS (15,410 rows × 32 `RAM_*` cols; `RAM_END_DATE` all-null)** maps onto
**SHS Ramp** (`Ramp_Design`→DESIGN_CODE/DESC — 16 codes incl. Rest Area/Vista
Point/Truck Scale, `Ramp_On_Off_Ind`→ON_OFF {ON,OFF,OTH}, `Area4_Ind`,
`Ramp_Description`/`Ramp_Name`→DESCRIPTION, `InventoryItemStartDate`→
CHANGE_DATE candidate) + **SHS Ramp Pt** (the point identity: district/county/
route/suffix/PM prefix/loc/suffix + ODMeasure→BEGIN_OFFSET, Alignment) +
**Traffic Volume Ramps** (`AADT`→RAM_ADT — a direct point pull, no
interpolation) + the shared City / SHS Population / SHS Highway Group / SHS
Inv Network Date. One derivation to settle at build time:
`RAM_PRIMARY_DIRECTION_CODE` {N,S,E,W} — candidates are SHS Ramp `Side_of_Hwy`
/ Ramp Pt `Alignment` (+ the **Route Direction** layer for the route's cardinal
pair; small, never exported — added to the manifest so no second export run).

**Deliberately NOT in the library** (no clean-road column reads them; the
2026-07-21 bundles hold copies if a build-time question ever needs one):
IM Roadway Segment (810k rows — the largest layer), Master Intersection AOI,
IM Filter Points (never exported), IM Intersection Manager Metadata (2 rows),
MS2 Traffic Station, Posted Speed Limit, the HPMS family, and the 8
empty-in-export SHS layers.

**SHIPPED v0.50.0 (2026-10-23).** Both builds exist and all three Clean Road rows
compare on the Reports vs ArcGIS matrix — see
[comparison-engine.md](../comparison-engine.md) §9n. They follow the site's own
pipelines (`clean_intersection.js` / `clean_ramp.js`, 9.1 capture) rather than the
schema-level mapping above where the two differ: intersection legs by ROUTE_ID, not
LEG_TYPE; ramp lookups by the point's AR measure. The vs-TSN comparison of these builds
is not wired (no request for it).

**Re-target (2026-09-02).** The site now exports all three Clean Road files
(v0.44.0), so the CA INTERSECTIONS and CA RAMPS builds above land as ROWS of the
ArcGIS tab's **Reports vs ArcGIS matrix** (*Reports vs layers* until v0.47.0; their
builds go on the **ArcGIS reports** tab; `arcgis_reports.py` already lists
`clean_highway` / `clean_intersection` / `clean_ramp`, greyed until their builds and
a build-vs-export comparator exist — the site-export consolidators shipped in
v0.48.0): our build vs the site's export, TSMIS vs TSMIS, one build per file
stamped with the layer drop it came from. The vs-TSN comparison of the same builds
stays on the Clean Road sub-tab pattern. The measured rules above are unchanged by
this — only where the comparison lives.

**The site exports themselves (v0.48.0).** Both editions of all three files now
consolidate and compare vs the TSN extracts (`compare_clean_road_tsn`; the
`clean_intersection` / `clean_ramp` TSN slots normalize verbatim) — see
[reports.md](../reports.md#clean-road-site-exports--consolidate-and-compare-v0480).
That is the SITE's file vs TSN, separate from this doc's builds; a build's
comparison against the site's export is what the greyed rows above still wait on.

## App integration — SHIPPED v0.29.0 (the ArcGIS tab)

The sketch above became the product on 2026-07-22: `clean_road_layers.py` (the
library substrate: INDEX manifest + dialects + LRS/PM algebra + the overlay),
`clean_highway_columns.py` (the 74-column contract + per-column PROVENANCE
tiers), `consolidate_clean_highway.py` (the build), a live
`tsn_load_clean_road.build_into_highway` (verbatim normalization, marker v1),
`compare_clean_highway_tsn.py` (both flavors via `compare_tsn_common`,
role-gated by the `ArcGIS Build` marker, the no-source/TSN-internal columns as
CONTEXT), and the **ArcGIS tab** (`gui_arcgis_api.py` + `ui-arcgis.js`:
library status vs the 40-layer manifest, Build with an as-of default from the
TSN extract, Compare with formulas+values). `build/check_clean_road.py` pins
all of it on a synthetic library.

**Build-time rules SETTLED against the real 40-layer drop + THY (route 001,
five probe rounds — these supersede the open items above):**

- **Coded attribute domains arrive as `CODE- Label`** ("J- Unpaved Median",
  "N - Non-Add", bare "Z") — `clean_road_layers.code_of` extracts the code.
  Toll/Forest are presence flags ("Toll Roads" / "Yes"); Route Break carries
  "Route Break"/"Route Resume" (+step-down variants) → R-BR/U-BR.
- **THY_DISTRICT_CODE paints from every covering span's own District column**
  (the labels dialect carries it everywhere); the SHS District layer's own
  route-length spans lose middle counties to any two-ended split.
- **Cross-county spans**: end PM < begin PM is their NORMAL shape (the end
  lives in the next county's PM space — never a degenerate row). The split
  walks the county CHAIN: first county to its extent end, odometer-covered
  middle counties WHOLLY, the continuation county from 0 (measured: MON
  101.178 → SCR 0.043 covers SCR from 0.000). Odometers order and apportion,
  never join.
- **A county line where the postmiles carry on** (v0.51.1). Some routes keep one
  postmile across a county line: 166 weaves along the SB / SLO line (SB 0–8.927,
  SLO 8.927–22.893, SB 22.893–R34.987, …), and 080, 062, 120, 132, 180, 035 and
  020 have stretches like it. The county rule painted Highway Group SLO 8.927 → SB
  23.537 over SB 0–23.537 and gave 166 rows where the road is in the other county.
  A part longer than the whole span's odometer length (× 1.5, plus 0.05 mi) marks
  such a line (`_too_long`). The span then covers the road windows of both counties
  between its two postmiles (`_cover_between`), never a stretch no window covers;
  where its postmiles run backwards (16 of 472 such spans), each part stays on its
  own county's visit. On the 2026-10-22 export: 1,007 fewer build-only cells, 41
  fewer build-only rows, 31 more export rows without a build row (24 on 166, where
  the export lists stretches of a few hundred feet under SLO that Highway Group puts
  in SB). Route 009's Traffic Volume layer flips between SCR and SM along the
  ridge with postmiles 0.006 mi off the other layers', which now leaves 8 sliver
  rows.
- **A span that changes postmile prefix inside ONE county** never takes the
  county rule (v0.51.1). The second prefix space does not start at 0 there: the
  rule painted 001 LA's Travel Way L R39.741 → 40.845 (33 ft) over plain 0–40.845,
  where the layer says 25 from 3.728. The span splits where its two spaces' road
  coverage puts it, each part within the span's odometer length of its own
  endpoint. The coverage is the road's on EITHER alignment. Every left-side layer
  of 001 SLO changes prefix at the same R60.261 → M61.15, so the Left alignment
  alone had no other span saying where its R space ends.
- **A cell the postmile overlay leaves blank is read at the row's AR** (v0.51.1),
  as the site reads every attribute. A layer can file a stretch under another
  prefix than the road's (SHS Population keeps 001 SLO's R34.716–R37.079 as plain
  32.634–37.19), or a prefix-changing span can split short of a row. The row's
  midpoint is translated to AR on its roadbed (`clean_road_layers.PostmileToAr`,
  from the ends of every span read), and the layer's span there fills the cell,
  never one of the row's own postmile space. On the 2026-10-22 export it brought
  5,848 cells into agreement with the site. Of the 118 it moved away from the
  site, TSN agrees with the build on 77 (routes 101 and 099, whose export leaves
  those blocks blank), 39 have no TSN row, and 2 are worse. This is not the
  odometer trap above: the calibration is each space's own span ends, and it
  only fills a blank.
- **An independent row carries the other roadbed's block — values and date**
  (v0.51.1). An R row reads the left-side layers at its midpoint, an L row the
  right-side layers (001 MON R81.505R carries the left roadbed's 1973-06-08 and
  its lanes). The site writes the block wherever the other side's layer has a span
  there (1,817 of 1,880 independent rows on the 2026-11-03 export); TSN nulls the
  values and keeps only the date, so against TSN those cells are drift. Blanking
  them as TSN does cost 8,093 false differences on one day's comparison, and the
  raw-layer check had been skipping that class instead of judging it. Since
  v0.51.2 the Left block of a divided row is read the same way, off the Left
  layers at the row's postmile whatever the Highway Group windows say — the site
  reads it at the row's AR on that alignment: 125 SD's divided 9.161–9.251 inside
  the Left roadbed's window carries its 24 ft / 2 lanes, 004 CC's R rows past the
  Left roadbed's end at R4.629 read the divided road's left side. Clipping the
  Left layers to the Left windows had left 170 cells blank where a span exists.
- **A span's part stays off another visit of the same postmiles** (v0.51.2). A
  postmile space can be visited twice: 155 KER's R0.039 → 0.427 runs odometer
  0.402–0.829 while R0.0 → R26.564 starts at 0.829; 680 SOL's independent right
  roadbed R0.0 → M1.496 precedes the divided road's R0.627 → R2.216; 190 INY's
  R63.7 → 68.274 and 116 SON's R41.906 → 44.557 begin where the plain space's road
  ends. The part of such a span landed on the other visit's postmiles and outranked
  that visit's own spans (design speed 35 for 60, a 1971 surface for 1970's,
  680 SOL's divided rows gone under its ramp roadbed). After the split, each part is
  kept off the postmiles of any plain span of its own layer and bucket whose
  odometers overlap less than half of the postmile overlap (`_other_visit`,
  `_clip_other_visits`) — the site, reading by AR, keeps that visit's own span.
  Traffic Volume Segments stay out of it: its vintages overlap one road by design
  and the ADT reduction chooses among them.
- **Every layer's span takes the odometer the Highway Group layer says at its AR**
  (v0.51.2, `_OdByAr`). Odometers go stale too: SHS Design Speed's R23.134 → M25.307
  on 084 ALA carries 48.366 → 50.539 where its AR, and Highway Group, say 52.056 →
  54.229, and the third-space rule below painted its 50 over plain 18.976 → 21.109,
  whose odometer that is; 005 LA's .16.871 → G16.901 carries 663 miles of odometer
  for 0.028 of AR (each prefix space's own odometer at its ends) and painted its
  Highway Group and Median over every county north of it, surfacing in every gap.
  Highway Group, the layer the rows and the offsets come from, calibrates odometer
  against AR per RouteID as it is read, and every later layer's span reads its
  odometer there; a span's own odometer length is trusted only within a factor of
  two of its AR length (`_odometer_length`). A span whose AR the Highway Group does
  not reach keeps its own.
- **A prefix change can run through a third space of its county** (v0.51.2). 246
  SB's SHS Population M24.21 → R34.601 runs through plain 26.104–33.332 (odometer
  17.4–24.6 of its 15.5–25.9). The chain walk covers only whole middle counties, and
  where a space of the span's own county had Highway Group odometers it painted the
  span over the space's whole extent (Population B over 246's plain 8.302–12.466,
  odometer 0–4). Such a stretch now takes the span window by window, where the
  window's odometers lie mostly inside the span's (`_cover_end_county_spaces`, for
  prefix changes and county crossings alike); a space of the span's own counties is
  never painted whole. Rows appear where the layers had them and the build did not
  (198 TUL's plain space, 035 SM's, 080 SAC's R, 133 ORA's M, 058 KER's M, 090 LA's,
  680 SOL's N); the export has rows in each.
- **A cell read at the row's AR translates on the other alignment where its own
  has no calibration** (v0.51.2). 005 SHA's M29.303 → M30.199 exists only on the
  Left alignment; its Left rows' Right block is read on the Right RouteID at the
  row's own AR, as the site reads a right-side layer at a Left row's AR
  (`_ArFill.span`, `_sibling_rid`). A part of a prefix-changing span that stops
  within 0.005 mi of its window's edge reaches it (005 LA's Median .16.619 →
  G16.901 runs 0.278 of AR for 0.282 of postmile).
- **The build is checked against the raw layers, not only against TSN and the site**
  (v0.51.1). `tools/clean_road_layer_check.py` reads, for a row, the current plain
  span of its route, county, prefix and side that contains its midpoint, straight
  from the layer file, with none of this module's placement logic. It judged the
  v0.51.0 build as differing from its own layers on 5.6% of the cells it can judge
  and 54% of the differences in the delivered 2026-10-22 comparison as the build's
  own; the v0.51.1 build differed on 0.03% (458 cells), the v0.51.2 build on 2 cells
  of 1,459,913 (a tie-break on a 0.001-mi span at 001 SB 12.504, a backwards span
  at 004 SJ 19.84). A build that differs from its layers on more than a fraction of
  a percent of judged cells does not ship, and the three-way's "ArcGIS differs"
  count above the explained residual is a blocker, not a residual. What the check
  cannot settle is the site's own reading of a layer gap: it reads a span's end
  point as inside it and snaps to the nearest span within 0.005 mi (005 TEH 33.34,
  038 SBD 4.252), and a layer whose postmiles sit off its AR (070 BUT's inner
  shoulder L, 0.035 mi) reads one span by postmile and the other by AR.
- **Rows are NOT merged where no value changes (measured 2026-11-03, rejected).**
  2,235 of the 3,658 rows only our build has repeat the previous row's values.
  Fusing every such cut fixed about 730 END PM / LENGTH cells against the site.
  It also lost 148 rows the site has and about 730 TSN has (136 of them both): both
  cut at points where the layers show no change, even inside one network segment.
  The build keeps cutting at every span boundary. A segmentation-insensitive
  comparison, not a merge rule, is what would close this. **Closed that way on
  2026-11-17 for the vs-TSMIS comparison only** ("The vendor's source list" below):
  the comparison, not the build, joins a segment that begins where nothing the file
  reports changes, and never one the site's file has a row at, so neither loss can
  happen. TSN and the three-way still read the build as written.
- **The ADT family**: reduce the Traffic Volume spans to the WINNING vintage
  (latest AADT_YEAR, then newest slice) BEFORE the overlay — every vintage's
  edges would otherwise fabricate cuts. The value at a span's BEGIN is its
  **AADT_AHEAD** count (AADT itself is the station's midpoint — anchoring on
  it shifted every row by a constant); slope = (BACK − AHEAD)/length; P marks
  a station row only on a CONTIGUOUS stretch (a span re-entering across a PM
  gap is S — ORA 17.461). **Compared and COUNTED since v0.29.1** (owner:
  wholesale column differences are the signal, not noise to hide); the two
  known model-fit classes inside the count — TSN's profiles continue ACROSS
  county lines (THY's LA-0.0 line fits endpoints at ORA 32.953 → LA 1.2035)
  and its overlap vintage pick isn't always latest-year — are named in the
  comparison Notes; fitting them exactly is a tracked refinement.
- **No X rows are fabricated.** The HG layer never says X, and TSN skips the
  unconstructed PM ranges entirely (ORA 001 14.057–17.461 has NO row). TSN's
  340 statewide HG=X inventory rows have no layer counterpart and surface
  one-sided in the comparison, by design.
- **Offsets are PM-continued per county**: BEGIN_OFFSET = the row's own begin
  PM plus the prior counties' cumulative largest end PMs (measured: every
  first-county offset IS the begin PM, gaps and prefix handoffs included).
- **City cuts AND paints (since v0.29.1)**: the City layer (SHS rows only)
  breaks rows at city limits, and its NAME values translate to the TASAS
  letter codes via `scripts/city_codes.py` — a 395-entry table DERIVED from
  statewide co-location with the extract (21,906 rows voted, 99.92%
  agreement; the one sub-90% vote, ALHAMBRA, resolved to its majority ALH —
  its boundary rows bleed into SPAS/PAS/LA). An unmapped name passes through
  verbatim so it surfaces. THY_CITY_CODE is compared.
- **Point layers may carry no lifecycle** (Equation Points ships blank
  LRSFromDates) — an undated point is always live; equate points cut rows and
  flag THY_EQUATE_CODE=E; a break/resume point that coincides with an equate
  reads as the equate.
- **Block effective dates**: each block's date is its **PRIMARY layer's own**
  `InventoryItemStartDate` — the travel way IS the roadbed, the median layer IS
  the median; surface, shoulders, special features, barrier and curb are
  attributes hanging off it. Falls back to the NEWEST of the remaining members
  only where the primary has no covering span (2.3% / 0.1% / 0.4%).
  **This CORRECTS the original "oldest member layer" composite**, which was
  fitted to route 001 alone and was wrong for the other 251: measured over
  44–46k joined rows against the real export it took `LB Eff` 56.7%→79.0%,
  `Med Eff` 55.5%→79.4%, `RB Eff` 57.4%→79.7% (v0.39.2, DA1), and the same rule
  then won against the TSN extract too, so it is the default for BOTH builds
  (v0.39.3, DA5). Every single-layer and pairwise alternative scored lower and
  all three blocks ranked identically and independently — a rule, not a fit.
  **Still open (DA5):** neither rule explains ~40% on the TSN side, where the
  same columns cap near 60% against the extract but reach 79–80% against the
  report. Something else differs between our reconstruction and TSN's table on
  these columns; the comparison surfaces the residual honestly.

Statewide acceptance numbers (the first full build + comparison) live in the
canary bindings doc once blessed; the residual known-diff classes are the
0.001-mi boundary slivers, the eff-date composites, TSN's X rows, and the
±5–10% attribute-coverage holes on multi-county spans whose odometers are
blank (chain-walk needs them).

## The vendor's source list (2026-11-16) and what it changed

The TSMIS vendor documented where each Clean Road attribute comes from: the
feature-service layer, by name and number, for the 34 Highway and 12 Ramp attributes
of the business requirement, one source for its 32 Intersection attributes, and the
service's 97 layers and tables. The app holds it as data in
`scripts/clean_road_sources.py`; the file itself and the requirement workbook stay in
the local corpus. Read with the site's own code (the 9.1 and the 11.17 dev captures,
whose three Clean Road modules, Highway Log engine, shared plumbing and data-source
config are identical) it settles three things.

**1. The Highway and Ramp builds read the layers the list names.**
`build/check_clean_road_sources.py` holds them to it column by column: all 29 Highway
columns the list covers and all 9 Ramp columns are built from the layer AND field it
names, and every library layer carries the number the list gives it. The list numbers
its layers on a development copy of the service; the numbers are the same on the
production service the library is refreshed from (`00_INDEX.xlsx` records each one).

| | Attributes on the list | With a column in the file | Built from the listed layer and field |
|---|---:|---:|---:|
| Highway | 34 | 30 (29 columns: toll and forest share one) | 29 of 29 |
| Ramp | 12 | 8 (9 columns: design code and description) | 9 of 9 |
| Intersection | 32 | n/a | another source, below |

**2. What the list names that the file or the site's code does not carry.**

- No column in the site's file: `Posted_Speed_Limit` (Posted Speed Limit #107, in the
  library's layer list since 2026-11-18), `Facility_Name` (SHS Tolls #138),
  `Thru_Num_Lanes_L` / `_R` (#139 / #140); for Ramp `Side_of_Hwy`, `Ramp_Length`,
  `Ramp_Num_Lanes` and the `Ramp_Name` key itself.
- `Toll_Type` (SHS Tolls #138): the list says the site's code does not read it, and the
  code agrees: `THY_TOLL_FOREST_CODE` is SHS Forest HWY's flag alone, `1` on a forest
  highway. The build writes the TSN extract's codes (toll 1, forest 2). Since
  2026-11-19 the two-sided ArcGIS vs TSMIS comparison reads them as the site writes the
  column (forest 2 as 1, toll as none: the owner's rule that the TSMIS version is the
  default), which took 1,638 cells of codes out of the 2026-11-19 count. The vs-TSN
  comparison and the three-way keep showing the site's code against TASAS's (finding
  TW8).
- `RAM_CREATE_DATE`: the list names SHS Ramp's `InventoryItemStartDate`. The site
  leaves the column blank (`src: null`; blank on all 15,215 rows of the 2026-11-06
  export). The Ramp build writes it since this change, shown as context and never
  counted, the way it already writes the odometer.
- Intersection: the list names one flat `Master Intersection` layer on another
  service, with a `Major_*` / `Minor_*` field per leg. The site's code reads the IM
  tables of the main service instead (#151, #0, #146, #150 and the three span layers),
  and so does the build, which reproduces the site's file to 3,073 differing cells on
  16,381 paired rows, 2,548 of them the two begin dates. Since 2026-11-22 the library
  refreshes the Master Intersection layer too, and a second build reads it (below). The list maps the
  two begin dates to the leg's own effective date, which is what the build writes and
  the site does not (TW4).

**The three layers added to the refresh (2026-11-18, owner's call).** Calibration Point
(#1), HPMS F System (#91) and Posted Speed Limit (#107) are library layers now, last in
the list so the first 40 keep their file numbers (`arcgis_refresh.LIBRARY_LAYERS`). No
build reads them yet: the site reads its postmile equations from the calibration points
(TW10) and `THY_FUNCTIONAL_CLASS_CODE` from HPMS F System, and both wait on a first
real export to be wired and measured. Calibration Point is exported in part
(`LAYER_FILTERS`: `NetworkId = 2`, the postmile network, the rows the site's own query
asks for), because the layer holds the calibration points of every LRS network in the
state; the manifest's Data Source states the clause. The work PC owes the first refresh.

**Master Intersection, the list's Intersection source (2026-11-22).**

- **Where it is.** The owner found the production copy on the state network:
  `https://rhapps-prod.dot.ca.gov/ars/rest/services/TSMIS/csms_supporting_prod/FeatureServer/204`.
  The list links the development copy, where the same layer is number 0. Its 72 fields
  are 13 for the intersection, 29 `Major_*` and 27 `Minor_*` (each leg with its own
  route, postmile, AR route and measure; no `Minor_District`), the two inventory dates
  and `OBJECTID`. It is a point layer, not versioned, with no editor tracking and no
  GlobalID. Every field the list found is one of them, and none of the three it could
  not find is (`clean_road_sources.MASTER_INTERSECTION_HEADER`, held by
  `check_clean_road_sources`).
- **How it is made.** The state's own `Update_Reports` toolbox, kept in the corpus with
  the example ArcGIS projects, has two steps named "Update Master Intersection". Step
  3.1 recalculates the route concurrencies from the LRS AllRoads network. Step 3.3
  appends a per-user `MASTER_INTERSECTION` build into the shared `MASTER_INTERSECTION`
  table; step 3.2 is not in that toolbox. So the layer is a copy that someone refreshes
  by running tools, and it can lag the IM tables the site's code reads.
- **The site still reads the IM tables** (the 11.17 capture mentions neither the service
  nor the layer).
- **The owner's decision (2026-11-22):** *"Beside it and label it, technically since
  they are claiming that this is what they use we want to default to this when we bring
  up problems to them"*. So a second Intersection comparison, built from Master
  Intersection, sits beside the IM one, each labeled by its source. The Master
  Intersection one is the default for anything raised with the vendor: first on the
  matrix and the lite copy that is sent. Its files are `clean_intersection_mi_vs_ArcGIS
  …`, so the IM-tables row keeps its key and its file names. The
  IM one stays as the cross-check of what the site's code reads. A difference that only
  the Master Intersection comparison shows lies in that copy, not in the report.
- **Done (v0.56.0):** the layer refresh reads it as library layer 44 from its own
  service (`arcgis_refresh.LAYER_SERVICES`; the worker reads that service's layer list
  too, and the manifest's Data Source records it). The builds' service note leaves it
  out (`clean_road_sources.library_service`). The work PC's first refresh, 2026-11-22:
  both layer lists read, the layer matched by name at 204, 24,296 rows (the server's
  count) in 40 s.
- **Done: the build and its comparison rows** (`arcgis_report_clean_intersection_mi`,
  [comparison-engine.md](../comparison-engine.md) §9n has the rules and the measured
  2026-11-19 numbers). Read first, from that real export: its IDs are the IM tables'
  INTERSECTION_IDs (18,831 of its 18,839 current rows are live IM intersections); it
  keeps 5,457 replaced rows with an end date and 2,727 current rows with no attributes;
  on the 16,028 rows the 2026-11-19 file lists under the major route it agrees with
  the file on 99% or more of most columns. The leg's lanes column is
  `Total_Num_Lanes` (99.0% against 2.4% for Pri). On that export the Master
  Intersection row reads 5,585 differing cells and 87 / 131 one-sided rows, the
  IM-tables row 3,071 and 94 / 3.

**3. Which data each side came from is stated, not assumed.** A build records the
feature service its layers were pulled from (the marker sheet's `Layer service`), and a
vs-TSMIS comparison prints it beside the data source the export's file name records. An
SSOR prod export against the default library reads the same service. An ARS prod export
is the same data up to a day behind (the owner, 2026-11-19: "the only difference
between it and ssor is ars takes 24 hrs to update"; the layers exist only on SSOR, and
exports are moving to ARS), and the Notes say so: a record edited in that day can
differ only because ARS had not caught up. An export from another environment is
flagged as a different service.

### Segments, read the way the file is defined

The business requirement asks for "homogenous segments of all the highway inventory
data items" that "break and restart at the county boundary" (items 8 and 10). The
site's Highway Log pipeline, the engine the list names for this file, works that way:
it drops an attribute-change boundary at which no reported field changes, and the four
block effective dates are not among the fields (`highway_log.js`). The build cuts at
every layer breakpoint, as the TSN table does.

So the **vs-TSMIS comparison** (`compare_clean_highway_arcgis._join_unchanged`) joins a
build segment to the one that ends where it begins when:

- it carries no landmark, equation or break,
- no counted column differs across the boundary, the END postmile, the length and the
  four block effective dates aside, and
- the site's file has no row keyed at that postmile.

The last condition is what the 2026-11-03 merge experiment lacked: a boundary the site
also cuts stays a paired row, so its cells still compare. A joined boundary carries no
attribute change, so no difference a reader could act on is hidden. A landmark,
equation or break belongs to the postmile a segment begins at, so the segment being
joined may carry none, while the one it joins may (and keeps it). The build workbook,
the vs-TSN comparison and the three-way are untouched. The Notes sheet states both
counts.

**Measured** on the 2026-11-06 dev-site export against the 2026-10-30 layers, build as
of 2026-11-06:

| | Without the join | With it |
|---|---:|---:|
| Paired segments | 49,146 | 49,146 |
| Only in the build | 8,174 | 3,744 |
| Only in the site's file | 4,065 | 4,065 |
| Differing cells | 48,204 | 41,865 |
| Segments joined | | 4,430 (3,794 nothing changes, 636 a date only) |

The 6,339 fewer cells are END PM and LENGTH on the segments that used to stop at a cut
the site does not have. What the remaining one-sided rows are is roadmap TW15. One more
thing the same data showed: the site's file is not the same from run to run on its
longest routes (TW13), so part of every count on those routes is the site's own.

### Points, read the way the site lists them (2026-11-18)

The site's file is the Highway Log's row list written flat (`clean_highway.js`
`clh_buildSegments`): one row for every landmark of the route (SHS Landmark #123), each
side of every postmile equation, the end of each county the route runs on from, each
city's begin, and each place an attribute changes. The build made a row only where a
Highway Group span put a segment, and its point layers did not read a point's roadbed.
Since 2026-11-18 it follows the site (`consolidate_clean_highway`, `site_points`, build
version 3):

- **A point belongs to its own roadbed** (its PM suffix). A landmark, equation or route
  break on the left roadbed of an independent alignment cuts and labels the L rows
  alone. Before, every point cut and labelled both roadbeds: 823 of the build-only rows
  of 2026-11-17.
- **Every landmark and each side of an equation is a row.** Where no segment of that
  roadbed begins there (the last postmile of a space, the plain road where an
  independent alignment begins, a place with no Highway Group span) it is a
  zero-length row. Several landmarks at one postmile are a row each in AR order, the
  last one the segment's (272 of the 293 such postmiles of the 2026-11-06 export follow
  that order). An equation's row carries the landmarks of every roadbed at its
  postmile, joined with ` / `, and they make no rows of their own (the site folds a
  landmark into an equation by PM prefix + postmile alone: 001 MON R82.739's
  left-roadbed `END LT INDEP ALIGN`).
- **A county the route runs on from ends in a zero-length row** carrying its last
  segment's values (`hl_queryCountyEnds`).
- **The names the site leaves out are left out**: a `BEGIN` / `END REALIGNMENT` marker
  where a named landmark shares its postmile (under the prefixes the site's pattern
  knows: none, H, M, N, R), and `SELF INTERSECT`.

A zero-length row is read at the point the way the site reads a row there
(`_seg_at_point`, `_ArFill.at`, from `queryRangeLayer` / `hl_queryRangeLayerS` /
`hl_queryLayer122Field`): the spans that hold the measure, ends included, with one that
begins within 1e-5 past it, and of those the one that begins last, so the last postmile
of a space reads the stretch that follows it; the nearest span within 0.005 where none
holds it; the left-roadbed layers at the postmile's own measure on the left alignment;
the network date and segment order by the spans that hold the measure and nothing
more; a landmark's own inventory date where no network-date span holds the row. A
landmark is placed by the layer's own RouteID and AR measure, an equation by its
postmile (277 of 743 equation rows agree on every value that way, 188 placed by their
landmark, 84 read a hair before). Segment rows keep the offsets and break markers they
had: a point row takes no part in either.

**In the vs-TSMIS comparison** (`compare_clean_highway_arcgis`):

- A `BEGIN` / `END REALIGNMENT` name reads with the row's PM prefix, as the site
  spells it (`END R REALIGNMENT`). The layer's own name has none.
- A zero-length row sits where two stretches meet, and either stretch's value is right
  at that point. Which one the site prints turns on a difference below its own 1e-5
  tolerance between the LRS's translation of the postmile and the layers' stored
  measures (001 SLO R7.531 reads the stretch that begins there, 001 MEN R56.064 the one
  that ends), which no layer file carries. So a cell of a zero-length row agrees when
  it shows the stretch that begins there or the one that ends there
  (`_settle_point_rows`). A value that is neither still counts.
- A zero-length row is never joined into a segment.

**Measured** on the 2026-11-06 dev-site export against the 2026-10-30 layers, build as
of 2026-11-06:

| | 2026-11-17 (segments joined) | 2026-11-18 (points) |
|---|---:|---:|
| Paired rows | 49,146 | 51,322 |
| Only in the build | 3,744 | 3,180 |
| Only in the site's file | 4,065 | 1,889 |
| Differing cells | 41,865 | 45,236 |
| Zero-length rows in the build | 0 | 2,365 |
| ... of them reading the ending stretch in some cell | | 366 rows, 1,532 cells |

The cells rise because 2,176 more rows are compared. On the newly paired
zero-length rows the largest columns are END PM and length (the site gives a
zero-length row the END of its next row, and its landmark row is the first segment of
the following space where the build keys that segment in its own space).

**Against the raw layers** (`tools/clean_road_layer_check.py file`, which shares none of
the build's logic): the build's segment rows differ from their own layers on 2 cells of
1,440,106 judged, the same two as before the change (TW12). It has 774 fewer segment
rows than the 2026-11-17 build, the cuts a point on the other roadbed used to make. Its
2,365 zero-length rows are judged apart, because that check reads a row's own county +
prefix space alone: 40,383 cells equal the stretch there and 6,713 read the stretch on
the other side of the postmile, which is the reading described above.

**Against TSN.** The TSN extract has no zero-length row at all (0 of 60,083), so most
of the build's stand alone on the ArcGIS side there, by the new file's layout. Measured
with the build as of 2025-09-08 against the extract, keyed pairing in memory: 235 more
rows pair (52,927 to 53,162), TSN-only rows fall 7,156 to 6,921, and the build has
2,429 zero-length rows, 1,789 of them with no TSN row. The vs-TSN and three-way Notes say so. **Highway Detail** keeps
the rule it was measured with (`site_points=False`): the build it reads is byte-equal
to the one the old code wrote for 2026-11-06, checked on the real layers.

### What is left on the Highway file, and whose it is (2026-11-18)

Every class below was measured on the 2026-11-06 export. The build is the LRS as of
the day; none of these is a build artifact any more.

| Class | Rows | Cause |
|---|---:|---|
| Only in the build: a layer the site does not scan changes | 660 | The site collects change points from 18 layers (`hl_queryAttributeChangePairs`). SHS Special Feature R #135 (290 rows), a city's END (356: only city begins are rows), tolls and forest (10) and non-add mileage (2) are not among them, though `clean_highway.js` lists every one of those columns as a record trigger. |
| Only in the build: a span ends and none follows | 200 | The site takes a span's begin as a change point, never its end. Where a median, barrier or left-roadbed span stops and the layer has a gap, the file carries the old value on (080 CC 2.538: the barrier ends, the site's row runs to 2.565 with it). |
| Only in the build: the left roadbed of an independent alignment | 142 | On `_S` routes the site scans the five left-roadbed layers alone, so a median or barrier change along the left roadbed makes no row (080 NEV R5.934L to R8.835L). |
| Only in the build: the site's duplicate check has no county | 180 | The site keeps one attribute-change point per PM prefix + postmile + suffix on a ROUTE and drops one whose key any landmark on the route holds (`_pmKeyOf`, `_attrPmBest`, `basePmKeys`). A change at CC 13.486 is lost when another county has a landmark at 13.486. 6.7% of the missing rows against 0.08% of paired ones (landmark) and 3.9% against 0.3% (another change). |
| Only in the build: lost on the day | 266 | 121 are in the site's file on 10-22, 10-30 and 11-03 and missing on 11-06 (78 on route 002); 145 come and go (route 001). TW13. |
| The same boundary, keyed up to 0.005 mi apart | 245 pairs | The build keys on the span's stored postmile, the site on the LRS's translation of its AR measure. |
| Only in the build: the first segment of a postmile space | 321 | The site lists it under another key: its row at that place is the landmark, equation or county end, which keeps its own postmile identity, and the change point at the same odometer is dropped as a duplicate. 120 match a site row of another prefix or roadbed exactly by END and length (001 SLO plain 7.531 to 7.59 is the site's `R7.531`); 146 are a county's first segment with no site row ending there in that county, the ends-of-county finding of 2026-11-13 seen from the build; 55 end where a site row of another length ends. |
| Only in the site's file: an attribute row the build has no boundary for | 1,222 | 1,134 of them lie under a retired LRS row. The site's range queries filter on inventory dates only, so beside 172,749 current rows of the 18 layers they keep 434,092 retired ones (`LRSToDate` set), which `_lrsDedupRetired` removes only where a current row has the same postmile extent and start date. Where a recalibration came with new inventory the retired rows stay, at stale AR measures (TW11 is this, seen in the values). 462 are on route 001. |
| Only in the site's file: equations | 275 | Equations the Equation Points layer does not carry (TW10). The site pairs the postmile network's calibration points; the library gets that layer on its next refresh. |
| Only in the site's file: no postmile | 186 | The site could not translate the row's AR to a postmile. Keyed on the odometer, unpairable. |

Not yet measured cell by cell: how many of the differing cells on paired rows are the
retired rows above. The rate is 0.71 differing value cells a row where a retired row
lies over the place and 0.42 where none does.
