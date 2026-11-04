"""Check a Clean Road Highway file, or a Reports vs ArcGIS Highway workbook,
against the raw ArcGIS layer files directly, the way a person checks one cell by
hand: open the layer, keep the rows current at the as-of date whose begin AND
end county and postmile prefix are both the row's, on the row's side, and read
the one whose postmiles contain the row's midpoint. It shares none of the
build's placement, overlay, split or AR logic (only the county name table), so
it can say whether a difference the comparison shows is the report's or the
build's. It covers the 28 span-layer columns: the roadbed values, median,
access, terrain, design speed, population and the four block dates. A cell
whose own space has no plain span (a layer that files a stretch under another
prefix, a crossing span) is reported as undetermined, never guessed.

Usage (from the repo root):

  python tools/clean_road_layer_check.py index  <layer library> <as-of> <cache.pkl>
  python tools/clean_road_layer_check.py file   <clean_road_highway.xlsx> <cache.pkl>
  python tools/clean_road_layer_check.py workbook <clean_highway_vs_layers.xlsx> <cache.pkl>
  python tools/clean_road_layer_check.py spot   <clean_highway_vs_layers.xlsx> <layer library>
                                                <as-of> <raw per-route export folder>
                                                [--n-diff 25] [--n-agree 15] [--seed 1] [--out spot.txt]

`file` judges every row of an export or a build (file = layer / file != layer /
no plain span); `workbook` judges every flagged cell of a comparison (false
difference = the layer says what TSMIS says, so the build was wrong; true
difference = the layer says what the build says) and a 5% sample of the
agreeing cells; `spot` draws random differing and agreeing cells and prints,
for each, the raw TSMIS per-route row, every raw layer row of that route near
that postmile (OBJECTID, postmiles, LRS dates) and the verdict a person reaches
from them. The as-of is the library's export date (its `_refresh/last_run.json`).
"""
import argparse
import collections
import datetime as dt
import pickle
import random
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from clean_road_layers import COUNTY_CODES  # noqa: E402  (a name table, not logic)

LAYERS = {
    "MED": ("21_SHS Median", ["Median_Type", "Median_Width", "Median_Variance"], "whole"),
    "BAR": ("30_SHS Barrier", ["Barrier_Type"], "whole"),
    "CURB": ("29_SHS Curb Landscape", ["Curb_Landscape"], "whole"),
    "ACC": ("31_SHS Access Control", ["SHS_Access_Control"], "whole"),
    "TER": ("10_Terrain Type", ["Terrain_Type"], "whole"),
    "DSP": ("28_SHS Design Speed", ["Design_Speed"], "whole"),
    "POP": ("17_SHS Population", ["Population_Code"], "whole"),
    "TWR": ("07_SHS Travel Way R", ["Travel_Way_Width_R", "Total_Num_Lanes_R"], "right"),
    "SUR": ("12_SHS Surface Type R", ["Surface_Type_R"], "right"),
    "SPR": ("14_SHS Special Feature R", ["Special_Feature_Type_R"], "right"),
    "OSR": ("18_SHS O Shld Width R", ["Shld_Width_Total_Out_R", "Shld_Width_Treated_Out_R"], "right"),
    "ISR": ("23_SHS I Shld Width R", ["Shld_Width_Total_In_R", "Shld_Width_Treated_In_R"], "right"),
    "TWL": ("08_SHS Travel Way L", ["Travel_Way_Width_L", "Total_Num_Lanes_L"], "left"),
    "SUL": ("13_SHS Surface Type L", ["Surface_Type_L"], "left"),
    "SPL": ("15_SHS Special Feature L", ["Special_Feature_Type_L"], "left"),
    "OSL": ("19_SHS O Shld Width L", ["Shld_Width_Total_Out_L", "Shld_Width_Treated_Out_L"], "left"),
    "ISL": ("24_SHS I Shld Width L", ["Shld_Width_Total_In_L", "Shld_Width_Treated_In_L"], "left"),
}
COLS = {
    "THY_MEDIAN_TYPE_CODE": ("MED", "code", 0), "THY_MEDIAN_WIDTH_AMT": ("MED", "num", 1),
    "THY_MEDIAN_WIDTH_VAR_CODE": ("MED", "code", 2), "THY_MEDIAN_BARRIER_CODE": ("BAR", "code", 0),
    "THY_CURB_LANDSCAPE_CODE": ("CURB", "code", 0), "THY_HIGHWAY_ACCESS_CODE": ("ACC", "code", 0),
    "THY_TERRAIN_CODE": ("TER", "code", 0), "THY_DESIGN_SPEED_AMT": ("DSP", "num", 0),
    "THY_POPULATION_CODE": ("POP", "code", 0),
    "THY_RT_TRAV_WAY_WIDTH_AMT": ("TWR", "num", 0), "THY_RT_LANES_AMT": ("TWR", "num", 1),
    "THY_RT_SURF_TYPE_CODE": ("SUR", "code", 0), "THY_RT_SPEC_FEATURES_CODE": ("SPR", "code", 0),
    "THY_RT_O_SHD_TOT_WIDTH_AMT": ("OSR", "num", 0), "THY_RT_O_SHD_TRT_WIDTH_AMT": ("OSR", "num", 1),
    "THY_RT_I_SHD_TOT_WIDTH_AMT": ("ISR", "num", 0), "THY_RT_I_SHD_TRT_WIDTH_AMT": ("ISR", "num", 1),
    "THY_LT_TRAV_WAY_WIDTH_AMT": ("TWL", "num", 0), "THY_LT_LANES_AMT": ("TWL", "num", 1),
    "THY_LT_SURF_TYPE_CODE": ("SUL", "code", 0), "THY_LT_SPEC_FEATURES_CODE": ("SPL", "code", 0),
    "THY_LT_O_SHD_TOT_WIDTH_AMT": ("OSL", "num", 0), "THY_LT_O_SHD_TRT_WIDTH_AMT": ("OSL", "num", 1),
    "THY_LT_I_SHD_TOT_WIDTH_AMT": ("ISL", "num", 0), "THY_LT_I_SHD_TRT_WIDTH_AMT": ("ISL", "num", 1),
    "THY_LEFT_ROAD_EFF_DATE": ("TWL", "date", -1), "THY_RIGHT_ROAD_EFF_DATE": ("TWR", "date", -1),
    "THY_MEDIAN_EFF_DATE": ("MED", "date", -1), "THY_ACCESS_EFF_DATE": ("ACC", "date", -1),
}
KEY_COLS = ["THY_COUNTY_CODE", "THY_ROUTE_NAME", "THY_ROUTE_SUFFIX_CODE", "THY_PM_PREFIX_CODE",
            "THY_END_PM_AMT", "THY_PM_SUFFIX_CODE"]
_CODE = re.compile(r"^\s*([A-Z0-9]{1,2})\s*-\s?")
_PM = re.compile(r"-?\d+(?:\.\d+)?")


# --------------------------------------------------------------------------- #
# values as a person writes them down
# --------------------------------------------------------------------------- #
def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s[:19], fmt).date()
        except ValueError:
            pass
    return None


def norm(kind, v):
    """The code before the dash, a whole number without leading zeros, a date
    as YYYY-MM-DD, blank as ''."""
    if v is None:
        return ""
    if kind == "date":
        d = to_date(v)
        return d.isoformat() if d else str(v).strip()
    s = str(v).strip()
    if s == "" or s.lower() == "none":
        return ""
    if kind == "code":
        m = _CODE.match(s)
        return m.group(1) if m else s
    try:
        f = float(s)
    except ValueError:
        return s
    return str(int(f)) if f == int(f) else str(f)


def pm_of(v):
    m = _PM.search(str(v)) if v is not None else None
    return float(m.group()) if m else None


def county_code(name):
    return COUNTY_CODES.get(str(name or "").strip().upper())


def route_number(value):
    """The route's number from a route name with or without its suffix ('005', '005S')."""
    m = re.match(r"\s*(\d+)", str(value or ""))
    return int(m.group(1)) if m else None


def row_query(route_name, route_suffix, county, prefix, begin, end, pm_suffix):
    """The lookup key of a Clean Road row from its own cells, or None."""
    route = route_number(route_name)
    if route is None:
        return None
    kind = str(pm_suffix or "").strip().upper()
    kind = kind if kind in ("R", "L") else "base"
    b, e = pm_of(begin), pm_of(end)
    if b is None:
        return None
    return (route, str(route_suffix or "").strip().upper(),
            str(county or "").strip().rstrip(".").upper(),
            str(prefix or "").strip().upper(), kind, b, b if e is None else e)


# --------------------------------------------------------------------------- #
# the raw-layer index
# --------------------------------------------------------------------------- #
def build_index(lib_root, asof, cache=None):
    """{tag: {(route, suffix, county, prefix, align): [(b, e, vals, item_start,
    OBJECTID)]}} over the plain (one county, one prefix) spans current at `asof`."""
    lib_root, asof = Path(lib_root), to_date(asof)
    if cache and Path(cache).exists():
        return pickle.loads(Path(cache).read_bytes())
    idx = {}
    for tag, (stem, attrs, _fam) in LAYERS.items():
        wb = openpyxl.load_workbook(lib_root / f"{stem}.xlsx", read_only=True, data_only=True)
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        hdr = [str(c).strip() if c is not None else "" for c in next(it)]
        ix = {h: i for i, h in enumerate(hdr)}
        spaces = {}
        kept = total = 0
        for r in it:
            if not r:
                continue
            total += 1
            frm, to = to_date(r[ix["LRSFromDate"]]), to_date(r[ix["LRSToDate"]])
            if frm is None or frm > asof or (to is not None and asof >= to):
                continue
            c1, c2 = county_code(r[ix["BeginCounty"]]), county_code(r[ix["EndCounty"]])
            p1 = str(r[ix["BeginPMPrefix"]] or "").strip().replace(".", "").upper()
            p2 = str(r[ix["EndPMPrefix"]] or "").strip().replace(".", "").upper()
            b, e = pm_of(r[ix["BeginPMMeasure"]]), pm_of(r[ix["EndPMMeasure"]])
            if c1 is None or c1 != c2 or p1 != p2 or b is None or e is None:
                continue
            try:
                route = int(r[ix["RouteNum"]])
            except (TypeError, ValueError):
                continue
            suffix = str(r[ix["RouteSuffix"]] or "").strip().replace(".", "").upper()
            align = str(r[ix["Alignment"]] or "").strip()[:1].upper()
            item = to_date(r[ix["InventoryItemStartDate"]]) if "InventoryItemStartDate" in ix else None
            spaces.setdefault((route, suffix, c1, p1, align), []).append(
                (min(b, e), max(b, e), tuple(r[ix[a]] for a in attrs), item,
                 r[ix["OBJECTID"]] if "OBJECTID" in ix else None))
            kept += 1
        wb.close()
        for v in spaces.values():
            v.sort()
        idx[tag] = spaces
        print(f"  {stem}: {kept:,} plain current spans of {total:,} rows", flush=True)
    if cache:
        Path(cache).write_bytes(pickle.dumps(idx))
    return idx


def lookup(idx, col, route, suffix, county, prefix, kind, b, e):
    """('ok', value, (b, e, OBJECTID)) | ('none',) | ('multi', [values]) |
    ('skip',) for one column on a row. A left-roadbed column on an R row (and
    the reverse) is blank by the file's convention, so it is skipped."""
    tag, vkind, pos = COLS[col]
    fam = LAYERS[tag][2]
    if fam == "left":
        align = "L"
    elif fam == "right":
        align = "R"
    else:
        align = "L" if kind == "L" else "R"
    spans = idx[tag].get((route, suffix, county, prefix, align), ())
    if not spans:
        return ("none",)
    mid = (b + e) / 2
    hits = [s for s in spans if s[0] <= mid < s[1]]
    if not hits and b == e:
        hits = [s for s in spans if s[0] <= mid <= s[1]]
    if not hits:
        return ("none",)
    vals = sorted({norm(vkind, s[3] if pos < 0 else s[2][pos]) for s in hits})
    if len(vals) > 1:
        return ("multi", vals)
    return ("ok", vals[0], (hits[0][0], hits[0][1], hits[0][4]))


def classify(idx, q, col, a_val, b_val):
    """(verdict, layer value, span) for one cell: a_val is the build's, b_val
    TSMIS's (equal for an agreeing cell)."""
    res = lookup(idx, col, *q)
    if res[0] != "ok":
        return (res[0], res[1] if len(res) > 1 else None, None)
    truth, ref = res[1], res[2]
    kind = COLS[col][1]
    a, b = norm(kind, a_val), norm(kind, b_val)
    if a == b:
        return ("agree-true" if truth == a else "agree-false", truth, ref)
    if truth == b:
        return ("false-diff", truth, ref)
    if truth == a:
        return ("true-diff", truth, ref)
    return ("neither", truth, ref)


# --------------------------------------------------------------------------- #
# reading a comparison workbook
# --------------------------------------------------------------------------- #
def _side(v, which):
    """The build (0) or TSMIS (1) half of a Comparison cell; a plain cell is both."""
    s = "" if v is None else str(v)
    if " ≠ " in s:
        s = s.split(" ≠ ", 1)[which]
    return "" if s == "(blank)" else s


def paired_rows(cmp_path):
    """Yield (key, {column: (build value, TSMIS value)}) per paired row of a
    Reports vs ArcGIS Highway workbook."""
    wb = openpyxl.load_workbook(cmp_path, read_only=True)
    ws = wb["Comparison"]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    col = {h: i for i, h in enumerate(hdr)}
    fields = [h for h in hdr if isinstance(h, str) and h.startswith("THY_") and h != "THY_BEGIN_PM_AMT"]
    try:
        for r in it:
            if r[col["Status"]] != "Both":
                continue
            key = {k: (_side(r[col[k]], 0), _side(r[col[k]], 1)) for k in KEY_COLS}
            key["route"] = r[col["Route"]]
            key["begin"] = str(r[col["THY_BEGIN_PM_AMT"]]).split(" / ")[-1]
            cells = {}
            for f in fields:
                v = r[col[f]]
                if v is not None:
                    cells[f] = (_side(v, 0), _side(v, 1))
            yield key, cells
    finally:
        wb.close()


def key_query(key):
    return row_query(key["route"], key["THY_ROUTE_SUFFIX_CODE"][1], key["THY_COUNTY_CODE"][1],
                     key["THY_PM_PREFIX_CODE"][1], key["begin"], key["THY_END_PM_AMT"][1],
                     key["THY_PM_SUFFIX_CODE"][1])


# --------------------------------------------------------------------------- #
# the three commands
# --------------------------------------------------------------------------- #
def cmd_file(xlsx, idx):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = [str(c) if c is not None else "" for c in next(it)]
    ix = {h: i for i, h in enumerate(hdr)}
    cols = [c for c in COLS if c in ix]
    verdicts = collections.Counter()
    by_col = collections.defaultdict(collections.Counter)
    n = 0
    for r in it:
        if not r or r[ix["THY_ROUTE_NAME"]] in (None, ""):
            continue
        n += 1
        q = row_query(r[ix["THY_ROUTE_NAME"]], r[ix["THY_ROUTE_SUFFIX_CODE"]], r[ix["THY_COUNTY_CODE"]],
                      r[ix["THY_PM_PREFIX_CODE"]], r[ix["THY_BEGIN_PM_AMT"]], r[ix["THY_END_PM_AMT"]],
                      r[ix["THY_PM_SUFFIX_CODE"]])
        if q is None:
            verdicts["unkeyed"] += 1
            continue
        for c in cols:
            v = classify(idx, q, c, r[ix[c]], r[ix[c]])[0]
            verdicts[v] += 1
            by_col[c][v] += 1
    wb.close()
    print(f"{Path(xlsx).name}: {n:,} rows")
    judged = verdicts["agree-true"] + verdicts["agree-false"]
    for k, m in verdicts.most_common():
        print(f"  {k:12s} {m:9,d}")
    print(f"  judged {judged:,}; the file differs from the raw layer on "
          f"{100 * verdicts['agree-false'] / max(judged, 1):.2f}% of them")
    print("  column                          file=layer  file!=layer  no plain span  several")
    for c in cols:
        v = by_col[c]
        print(f"  {c:30s} {v['agree-true']:10d} {v['agree-false']:12d} {v['none']:14d} {v['multi']:8d}")


def cmd_workbook(cmp_path, idx):
    rng = random.Random(0)
    flagged, agree = collections.Counter(), collections.Counter()
    by_col = collections.defaultdict(collections.Counter)
    n_rows = 0
    for key, cells in paired_rows(cmp_path):
        n_rows += 1
        q = key_query(key)
        sample = rng.random() < 0.05
        for c, (a, b) in cells.items():
            if c not in COLS:
                if a != b:
                    flagged["not a span-layer column"] += 1
                continue
            if q is None:
                flagged["unkeyed"] += 1
                continue
            if a != b:
                v = classify(idx, q, c, a, b)[0]
                flagged[v] += 1
                by_col[c][v] += 1
            elif sample:
                agree[classify(idx, q, c, a, b)[0]] += 1
    judged = sum(flagged[k] for k in ("false-diff", "true-diff", "neither"))
    print(f"{Path(cmp_path).name}: {n_rows:,} paired rows")
    print("Flagged cells (what the workbook shows as differences):")
    for k, m in flagged.most_common():
        print(f"  {k:26s} {m:9,d}")
    print(f"  judged {judged:,}: false (the layer says what TSMIS says; the build was wrong) "
          f"{100 * flagged['false-diff'] / max(judged, 1):.1f}%, true "
          f"{100 * flagged['true-diff'] / max(judged, 1):.1f}%, all three apart "
          f"{100 * flagged['neither'] / max(judged, 1):.1f}%")
    print(f"Agreeing cells, 5% row sample: {dict(agree)}")
    print("  column                         false-diff  true-diff  neither  none  multi")
    for c in sorted(by_col, key=lambda c: -by_col[c]["false-diff"]):
        v = by_col[c]
        print(f"  {c:30s} {v['false-diff']:10d} {v['true-diff']:10d} {v['neither']:8d} "
              f"{v['none']:5d} {v['multi']:6d}")


def cmd_spot(cmp_path, lib, asof, raw_dir, n_diff, n_agree, seed, out_path):
    lib, asof, raw_dir = Path(lib), to_date(asof), Path(raw_dir)
    rng = random.Random(seed)
    out = open(out_path, "w", encoding="utf-8") if out_path else None

    def say(*a):
        s = " ".join(str(x) for x in a)
        print(s)
        if out:
            out.write(s + "\n")

    diff_pool, agree_pool, seen_d, seen_a = [], [], 0, 0
    for key, cells in paired_rows(cmp_path):
        for c, (a, b) in cells.items():
            if c not in COLS:
                continue
            if a != b:
                seen_d += 1
                if len(diff_pool) < 400:
                    diff_pool.append((key, c, a, b))
                else:
                    j = rng.randrange(seen_d)
                    if j < 400:
                        diff_pool[j] = (key, c, a, b)
            else:
                seen_a += 1
                if len(agree_pool) < 400:
                    agree_pool.append((key, c, a, b))
                else:
                    j = rng.randrange(seen_a)
                    if j < 400:
                        agree_pool[j] = (key, c, a, b)
    picks = (rng.sample(diff_pool, min(n_diff, len(diff_pool)))
             + rng.sample(agree_pool, min(n_agree, len(agree_pool))))
    say(f"# Spot check of {Path(cmp_path).name}")
    say(f"# {seen_d:,} differing and {seen_a:,} agreeing cells in the span-layer columns; "
        f"{min(n_diff, len(diff_pool))} + {min(n_agree, len(agree_pool))} drawn at random (seed {seed}).")
    say(f"# Layers: {lib} as of {asof}. Raw TSMIS rows: {raw_dir}")

    routes = {route_number(k["route"]) for k, *_ in picks}
    layer_rows = {}
    for tag in {COLS[c][0] for _k, c, *_ in picks}:
        stem, attrs, _fam = LAYERS[tag]
        wb = openpyxl.load_workbook(lib / f"{stem}.xlsx", read_only=True, data_only=True)
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        h = [str(c).strip() if c is not None else "" for c in next(it)]
        ix = {x: i for i, x in enumerate(h)}
        want = ("OBJECTID", "RouteSuffix", "Alignment", "BeginCounty", "BeginPMPrefix", "BeginPMMeasure",
                "EndCounty", "EndPMPrefix", "EndPMMeasure", "InventoryItemStartDate",
                "InventoryItemEndDate", "LRSFromDate", "LRSToDate", *attrs)
        rows = collections.defaultdict(list)
        for r in it:
            if not r:
                continue
            try:
                rt = int(r[ix["RouteNum"]])
            except (TypeError, ValueError):
                continue
            if rt in routes:
                rows[rt].append({x: r[ix[x]] for x in want if x in ix})
        wb.close()
        layer_rows[tag] = rows

    raw_cache = {}

    def raw_rows(route, suffix):
        k = (route, suffix)
        if k not in raw_cache:
            fs = sorted(raw_dir.glob(f"*clean_highway_route_{route:03d}{suffix}.xlsx"))
            rows = []
            if fs:
                wb = openpyxl.load_workbook(fs[0], read_only=True)
                ws = wb.worksheets[0]
                it = ws.iter_rows(values_only=True)
                h = [str(c) if c is not None else "" for c in next(it)]
                rows = [(n, dict(zip(h, r))) for n, r in enumerate(it, start=2)]
                wb.close()
            raw_cache[k] = (fs[0].name if fs else None, rows)
        return raw_cache[k]

    tally = collections.Counter()
    for i, (key, c, a, b) in enumerate(picks, start=1):
        route, suffix = route_number(key["route"]), key["THY_ROUTE_SUFFIX_CODE"][1]
        county, prefix = key["THY_COUNTY_CODE"][1], key["THY_PM_PREFIX_CODE"][1]
        begin, end = pm_of(key["begin"]), pm_of(key["THY_END_PM_AMT"][1])
        kind = key["THY_PM_SUFFIX_CODE"][1]
        tag, vkind, pos = COLS[c]
        attrs = LAYERS[tag][1]
        mid = (begin + (begin if end is None else end)) / 2
        say("")
        say(f"## {i}. {'DIFFERING' if a != b else 'AGREEING'} cell — {route:03d}{suffix} {county} "
            f"{prefix}{begin:g}{kind or ''} → {'' if end is None else f'{end:g}'}  column {c}")
        say(f"   comparison shows: ArcGIS build = {a!r}   TSMIS = {b!r}")
        fname, rows = raw_rows(route, suffix)
        hit = [(n, r) for n, r in rows
               if str(r.get("THY_COUNTY_CODE") or "").rstrip(".") == county
               and pm_of(r.get("THY_BEGIN_PM_AMT")) == begin
               and str(r.get("THY_PM_PREFIX_CODE") or "").strip() == prefix
               and str(r.get("THY_PM_SUFFIX_CODE") or "").strip() == (kind or "")]
        loose = not hit
        if loose:
            hit = [(n, r) for n, r in rows if str(r.get("THY_COUNTY_CODE") or "").rstrip(".") == county
                   and pm_of(r.get("THY_BEGIN_PM_AMT")) == begin]
        tsmis_raw = None
        if hit:
            n, r = hit[0]
            say(f"   raw TSMIS file {fname} row {n}: {c} = {r.get(c)!r}  (prefix {r.get('THY_PM_PREFIX_CODE')!r}, "
                f"END PM {r.get('THY_END_PM_AMT')!r}, suffix {r.get('THY_PM_SUFFIX_CODE')!r}"
                + (", matched by county + begin PM only" if loose else "") + ")")
            tsmis_raw = norm(vkind, r.get(c))
        else:
            say(f"   raw TSMIS file {fname}: row NOT found by county + begin PM")
        say(f"   raw layer {LAYERS[tag][0]}.xlsx, current rows of route {route:03d}{suffix} in {county} "
            f"within 0.3 mi of PM {mid:g}:")
        want_align = "L" if (LAYERS[tag][2] == "left" or kind == "L") else "R"
        truth = []
        near = []
        for r in layer_rows[tag].get(route, []):
            if str(r.get("RouteSuffix") or "").replace(".", "").strip().upper() != suffix:
                continue
            frm, to = to_date(r["LRSFromDate"]), to_date(r["LRSToDate"])
            if frm is None or frm > asof or (to is not None and asof >= to):
                continue
            c1, c2 = county_code(r.get("BeginCounty")), county_code(r.get("EndCounty"))
            if county not in (c1, c2):
                continue
            lo, hi = pm_of(r["BeginPMMeasure"]), pm_of(r["EndPMMeasure"])
            if lo is None or hi is None or not (min(lo, hi) - 0.3 <= mid <= max(lo, hi) + 0.3):
                continue
            near.append((lo, hi, c1, c2, r))
        for lo, hi, c1, c2, r in sorted(near, key=lambda x: x[0]):
            p1 = str(r.get("BeginPMPrefix") or "").replace(".", "").strip()
            p2 = str(r.get("EndPMPrefix") or "").replace(".", "").strip()
            val = r[attrs[pos]] if pos >= 0 else r.get("InventoryItemStartDate")
            contains = c1 == c2 == county and p1 == p2 == prefix and lo <= mid < hi
            align = str(r.get("Alignment") or "")[:1]
            mark = ("<== contains the PM, same county/prefix/side" if contains and align == want_align
                    else "(contains, other side)" if contains else "")
            say(f"      OBJECTID {r.get('OBJECTID')}: {r.get('Alignment')} {c1}/{p1 or '-'}{lo:g} → "
                f"{c2}/{p2 or '-'}{hi:g}  {attrs[pos] if pos >= 0 else 'InventoryItemStartDate'} = {val!r}  "
                f"item {str(r.get('InventoryItemStartDate'))[:10]}  "
                f"LRS {str(r.get('LRSFromDate'))[:10]}→{str(r.get('LRSToDate'))[:10]}  {mark}")
            if contains and align == want_align:
                truth.append(norm(vkind, val))
        truth = sorted(set(truth))
        na, nb = norm(vkind, a), norm(vkind, b)
        if str(a).startswith("(unavailable:") or str(b).startswith("(unavailable:"):
            verdict = ("UNAVAILABLE: the build could not place this span and marks the cell "
                       "on both sides (its ArcGIS Build sheet lists the skipped span)")
        elif tsmis_raw is not None and tsmis_raw != nb:
            verdict = f"COMPARISON MISREAD TSMIS (raw {tsmis_raw!r} vs shown {nb!r})"
        elif not truth:
            verdict = "UNDETERMINED by a plain layer row (a span of another county/prefix or none covers it)"
        elif len(truth) > 1:
            verdict = f"UNDETERMINED: several plain rows contain it with different values {truth}"
        elif a != b:
            verdict = ("TRUE DIFFERENCE: the layer says what our build says" if truth[0] == na else
                       "FALSE DIFFERENCE: the layer says what TSMIS says — our build is wrong" if truth[0] == nb
                       else f"ALL THREE DIFFER: layer {truth[0]!r}")
        else:
            verdict = ("AGREEMENT CONFIRMED by the layer" if truth[0] == na else
                       f"FALSE AGREEMENT: the layer says {truth[0]!r}, both files say {na!r}")
        tally[verdict.split(":")[0].split(" (")[0]] += 1
        say(f"   VERDICT: {verdict}")
    say("")
    say("# Tally")
    for k, m in tally.most_common():
        say(f"  {m:3d}  {k}")
    if out:
        out.close()


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("index")
    s.add_argument("lib_root")
    s.add_argument("asof")
    s.add_argument("cache")
    s = sub.add_parser("file")
    s.add_argument("xlsx")
    s.add_argument("cache")
    s = sub.add_parser("workbook")
    s.add_argument("xlsx")
    s.add_argument("cache")
    s = sub.add_parser("spot")
    s.add_argument("xlsx")
    s.add_argument("lib_root")
    s.add_argument("asof")
    s.add_argument("raw_dir")
    s.add_argument("--n-diff", type=int, default=25)
    s.add_argument("--n-agree", type=int, default=15)
    s.add_argument("--seed", type=int, default=1)
    s.add_argument("--out")
    a = p.parse_args(argv)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    if a.cmd == "index":
        idx = build_index(a.lib_root, a.asof, cache=a.cache)
        print("spaces per layer:", {t: len(s) for t, s in idx.items()})
    elif a.cmd == "file":
        cmd_file(a.xlsx, pickle.loads(Path(a.cache).read_bytes()))
    elif a.cmd == "workbook":
        cmd_workbook(a.xlsx, pickle.loads(Path(a.cache).read_bytes()))
    else:
        cmd_spot(a.xlsx, a.lib_root, a.asof, a.raw_dir, a.n_diff, a.n_agree, a.seed, a.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
