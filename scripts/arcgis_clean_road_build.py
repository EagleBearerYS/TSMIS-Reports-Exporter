"""The shared substrate of the Clean Road Intersection and Clean Road Ramp
builds from the ArcGIS layer library (v0.50.0).

The site renders its three Clean Road files flat off the same pipelines its
detail reports use (`clean_intersection.js` re-renders the Intersection Detail
record set, `clean_ramp.js` the Ramp Detail one), so each build here is a POINT
projection — one row per intersection / ramp, no segmentation — written in the
site export's own shape: the sheet the consolidated export uses, "Route" plus
the export's exact header, values in the site's print forms. That keeps the
build readable side by side with the site's file and lets the comparison read
both through the one Clean Road projection (`compare_clean_road_tsn`).

A build module supplies the report's spec (`clean_road_columns`), its layers,
its PROVENANCE table and `build_rows(lib, index, *, asof, events)`; this module
owns the workbook (data, Provenance, the marker sheet), the outcome sidecar, the
build's role identity and the shared cell forms. Console-free; openpyxl loads
lazily.
"""
from __future__ import annotations

import logging
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

import arcgis_layers
import clean_road_layers as crl
import consolidate_clean_highway as cch
import consolidation_meta
import outcome
from events import ConsolidateResult, Events

log = logging.getLogger(__name__)

MARKER_SHEET = "ArcGIS Clean Road Build"
PROVENANCE_SHEET = "Provenance"
ROUTE_COL = "Route"
NO_SOURCE = "no source on the site (blank there by design)"


# --------------------------------------------------------------------------- #
# cell forms — the site's print forms, one home
# --------------------------------------------------------------------------- #
def s(v):
    return "" if v is None else str(v).strip()


def _dec(v):
    try:
        return Decimal(str(v))
    except (InvalidOperation, TypeError, ValueError):  # silent-ok: a pure numeric predicate — a non-numeric cell has no number and every caller renders that as blank or verbatim
        return None


def mdy(v):
    """A layer date as the site's MM/DD/YYYY."""
    serial = crl.to_serial(v)
    if serial is None:
        return s(v)
    d = crl.serial_to_date(serial)
    return f"{d.month:02d}/{d.day:02d}/{d.year:04d}"


def num(v):
    """A measure as the site's natural decimal (`String(parseFloat(v))`)."""
    d = _dec(v)
    if d is None:
        return s(v)
    return str(int(d)) if d == int(d) else f"{d.normalize():f}"


def integer(v):
    d = _dec(v)
    if d is None:
        return s(v)
    return str(int(d))


def yn(v):
    """The layers' Yes/No (or 1/0) indicator as the site's Y/N."""
    t = s(v).upper()
    return {"YES": "Y", "NO": "N", "1": "Y", "0": "N", "TRUE": "Y",
            "FALSE": "N", "Y": "Y", "N": "N"}.get(t, t)


def code(v):
    return crl.code_of(v)


def pfx(v):
    """A prefix/suffix cell: the layers write '.' for none."""
    return crl.dot_none(v).strip().upper()


def route_token(route_num, route_suffix):
    return crl.norm_route(route_num) + pfx(route_suffix)


# --------------------------------------------------------------------------- #
# the site's range-layer lookup: a value AT one LRS measure on one route
# --------------------------------------------------------------------------- #
class MeasureOverlay:
    """One span layer indexed the way the site's `queryRangeLayer` /
    `hl_queryLayer122Field` look a point up: by RouteID and AR measure, ends
    inclusive. Where several spans match, the current calibration (no
    LRSToDate at the as-of date) wins, then the span starting highest — the one
    that BEGINS at a shared boundary. A retired copy is dropped when a current
    copy shares its route, postmile extent and start date (`_lrsDedupRetired`).
    """

    __slots__ = ("_by_route",)
    _FWD = 0.00001
    _SNAP = 0.005

    def __init__(self):
        self._by_route = defaultdict(list)

    def value_at(self, route_id, measure):
        best = None
        for frm, to, current, value in self._by_route.get(route_id, ()):
            if frm <= measure <= to:
                rank = (current, frm)
                if best is None or rank > best[0]:
                    best = (rank, value)
        return "" if best is None else best[1]

    def site_value(self, route_id, measure, *, snap=False):
        """The site's `queryRangeLayer` pick for one point. A secondary-
        alignment point ('…_S') matches its own route first, then the primary
        one. With `snap` (the intersection lookups pass it) a point at the end
        of its span also pulls in the span that begins within 0.00001 after it
        (the layers store consecutive spans a hair apart), and a point just past
        every span reads the nearest one within 0.005. Then the current
        calibration, then the span starting highest."""
        if measure is None or not route_id:
            return ""
        own = self._by_route.get(route_id, ()) if route_id.endswith("_S") else ()
        primary = self._by_route.get(primary_route_id(route_id), ())
        matches = ([sp for sp in own if sp[0] <= measure <= sp[1]]
                   or [sp for sp in primary if sp[0] <= measure <= sp[1]])
        if snap and matches:
            matches += [sp for sp in (*primary, *own)
                        if measure < sp[0] <= measure + self._FWD and sp not in matches]
        if snap and not matches:
            best = None
            for sp in own or primary:
                dist = sp[0] - measure if measure < sp[0] else max(measure - sp[1], 0)
                if dist <= self._SNAP and (best is None or dist < best[0]):
                    best = (dist, sp)
            matches = [best[1]] if best is not None else []
        if not matches:
            return ""
        return max(matches, key=lambda sp: (sp[2], sp[0]))[3]

    @classmethod
    def read(cls, path, layer, attr, *, asof, expected_rows=None,
             start="InventoryItemStartDate", end="InventoryItemEndDate"):
        """The spans live at `asof` by the site's date filter (`start <= D` and
        no end or an end after D), never a calibration that starts after D."""
        cols = ["RouteID", "FromARMeasure", "ToARMeasure", "BeginPMMeasure",
                "EndPMMeasure", start, end, "LRSFromDate", "LRSToDate", attr]
        groups = defaultdict(list)
        for r in crl.stream_layer(path, cols, layer_name=layer,
                                  expected_rows=expected_rows):
            frm, to = _dec(r["FromARMeasure"]), _dec(r["ToARMeasure"])
            rid = s(r["RouteID"])
            if frm is None or to is None or not rid:
                continue
            first = crl.to_serial(r[start])
            last = crl.to_serial(r[end])
            if first is not None and first > asof or last is not None and last <= asof:
                continue
            lrs_from = crl.to_serial(r["LRSFromDate"])
            if lrs_from is not None and lrs_from > asof:
                continue
            lrs_to = crl.to_serial(r["LRSToDate"])
            current = lrs_to is None or lrs_to > asof
            key = (rid, s(r["BeginPMMeasure"]), s(r["EndPMMeasure"]), first)
            groups[key].append((float(frm), float(to), current, s(r[attr])))
        ov = cls()
        for (rid, _b, _e, _first), spans in groups.items():
            if any(current for _f, _t, current, _v in spans):
                spans = [sp for sp in spans if sp[2]]
            ov._by_route[rid].extend(spans)
        return ov


class PointLayer:
    """One span layer read for a point that carries no LRS measure (an
    intersection), the way the site's AR lookup lands on it. Roadbed first, as
    the site matches every `_S` span before any `_P` one: a left-roadbed point
    (postmile suffix L) tries the Left alignment, then the Right one. On each:

      * a span of the layer in the point's own county + postmile-prefix space
        that contains it — exact at the span's own boundaries, where the site's
        pick of the span STARTING there applies (current calibration, then the
        highest begin);
      * else any other span of the layer at the point's AR
        (`clean_road_layers.PostmileToAr`):
        one that changes county or prefix part-way, which postmile space cannot
        express (014 KER L16.146 lies inside SHS Population's R6.116 → 20.2;
        101 DN's left roadbed reads Highway Group 'L' off a span the layer ends
        in prefix L), or one the layer records under another prefix (SHS
        Population still has 120 SJ R16.748's stretch as plain 15.86 → 18.7).
        Never a span of the point's own space: postmile space has said none of
        those contains it, so a point in a gap between two of them stays blank
        (009 SCL 6.372, between two Saratoga spans).

    A point just past every span snaps to the nearest one within 0.005 mi in
    its space (the site's snap), the Left alignment's when it has any."""

    __slots__ = ("_by_pm", "_by_ar")
    _SNAP = 0.005

    def __init__(self):
        self._by_pm = defaultdict(list)
        self._by_ar = defaultdict(list)

    def value(self, route, county, prefix, pm, align, rid, ar, *, snap=True):
        sides = ("L", "R") if align == "L" else ("R",)
        for side in sides:
            space = (route, county, prefix, side)
            best = None
            if pm is not None:
                for b, e, current, value in self._by_pm.get(space, ()):
                    if b <= pm <= e and (best is None or (current, b) > best[0]):
                        best = ((current, b), value)
            if best is None and rid and ar is not None:
                side_rid = rid if side == "L" else primary_route_id(rid)
                for frm, to, current, value, home in self._by_ar.get(side_rid, ()):
                    if home != space and frm <= ar <= to and (
                            best is None or (current, frm) > best[0]):
                        best = ((current, frm), value)
            if best is not None:
                return best[1]
        if snap and pm is not None:
            for side in sides:
                spans = self._by_pm.get((route, county, prefix, side), ())
                if not spans:
                    continue
                near = None
                for b, e, _current, value in spans:
                    dist = b - pm if pm < b else max(pm - e, 0)
                    if dist <= self._SNAP + 1e-9 and (near is None or dist < near[0]):
                        near = (dist, value)
                return near[1] if near is not None else ""
        return ""

    @classmethod
    def read(cls, path, layer, attr, *, asof, expected_rows=None,
             start="InventoryItemStartDate", end="InventoryItemEndDate",
             calibration=None):
        """Both indexes in one pass, the spans live at `asof` by the site's
        date filter (as `MeasureOverlay.read`, retired copies dropped the same
        way). With `calibration` (a `crl.PostmileToAr`) every span read also
        calibrates it, the pieces one recalibration split it into as one."""
        cols = ["RouteID", "RouteNum", "RouteSuffix", "Alignment", "BeginCounty",
                "EndCounty", "BeginPMPrefix", "EndPMPrefix", "BeginPMMeasure",
                "EndPMMeasure", "FromARMeasure", "ToARMeasure", start, end,
                "LRSFromDate", "LRSToDate", attr]
        groups = defaultdict(list)
        for r in crl.stream_layer(path, cols, layer_name=layer,
                                  expected_rows=expected_rows):
            first, last = crl.to_serial(r[start]), crl.to_serial(r[end])
            if first is not None and first > asof or last is not None and last <= asof:
                continue
            lrs_from = crl.to_serial(r["LRSFromDate"])
            if lrs_from is not None and lrs_from > asof:
                continue
            lrs_to = crl.to_serial(r["LRSToDate"])
            measures = [_dec(r[k]) for k in ("BeginPMMeasure", "EndPMMeasure",
                                              "FromARMeasure", "ToARMeasure")]
            key = (s(r["RouteID"]), s(r["BeginPMMeasure"]), s(r["EndPMMeasure"]), first)
            groups[key].append((
                lrs_to is None or lrs_to > asof, lrs_from, s(r[attr]),
                route_token(r["RouteNum"], r["RouteSuffix"]),
                crl.norm_alignment(r["Alignment"]) or "R",
                (crl.norm_county(r["BeginCounty"]), pfx(r["BeginPMPrefix"])),
                (crl.norm_county(r["EndCounty"]), pfx(r["EndPMPrefix"])),
                [None if v is None else float(v) for v in measures]))
        ov = cls()
        for (rid, _b, _e, _first), spans in groups.items():
            if any(sp[0] for sp in spans):
                spans = [sp for sp in spans if sp[0]]
            pieces = defaultdict(list)
            for current, lrs_from, value, route, align, begin, finish, (b, e, frm, to) in spans:
                home = None
                if b is not None and e is not None and begin == finish and e >= b:
                    home = (route, *begin, align)
                    ov._by_pm[home].append((b, e, current, value))
                if rid and frm is not None and to is not None:
                    ov._by_ar[rid].append((frm, to, current, value, home))
                    if b is not None and e is not None:
                        pieces[(lrs_from, current, route, begin, finish, b, e)].append(
                            (frm, to))
            if calibration is not None:
                for (_lrs_from, current, route, begin, finish, b, e), ars in pieces.items():
                    calibration.add(route, begin, finish, (b, e),
                                    (min(frm for frm, _to in ars), max(to for _frm, to in ars)),
                                    rid, current)
        return ov


def primary_route_id(route_id):
    """The `_P` route a secondary-alignment id looks up on ('SHS_001._S' ->
    'SHS_001._P'), the rewrite the site's seg-order lookup applies."""
    return route_id[:-3] + "._P" if route_id.endswith("._S") else route_id


# --------------------------------------------------------------------------- #
# reading a build back
# --------------------------------------------------------------------------- #
def is_build(path, spec):
    """Is `path` a Clean Road build of `spec`'s report made by this app? The
    marker sheet names the report, so an intersection build can never stand as
    a ramp one. An unreadable file answers False — the loader that follows
    reports the real problem."""
    from openpyxl import load_workbook

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:  # silent-ok: a role probe — an unopenable file is simply not one of ours, and the loader that follows reports the real problem
        log.info("clean road build probe: %s unreadable (%s: %s)",
                 path, type(e).__name__, e)
        return False
    try:
        if MARKER_SHEET not in wb.sheetnames:
            return False
        for r in wb[MARKER_SHEET].iter_rows(values_only=True):
            if r and s(r[0]) == "Report":
                return s(r[1] if len(r) > 1 else None) == spec.key
        return False
    finally:
        wb.close()


def report_facts(path):
    """The facts the marker sheet records ({} when absent — never invented)."""
    from openpyxl import load_workbook

    keys = {"As-of date": "asof", "Build version": "build_version",
            "Report": "report", "Records written": "rows", "Routes": "routes",
            "Layer library": "layer_library",
            "Layer drop exported": "drop_exported",
            "Layer drop fingerprint": "drop_fingerprint"}
    facts = {}
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if MARKER_SHEET not in wb.sheetnames:
            return facts
        for r in wb[MARKER_SHEET].iter_rows(values_only=True):
            if r and r[0] is not None and s(r[0]) in keys:
                facts[keys[s(r[0])]] = s(r[1] if len(r) > 1 else None)
        return facts
    finally:
        wb.close()


def read_rows(path, spec):
    """The build's data rows as [route, *values in spec.header order], header-
    gated. Raises ValueError naming what to do."""
    from openpyxl import load_workbook

    name = Path(path).name
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if spec.sheet not in wb.sheetnames:
            raise ValueError(f"{name} has no '{spec.sheet}' sheet — rebuild the "
                             f"{spec.label} report.")
        it = wb[spec.sheet].iter_rows(values_only=True)
        header = [s(c) for c in (next(it, None) or ())]
        if header != [ROUTE_COL] + list(spec.header):
            raise ValueError(
                f"{name} does not carry the {len(spec.header)}-column "
                f"{spec.label} header — rebuild it with this version.")
        out = []
        for r in it:
            if r and any(c not in (None, "") for c in r):
                vals = list(r[1:1 + len(spec.header)])
                vals += [None] * (len(spec.header) - len(vals))
                out.append(vals)
        return out
    finally:
        wb.close()


# --------------------------------------------------------------------------- #
# the workbook
# --------------------------------------------------------------------------- #
def _write_workbook(module, out_path, rows, stats, asof_text, lib_root, drop):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    spec = module.SPEC
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = spec.sheet
    head_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="1F3864")
    grey_fill = PatternFill("solid", start_color="808080")
    header = [ROUTE_COL] + list(spec.header)
    ws.append(header)
    for c in next(ws.iter_rows(min_row=1, max_row=1)):
        c.font = head_font
        c.fill = grey_fill if c.value in spec.unsourced else head_fill
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
    for r in rows:
        ws.append(r)
    ws.freeze_panes = "B2"
    for i, col in enumerate(header, start=1):
        ws.column_dimensions[get_column_letter(i)].width = (
            30 if col in module.TEXT_COLUMNS else 12)

    prov = wb.create_sheet(PROVENANCE_SHEET)
    prov.sheet_properties.tabColor = "1F3864"
    prov.append(["Column", "ArcGIS layer field(s)", "How", "On the site"])
    for c in next(prov.iter_rows(min_row=1, max_row=1)):
        c.font, c.fill = head_font, head_fill
    for col in spec.header:
        src, how = module.PROVENANCE.get(col, ("", ""))
        prov.append([col, src or "(none)", how,
                     NO_SOURCE if col in spec.unsourced else "sourced"])
    for col, w in (("A", 32), ("B", 58), ("C", 52), ("D", 40)):
        prov.column_dimensions[col].width = w

    mk = wb.create_sheet(MARKER_SHEET)
    mk.sheet_properties.tabColor = "1F3864"
    for k, v in [
        ("Report", spec.key),
        ("Build version", module.BUILD_VERSION),
        ("As-of date", asof_text),
        ("Layer library", str(lib_root)),
        ("Layer drop exported", drop.get("exported") or "(unknown)"),
        ("Layer drop fingerprint", drop.get("fingerprint") or "(unknown)"),
        ("Records written", stats["rows"]),
        ("Routes", stats["routes"]),
    ] + list(module.marker_rows(stats)):
        mk.append([k, v])
    mk.column_dimensions["A"].width = 30
    mk.column_dimensions["B"].width = 80
    for row in mk.iter_rows(min_col=2, max_col=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(out_path)


# --------------------------------------------------------------------------- #
# the shared entry point
# --------------------------------------------------------------------------- #
def consolidate(module, events=None, confirm_overwrite=None, *, out_path=None,
                asof=None, lib_root=None):
    """Build `module`'s report from the layers as of `asof` (the export day
    being compared, or the comparison measures network change). Returns a
    ConsolidateResult; the destination is app-owned and a rebuild replaces it."""
    events = events or Events()
    out_path = Path(out_path) if out_path else module.OUT_PATH
    try:
        return _consolidate(module, events, confirm_overwrite, out_path, asof,
                            lib_root)
    except ValueError as e:
        return ConsolidateResult(status="error", message=str(e))


def _consolidate(module, events, confirm_overwrite, out_path, asof, lib_root):
    if out_path.exists() and confirm_overwrite is not None \
            and not confirm_overwrite(out_path):
        return ConsolidateResult(status="cancelled",
                                 message="Cancelled. Existing file kept.",
                                 completion=outcome.CANCELLED)
    lib = crl.inventory(lib_root)
    missing = [n for n in module.REQUIRED_LAYERS if n not in lib["present"]]
    if missing:
        raise ValueError(
            "The ArcGIS layer library is missing the layer(s) this report is "
            "built from:\n\n  " + "\n  ".join(missing)
            + "\n\nRefresh them on the ArcGIS ▸ Layers tab and build again.")
    index = crl.read_index(lib_root)
    asof_date = cch._resolve_asof(asof)
    asof_text = asof_date.isoformat()
    drop = arcgis_layers.drop_info(lib_root)
    events.on_log("=" * 60)
    events.on_log(f"{module.REPORT_NAME} — built from the ArcGIS layers "
                  f"(as of {asof_text})"
                  + (f" (drop exported {drop['exported']})"
                     if drop.get("exported") else ""))
    events.on_log("=" * 60)

    rows, stats = module.build_rows(lib, index, asof=crl.to_serial(asof_date),
                                    asof_day=asof_date, events=events)
    if events.is_cancelled():
        return ConsolidateResult(status="cancelled", message="Cancelled.",
                                 completion=outcome.CANCELLED)
    if not rows:
        return ConsolidateResult(
            status="error",
            message=f"The layers produced no {module.SPEC.noun_plural} to "
                    "project — check the layer library and the as-of date.")
    root = Path(lib_root) if lib_root else crl.root()
    _write_workbook(module, out_path, rows, stats, asof_text, root, drop)
    spec = module.SPEC
    result = ConsolidateResult(
        status="ok",
        message=(f"Built {stats['rows']:,} {spec.label} records "
                 f"({stats['routes']} routes) from the ArcGIS layers as of "
                 f"{asof_text}."),
        output_path=str(out_path),
        summary_lines=[f"{module.REPORT_NAME}: {stats['rows']:,} records, "
                       f"{stats['routes']} routes -> {out_path.name}"],
        completion=outcome.COMPLETE)
    if not consolidation_meta.write_outcome(
            out_path, result,
            extra={module.SIDECAR_KEY: {
                "report": spec.key,
                "build_version": module.BUILD_VERSION,
                "asof": asof_text,
                "layer_drop": {"fingerprint": drop.get("fingerprint"),
                               "exported": drop.get("exported")},
                "layer_library": str(root),
                **stats,
            }}):
        return ConsolidateResult(
            status="error",
            message="The build finished but its outcome sidecar could not be "
                    "published — rebuild before comparing.")
    result.sidecar_published = True
    events.on_log(result.message)
    return result
