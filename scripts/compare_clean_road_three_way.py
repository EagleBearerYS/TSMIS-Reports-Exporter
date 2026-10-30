"""Clean Road Highway / Intersection / Ramp — the site's export vs the TSN extract
vs OUR build from the ArcGIS layers, three ways (v0.51.0).

The three Clean Road files are the first reports the three-way lane serves,
because they are the only ones with all three sources: the site exports them, TSN
left the legacy CA HIGHWAYS / CA INTERSECTIONS / CA RAMPS extracts, and the
ArcGIS reports tab builds each one from the layers (`arcgis_reports`). The engine
is `three_way`; this module is each report's adapter — its loaders, its two
pairing schemas, the layer every column is traced to, and where to look a record
up.

Every side goes through the ONE Clean Road projection (`compare_clean_road_tsn`):
the same key (route · county · PM prefix · postmile · roadbed) and the same
format-only reconciliations the two-sided comparisons use, then the same-source
render rule on all three (edge control characters fold away), so equality is one
relation across the three. Records pair with the TSMIS-vs-TSN comparison's schema
on the TSN side and the Reports-vs-ArcGIS comparison's schema on the ArcGIS side.
A column is compared three ways only where both of those comparisons count it:
Intersection's INX_CONNECTION_ID and Ramp's RAM_DESIGN_DESC are counted only
against ArcGIS (TSN numbers intersections in another id space and words its design
codes its own way), so here they are shown, not counted.

Role gates come from the loaders: side 1 must be the site's consolidated export
(never our build), side 2 the TSN extract or the library's normalized copy, side 3
the app's own ArcGIS build of the same report. Console-free.
"""
import logging
import re
from collections import Counter
from dataclasses import replace
from pathlib import Path

import clean_highway_columns as chc
import clean_road_layers as crl
import compare_clean_highway_arcgis as cha
import compare_clean_intersection_arcgis as cia
import compare_clean_ramp_arcgis as cra
import compare_clean_road_tsn as crt
import compare_tsn_common as ctc
import consolidate_clean_highway as cch
import outcome
import three_way
from events import ConsolidateResult, Events
from paths import today_str
from pdf_table_lib import pdf_source_marker_state
from three_way import ColumnTrace

log = logging.getLogger("tsmis.compare")


def _s(v):
    return "" if v is None else str(v).strip()


def _export_day(path):
    m = re.search(r"(\d{4}-\d{2}-\d{2})", Path(path).name)
    return m.group(1) if m else ""


# --------------------------------------------------------------------------- #
# the layer trace — every column to the layer it is built from
# --------------------------------------------------------------------------- #
_LAYER_NAMES = sorted(crl.EXPECTED_LAYERS, key=len, reverse=True)


def trace_from_provenance(source, how):
    """A ColumnTrace from a point build's PROVENANCE entry ("<layer>[ (which
    row)].<field>[ (note)]", "how"). The layer is matched against the library
    manifest's own names, so a typo in a build's table fails the trace check
    instead of reading as a layer nobody can find."""
    text = _s(source)
    layer = next((n for n in _LAYER_NAMES if text.startswith(n)), None)
    if layer is None:
        return ColumnTrace(rule=f"{text} — {how}" if text and how else text or how)
    rest = text[len(layer):]
    which = ""
    if rest.startswith(" ("):
        close = rest.index(")")
        which, rest = rest[2:close], rest[close + 1:]
    fields = rest[1:] if rest.startswith(".") else rest.strip()
    rule = "; ".join(p for p in (which, how) if p)
    return ColumnTrace(layers=(layer,), fields=fields, rule=rule)


def _point_trace(module, extra_layers=None):
    """{column: ColumnTrace} for a point build (Intersection / Ramp): its
    PROVENANCE table, plus the columns the site gives no source."""
    spec = module.SPEC
    out = {}
    for col in spec.header:
        if col in module.PROVENANCE:
            tr = trace_from_provenance(*module.PROVENANCE[col])
            extra = (extra_layers or {}).get(tr.layers[0]) if tr.layers else None
            if extra:
                tr = replace(tr, layers=tr.layers + extra[0],
                             rule="; ".join(p for p in (tr.rule, extra[1]) if p))
            out[col] = tr
        else:
            out[col] = ColumnTrace(rule="no source on the site (blank in its file "
                                        "by design); the build leaves it blank too")
    return out


_INX_LEG_RULE = ("the leg is picked in IM Intersection Approach Segments by "
                 "INTERSECTION_ID: ROUTE_ID on the file's route for the mainline, "
                 "another route for the cross street; Increasing, then _P, then "
                 "the lowest APPROACH_ID")


def _intersection_trace():
    build = cia._build
    return _point_trace(build, {build.APP_LAYER: ((build.SEG_LAYER,), _INX_LEG_RULE)})


def _ramp_trace():
    return _point_trace(cra._build)


# Highway: the CA HIGHWAYS build is an overlay of span layers, so a column is
# painted from one layer, a block of them, or the overlay's own structure.
_BLOCK_TAGS = {
    "(left-road block)": ("L", cch._LEFT_TAGS),
    "(right-road block)": ("R", cch._RIGHT_TAGS),
    "(median block)": ("MED", ("MED", "BAR", "CURB")),
}


def _layer_of(tag):
    if tag in cch.SPAN_LAYERS:
        return cch.SPAN_LAYERS[tag][0]
    return cch.POINT_LAYERS[tag][0]


def _block_trace(block, tags):
    primary = cch.BLOCK_PRIMARY[block]
    others = [_layer_of(t) for t in tags if t != primary]
    return ColumnTrace(
        layers=(_layer_of(primary),) + tuple(others),
        fields="InventoryItemStartDate",
        rule=(f"{_layer_of(primary)}'s own InventoryItemStartDate on the span "
              "covering the segment's begin; where it has no span, the newest "
              f"of {', '.join(others)}"))


def _span_rule(col, field):
    """How the overlay paints a plain one-layer column."""
    if field == "InventoryItemStartDate":
        return "the covering span's own InventoryItemStartDate"
    if field == "Network_Start_Date":
        return "the covering network span's Network_Start_Date"
    kind = "number" if col.endswith("_AMT") or col.endswith("_ID") else "code"
    return f"the span covering the segment's begin — its {kind}"


def _highway_trace():
    out = {}
    for col in chc.HEADER:
        tier, layer, column, note = chc.PROVENANCE[col]
        if layer in _BLOCK_TAGS:
            out[col] = _block_trace(*_BLOCK_TAGS[layer])
        elif layer in ("(every span layer)",):
            out[col] = ColumnTrace(
                fields=column,
                rule="every span layer's own location fields — the segment's "
                     "identity (the key), never a painted value"
                     + (f"; {note}" if note else ""))
        elif layer in ("(overlay)", "(build)"):
            out[col] = ColumnTrace(rule=note or "derived by the build")
        elif layer:
            names = tuple(n for n in re.split(r"\s*\+\s*", layer)
                          if n in crl.EXPECTED_LAYERS)
            out[col] = ColumnTrace(layers=names, fields=column,
                                   rule=note or _span_rule(col, column))
        else:
            out[col] = ColumnTrace(rule=f"{tier} — {note}" if note else tier)
    return out


# --------------------------------------------------------------------------- #
# where to look a record up
# --------------------------------------------------------------------------- #
def _pm_text(prefix, pm, roadbed=""):
    return f"{prefix}{pm}{roadbed}" if pm else "(no postmile)"


def _locate_highway(route, key, sides):
    m, n, a = sides
    v = m or a or n
    end = v.get("THY_END_PM_AMT", "")
    return (f"route {route} · county {v.get('THY_COUNTY_CODE', '')} · PM {key}"
            + (f" to {v.get('THY_PM_PREFIX_CODE', '')}{end}" if end else "")
            + " — filter the layer by RouteNum, BeginCounty and the "
              "BeginPMMeasure/EndPMMeasure span covering it")


def _locate_intersection(route, key, sides):
    m, n, a = sides
    v = m or a or n
    ident = (m or a or {}).get("INX_CONNECTION_ID", "")
    where = (f"route {route} · county {v.get('INX_COUNTY_CODE', '')} · PM {key}")
    return (f"INTERSECTION_ID {ident} · {where}" if ident
            else f"{where} (no INTERSECTION_ID on this side)")


def _locate_ramp(route, key, sides):
    m, n, a = sides
    v = m or a or n
    return (f"route {route} · county {v.get('RAM_COUNTY_CODE', '')} · PM {key} — "
            "the SHS Ramp Pt row there (Ramp_Name joins it to SHS Ramp)")


# --------------------------------------------------------------------------- #
# the three loaders, each role-gated
# --------------------------------------------------------------------------- #
def _is_print(path):
    return pdf_source_marker_state(path) != 0


class _Report:
    """One Clean Road report's three-way adapter."""

    def __init__(self, key, arcgis_compare, build_module, load_build, trace, locate,
                 *, token="", extra_notes=()):
        self.key = key
        self.extra_notes = tuple(extra_notes)
        self.profile = crt.profile_for(key)
        self.spec = self.profile.spec
        self.build = build_module
        self._arcgis = arcgis_compare
        self._load_build = load_build
        self.trace = trace
        self.locate = locate
        self.token = token
        self.label = self.spec.label
        self.REPORT_NAME = f"{self.label} — TSMIS vs TSN vs ArcGIS"

    # The two two-sided comparisons this one combines, by edition.
    def pair_schemas(self, pdf):
        vs_tsn = {"clean_highway": (crt.HIGHWAY_VS_TSN, crt.HIGHWAY_PDF_VS_TSN),
                  "clean_intersection": (crt.INTERSECTION_VS_TSN,
                                         crt.INTERSECTION_PDF_VS_TSN),
                  "clean_ramp": (crt.RAMP_VS_TSN, crt.RAMP_PDF_VS_TSN)}[self.key]
        return vs_tsn[1 if pdf else 0]._schema, self._arcgis_schema()

    def _arcgis_schema(self):
        if self.key == "clean_highway":
            return cha._BASE_SCHEMA
        return self._arcgis._COMPARE._schema

    def counted(self, pdf):
        sc_tsn, sc_arc = self.pair_schemas(pdf)
        return frozenset(c for c in self.spec.header
                         if c != self.spec.pm_col
                         and c not in set(sc_tsn.context_fields)
                         and c not in set(sc_arc.context_fields))

    def schema(self, pdf):
        sc_tsn, sc_arc = self.pair_schemas(pdf)
        return three_way.ThreeWaySchema(
            report_name=self.label, noun=self.spec.noun,
            noun_plural=self.spec.noun_plural, header=tuple(self.spec.header),
            key_field=self.spec.header.index(self.spec.pm_col),
            counted=self.counted(pdf), pair_tsn=sc_tsn, pair_arcgis=sc_arc,
            trace=self.trace, locate=self.locate, key_label="Location",
            unavailable_token=self.token, text_columns=tuple(self.profile.text))

    def load(self, export_path, tsn_path, build_path):
        """(rows TSMIS, rows TSN, rows ArcGIS, pdf) — each through its own role
        gate and the one projection, then the same-source render rule."""
        pdf = _is_print(export_path)
        rows_m = crt.load_site(export_path, self.profile, pdf=pdf, collapse=pdf)
        rows_n = crt.load_tsn(tsn_path, self.profile, collapse=pdf)
        rows_a = self._load_build(build_path, pdf)
        return (ctc.same_source_render_rows(rows_m),
                ctc.same_source_render_rows(rows_n),
                ctc.same_source_render_rows(rows_a), pdf)

    # ---------------------------------------------------------------- notes
    def _tsn_extract_date(self, rows_n):
        col = next((c for c in self.spec.header if c.endswith("_EXTRACT_DATE")), None)
        if col is None:
            return ""
        i = 1 + self.spec.header.index(col)
        dates = Counter(_s(r[i]) for r in rows_n if _s(r[i]))
        return dates.most_common(1)[0][0] if dates else ""

    def _facts(self, build_path):
        try:
            return self.build.report_facts(build_path)
        except Exception as e:  # silent-ok: the Notes are documentation; an unreadable marker costs one stated date, never the comparison (the loader gates the workbook itself)
            log.info("three-way notes: build facts unreadable (%s: %s)",
                     type(e).__name__, e)
            return {}

    def meta(self, export_path, tsn_path, build_path, rows_n, pdf, counted):
        facts = self._facts(build_path)
        day = _export_day(export_path) or "an unrecorded day"
        asof = facts.get("asof") or ""
        tsn_day = self._tsn_extract_date(rows_n)
        edition = "print (read back from its PDF)" if pdf else "Excel export"
        inputs = (
            f"{Path(export_path).name} — the site's {self.label} {edition} of {day}",
            f"{Path(tsn_path).name} — the TSN {self.spec.tsn_label} extract"
            + (f" of {tsn_day}" if tsn_day else "") + " (frozen at the cutover)",
            f"{Path(build_path).name} — our build from the ArcGIS layers as of "
            f"{asof or '(unknown)'} (layers exported "
            f"{facts.get('drop_exported') or '(unknown)'})",
        )
        warnings = []
        if asof and day[:10] != asof[:10]:
            warnings.append(
                f"VINTAGE: the ArcGIS build is as of {asof} and the export is of "
                f"{day}. A layer edited in between makes the report follow the "
                "newer value while the build still shows the older one — that "
                "reads as 'report differs'. For a clean reading, refresh the "
                "layers on the export day (ArcGIS ▸ Layers), rebuild the report "
                "(ArcGIS ▸ ArcGIS reports) and compare again.")
        return {"inputs": inputs, "warnings": tuple(warnings),
                "notes": self._notes(day, asof, tsn_day, pdf, counted),
                "layer_index": _layer_index(facts.get("layer_library"))}

    def _notes(self, day, asof, tsn_day, pdf, counted):
        spec = self.spec
        shown = [c for c in spec.header if c not in counted and c != spec.pm_col]
        return (
            f"Three sources of the {self.label} file: the site's export of {day} "
            f"(TSMIS), the TSN {spec.tsn_label} extract{' of ' + tsn_day if tsn_day else ''} "
            f"(TSN) and our build of the file from the ArcGIS layers as of "
            f"{asof or '(unknown)'} (ArcGIS).",
            "READ 'Report differs' FIRST: those are the cells where TSN and the "
            "layers agree and the report does not — two independent sources against "
            "one. 'TSN differs' is expected: TSMIS replaced TSN, the TSN side is "
            "frozen at the September-2025 cutover, and there the report follows the "
            "layers. 'ArcGIS differs' points at our build (or at the report reading "
            "something other than the layer the column is traced to); the Layer "
            "trace sheet names that layer for every column.",
            "VINTAGE: the build is a reconstruction as of a date. When that date is "
            "not the export day, a layer edited in between shows as 'report "
            "differs' (the report has the newer value, the build and TSN the "
            "older). Rebuild the report from layers exported on the export day "
            "before treating a cell as a defect.",
            f"Records pair as the two two-sided comparisons pair them: route + "
            f"county + PM prefix + the {spec.noun}'s postmile + roadbed, the "
            "report's rows against TSN as the TSMIS-vs-TSN comparison pairs them "
            "and against the build as the Reports-vs-ArcGIS comparison does; TSN "
            "and ArcGIS rows neither paired with a report row pair with each other "
            "('TSN + ArcGIS, missing from the TSMIS report').",
            "Values compare after the format-only reconciliations the two-sided "
            "comparisons use (dates as ISO dates, amounts as canonical numbers, "
            "county codes without the trailing period, TSN's design-code domain "
            "letter, the control-type and on/off crosswalks) and with edge control "
            "characters folded on every side (the Reports-vs-ArcGIS reading); a "
            "run of spaces is one space."
            + (" This export is the PRINT, so tabs and line breaks fold the same way "
               "on all three sides." if pdf else ""),
            f"{len(counted)} columns are compared three ways — every column both "
            "two-sided comparisons count. SHOWN, NEVER COUNTED: "
            + ", ".join(shown) + ". The columns the site has no source for are "
            "blank in its file by design; TSN's bookkeeping and the extract date "
            "are each system's own; "
            + ("INX_CONNECTION_ID is counted in the Reports-vs-ArcGIS comparison "
               "(TSN uses another id space). " if self.key == "clean_intersection"
               else "RAM_DESIGN_DESC is counted in the Reports-vs-ArcGIS comparison "
               "(TSN words its codes its own way). " if self.key == "clean_ramp"
               else "the two offset columns are each system's own cumulative. "),
            "A grid cell leads with the TSMIS value, then the source that "
            "disagrees: '3 ≠ 2 (TSN & ArcGIS)' (report differs), '3 (TSN: 2)' "
            "(TSN differs), '3 (ArcGIS: 1)' (ArcGIS differs), '3 ≠ TSN 2 ≠ ArcGIS "
            "1' (all differ). A record only two sources have compares those two "
            "(yellow when they differ). Filter the grid on its count columns.",
            "Each finding list names, per cell, the ArcGIS layer and field the build "
            "reads and where to look the record up; the Layer trace sheet adds the "
            "library file and the feature service each layer was exported from. "
            "Filter a list on its Column header to read one column at a time, and "
            "the grid's Presence column for the records one source leaves out.",
        ) + self.extra_notes

    # -------------------------------------------------------------- compare
    def suggest_name(self, _path=None):
        return (f"ThreeWay_{self.label.replace(' ', '')}_Comparison "
                f"{today_str()}.xlsx")

    def compare(self, export_path, tsn_path, build_path, out_path, events=None,
                commit_guard=None):
        """Build the three-way workbook for one export / TSN extract / build.
        Returns a ConsolidateResult (ok / cancelled / error)."""
        events = events or Events()
        for side, p in zip(three_way.SIDES, (export_path, tsn_path, build_path)):
            if not Path(p).is_file():
                return ConsolidateResult(status="error",
                                         message=f"The {side} file doesn't exist:\n{p}")
        events.on_log("=" * 60)
        events.on_log(f"{self.REPORT_NAME} (three-way)")
        events.on_log(f"  TSMIS:  {export_path}")
        events.on_log(f"  TSN:    {tsn_path}")
        events.on_log(f"  ArcGIS: {build_path}")
        events.on_log("=" * 60)
        try:
            rows_m, rows_n, rows_a, pdf = self.load(export_path, tsn_path, build_path)
            notes, completion, _sk, _f, _failures = ctc._merge_input_outcomes(
                tuple(zip(three_way.SIDES, (export_path, tsn_path, build_path))), None)
        except ValueError as e:
            return ConsolidateResult(status="error", message=str(e))
        counted = self.counted(pdf)
        meta = self.meta(export_path, tsn_path, build_path, rows_n, pdf, counted)
        if completion == outcome.PARTIAL:
            meta["partial_reasons"] = tuple(notes or ())
        return three_way.run(self.schema(pdf), rows_m, rows_n, rows_a, out_path,
                             meta=meta, events=events, commit_guard=commit_guard,
                             input_completion=completion)


def _layer_index(library):
    """{layer: {file, source, …}} from the build's own layer library manifest
    (the folder its marker sheet names, else the library now), or {} — the trace
    then names the layer without its file and service."""
    for root in ((library,) if library else ()) + (None,):
        try:
            return crl.read_index(root)
        except (ValueError, OSError) as e:  # silent-ok: the trace sheet still names every layer and field; only the file/service columns go blank, and it says so in the log
            log.info("three-way: layer index unreadable at %s (%s: %s)",
                     root or "the library", type(e).__name__, e)
    return {}


def _highway_build_rows(path, pdf):
    return cha._load_build(path, collapse=pdf)


def _point_build_rows(compare):
    return lambda path, pdf: compare._COMPARE._load_build(path, collapse=pdf)


_HIGHWAY_NOTES = (
    "Highway segments pair on their BEGIN postmile, so a segment the three cut "
    "differently still pairs and shows its END postmile and length as the "
    "difference. The site cuts a segment at every event point of its Highway Log "
    "engine, TSN and the build where an attribute changes — so most END PM / "
    "LENGTH findings are the report's segmentation. Filter the Report differs "
    "list's Column header past them to read the attribute findings.",
    "Three of the site's file-format habits show as 'report differs' by design, "
    "because the legacy file (and the build) spell them otherwise: the equation "
    "marker 'E' in THY_PM_SUFFIX_CODE, the route label ('005S') in THY_ROUTE_NAME, "
    "and a route suffix filled on some rows only.",
    "Cells where the build could not place its source span (the HF-01 anchors) "
    "are 'undecided' — never counted as a finding against any side.",
)

HIGHWAY = _Report("clean_highway", cha, cch, _highway_build_rows, _highway_trace(),
                  _locate_highway, token=cch.UNAVAILABLE_TOKEN,
                  extra_notes=_HIGHWAY_NOTES)
INTERSECTION = _Report("clean_intersection", cia, cia._build, _point_build_rows(cia),
                       _intersection_trace(), _locate_intersection)
RAMP = _Report("clean_ramp", cra, cra._build, _point_build_rows(cra), _ramp_trace(),
               _locate_ramp)
REPORTS = {r.key: r for r in (HIGHWAY, INTERSECTION, RAMP)}


def for_family(key):
    """The adapter of a Clean Road report (its family key), or None."""
    return REPORTS.get(key)


__all__ = ["HIGHWAY", "INTERSECTION", "RAMP", "REPORTS", "for_family",
           "trace_from_provenance"]
