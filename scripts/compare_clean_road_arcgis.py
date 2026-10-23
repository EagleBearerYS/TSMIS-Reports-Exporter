"""Clean Road Intersection / Ramp — the ArcGIS-built file vs the SITE's export
(v0.50.0).

The Reports vs ArcGIS matrix's Clean Road Intersection and Ramp rows (Highway
has its own module, `compare_clean_highway_arcgis`, because its build is the
74-column CA HIGHWAYS workbook with its own unplaceable-span rule). Side A is
OUR file built from the layers (`arcgis_report_clean_intersection` /
`arcgis_report_clean_ramp`, the ArcGIS reports tab); side B is the site's own
file for one day, consolidated — the Excel export or the print read back from
its PDF, one shape, so one comparison serves both editions.

TSMIS vs TSMIS: both sides come from the same layers, so a difference is a gap
in our reconstruction or a finding about how the site builds its file — never
migration drift. Both sides go through the one Clean Road projection and key
(`compare_clean_road_tsn`): rows pair on route + county + PM prefix + postmile
+ roadbed, dates compare as ISO dates, amounts as canonical numbers, county
codes without the trailing period. The build writes the site's own codes, so
none of the vs-TSN crosswalks has anything to bridge here. CONTEXT is the
vs-TSN set: what the site has no source for (blank in its file by design — the
Ramp build still writes the odometer the requirements ask for), TSN-shaped
bookkeeping and the extract date. Console-free; engine in compare_core.
"""
import logging
import re
from dataclasses import replace
from pathlib import Path

import arcgis_clean_road_build as acb
import compare_clean_road_tsn as crt
import compare_tsn_common as ctc
from paths import today_str
from pdf_table_lib import pdf_source_marker_state

log = logging.getLogger("tsmis.compare")

SIDE_A = "ArcGIS"
SIDE_B = "TSMIS"


def _is_print(path):
    return pdf_source_marker_state(path) != 0


def _export_day(path):
    m = re.search(r"(\d{4}-\d{2}-\d{2})", Path(path).name)
    return m.group(1) if m else ""


class CleanRoadArcgisCompare:
    """compare(build, export, out, …) + suggest_name for one Clean Road report."""

    def __init__(self, build_module):
        self._build = build_module
        self._spec = build_module.SPEC
        self._profile = crt.profile_for(self._spec.key)
        self.REPORT_NAME = f"{self._spec.label} — ArcGIS vs TSMIS"
        # Only what the site has no source for, and the extract date. The
        # vs-TSN extras (TSN's other id space, its own design wording) are one
        # source on both sides here, so they count.
        spec = self._spec
        self.context = tuple(c for c in spec.header
                             if c in spec.unsourced or c in spec.site_only
                             or c.endswith("_EXTRACT_DATE"))
        base = crt._schema(
            self._profile, SIDE_A, SIDE_B,
            notes_title=f"{self.REPORT_NAME}: comparison notes", notes=(),
            context=self.context, representation=self._profile.text)
        self._schema = replace(
            base, report_name=self.REPORT_NAME, sides_noun="builds",
            one_sided_note_extra=(f" ({self._spec.noun_plural} one side records "
                                  "at a postmile the other doesn't — a row-"
                                  "universe difference, or network change "
                                  "across the two dates)"))

    # ----------------------------------------------------------------- load
    def _load_build(self, path, collapse):
        if not self._build.is_arcgis_report(path):
            raise ValueError(
                f"The {SIDE_A} side is not an ArcGIS-built {self._spec.label} "
                f"file (no '{acb.MARKER_SHEET}' sheet naming it):\n{path}\n\n"
                "Build it on the ArcGIS ▸ ArcGIS reports tab first.")
        hint = f"{Path(path).name} ({self._spec.sheet})"
        return [crt._row(self._profile, vals, "site", hint, collapse=collapse)
                for vals in acb.read_rows(path, self._spec)]

    def _load_pair(self, build_path, export_path):
        if acb.is_build(export_path, self._spec):
            raise ValueError(
                f"The {SIDE_B} side is an ArcGIS-built file, not the site's "
                f"export:\n{export_path}\n\nPick the consolidated "
                f"{self._spec.label} export of the day to compare against.")
        pdf = _is_print(export_path)
        rows_a = self._load_build(build_path, collapse=pdf)
        rows_b = crt.load_site(export_path, self._profile, pdf=pdf, collapse=pdf)
        return (ctc.same_source_render_rows(rows_a),
                ctc.same_source_render_rows(rows_b), None)

    # ---------------------------------------------------------------- notes
    def _notes_lines(self, build_path, export_path, unplaced):
        spec = self._spec
        try:
            facts = self._build.report_facts(build_path)
        except Exception as e:  # silent-ok: the Notes sheet is documentation; an unreadable marker costs one stated date, never the comparison (the loaders gate the workbooks themselves)
            log.info("clean-road arcgis notes: build facts unreadable (%s: %s)",
                     type(e).__name__, e)
            facts = {}
        edition = ("print (read back from its PDF)" if _is_print(export_path)
                   else "Excel export")
        sourced = [c for c in spec.header if c in self._build.PROVENANCE
                   and c not in spec.unsourced]
        table = tuple(
            f"    {c}  ·  {self._build.PROVENANCE[c][0]} ({self._build.PROVENANCE[c][1]})"
            for c in sourced)
        return (
            f"Side A ({SIDE_A}) is OUR {spec.label} file, built from the ArcGIS "
            f"layer library as of {facts.get('asof') or '(unknown)'} (layers "
            f"exported {facts.get('drop_exported') or '(unknown)'}): "
            f"{facts.get('rows') or '?'} {spec.noun_plural} on "
            f"{facts.get('routes') or '?'} routes. Side B ({SIDE_B}) is the "
            f"site's own {spec.label} file — its {edition} of "
            f"{_export_day(export_path) or 'an unrecorded day'}, consolidated.",
            "VINTAGE: these two dates must match for the counts to measure "
            "correctness. Across a gap they include network change — refresh "
            "the layers (ArcGIS ▸ Layers), rebuild the report as of the export "
            "day (ArcGIS ▸ ArcGIS reports) and compare again.",
            "Both sides come from the same layers, so a difference here is a gap "
            "in our reconstruction or a finding about how the site builds its "
            "file — never migration drift.",
            f"Rows pair on route + county + PM prefix + the {spec.noun}'s "
            "postmile + roadbed (the R/L postmile suffix) — the key the "
            "site-vs-TSN comparison uses.",
            "Format-only reconciliations: dates compare as ISO dates; amounts as "
            "canonical numbers ('02' = '2'); county codes drop the trailing "
            "period the site writes for two-letter counties ('SD.' = 'SD'). A "
            "print-edition export folds runs of spaces on both sides, as its "
            "HTML does.",
            "CONTEXT (shown, never counted): the columns the site has no source "
            f"for ({', '.join(spec.unsourced)}) and the extract date."
            + (" The build writes RAM_BEGIN_OFFSET_AMT from the ramp point's "
               "odometer, which the site leaves blank."
               if spec.key == "clean_ramp" else
               " INX_CONNECTION_ID is the layer's INTERSECTION_ID on both "
               "sides, so unlike the vs-TSN comparison it counts."),
        ) + tuple(unplaced) + (
            "Where each counted column comes from in the build:",
        ) + table

    # -------------------------------------------------------------- compare
    def suggest_name(self, _build_path=None):
        tag = self._spec.label.replace(" ", "")
        return f"ArcGIS_vs_TSMIS_{tag}_Comparison {today_str()}.xlsx"

    def compare(self, build_path, export_path, out_path, events=None,
                confirm_overwrite=None, mode="formulas", commit_guard=None):
        found = {}

        def loader(a, b):
            rows_a, rows_b, extra = self._load_pair(a, b)
            found["unplaced"] = crt.unplaced_note_lines(
                self._spec, rows_a, rows_b, SIDE_A, SIDE_B)
            return rows_a, rows_b, extra

        def notes(wb):
            return ctc.make_notes_writer(
                f"{self.REPORT_NAME}: comparison notes",
                self._notes_lines(build_path, export_path,
                                  found.get("unplaced", ())))(wb)

        return ctc.run_files_compare(
            replace(self._schema, legend_writer=notes), build_path, export_path,
            out_path,
            banner=f"{self._spec.label} Comparison — {SIDE_A} vs {SIDE_B}",
            has_route=True, loader=loader, deps_ok=crt._DEPS_OK,
            deps_msg="Required components are missing (openpyxl).",
            side_a=SIDE_A, side_b=SIDE_B,
            events=events, confirm_overwrite=confirm_overwrite, mode=mode,
            commit_guard=commit_guard)
