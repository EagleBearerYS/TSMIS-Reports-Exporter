"""ArcGIS-tab "Three-way" matrix engine (v0.51.0) — each exported day of a report
compared against the TSN extract AND the report's ArcGIS build at once.

The sixth matrix. Rows = every export EDITION of a report that has all three
sources (the Excel export and the print are separate rows; today the three Clean
Road files — `compare_clean_road_three_way.REPORTS`), columns = exported DAYS the
user adds, each cell = that day's consolidated export vs the report's TSN dataset
vs the report's ONE ArcGIS build, classified by `three_way`: where TSN and the
layers agree and the report does not ("report differs") stands apart from where
the report follows the layers and only TSN is behind.

Each side is the one the two two-sided matrices already use, never a second copy:
the TSN dataset is the by-day vs-TSN matrix's (`matrix.tsn_source`, healed by
`tsn_library.ensure_current` before a build, the user's TSN pick honoured), the
ArcGIS build is the Reports-vs-ArcGIS matrix's (`arcgis_matrix.arcgis_side`, built
and refreshed on ArcGIS ▸ ArcGIS reports), and the export is consolidated through
the shared `matrix._ensure_consolidated`.

Truth: the counts are the three-way workbook's own published sidecar
(`three_way.read_counts`); this module's results cache records only what the
cell was built from (the export folder's fingerprint, the TSN identity token and
the build's content identity) so a moved input reads the cell stale.

Store: output/comparisons/three-way-by-day/<run>/<row>_three_way <run>.xlsx, the
counts cache + attempts overlay under its _state. Console-free.
"""
import json
import logging
import os
import time
from collections import namedtuple
from pathlib import Path

import arcgis_matrix
import arcgis_reports
import artifact_store
import cache_envelope
import consolidation_meta
import matrix
import output_state
import three_way
from paths import (comparisons_root, day_run_label, parse_run_folder, run_days_for,
                   today_str)

log = logging.getLogger("tsmis.three_way_matrix")

SOURCE_DEFAULT = "ssor-prod"
TW_DIRNAME = "three-way-by-day"              # under output/comparisons/
_RESULTS_FILE = "_results.json"
_CACHE_IDENTITY = "three-way-by-day"
_SIDES = ("export", "tsn", "layers")         # the order a missing side is named in

# One matrix row: an export edition of a report the three-way lane serves.
Row = namedtuple("Row", "key label code subdir family tsn_key")


def _adapters():
    import compare_clean_road_three_way as t3   # lazy: pulls the three loaders
    return t3.REPORTS


def _families():
    """The report families with a three-way adapter, in registry order."""
    import compare_clean_road_three_way as t3
    return [k for k in arcgis_reports.KEYS if k in t3.REPORTS]


def _rows():
    """[Row] — every consolidated export edition of each three-way report."""
    out = []
    for family in _families():
        for e in arcgis_reports.editions_of(family):
            if e.consolidator:
                out.append(Row(e.key, e.label, e.code, e.subdir, family,
                               matrix.tsn_subdir_for(e.key, e.subdir)))
    return out


def row_keys():
    return {r.key for r in _rows()}


def sources():
    return matrix.day_source_keys()


# --------------------------------------------------------------------------- #
# paths + the results cache
# --------------------------------------------------------------------------- #
def tw_root():
    return comparisons_root() / TW_DIRNAME


def day_folder_name(date, source):
    return day_run_label(date, source)


def day_out_path(date, source, row_key):
    """The three-way workbook of one (day, edition): the name says the report,
    the day and the source (TSMIS site tag included) so two can be open at once."""
    name = day_folder_name(date, source)
    return tw_root() / name / f"{row_key}_three_way {name}.xlsx"


def _results_path():
    return output_state.state_file(tw_root(), _RESULTS_FILE)


def load_results():
    """{"<run>|<row>": {built_at_mtime, completion, input_fingerprint,
    source_identities, producer_versions}} — what each cell was built from."""
    p = output_state.named_read_file(tw_root(), _RESULTS_FILE)
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return cache_envelope.unwrap(data, output_identity=_CACHE_IDENTITY)
    except OSError:  # silent-ok: no cache yet (first run) — the empty map is the correct state
        return {}
    except ValueError as e:                  # corrupt JSON: surface it, then degrade
        log.warning("three_way_matrix: corrupt results cache %s (%s: %s); treating "
                    "as empty", p, type(e).__name__, e)
        return {}


def record_result(date, source, row_key, built_at_mtime, completion,
                  input_fingerprint, source_identities, commit_guard=None):
    data = load_results()
    data[f"{day_folder_name(date, source)}|{row_key}"] = {
        "built_at_mtime": built_at_mtime, "completion": completion,
        "input_fingerprint": input_fingerprint,
        "source_identities": source_identities or {},
        "producer_versions": matrix.producer_identity(),
    }
    p = _results_path()
    tmp = p.with_name(p.name + ".tmp")
    if output_state.ensure_state_dir(tw_root(), commit_guard) != p.parent:
        raise ValueError("The three-way matrix state directory is unavailable.")
    for target in (p.parent, p, tmp):
        if not consolidation_meta.guard_allows(commit_guard, target):
            raise ValueError("A three-way matrix input or destination changed before "
                             "its cache was written; refresh the comparison.")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cache_envelope.wrap(data, output_identity=_CACHE_IDENTITY), f)
        os.replace(tmp, p)
    except OSError as e:
        log.warning("three_way_matrix: could not write results cache %s: %s: %s",
                    p, type(e).__name__, e)
        raise ValueError("The three-way workbook was created, but its matrix cache "
                         "could not be written. Refresh the cell.") from e


# --------------------------------------------------------------------------- #
# the sides
# --------------------------------------------------------------------------- #
def _export_side(date, source, subdir):
    """(folder, newest mtime or None) of the edition's export that day."""
    d = arcgis_matrix.tsmis_dir(date, source, subdir)
    return d, artifact_store.newest_report_file_mtime(d)


def _tsn_side(dest, tsn_key, tsn_files, cache):
    if tsn_key not in cache:
        cache[tsn_key] = matrix.tsn_source(dest, tsn_key, tsn_files.get(tsn_key))
    return cache[tsn_key]


def _tsn_ready(src):
    return src.get("kind") in ("file", "consolidated")


def available_days(source):
    """Dates (newest first) with an export of any row's edition for `source`."""
    subs = [r.subdir for r in _rows()]
    out = []
    for date, base in run_days_for(source):
        if any(artifact_store.newest_report_file_mtime(base / s) is not None
               for s in subs):
            out.append(date)
    return out


def available_day_reports(source):
    """{date: [code, ...]} — which rows' editions each offered day holds."""
    rows = _rows()
    present = artifact_store.exported_subdirs_by_day(source, [r.subdir for r in rows])
    return {date: [r.code for r in rows if r.subdir in found]
            for date, found in present.items()}


# --------------------------------------------------------------------------- #
# cell state
# --------------------------------------------------------------------------- #
def _counts_view(counts):
    """The cell's render model of the published counts."""
    c, r = counts.cells, counts.records
    return {"report": c[three_way.REPORT], "arcgis": c[three_way.ARCGIS],
            "all": c[three_way.ALL], "tsn": c[three_way.TSN],
            "match": c[three_way.MATCH], "undecided": c[three_way.UNDECIDED],
            "missing_from_report": r["tsn_arcgis"], "only_in_report": r["tsmis"],
            "records": dict(r), "counted_columns": counts.counted_columns}


def cell_state(out_path, sources, rec, fp_folders=()):
    """{supported, built, stale, reason, missing_side, missing_sides, completion,
    three_way} for one cell. Freshness rides the shared `matrix._staleness` (the
    sources' mtimes and identities, the export folder's fingerprint, the producer
    version); the counts come only from the workbook's own published sidecar."""
    missing = [s["name"] for s in sources if not s.get("present")]
    st = matrix._staleness(matrix._safe_mtime(out_path), sources, rec, fp_folders,
                           missing[0] if missing else None)
    st["missing_sides"] = missing
    st.pop("verdict", None)
    st.pop("diff_cells", None)
    st.pop("one_sided", None)
    counts = three_way.read_counts(out_path) if st["built"] else None
    if st["built"] and counts is None:
        st["stale"], st["reason"] = True, "outcome_missing"
    if counts is not None:
        st["three_way"] = _counts_view(counts)
        st["completion"] = counts.completion
        if not st["stale"] and counts.completion != "complete":
            st["stale"], st["reason"] = True, "partial"
    return st


# --------------------------------------------------------------------------- #
# the snapshot the GUI renders (pure filesystem read)
# --------------------------------------------------------------------------- #
def tw_matrix_snapshot(source, days, hidden=None, row_order=None, tsn_files=None,
                       dest=None, now=None, today=None, library=None):
    """Full render model for the three-way matrix. Pure stat — counts come from
    each workbook's sidecar, no workbook opened. `library` is the ArcGIS reports
    library (the drop + every report's build), as the Reports-vs-ArcGIS matrix
    reads it."""
    import tsn_library                              # lazy: canonical selections
    now = now if now is not None else time.time()
    today = today if today is not None else today_str()
    source = matrix.normalize_day_source(source, SOURCE_DEFAULT)
    days = [d for d in (days or []) if isinstance(d, str)]
    hidden = set(hidden or [])
    tsn_files, _changed = tsn_library.canonicalize_selections(tsn_files or {})
    all_rows = _rows()
    visible = {r.key: r for r in all_rows if r.key not in hidden}
    rows = [visible[k] for k in matrix.apply_order(list(visible), row_order)]
    results = load_results()
    attempts = matrix.load_attempts(tw_root())
    library = library if library is not None else arcgis_matrix.library_snapshot()
    builds = library.get("builds", {})
    tsn_cache = {}

    # The three-sides card reads every family's TSN state, hidden rows included,
    # so a report whose rows are all hidden still shows its TSN as it is.
    tsn_meta, tsn_sides, cells = {}, {}, {}
    for r in all_rows:
        src = tsn_sides[r.key] = _tsn_side(dest, r.tsn_key, tsn_files, tsn_cache)
        tsn_meta[r.key] = {"tsn_subdir": r.tsn_key, "source_kind": src.get("kind"),
                           "source_path": src.get("path"),
                           "ready": _tsn_ready(src)}
    for r in rows:
        src = tsn_sides[r.key]
        bs = builds.get(r.family, {})
        arcgis_ok = bool(bs.get("built") and bs.get("comparable_now"))
        per = {}
        for date in days:
            export_dir, export_m = _export_side(date, source, r.subdir)
            export = {"present": export_m is not None, "mtime": export_m,
                      "age_seconds": (now - export_m) if export_m is not None else None,
                      "subdir": r.subdir}
            srcs = [{"name": "export", "present": export_m is not None,
                     "mtime": export_m},
                    {"name": "tsn", "present": _tsn_ready(src),
                     "mtime": src.get("mtime"), "identity": src.get("identity_token"),
                     "identity_required": True},
                    {"name": "layers", "present": arcgis_ok, "mtime": bs.get("mtime"),
                     "identity": bs.get("identity")}]
            rec = results.get(f"{day_folder_name(date, source)}|{r.key}")
            cmp = cell_state(day_out_path(date, source, r.key), srcs, rec,
                             fp_folders=(export_dir,))
            attempt = matrix._last_attempt_for(attempts, f"{r.key}|{source}", date, cmp)
            if attempt is not None:
                cmp["last_attempt"] = attempt
            per[date] = {"export": export, "cmp": cmp}
        cells[r.key] = per

    return {
        "source": source,
        "sources": matrix.day_source_options(),
        "source_env": matrix.day_source_env(source),
        "source_label": matrix.day_source_label(source),
        "day_hosts": matrix.day_hosts(source, days),
        "days": days,
        "today": today,
        "rows": [r.key for r in rows],
        "row_labels": {r.key: r.label for r in rows},
        "row_family": {r.key: r.family for r in all_rows},
        "all_rows": [{"key": r.key, "label": r.label, "code": r.code,
                      "family": r.family} for r in all_rows],
        "hidden": sorted(hidden),
        "tsn_meta": tsn_meta,
        "cells": cells,
        "library": library,
    }


def cells_to_rebuild(snapshot, scope="stale", row=None, date=None):
    """[(date, row_key)] to (re)build: 'all' = every cell with its three sides
    present, 'stale' = only the missing/stale ones; optional row/date filters."""
    todo = []
    for row_key in snapshot["rows"]:
        if row and row_key != row:
            continue
        for d in snapshot["days"]:
            if date and d != date:
                continue
            cmp = snapshot["cells"][row_key][d]["cmp"]
            if not matrix.cell_buildable(cmp):
                continue
            if scope == "all" or cmp.get("stale"):
                todo.append((d, row_key))
    return todo


# --------------------------------------------------------------------------- #
# one cell
# --------------------------------------------------------------------------- #
def _row(row_key):
    for r in _rows():
        if r.key == row_key:
            return r
    raise ValueError(f"unknown three-way matrix row: {row_key}")


def _resolve_tsn(dest, tsn_key, tsn_files, events):
    """The TSN dataset the cell reads, healed first exactly as the by-day vs-TSN
    matrix heals it: (path, identity token, current-identity predicate)."""
    import tsn_library
    tsn_files, _changed = tsn_library.canonicalize_selections(tsn_files or {})
    src = matrix.tsn_source(dest, tsn_key, tsn_files.get(tsn_key))
    if src.get("kind") == "missing_explicit":
        raise ValueError(tsn_library.explicit_selection_problem(src))
    if not _tsn_ready(src):
        raise ValueError("This report's TSN dataset isn't in place yet — import the "
                         "TSN extract on Settings ▸ TSN reports.")
    if src.get("kind") == "consolidated":
        healed = tsn_library.ensure_current(tsn_key, events, source=src)
        if healed is not None and healed.status != "ok":
            raise ValueError(healed.message or "the TSN library is not certifiably "
                                               "current; comparison stopped")
        src = matrix.tsn_source(dest, tsn_key, tsn_files.get(tsn_key))
        if src.get("kind") != "consolidated":
            raise ValueError("the canonical TSN source changed while it was being "
                             "certified")
    token, check = matrix.tsn_identity_check_for(tsn_key, src)
    return Path(src["path"]), token, check


def build_cell(source, date, row_key, dest, events, tsn_files=None,
               force_consolidate=False, commit_guard=None):
    """Build ONE (day, edition) three-way comparison: consolidate the day's export
    of the edition, heal + resolve the TSN dataset, take the report's ArcGIS
    build, compare the three and cache what the cell was built from. Returns the
    ConsolidateResult; raises ValueError for a missing side or a bad cell."""
    r = _row(row_key)
    if not parse_run_folder(day_folder_name(date, source)):
        raise ValueError(f"invalid date/source for the three-way matrix: {date!r} / "
                         f"{source!r}")
    bs = arcgis_matrix.arcgis_side(r.family)
    if not bs.get("built"):
        raise ValueError(f"Build the {arcgis_reports.label_of(r.family)} ArcGIS report "
                         "first (ArcGIS ▸ ArcGIS reports) — it is the ArcGIS side.")
    if not bs.get("comparable_now"):
        raise ValueError(f"The {arcgis_reports.label_of(r.family)} ArcGIS report's "
                         "outcome record is missing or untrusted — refresh it first.")
    export_dir, export_m = _export_side(date, source, r.subdir)
    if export_m is None:
        raise ValueError(f"No {r.label} export for {date} {source}.")
    tsn_path, tsn_token, tsn_check = _resolve_tsn(dest, r.tsn_key, tsn_files, events)
    out_path = day_out_path(date, source, row_key)
    fp_before = matrix._cell_input_fingerprint(export_dir)
    layers_identity = bs.get("identity")
    build_path = Path(bs["path"])

    side_export, _comp = matrix._ensure_consolidated(
        export_dir, r.subdir, events, force_consolidate, commit_guard=commit_guard)
    events.on_log(f"  {r.label}: the {date} {source} export vs TSN vs the ArcGIS "
                  f"report as of {bs.get('asof') or '?'}")
    guard = matrix._compose_source_guard(commit_guard, tsn_check)
    result = _adapters()[r.family].compare(
        side_export, tsn_path, build_path, out_path, events=events,
        commit_guard=guard)
    if result.status == "ok" and out_path.exists():
        matrix._require_source_identity(tsn_check, "recording the three-way cache")
        identities = {"tsn": tsn_token}
        if layers_identity:
            identities["layers"] = layers_identity
        record_result(date, source, row_key, matrix._safe_mtime(out_path),
                      result.completion,
                      matrix._fingerprint_for_record(fp_before, (export_dir,),
                                                     out_path.name, events),
                      identities, commit_guard=guard)
    return result
