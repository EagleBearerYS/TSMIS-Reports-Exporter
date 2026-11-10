"""Compare-tab "Three Days" matrix engine (v0.52.0) — one report on three exported
days of the same source, classified by `three_day`.

The vs Baseline matrix answers "what differs between two pulls". This one adds a
third pull, so each difference reads as a history: changed on the newest day,
changed on the middle day and kept, changed and put back, or changed both times
(and each record as added, removed, or there on the middle day only).

Rows = the report editions whose three-day comparison has been verified against
their two-day comparison on real exports (`ROW_KEYS` — the Clean Road files,
Excel and print). Columns = the SETS of three exported days the user adds, each
"<oldest> → <middle> → <newest>". Each cell loads its three days through the
row's own cross-environment adapter — the vs Baseline matrix's adapter — so a
cell keys and compares a record exactly as the two-day comparison does, and its
pair totals are that comparison's own numbers (`three_day`).

Truth: the counts are the three-day workbook's own published sidecar
(`three_day.read_counts`); this module's results cache records only what the cell
was built from (the three export folders' fingerprint and the producer version),
so a route added to, removed from or re-exported on any of the three days reads
the cell stale.

Store: output/comparisons/three-days/<newest day>/<row>_three_days <oldest>
<middle> <newest> <source>.xlsx, the results cache + attempts overlay under its
_state. A workbook's name carries all three days and the source the matrix was
reading, so two sets that share a newest day, or two sources that share a
folder from before the site was recorded, never overwrite each other.
Console-free.
"""
import json
import logging
import os
import time
from collections import namedtuple

import artifact_store
import cache_envelope
import consolidation_meta
import matrix
import output_state
import reports
import three_day
from paths import (comparisons_root, day_run_label, day_source_dir, parse_run_folder,
                   run_days_for, today_str, valid_calendar_date)

log = logging.getLogger("tsmis.three_day_matrix")

SOURCE_DEFAULT = "ssor-prod"
TD_DIRNAME = "three-days"                    # under output/comparisons/
_RESULTS_FILE = "_results.json"
_CACHE_IDENTITY = "three-days"
_SET_SEP = "_"
_DAY_NAMES = ("day1", "day2", "day3")        # the cell's three sides, oldest first

# The rows this matrix offers: every report whose three-day comparison has been
# held against its two-day comparison on real exports. The engine reads any flat
# cross-environment adapter; a report is added here once it has had that pass.
ROW_KEYS = ("clean_highway", "clean_highway_pdf", "clean_intersection",
            "clean_intersection_pdf", "clean_ramp", "clean_ramp_pdf")

# One matrix row: a report edition and the adapter that reads its exports.
Row = namedtuple("Row", "key label code subdir adapter")


def _rows():
    """[Row] — the offered editions, in the cross-environment registry's order."""
    import report_catalog                        # already loaded via reports
    return [Row(key, label, report_catalog.short_code(key), subdir, adapter)
            for key, label, subdir, _idx, adapter in reports.matrix_rows()
            if key in ROW_KEYS and adapter is not None
            and three_day.is_flat(adapter)]


def row_keys():
    return {r.key for r in _rows()}


def sources():
    return matrix.day_source_keys()


# --------------------------------------------------------------------------- #
# a set of three days
# --------------------------------------------------------------------------- #
def make_set(days):
    """The three days as an ascending tuple, or None when they are not three
    different real dates. The order they were picked in does not matter."""
    days = tuple(sorted(d for d in (days or ()) if isinstance(d, str)))
    if len(days) != 3 or len(set(days)) != 3:
        return None
    if not all(valid_calendar_date(d) for d in days):
        return None
    return days


def set_id(days):
    """The id of a set of three days: '<oldest>_<middle>_<newest>'."""
    return _SET_SEP.join(days)


def parse_set(value):
    """The (oldest, middle, newest) days of a set id, or None for anything that
    is not three different real dates in ascending order."""
    if not isinstance(value, str):
        return None
    parts = tuple(value.split(_SET_SEP))
    return parts if make_set(parts) == parts else None


def clean_sets(values):
    """The valid set ids of `values`, without repeats, oldest set first."""
    seen = {v for v in (values or []) if parse_set(v)}
    return sorted(seen, key=lambda v: tuple(reversed(parse_set(v))))


# --------------------------------------------------------------------------- #
# paths + the results cache
# --------------------------------------------------------------------------- #
def td_root():
    return comparisons_root() / TD_DIRNAME


def set_out_path(days, source, row_key):
    """The three-day workbook of one (set, edition), filed under its newest day.
    The name carries all three days and the matrix source in its explicit site
    form ('ssor-prod dev-site'), so two sets — or two reports, or the same days
    read as two sources — never share a file and can be open in Excel at once.
    The folder is the bare date to keep the path short (the sidecar beside the
    workbook is the longest path a cell writes)."""
    name = f"{row_key}_three_days {days[0]} {days[1]} {days[2]} {source}.xlsx"
    return td_root() / days[2] / name


def _result_key(days, source, row_key):
    return f"{days[0]}|{days[1]}|{days[2]}|{source}|{row_key}"


def _results_path():
    return output_state.state_file(td_root(), _RESULTS_FILE)


def load_results():
    """{"<oldest>|<middle>|<newest>|<source>|<row>": {built_at_mtime,
    completion, input_fingerprint, producer_versions}} — what each cell was
    built from."""
    p = output_state.named_read_file(td_root(), _RESULTS_FILE)
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return cache_envelope.unwrap(data, output_identity=_CACHE_IDENTITY)
    except OSError:  # silent-ok: no cache yet (first run) — the empty map is the correct state
        return {}
    except ValueError as e:                  # corrupt JSON: surface it, then degrade
        log.warning("three_day_matrix: corrupt results cache %s (%s: %s); treating "
                    "as empty", p, type(e).__name__, e)
        return {}


def record_result(days, source, row_key, built_at_mtime, completion,
                  input_fingerprint, commit_guard=None):
    data = load_results()
    data[_result_key(days, source, row_key)] = {
        "built_at_mtime": built_at_mtime, "completion": completion,
        "input_fingerprint": input_fingerprint,
        "producer_versions": matrix.producer_identity(),
    }
    p = _results_path()
    tmp = p.with_name(p.name + ".tmp")
    if output_state.ensure_state_dir(td_root(), commit_guard) != p.parent:
        raise ValueError("The three-day matrix state directory is unavailable.")
    for target in (p.parent, p, tmp):
        if not consolidation_meta.guard_allows(commit_guard, target):
            raise ValueError("A three-day matrix input or destination changed before "
                             "its cache was written; refresh the comparison.")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cache_envelope.wrap(data, output_identity=_CACHE_IDENTITY), f)
        os.replace(tmp, p)
    except OSError as e:
        log.warning("three_day_matrix: could not write results cache %s: %s: %s",
                    p, type(e).__name__, e)
        raise ValueError("The three-day workbook was created, but its matrix cache "
                         "could not be written. Refresh the cell.") from e


# --------------------------------------------------------------------------- #
# the sides
# --------------------------------------------------------------------------- #
def export_dir(date, source, subdir):
    """The edition's per-route export on one day, in the REAL run folder (the
    TSMIS site `source` names, else one from before the site was recorded)."""
    return day_source_dir(date, source) / subdir


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
    return {"newest": c[three_day.NEWEST], "middle": c[three_day.MIDDLE],
            "back": c[three_day.BACK], "both": c[three_day.BOTH],
            "same": c[three_day.SAME], "changed": counts.changed_cells(),
            "added": counts.added(), "removed": counts.removed(),
            "middle_only": counts.middle_only(), "records": dict(r),
            "two_day": sum(counts.two_day.values()),
            "counted_columns": counts.counted_columns,
            "pairs": {pair: {k: v for k, v in p.items() if k != "column_diffs"}
                      for pair, p in counts.pairs().items()}}


def _record_current(rec, built_mtime):
    """True when the results cache holds this very workbook's record (its mtime
    and an input fingerprint) — without it nothing certifies what the cell read."""
    if not isinstance(rec, dict) or built_mtime is None:
        return False
    try:
        recorded = float(rec.get("built_at_mtime"))
    except (TypeError, ValueError, OverflowError):  # silent-ok: a malformed stamp makes the record unusable
        return False
    fingerprint = rec.get("input_fingerprint")
    return (abs(recorded - built_mtime) < matrix._MTIME_TOL_S
            and isinstance(fingerprint, str) and bool(fingerprint))


def cell_state(out_path, sides, rec, fp_folders=()):
    """{supported, built, stale, reason, missing_side, missing_sides, completion,
    three_day} for one cell. Freshness rides the shared `matrix._staleness` (the
    three exports' mtimes, their folders' fingerprint, the producer version);
    the counts come only from the workbook's own published sidecar."""
    missing = [s["name"] for s in sides if not s.get("present")]
    built_mtime = matrix._safe_mtime(out_path)
    st = matrix._staleness(built_mtime, sides, rec, fp_folders,
                           missing[0] if missing else None)
    st["missing_sides"] = missing
    for key in ("verdict", "diff_cells", "one_sided"):
        st.pop(key, None)
    if st["built"] and not st["stale"] and not _record_current(rec, built_mtime):
        st["stale"], st["reason"] = True, "cache_missing_or_mismatched"
    counts = three_day.read_counts(out_path) if st["built"] else None
    if st["built"] and counts is None:
        st["stale"], st["reason"] = True, "outcome_missing"
    if counts is not None:
        st["three_day"] = _counts_view(counts)
        st["completion"] = counts.completion
        if not st["stale"] and counts.completion != "complete":
            st["stale"], st["reason"] = True, "partial"
    return st


# --------------------------------------------------------------------------- #
# the snapshot the GUI renders (pure filesystem read)
# --------------------------------------------------------------------------- #
def matrix_snapshot(source, sets, hidden=None, row_order=None, now=None, today=None):
    """Full render model for the Three Days matrix. Pure stat — counts come from
    each workbook's sidecar, no workbook opened. `sets` is the set ids the user
    added; invalid ones are dropped."""
    now = now if now is not None else time.time()
    today = today if today is not None else today_str()
    source = matrix.normalize_day_source(source, SOURCE_DEFAULT)
    set_ids = clean_sets(sets)
    hidden = set(hidden or [])
    all_rows = _rows()
    visible = {r.key: r for r in all_rows if r.key not in hidden}
    rows = [visible[k] for k in matrix.apply_order(list(visible), row_order)]
    results = load_results()
    attempts = matrix.load_attempts(td_root())

    cells = {}
    for r in rows:
        per = {}
        for sid in set_ids:
            days = parse_set(sid)
            folders = [export_dir(d, source, r.subdir) for d in days]
            mtimes = [artifact_store.newest_report_file_mtime(f) for f in folders]
            exports = [{"date": d, "present": m is not None, "mtime": m,
                        "age_seconds": (now - m) if m is not None else None}
                       for d, m in zip(days, mtimes)]
            sides = [{"name": name, "date": d, "present": m is not None, "mtime": m}
                     for name, d, m in zip(_DAY_NAMES, days, mtimes)]
            cmp = cell_state(set_out_path(days, source, r.key), sides,
                             results.get(_result_key(days, source, r.key)),
                             fp_folders=tuple(folders))
            cmp["missing_days"] = [s["date"] for s in sides if not s["present"]]
            attempt = matrix._last_attempt_for(attempts, f"{r.key}|{source}", sid, cmp)
            if attempt is not None:
                cmp["last_attempt"] = attempt
            per[sid] = {"exports": exports, "cmp": cmp}
        cells[r.key] = per

    all_days = sorted({d for sid in set_ids for d in parse_set(sid)})
    return {
        "source": source,
        "sources": matrix.day_source_options(),
        "source_env": matrix.day_source_env(source),
        "source_label": matrix.day_source_label(source),
        "day_hosts": matrix.day_hosts(source, all_days),
        "sets": [{"id": sid, "days": list(parse_set(sid))} for sid in set_ids],
        "today": today,
        "rows": [r.key for r in rows],
        "row_labels": {r.key: r.label for r in all_rows},
        "all_rows": [{"key": r.key, "label": r.label, "code": r.code}
                     for r in all_rows],
        "hidden": sorted(hidden),
        "cells": cells,
    }


def cells_to_rebuild(snapshot, scope="stale", row=None, set_key=None):
    """[(set id, row_key)] to (re)build: 'all' = every cell with its three
    exports present, 'stale' = only the missing/stale ones; optional row / set
    filters."""
    todo = []
    for row_key in snapshot["rows"]:
        if row and row_key != row:
            continue
        for s in snapshot["sets"]:
            sid = s["id"]
            if set_key and sid != set_key:
                continue
            cmp = snapshot["cells"][row_key][sid]["cmp"]
            if not matrix.cell_buildable(cmp):
                continue
            if scope == "all" or cmp.get("stale"):
                todo.append((sid, row_key))
    return todo


# --------------------------------------------------------------------------- #
# one cell
# --------------------------------------------------------------------------- #
def _row(row_key):
    for r in _rows():
        if r.key == row_key:
            return r
    raise ValueError(f"unknown three-day matrix row: {row_key}")


def build_cell(source, set_key, row_key, events, commit_guard=None):
    """Build ONE (set of three days, edition) comparison: load the three days'
    per-route exports through the row's cross-environment adapter, classify them
    and cache what the cell was built from. Returns the ConsolidateResult; raises
    ValueError for a bad set, an unknown row or a day with no export."""
    r = _row(row_key)
    days = parse_set(set_key)
    if days is None:
        raise ValueError(f"invalid set of three days for the matrix: {set_key!r}")
    # The explicit site form names the cell's file and cache record, exactly as
    # the snapshot reads them; anything that is not a source is refused.
    source = matrix.day_source_input(source)
    if source not in sources():
        raise ValueError("invalid source for the three-day matrix")
    for d in days:
        if not parse_run_folder(day_run_label(d, source)):
            raise ValueError(f"invalid date/source for the three-day matrix: {d!r} / "
                             f"{source!r}")
    folders = [export_dir(d, source, r.subdir) for d in days]
    for d, folder in zip(days, folders):
        if artifact_store.newest_report_file_mtime(folder) is None:
            raise ValueError(f"No {r.label} export for {d} {source}.")
    out_path = set_out_path(days, source, row_key)
    fp_before = matrix._cell_input_fingerprint(*folders)
    events.on_log(f"  {r.label}: {days[0]} → {days[1]} → {days[2]} ({source})")
    result = three_day.compare_days(
        r.adapter, [day_source_dir(d, source) for d in days], days, out_path,
        events=events, commit_guard=commit_guard)
    if result.status == "ok" and out_path.exists():
        record_result(days, source, row_key, matrix._safe_mtime(out_path),
                      result.completion,
                      matrix._fingerprint_for_record(fp_before, folders,
                                                     out_path.name, events),
                      commit_guard=commit_guard)
    return result
