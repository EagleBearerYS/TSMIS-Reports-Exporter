"""GuiThreeDayMixin — the Compare tab's "Three Days" sub-tab (v0.52.0).

A MATRIX like vs Baseline, one day wider: rows = the report editions the
three-day comparison serves (the Clean Road files today), columns = the SETS of
three exported days the user adds, each cell = that report's three exports
classified cell by cell (`three_day`) — changed on the newest day, changed on the
middle day and kept, changed and back, changed both times. Each day loads through
the row's own cross-environment adapter, so nothing here has a control of its
own beyond the source, the sets and the live-formulas option (its own setting,
like every matrix's). Engine: `three_day_matrix`; compares ride the shared
matrix queue (`which:"three_day"`). Composition only — every `self._*` it
touches lives on GuiApi.
"""
import logging

import settings
from gui_endpoint import _api_method
from gui_worker_matrix import ThreeDayMatrixCompareWorker

ui_log = logging.getLogger("tsmis.ui")


def _arrow(set_key):
    """'2026-10-30_2026-11-03_2026-11-06' as '2026-10-30 → 2026-11-03 → 2026-11-06'."""
    return " → ".join(str(set_key).split("_"))


class GuiThreeDayMixin:
    @staticmethod
    def _three_day_source():
        """The matrix source every endpoint, the snapshot and the worker read:
        the saved one when it is still offered, else the default — so the grid
        can never show one source while Add or Build acts on another."""
        import matrix
        import three_day_matrix

        return matrix.normalize_day_source(settings.get_three_day_matrix_source(),
                                           three_day_matrix.SOURCE_DEFAULT)

    def _three_day_snapshot(self):
        """The Three Days snapshot with the user's source / sets / hidden rows /
        drag order applied."""
        import three_day_matrix

        return three_day_matrix.matrix_snapshot(
            self._three_day_source(),
            settings.get_three_day_matrix_sets(),
            hidden=settings.get_three_day_matrix_hidden(),
            row_order=settings.get_three_day_matrix_row_order())

    @staticmethod
    def _three_day_set_ids(snap):
        return {s["id"] for s in snap["sets"]}

    def _three_day_job_label(self, scope, row=None, set_key=None):
        import three_day_matrix

        labels = {r.key: r.label for r in three_day_matrix._rows()}
        rl = labels.get(row, row) if row else None
        if scope == "cell":
            return f"Three days {rl} — {_arrow(set_key)}"
        if scope == "row":
            return f"Three days {rl} — every set of days"
        if scope == "column":
            return f"Three days all reports — {_arrow(set_key)}"
        if scope == "stale":
            return "Refresh stale three-day comparisons"
        return "Rebuild all three-day comparisons"

    @_api_method
    def three_day_matrix_info(self):
        """The Three Days snapshot + the day pickers' available days."""
        import matrix
        import three_day_matrix

        snap = self._three_day_snapshot()
        snap["available_days"] = three_day_matrix.available_days(snap["source"])
        snap["available_day_reports"] = three_day_matrix.available_day_reports(
            snap["source"])
        snap["day_hosts"].update(matrix.day_hosts(snap["source"], snap["available_days"]))
        self._push_state()
        return snap

    @_api_method
    def set_three_day_matrix_source(self, source):
        import matrix
        import three_day_matrix

        source = matrix.day_source_input(source)
        if source not in three_day_matrix.sources():
            return {"error": "Unknown data source."}
        settings.set_three_day_matrix_source(source)
        avail = set(three_day_matrix.available_days(source))
        current = three_day_matrix.clean_sets(settings.get_three_day_matrix_sets())
        kept = [s for s in current
                if set(three_day_matrix.parse_set(s)) <= avail]
        if kept != settings.get_three_day_matrix_sets():
            settings.set_three_day_matrix_sets(kept)
        self._emit_log("Three Days matrix source set to "
                       f"{matrix.day_source_label(source)}.")
        self._push_state()
        return {"ok": True, "source": source, "sets": kept}

    @_api_method
    def add_three_day_matrix_set(self, day_a, day_b, day_c):
        """Add a set of three exported days as a column. The three may be picked
        in any order; they are compared oldest → newest."""
        import three_day_matrix

        days = three_day_matrix.make_set([day_a, day_b, day_c])
        if days is None:
            return {"error": "Pick three different days."}
        avail = three_day_matrix.available_days(self._three_day_source())
        missing = [d for d in days if d not in avail]
        if missing:
            return {"error": f"{', '.join(missing)} has no export for this source "
                             "(only exported days can be compared)."}
        sid = three_day_matrix.set_id(days)
        sets = three_day_matrix.clean_sets(settings.get_three_day_matrix_sets())
        if sid not in sets:
            settings.set_three_day_matrix_sets(three_day_matrix.clean_sets(sets + [sid]))
        self._push_state()
        return {"ok": True, "set": sid,
                "sets": settings.get_three_day_matrix_sets()}

    @_api_method
    def remove_three_day_matrix_set(self, set_key):
        settings.set_three_day_matrix_sets(
            [s for s in settings.get_three_day_matrix_sets() if s != set_key])
        self._push_state()
        return {"ok": True, "sets": settings.get_three_day_matrix_sets()}

    @_api_method
    def set_three_day_matrix_report(self, row_key, visible):
        """Show/hide an edition ROW. At least one stays on."""
        import three_day_matrix

        keys = three_day_matrix.row_keys()
        if row_key not in keys:
            return {"error": "Unknown report for the matrix."}
        hidden = set(settings.get_three_day_matrix_hidden())
        if visible:
            hidden.discard(row_key)
        else:
            hidden.add(row_key)
        if len(hidden & keys) >= len(keys):
            return {"error": "Keep at least one report on the matrix."}
        settings.set_three_day_matrix_hidden(sorted(hidden))
        self._push_state()
        return {"ok": True, "hidden": sorted(hidden)}

    @_api_method
    def set_three_day_matrix_row_order(self, keys):
        import three_day_matrix

        valid = three_day_matrix.row_keys()
        clean = [k for k in (keys or []) if isinstance(k, str) and k in valid]
        settings.set_three_day_matrix_row_order(clean)
        self._push_state()
        return {"ok": True, "order": clean}

    @_api_method
    def build_three_day_matrix_cell(self, row_key, set_key):
        """Queue a (re)build of ONE (edition, set of three days) comparison."""
        import matrix

        snap = self._three_day_snapshot()
        if row_key not in {r["key"] for r in snap["all_rows"]}:
            return {"error": "Unknown report for the matrix."}
        if set_key not in self._three_day_set_ids(snap):
            return {"error": "Add that set of days first."}
        cmp = snap["cells"].get(row_key, {}).get(set_key, {}).get("cmp")
        missing = (cmp or {}).get("missing_days")
        if missing:
            return {"error": f"{snap['row_labels'].get(row_key, row_key)} has no "
                             f"export for {', '.join(missing)} — a cell needs the "
                             "report exported on all three days."}
        reason = matrix.cell_unbuildable_reason(cmp)
        if reason:
            return {"error": reason}
        job = self._make_job("compare", "cell",
                             self._three_day_job_label("cell", row_key, set_key),
                             row=row_key, env=set_key, which="three_day")
        return self._enqueue_matrix_job(job)

    @_api_method
    def rebuild_three_day_matrix(self, scope="stale", row=None, set_key=None):
        """Queue a rebuild in scope ('stale'/'all'), optionally one row or one
        set of days. {nothing:True} when idle with nothing to do."""
        import three_day_matrix

        snap = self._three_day_snapshot()
        scope = scope if scope in ("stale", "all") else "stale"
        if row and row not in {r["key"] for r in snap["all_rows"]}:
            return {"error": "Unknown report row for the rebuild."}
        if set_key and set_key not in self._three_day_set_ids(snap):
            return {"error": "Unknown set of days for the rebuild."}
        row, set_key = row or None, set_key or None
        job_scope = "row" if row else "column" if set_key else scope
        with self._lock:
            idle = not self._task and not self._queue
        if idle and not three_day_matrix.cells_to_rebuild(
                snap, scope=scope, row=row, set_key=set_key):
            return {"ok": True, "nothing": True}
        job = self._make_job("compare", job_scope,
                             self._three_day_job_label(job_scope, row, set_key),
                             row=row, env=set_key, which="three_day")
        return self._enqueue_matrix_job(job)

    @_api_method
    def set_three_day_matrix_formulas(self, on):
        """Toggle the Three Days matrix's OWN live-formulas option (persisted,
        independent of the other matrices')."""
        val = settings.set_three_day_matrix_formulas(bool(on))
        self._emit_log("Three Days live-formulas workbook " + ("on" if val else "off")
                       + (" — each comparison also writes a '… (formulas).xlsx'."
                          if val else "."))
        self._push_state()
        return {"ok": True, "on": val}

    @_api_method
    def open_three_day_cell_comparison(self, row_key, set_key):
        """Open ONE three-day workbook."""
        import three_day_matrix

        snap = self._three_day_snapshot()
        if (set_key not in self._three_day_set_ids(snap)
                or row_key not in snap["rows"]):
            return {"error": "Unknown cell."}
        path = three_day_matrix.set_out_path(
            three_day_matrix.parse_set(set_key), snap["source"], row_key)
        if not path.exists():
            return {"error": "No comparison built yet — build the cell first."}
        self._open_file(path)
        return {"ok": True}

    @_api_method
    def open_three_day_comparisons_folder(self):
        import three_day_matrix

        root = three_day_matrix.td_root()
        root.mkdir(parents=True, exist_ok=True)
        self._open_folder(root)
        return {"ok": True}

    # -- the queue's dispatch hooks (called by gui_matrix's dispatcher) -------- #
    def _resolve_three_day_cells(self, job):
        """[(set id, row)] for a three-day compare job."""
        import matrix
        import three_day_matrix

        snap = self._three_day_snapshot()
        if job["scope"] == "cell":
            row, set_key = job["row"], job["env"]
            if row not in snap["rows"] or set_key not in self._three_day_set_ids(snap):
                return []
            cmp = snap["cells"].get(row, {}).get(set_key, {}).get("cmp")
            return [(set_key, row)] if matrix.cell_buildable(cmp) else []
        scope = "stale" if job["scope"] == "stale" else "all"
        return three_day_matrix.cells_to_rebuild(snap, scope=scope, row=job.get("row"),
                                                 set_key=job.get("env"))

    def _dispatch_three_day_compare_job(self, job):
        cells = self._resolve_three_day_cells(job)
        if not cells:
            return False
        with self._lock:
            self._matrix = {"phase": "comparing", "row": job.get("row"),
                            "cell": job.get("env"), "done": 0, "total": len(cells)}
        self._emit_log(f"{job['label']} — {len(cells)} three-day comparison(s)…")
        self._set_dot("busy", "Comparing…")
        self._emit({"t": "run_started", "mode": "consolidate", "label": "Comparing…",
                    "workers": 1})
        ThreeDayMatrixCompareWorker(
            self._three_day_source(), cells,
            self._gated_queue(), self.cancel_event,
            also_formulas=settings.get_three_day_matrix_formulas()).start()
        return True
