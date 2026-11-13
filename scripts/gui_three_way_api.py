"""GuiThreeWayMixin — the ArcGIS tab's "Three-way" sub-tab (v0.51.0).

A by-day MATRIX like Reports vs ArcGIS: rows = every export edition of a report
with all three sources (the three Clean Road files today), columns = exported
days, each cell = that day's export vs the report's TSN dataset vs its ArcGIS
build, classified cell by cell (`three_way`). The TSN side is the vs-TSN
matrices' (the user's TSN pick honoured) and the ArcGIS side the report's one
build from ArcGIS ▸ ArcGIS reports, so neither has a control here. Engine:
`three_way_matrix`; compares ride the shared matrix queue (`which:"three_way"`).
Composition only — every `self._*` it touches lives on GuiApi.
"""
import logging

import settings
import job_host
from gui_endpoint import _api_method, _snapshot_endpoint
from gui_worker_matrix import ThreeWayMatrixCompareWorker

ui_log = logging.getLogger("tsmis.ui")


class GuiThreeWayMixin:
    def _three_way_snapshot(self):
        """The three-way snapshot with the user's source / days / hidden rows /
        drag order and TSN pick applied."""
        import three_way_matrix

        return three_way_matrix.tw_matrix_snapshot(
            settings.get_three_way_matrix_source(),
            settings.get_three_way_matrix_days(),
            hidden=settings.get_three_way_matrix_hidden(),
            row_order=settings.get_three_way_matrix_row_order(),
            tsn_files=self._matrix_tsn_selections(),
            dest=settings.get_batch_dest())

    @staticmethod
    def _three_way_row_label(row_key):
        import arcgis_reports
        return arcgis_reports.label_of(row_key)

    def _three_way_job_label(self, scope, row=None, date=None):
        rl = self._three_way_row_label(row) if row else None
        if scope == "cell":
            return f"Three-way {rl} — {date}"
        if scope == "row":
            return f"Three-way {rl} — all days"
        if scope == "column":
            return f"Three-way all reports — {date}"
        if scope == "stale":
            return "Refresh stale three-way comparisons"
        return "Rebuild all three-way comparisons"

    @_snapshot_endpoint
    def three_way_matrix_info(self):
        """The three-way snapshot (with the ArcGIS library) + the add-day picker's
        available days."""
        import matrix
        import three_way_matrix

        snap = self._three_way_snapshot()
        snap["available_days"] = three_way_matrix.available_days(snap["source"])
        snap["available_day_reports"] = three_way_matrix.available_day_reports(
            snap["source"])
        snap["day_hosts"].update(matrix.day_hosts(snap["source"], snap["available_days"]))
        self._push_state()
        return snap

    @_api_method
    def set_three_way_matrix_source(self, source):
        import matrix
        import three_way_matrix

        source = matrix.day_source_input(source)
        if source not in three_way_matrix.sources():
            return {"error": "Unknown data source."}
        settings.set_three_way_matrix_source(source)
        avail = set(three_way_matrix.available_days(source))
        kept = [d for d in settings.get_three_way_matrix_days() if d in avail]
        if kept != settings.get_three_way_matrix_days():
            settings.set_three_way_matrix_days(kept)
        self._emit_log(f"Three-way matrix source set to {matrix.day_source_label(source)}.")
        self._push_state()
        return {"ok": True, "source": source, "days": kept}

    @_api_method
    def add_three_way_matrix_day(self, date):
        import three_way_matrix

        if date not in three_way_matrix.available_days(
                settings.get_three_way_matrix_source()):
            return {"error": "That day has no export for this source (only exported "
                             "days can be added)."}
        days = settings.get_three_way_matrix_days()
        if date not in days:
            settings.set_three_way_matrix_days(days + [date])
        self._push_state()
        return {"ok": True, "days": settings.get_three_way_matrix_days()}

    @_api_method
    def remove_three_way_matrix_day(self, date):
        settings.set_three_way_matrix_days(
            [d for d in settings.get_three_way_matrix_days() if d != date])
        self._push_state()
        return {"ok": True, "days": settings.get_three_way_matrix_days()}

    @_api_method
    def set_three_way_matrix_report(self, row_key, visible):
        """Show/hide an edition ROW."""
        import three_way_matrix

        keys = three_way_matrix.row_keys()
        if row_key not in keys:
            return {"error": "Unknown report for the matrix."}
        hidden = set(settings.get_three_way_matrix_hidden())
        if visible:
            hidden.discard(row_key)
        else:
            hidden.add(row_key)
        settings.set_three_way_matrix_hidden(sorted(hidden))
        self._push_state()
        return {"ok": True, "hidden": sorted(hidden)}

    @_api_method
    def set_three_way_matrix_row_order(self, keys):
        import three_way_matrix

        valid = three_way_matrix.row_keys()
        clean = [k for k in (keys or []) if isinstance(k, str) and k in valid]
        settings.set_three_way_matrix_row_order(clean)
        self._push_state()
        return {"ok": True, "order": clean}

    @_api_method
    def set_three_way_matrix_day_order(self, days):
        """Persist the drag-to-reorder day order; an omitted day keeps its place."""
        current = settings.get_three_way_matrix_days()
        clean = [d for d in (days or []) if isinstance(d, str) and d in current]
        clean += [d for d in current if d not in clean]
        settings.set_three_way_matrix_days(clean)
        self._push_state()
        return {"ok": True, "days": clean}

    @_api_method
    def build_three_way_matrix_cell(self, row_key, date):
        """Queue a (re)build of ONE (edition, day) three-way comparison."""
        import matrix

        snap = self._three_way_snapshot()
        if row_key not in {r["key"] for r in snap["all_rows"]}:
            return {"error": "Unknown report for the matrix."}
        if date not in snap["days"]:
            return {"error": "Add that day first."}
        cmp = snap["cells"].get(row_key, {}).get(date, {}).get("cmp")
        reason = matrix.cell_unbuildable_reason(cmp)
        if reason:
            return {"error": reason}
        job = self._make_job("compare", "cell",
                             self._three_way_job_label("cell", row_key, date),
                             row=row_key, env=date, which="three_way")
        return self._enqueue_matrix_job(job)

    @_api_method
    def rebuild_three_way_matrix(self, scope="stale", row=None, date=None, force=False):
        """Queue a rebuild in scope ('stale'/'all'), optionally one row or one
        day. {nothing:True} when idle with nothing to do."""
        import three_way_matrix

        snap = self._three_way_snapshot()
        scope = scope if scope in ("stale", "all") else "stale"
        if row and row not in {r["key"] for r in snap["all_rows"]}:
            return {"error": "Unknown report row for the rebuild."}
        if date and date not in snap["days"]:
            return {"error": "Unknown day for the rebuild."}
        row, date = row or None, date or None
        job_scope = "row" if row else "column" if date else scope
        with self._lock:
            idle = not self._task and not self._queue
        if idle and not three_way_matrix.cells_to_rebuild(
                snap, scope="all" if force else scope, row=row, date=date):
            return {"ok": True, "nothing": True}
        job = self._make_job("compare", job_scope,
                             self._three_way_job_label(job_scope, row, date),
                             row=row, env=date, which="three_way", force=force)
        return self._enqueue_matrix_job(job)

    @_api_method
    def open_three_way_cell_comparison(self, row_key, date):
        """Open ONE three-way workbook."""
        import three_way_matrix

        snap = self._three_way_snapshot()
        if date not in snap["days"] or row_key not in snap["rows"]:
            return {"error": "Unknown cell."}
        path = three_way_matrix.day_out_path(date, snap["source"], row_key)
        if not path.exists():
            return {"error": "No comparison built yet — build the cell first."}
        self._open_file(path)
        return {"ok": True}

    @_api_method
    def open_three_way_comparisons_folder(self):
        import three_way_matrix

        root = three_way_matrix.tw_root()
        root.mkdir(parents=True, exist_ok=True)
        self._open_folder(root)
        return {"ok": True}

    # -- the queue's dispatch hooks (called by gui_matrix's dispatcher) -------- #
    def _resolve_three_way_cells(self, job):
        """[(date, row)] for a three-way compare job."""
        import matrix
        import three_way_matrix

        snap = self._three_way_snapshot()
        if job["scope"] == "cell":
            row, date = job["row"], job["env"]
            if row not in snap["rows"] or date not in snap["days"]:
                return []
            cmp = snap["cells"].get(row, {}).get(date, {}).get("cmp")
            return [(date, row)] if matrix.cell_buildable(cmp) else []
        scope = "stale" if job["scope"] == "stale" else "all"
        return three_way_matrix.cells_to_rebuild(snap, scope=scope, row=job.get("row"),
                                                 date=job.get("env"))

    def _dispatch_three_way_compare_job(self, job):
        cells = self._resolve_three_way_cells(job)
        if not cells:
            return False
        with self._lock:
            self._matrix = {"phase": "comparing", "row": job.get("row"),
                            "cell": job.get("env"), "done": 0, "total": len(cells)}
        self._emit_log(f"{job['label']} — {len(cells)} three-way comparison(s)…")
        self._set_dot("busy", "Comparing…")
        self._emit({"t": "run_started", "mode": "consolidate", "label": "Comparing…",
                    "workers": 1})
        job_host.hosted(ThreeWayMatrixCompareWorker,
            settings.get_three_way_matrix_source(), cells, settings.get_batch_dest(),
            self._gated_queue(), self.cancel_event,
            tsn_files=self._matrix_tsn_selections(),
            force_consolidate=job.get("force", False)).start()
        return True
