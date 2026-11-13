"""The shared js_api endpoint decorator (P7c).

Extracted from `gui_api` so both `gui_api.GuiApi` and the new `gui_matrix.GuiMatrixMixin`
can decorate their pywebview-exposed methods without a `gui_api` ↔ `gui_matrix` import
cycle. Behavior-identical to the original `gui_api._api_method`.
"""
import logging


def _api_method(fn):
    """Wrap a js_api method: an uncaught exception in a windowed .exe would
    vanish (no stderr) and leave the UI hanging on a dead Promise, so log the
    full traceback and hand JS a structured error instead."""
    def wrapper(self, *args, **kwargs):
        try:
            return fn(self, *args, **kwargs)
        except Exception as e:
            logging.getLogger("tsmis.crash").critical(
                "uncaught exception in GUI api %s", fn.__name__, exc_info=True)
            try:
                self._emit_log(f"ERROR: {type(e).__name__}: {e} (details in the log file)")
            except Exception:
                pass
            return {"error": f"{type(e).__name__}: {e} (details are in the log file)"}
    wrapper.__name__ = fn.__name__
    return wrapper


def _snapshot_endpoint(fn):
    """A read-only js_api endpoint that returns a matrix snapshot: `_api_method`
    plus one `artifact_store.read_scope`, so the snapshot, its add-day picker and
    its day tags read each export folder ONCE between them instead of once per
    question. Never use it on an endpoint that builds or writes: a build's
    freshness checks must read the disk every time."""
    def scoped(self, *args, **kwargs):
        import artifact_store            # lazy: keeps this module import-light
        import consolidation_meta
        try:
            with artifact_store.read_scope():
                return fn(self, *args, **kwargs)
        finally:
            # Whatever this snapshot had to hash or decode is now known; keep it
            # for the next launch (throttled, and a no-op when nothing is new).
            consolidation_meta.flush_persistent_caches()
    scoped.__name__ = fn.__name__
    return _api_method(scoped)


def _task_endpoint(kind):
    """A js_api endpoint that RUNS UNDER the single-task gate (S3 / DUP-03):
    wraps `_api_method`, claims `kind` via `self._claim_task_error` — B8's
    background-check exclusion and the soft busy message live THERE, the one
    enforcement point — and only then calls the body. The body hands the gate
    to a worker (via `self._gated_queue()`) or must `self._release_task()` on
    every early return. An endpoint that validates its INPUT before claiming
    keeps the explicit two-line claim instead (claim order is part of its
    contract); new claim-first endpoints should use this decorator."""
    def deco(fn):
        @_api_method
        def wrapper(self, *args, **kwargs):
            err = self._claim_task_error(kind)
            if err:
                return err
            return fn(self, *args, **kwargs)
        wrapper.__name__ = fn.__name__
        return wrapper
    return deco


def pick_path(window, dialog_kind, **kwargs):
    """One native file/folder dialog -> a single path STRING, or None when the
    user cancels. pywebview returns a list/tuple on some backends and a bare
    string on others (pywebview trap) — this is the ONE unwrap point every
    endpoint uses instead of hand-rolling `picked[0] if isinstance(...)`."""
    picked = window.create_file_dialog(dialog_kind, **kwargs)
    if not picked:
        return None
    return str(picked[0] if isinstance(picked, (list, tuple)) else picked)


def pick_paths(window, dialog_kind, **kwargs):
    """The multi-select variant: a list of path strings ([] when cancelled)."""
    picked = window.create_file_dialog(dialog_kind, **kwargs)
    if not picked:
        return []
    if isinstance(picked, (list, tuple)):
        return [str(p) for p in picked]
    return [str(picked)]
