"""Stand-in workers for check_job_host — importable by a worker process.

They have the shape every hosted GUI worker has: plain-data arguments plus a
queue to ``put`` messages on and a cancel event, a ``run()`` that posts
``matrix_cell`` progress and ends with one ``matrix_done``.
"""
import logging
import os
import threading
import time
from pathlib import Path

log = logging.getLogger("tsmis.fixture")


class CountingWorker(threading.Thread):
    """Does `cells` steps of `step_s` seconds, writing one marker file each."""

    def __init__(self, out_dir, cells, queue, cancel_event, step_s=0.0, mode="ok"):
        super().__init__(daemon=True, name="fixture-counting")
        self.out_dir = out_dir
        self.cells = list(cells)
        self.q = queue
        self.cancel = cancel_event
        self.step_s = step_s
        self.mode = mode

    def run(self):
        done = 0
        total = len(self.cells)
        try:
            log.info("fixture worker starting in pid %s", os.getpid())
            for name in self.cells:
                if self.cancel.is_set():
                    break
                self.q.put(("matrix_cell", {"row": name, "cell": "c", "status": "running",
                                            "done": done, "total": total}))
                self.q.put(("log", f"  working on {name}"))
                if self.mode == "die" and done == 1:
                    os._exit(7)                    # a hard crash: no finally, no terminal
                if self.mode == "raise" and done == 1:
                    raise RuntimeError("fixture failure")
                time.sleep(self.step_s)
                Path(self.out_dir, f"{name}.{os.getpid()}.txt").write_text("done", "utf-8")
                done += 1
                self.q.put(("matrix_cell", {"row": name, "cell": "c", "status": "ok",
                                            "done": done, "total": total}))
        finally:
            if self.mode != "raise":
                self.q.put(("matrix_done", {"done": done, "total": total, "errors": 0,
                                            "cancelled": self.cancel.is_set(),
                                            "pid": os.getpid()}))


class LabelledWorker(CountingWorker):
    """A worker that names its run, as the ArcGIS report refresh does."""

    JOB_LABEL = "Fixture refresh"


class UnpicklableArgWorker(threading.Thread):
    """Takes a callable — not plain data, so it must keep its thread."""

    def __init__(self, fn, queue, cancel_event):
        super().__init__(daemon=True, name="fixture-callable")
        self.fn = fn
        self.q = queue
        self.cancel = cancel_event

    def run(self):
        self.q.put(("matrix_done", {"done": self.fn(), "total": 1, "errors": 0,
                                    "cancelled": False, "pid": os.getpid()}))
