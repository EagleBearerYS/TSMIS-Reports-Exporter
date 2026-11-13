"""Process-hosted GUI workers (job_host).

Comparison and build jobs run in a child process so they stop competing with the
window for the interpreter. This holds the contract that makes that safe to
ship without being able to test every locked-down PC it will run on:

  * hosting is OFF until the app enables it — the worker is the plain thread;
  * a hosted job really runs in another process, and its messages, its log
    records and the Stop button all cross the boundary;
  * a worker whose arguments are not plain data keeps its thread;
  * a child that dies mid-job still ends the job with ONE terminal message, so
    the task gate is always released;
  * a child that cannot start never costs the job: it runs in-process, and
    hosting switches itself off for the session.

Run from the repo root:  build\\.venv\\Scripts\\python.exe build\\check_job_host.py
"""
from __future__ import annotations

import logging
import os
import queue
import sys
import tempfile
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT / "scripts"), str(ROOT / "build"), str(ROOT)]

import job_host  # noqa: E402
from _job_host_fixture import (CountingWorker, LabelledWorker,  # noqa: E402
                               UnpicklableArgWorker)

_failures: list[str] = []


def check(name, cond, detail=""):
    print(f"  [{'OK ' if cond else 'FAIL'}] {name}")
    if not cond:
        if detail:
            print(f"       {detail}")
        _failures.append(name)


def _drain(q, timeout=60.0):
    """Every message up to and including the terminal one."""
    out = []
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            item = q.get(timeout=0.2)
        except queue.Empty:
            continue
        out.append(item)
        if item[0] == "matrix_done":
            return out
    return out


def _reset():
    job_host._state.update(enabled=False, disabled_reason=None, context=None)
    os.environ.pop("TSMIS_INPROCESS_WORKERS", None)


class _Capture(logging.Handler):
    def __init__(self):
        super().__init__()
        self.lines = []

    def emit(self, record):
        self.lines.append((record.name, record.getMessage()))


def main():
    tmp = Path(tempfile.mkdtemp(prefix="tsmis_job_host_"))
    me = os.getpid()

    print("hosting is opt-in:")
    _reset()
    q, cancel = queue.Queue(), threading.Event()
    worker = job_host.hosted(CountingWorker, str(tmp), ["a"], q, cancel)
    check("before enable(), hosted() returns the worker itself (a thread)",
          isinstance(worker, CountingWorker))
    worker.start()
    msgs = _drain(q)
    check("...and it runs in THIS process",
          msgs and msgs[-1][1].get("pid") == me, repr(msgs[-1:] if msgs else msgs))

    print("a hosted job runs in another process and reports back:")
    _reset()
    job_host.enable(lambda: {"log_level": logging.INFO})
    capture = _Capture()
    fixture_log = logging.getLogger("tsmis.fixture")
    fixture_log.addHandler(capture)
    fixture_log.setLevel(logging.INFO)
    q, cancel = queue.Queue(), threading.Event()
    relay = job_host.hosted(CountingWorker, str(tmp), ["a", "b", "c"], q, cancel)
    check("hosted() returns a relay, not the worker",
          not isinstance(relay, CountingWorker) and isinstance(relay, threading.Thread))
    relay.start()
    msgs = _drain(q)
    relay.join(30)
    kinds = [m[0] for m in msgs]
    terminal = msgs[-1][1] if msgs else {}
    check("every message arrived in order, ending with the terminal",
          kinds.count("matrix_cell") == 6 and kinds.count("log") == 3
          and kinds[-1] == "matrix_done" and kinds.count("matrix_done") == 1, repr(kinds))
    check("the work happened in a DIFFERENT process",
          isinstance(terminal.get("pid"), int) and terminal["pid"] != me,
          f"worker pid {terminal.get('pid')} vs this pid {me}")
    check("the terminal carries the worker's own counts",
          terminal.get("done") == 3 and terminal.get("total") == 3)
    check("the files it wrote exist",
          len(list(tmp.glob("*.txt"))) >= 4)
    check("the child's log records were replayed through this process's logging",
          any("fixture worker starting" in text for _n, text in capture.lines),
          repr(capture.lines))
    check("a healthy run leaves hosting on", job_host.enabled())

    print("Stop crosses the boundary:")
    q, cancel = queue.Queue(), threading.Event()
    relay = job_host.hosted(CountingWorker, str(tmp), [f"s{i}" for i in range(200)],
                            q, cancel, step_s=0.05)
    relay.start()
    seen = []
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        item = q.get(timeout=30)
        seen.append(item)
        if item[0] == "matrix_cell" and item[1].get("done", 0) >= 3:
            break
    cancel.set()
    rest = _drain(q)
    relay.join(30)
    terminal = rest[-1][1] if rest else {}
    check("the job stopped early and said so",
          rest and rest[-1][0] == "matrix_done" and terminal.get("cancelled") is True
          and terminal.get("done", 200) < 200, repr(terminal))

    print("a worker that is not plain data keeps its thread:")
    q, cancel = queue.Queue(), threading.Event()
    worker = job_host.hosted(UnpicklableArgWorker, lambda: 1, q, cancel)
    check("hosted() returns the worker itself", isinstance(worker, UnpicklableArgWorker))
    worker.start()
    msgs = _drain(q)
    check("...and it runs here", msgs and msgs[-1][1].get("pid") == me)

    print("a child that dies still ends the job exactly once:")
    q, cancel = queue.Queue(), threading.Event()
    relay = job_host.hosted(CountingWorker, str(tmp), ["x", "y", "z"], q, cancel,
                            mode="die")
    relay.start()
    msgs = _drain(q)
    relay.join(30)
    kinds = [m[0] for m in msgs]
    terminal = msgs[-1][1] if msgs else {}
    check("one terminal message was posted for the dead child",
          kinds.count("matrix_done") == 1 and kinds[-1] == "matrix_done", repr(kinds))
    check("it reports what was finished and what was not",
          terminal.get("done") == 1 and terminal.get("total") == 3
          and terminal.get("errors") == 2 and terminal.get("failed") == 2, repr(terminal))
    check("the user is told the worker stopped",
          any(m[0] == "log" and "stopped unexpectedly" in m[1] for m in msgs))
    check("a worker with no name for its run gets none invented",
          "label" not in terminal, repr(terminal))

    q, cancel = queue.Queue(), threading.Event()
    relay = job_host.hosted(LabelledWorker, str(tmp), ["x", "y", "z"], q, cancel,
                            mode="die")
    relay.start()
    msgs = _drain(q)
    relay.join(30)
    check("a worker that names its run keeps the name when its process dies",
          msgs and msgs[-1][0] == "matrix_done"
          and msgs[-1][1].get("label") == LabelledWorker.JOB_LABEL, repr(msgs[-1:]))

    q, cancel = queue.Queue(), threading.Event()
    relay = job_host.hosted(CountingWorker, str(tmp), ["x", "y", "z"], q, cancel,
                            mode="raise")
    relay.start()
    msgs = _drain(q)
    relay.join(30)
    check("a worker that raises without a terminal is ended the same way",
          [m[0] for m in msgs].count("matrix_done") == 1
          and msgs[-1][1].get("failed") == 2, repr(msgs[-2:]))
    check("a job that failed AFTER it started does not switch hosting off",
          job_host.enabled())

    print("a child that cannot start never costs the job:")
    _reset()
    job_host.enable(lambda: {})
    real_get_context = job_host.multiprocessing.get_context

    class _NoSpawn:
        def Pipe(self, duplex=True):
            raise OSError("process creation is not permitted")

    job_host.multiprocessing.get_context = lambda _kind: _NoSpawn()
    try:
        q, cancel = queue.Queue(), threading.Event()
        relay = job_host.hosted(CountingWorker, str(tmp), ["f1", "f2"], q, cancel)
        relay.start()
        msgs = _drain(q)
        relay.join(30)
    finally:
        job_host.multiprocessing.get_context = real_get_context
    terminal = msgs[-1][1] if msgs else {}
    check("the job ran anyway, in this process",
          msgs and msgs[-1][0] == "matrix_done" and terminal.get("done") == 2
          and terminal.get("pid") == me, repr(terminal))
    check("hosting switched itself off for the session, with a reason",
          not job_host.enabled() and "could not be started" in (job_host.disabled_reason() or ""))
    q, cancel = queue.Queue(), threading.Event()
    check("later jobs are plain threads again",
          isinstance(job_host.hosted(CountingWorker, str(tmp), [], q, cancel),
                     CountingWorker))

    print("the self-test's probe:")
    _reset()
    job_host.enable(lambda: {})
    probed = job_host.probe(modules=("json", "_job_host_fixture"))
    check("a child imports the modules it is given and answers from another process",
          probed["ok"] and probed["pid"] != me and not probed["refused"]
          and probed["detail"] == "", repr(probed))
    probed = job_host.probe(modules=("json", "tsmis_no_such_module"))
    check("a module the child cannot import fails the probe, and says which",
          not probed["ok"] and not probed["refused"]
          and "tsmis_no_such_module" in probed["detail"], repr(probed))
    check("...without switching hosting off (the process itself started fine)",
          job_host.enabled())

    print("the session switch:")
    _reset()
    job_host.enable(lambda: {})
    os.environ["TSMIS_INPROCESS_WORKERS"] = "1"
    q, cancel = queue.Queue(), threading.Event()
    check("TSMIS_INPROCESS_WORKERS=1 keeps every job in-process",
          isinstance(job_host.hosted(CountingWorker, str(tmp), [], q, cancel),
                     CountingWorker))
    _reset()


if __name__ == "__main__":
    print("job_host: GUI workers in a child process")
    main()
    if _failures:
        print(f"\n{len(_failures)} check(s) FAILED")
        raise SystemExit(1)
    print("all good")
