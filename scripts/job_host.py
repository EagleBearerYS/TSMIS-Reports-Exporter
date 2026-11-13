"""Run a GUI worker in its own process.

Every comparison, consolidation and build used to run as a thread INSIDE the
window's process. Those jobs are pure-Python and CPU-bound, so for as long as one
ran it held the interpreter lock for most of every 5 ms slice, and the threads
that make the window feel alive — the bridge call behind each click, the pump,
the sender that pushes state to the page — queued up behind it. That is the
"everything is sluggish while a comparison runs" the app had, and no amount of
tuning the job itself removes it.

`hosted(WorkerClass, *args, **kwargs)` builds the same worker the caller would
have built, but when hosting is enabled its ``run()`` happens in a child process:

  * the child constructs the worker with a stand-in queue and a stand-in cancel
    event, calls ``run()`` and exits — so the job's memory goes back to the OS
    when it ends, too;
  * a small relay thread in the window's process forwards every message the
    worker posts onto the real (epoch-stamped) queue, forwards Stop the other
    way, and re-emits the child's log records through the app's own logging, so
    ``tsmis.log`` reads exactly as before;
  * nothing about the job changes: same worker class, same arguments, same files.

It is deliberately conservative about when it applies:

  * OFF unless the running app turns it on (`enable`): the check suite and the
    console flow keep plain threads, and so does anything that patches a worker's
    collaborators in this process;
  * only for workers whose arguments and messages are plain data — it verifies
    the arguments pickle before starting a process and otherwise returns the
    ordinary thread. A hosted worker's ``__init__`` runs in BOTH processes (here,
    so a bad argument still raises at the call site and the fallback has an
    object to run; and again in the child), so it must only store its arguments;
  * if the child cannot be started, or does not report ready in time (a locked-
    down PC may refuse the spawn), the job runs in-process as it always has and
    hosting stays off for the rest of the session. A failed spawn can cost a
    delay; it can never cost the job.

If the child dies after it started work, the relay posts the job's terminal
message itself (naming the exit code), so the task gate is always released.
``TSMIS_INPROCESS_WORKERS=1`` forces the in-process path.
"""
from __future__ import annotations

import importlib
import logging
import multiprocessing
import os
import pickle
import threading
import time
import traceback

import contract

log = logging.getLogger("tsmis.jobhost")

__all__ = ["enable", "enabled", "hosted", "disabled_reason"]

_QUEUE_TOKEN = "__tsmis_job_queue__"
_CANCEL_TOKEN = "__tsmis_job_cancel__"
_READY = "__job_ready__"
_EXIT = "__job_exit__"
_LOG_RECORD = "__job_log__"

# How long a child may take to report ready. A cold start of the packaged exe
# under an on-access scanner is a few seconds; this is generous on purpose —
# missing it only moves the job in-process, it does not fail it.
START_TIMEOUT_S = 90.0
_POLL_S = 0.1

_state = {"enabled": False, "disabled_reason": None, "context": None}
_state_lock = threading.Lock()


def enable(context=None):
    """Turn process hosting on for this (GUI) process. `context()` returns the
    process-wide facts a child must start with (see `_child_context`)."""
    with _state_lock:
        _state["enabled"] = True
        _state["context"] = context


def enabled():
    """True when a new job would be hosted in a child process."""
    if os.environ.get("TSMIS_INPROCESS_WORKERS", "").strip().lower() in ("1", "true", "yes"):
        return False
    with _state_lock:
        return bool(_state["enabled"] and not _state["disabled_reason"])


def disabled_reason():
    """Why hosting switched itself off this session, or None."""
    with _state_lock:
        return _state["disabled_reason"]


def _disable(reason):
    with _state_lock:
        if not _state["disabled_reason"]:
            _state["disabled_reason"] = reason
    log.warning("worker processes disabled for this session: %s", reason)


def _is_queue(value):
    return (not isinstance(value, (str, bytes, dict, list, tuple, set))
            and callable(getattr(value, "put", None)))


def _extract(args, kwargs):
    """Swap the queue and the cancel event for tokens; return the real pair."""
    found = {"queue": [], "cancel": []}

    def swap(value):
        if isinstance(value, threading.Event):
            found["cancel"].append(value)
            return _CANCEL_TOKEN
        if _is_queue(value):
            found["queue"].append(value)
            return _QUEUE_TOKEN
        return value

    new_args = tuple(swap(a) for a in args)
    new_kwargs = {k: swap(v) for k, v in kwargs.items()}
    if len(found["queue"]) != 1 or len(found["cancel"]) != 1:
        return None
    return new_args, new_kwargs, found["queue"][0], found["cancel"][0]


def hosted(worker_cls, *args, **kwargs):
    """`worker_cls(*args, **kwargs)`, to be ``.start()``-ed by the caller.

    Returns a relay thread that runs the worker in a child process when hosting
    is on and the worker's arguments are plain data; otherwise the worker
    itself."""
    # Built here, by the caller's thread, exactly as before: a bad argument still
    # raises where it always did, and this is the object that runs if the job
    # ends up in-process after all.
    worker = worker_cls(*args, **kwargs)
    if not enabled():
        return worker
    extracted = _extract(args, kwargs)
    if extracted is None:
        return worker
    new_args, new_kwargs, real_queue, real_cancel = extracted
    try:
        pickle.dumps((new_args, new_kwargs))
    except Exception as e:  # noqa: BLE001  # silent-ok: a non-plain argument keeps the thread
        log.info("%s keeps its thread: arguments are not plain data (%s: %s)",
                 worker_cls.__name__, type(e).__name__, e)
        return worker
    return _Relay(worker, new_args, new_kwargs, real_queue, real_cancel)


class _Relay(threading.Thread):
    """Window-side half of one hosted job."""

    def __init__(self, worker, child_args, child_kwargs, real_queue, real_cancel):
        worker_cls = type(worker)
        super().__init__(daemon=True, name=f"host-{worker_cls.__name__}")
        self._worker = worker
        self._worker_cls = worker_cls
        self._child_args = child_args
        self._child_kwargs = child_kwargs
        self._queue = real_queue
        self._cancel = real_cancel
        self._progress = {"done": 0, "total": 0}
        self.crash_text = None       # the child's traceback, when it reported one

    def _run_inline(self):
        """The pre-hosting behaviour: the worker's own ``run()``, on this thread."""
        self._worker.run()

    def run(self):
        try:
            ctx = multiprocessing.get_context("spawn")
            context = _state["context"]() if _state["context"] else {}
            # A one-way pipe the child writes SYNCHRONOUSLY: a message the
            # worker has posted is in the pipe before its `put` returns, so a
            # child that dies outright loses nothing it had already reported.
            # (multiprocessing.Queue hands messages to a feeder thread, and a
            # hard exit drops whatever that thread had not written yet.)
            inbox, outbox = ctx.Pipe(duplex=False)
            cancel = ctx.Event()
            proc = ctx.Process(
                target=_child_main, daemon=True,
                name=f"tsmis-{self._worker_cls.__name__}",
                args=(self._worker_cls.__module__, self._worker_cls.__name__,
                      self._child_args, self._child_kwargs, outbox, cancel, context))
            proc.start()
            outbox.close()               # the child holds the only write end now
        except Exception as e:  # noqa: BLE001 — any spawn failure means "run it here"
            _disable(f"a worker process could not be started ({type(e).__name__}: {e})")
            self._run_inline()
            return

        ready = terminal = False
        started = time.monotonic()
        crash_text = None
        try:
            while True:
                if self._cancel.is_set() and not cancel.is_set():
                    cancel.set()
                try:
                    if not inbox.poll(_POLL_S):
                        if not ready and time.monotonic() - started > START_TIMEOUT_S:
                            break
                        continue
                    raw = inbox.recv_bytes()
                except (EOFError, OSError):  # silent-ok: the child closed its end — it is gone
                    break
                try:
                    kind, payload = pickle.loads(raw)
                except Exception as e:  # noqa: BLE001
                    log.warning("dropped an undecodable worker message (%s: %s)",
                                type(e).__name__, e)
                    continue
                if kind == _READY:
                    ready = True
                    log.info("%s running in worker process %s",
                             self._worker_cls.__name__, payload)
                elif kind == _LOG_RECORD:
                    _replay_log(payload)
                elif kind == _EXIT:
                    crash_text = payload
                    break
                else:
                    if kind == contract.Msg.MATRIX_CELL and isinstance(payload, dict):
                        self._progress["done"] = payload.get("done", 0) or 0
                        self._progress["total"] = payload.get("total", 0) or 0
                    if kind in contract.TERMINAL:
                        terminal = True
                    self._queue.put((kind, payload))
        finally:
            if not ready:
                # Never reported ready: stop it before the job is run here, so
                # two copies can never work on the same files.
                self._kill_child(proc)
            else:
                proc.join(timeout=10)
                if proc.is_alive():
                    self._kill_child(proc)
            try:
                inbox.close()
            except OSError:  # silent-ok: already closed with the child
                pass
        if not ready:
            _disable("a worker process did not start in time"
                     if proc.exitcode is None or proc.exitcode == 0
                     else f"a worker process exited at startup (code {proc.exitcode})")
            self._run_inline()
            return
        self.crash_text = crash_text
        self._after_job()
        if not terminal:
            self._post_crash_terminal(proc.exitcode, crash_text)

    @staticmethod
    def _kill_child(proc):
        try:
            proc.terminate()
            proc.join(timeout=10)
        except Exception as e:  # noqa: BLE001
            log.warning("worker process could not be stopped (%s: %s)",
                        type(e).__name__, e)

    @staticmethod
    def _after_job():
        """Pick up what the child learned (its digests and comparison headlines
        went to the shared cache files when it exited)."""
        try:
            import consolidation_meta
            consolidation_meta.reload_persistent_caches()
        except Exception as e:  # noqa: BLE001
            log.info("cache reload after a hosted job skipped (%s: %s)",
                     type(e).__name__, e)

    def _post_crash_terminal(self, exitcode, crash_text):
        """The child ended without its terminal message — release the gate."""
        done, total = self._progress["done"], self._progress["total"]
        log.error("%s worker process ended without finishing (exit code %s)%s",
                  self._worker_cls.__name__, exitcode,
                  f":\n{crash_text}" if crash_text else "")
        self._queue.put((contract.Msg.LOG,
                         "  The worker process stopped unexpectedly"
                         + (f" (exit code {exitcode})" if exitcode not in (None, 0) else "")
                         + " — what it had not finished was not built. "
                           "Details are in the log file."))
        remaining = max(0, total - done)
        payload = {
            "done": done, "total": total, "errors": remaining,
            "cancelled": bool(self._cancel.is_set()), "attempted": done,
            "succeeded": 0, "failed": remaining, "cancelled_cells": 0,
            "partial_cells": 0}
        # A worker that names its run in its own terminal message (JOB_LABEL)
        # is named the same way when the relay has to end it.
        label = getattr(self._worker_cls, "JOB_LABEL", None)
        if label:
            payload["label"] = label
        self._queue.put((contract.Msg.MATRIX_DONE, payload))


class ProbeWorker(threading.Thread):
    """A worker that imports the modules it is given, then says which process
    it ran in (`probe`)."""

    def __init__(self, token, modules, queue, cancel_event):
        super().__init__(daemon=True, name="worker-probe")
        self.token = token
        self.modules = tuple(modules or ())
        self.q = queue
        self.cancel = cancel_event

    def run(self):
        for name in self.modules:
            importlib.import_module(name)
        self.q.put((contract.Msg.MATRIX_DONE, {
            "done": 1, "total": 1, "errors": 0, "cancelled": False,
            "pid": os.getpid(), "token": self.token,
            "imported": list(self.modules)}))


def probe(modules=(), timeout=START_TIMEOUT_S + 30.0):
    """Start one worker in a child process, have it import `modules`, and wait
    for its answer.

    The packaged build's ``--self-test`` calls this, because a frozen exe hosts
    a child by re-launching ITSELF, and only the real artifact can prove that
    works — and that the modules a real job needs import in that child. Returns
    ``{ok, pid, seconds, reason, refused, detail}``: `ok` means a different
    process imported everything and answered; `refused` means the operating
    system would not create the process at all (an environment restriction,
    where the app simply keeps its jobs in-process) as opposed to a child that
    started and failed; `detail` is the last line of the child's traceback when
    it failed."""
    import queue as queue_module
    answers = queue_module.Queue()
    modules = tuple(modules or ())
    worker = ProbeWorker("probe", modules, answers, threading.Event())
    relay = _Relay(worker, ("probe", modules, _QUEUE_TOKEN, _CANCEL_TOKEN), {},
                   answers, worker.cancel)
    started = time.monotonic()
    relay.start()
    payload = {}
    try:
        while True:
            kind, body = answers.get(timeout=timeout)[:2]
            if kind == contract.Msg.MATRIX_DONE:
                payload = body if isinstance(body, dict) else {}
                break
    except queue_module.Empty:  # silent-ok: reported as a failed probe just below
        pass
    seconds = time.monotonic() - started
    reason = disabled_reason()
    pid = payload.get("pid")
    ok = (isinstance(pid, int) and pid != os.getpid() and payload.get("token") == "probe"
          and tuple(payload.get("imported") or ()) == modules)
    crash = (relay.crash_text or "").strip()
    return {"ok": ok, "pid": pid, "seconds": seconds, "reason": reason,
            "refused": bool(reason and "could not be started" in reason),
            "detail": crash.splitlines()[-1] if crash else ""}


def _replay_log(record):
    """Re-emit one child log record through this process's handlers."""
    try:
        logger = logging.getLogger(record.get("name") or "tsmis.worker")
        level = int(record.get("level", logging.INFO))
        if logger.isEnabledFor(level):
            text = str(record.get("msg", ""))
            if record.get("exc"):
                text = f"{text}\n{record['exc']}"
            logger.log(level, "%s", text)
    except Exception:  # noqa: BLE001  # silent-ok: a log line must never break the relay
        pass


# --------------------------------------------------------------------------- #
# the child
# --------------------------------------------------------------------------- #
class _Outbox:
    """The child's end of the pipe: one whole message per ``send``, written
    before ``send`` returns, from whichever thread posts it."""

    def __init__(self, connection):
        self._connection = connection
        self._lock = threading.Lock()

    def send(self, kind, payload):
        raw = pickle.dumps((kind, payload))
        with self._lock:
            self._connection.send_bytes(raw)

    def close(self):
        with self._lock:
            self._connection.close()


class _ChildQueue:
    """What the worker sees as its queue: ``put`` ships the message to the relay.

    A payload that cannot cross the process boundary fails HERE, at the ``put``
    that sent it; a terminal message is then still delivered (naming the fault)
    so the job can end."""

    def __init__(self, outbox):
        self._outbox = outbox

    def put(self, item):
        kind, payload = item[0], item[1]
        try:
            self._outbox.send(kind, payload)
        except (pickle.PicklingError, TypeError, AttributeError) as e:
            logging.getLogger("tsmis.jobhost").error(
                "worker message %r could not be sent (%s: %s)", kind,
                type(e).__name__, e)
            if kind in contract.TERMINAL:
                self._outbox.send(kind, {"error": f"{type(e).__name__}: {e}"})


class _ForwardLogs(logging.Handler):
    def __init__(self, outbox):
        super().__init__()
        self._outbox = outbox

    def emit(self, record):
        try:
            exc = None
            if record.exc_info:
                exc = "".join(traceback.format_exception(*record.exc_info)).rstrip()
            self._outbox.send(_LOG_RECORD, {
                "name": record.name, "level": record.levelno,
                "msg": record.getMessage(), "exc": exc})
        except Exception:  # noqa: BLE001  # silent-ok: logging must never raise
            pass


def _watch_parent():
    """Exit if the window's process goes away (a killed app must not leave a
    comparison grinding on with nobody to report to)."""
    parent = multiprocessing.parent_process()
    if parent is None:
        return

    def wait():
        parent.join()
        os._exit(3)

    threading.Thread(target=wait, daemon=True, name="parent-watch").start()


def _child_context():
    """The default `enable` context: what a child needs from this process."""
    context = {"log_level": logging.getLogger().getEffectiveLevel()}
    try:
        from common import get_site
        context["site"] = tuple(get_site())
    except Exception as e:  # noqa: BLE001
        log.info("worker context: site unavailable (%s: %s)", type(e).__name__, e)
    try:
        from paths import CACHE_DIR
        context["cache_dir"] = str(CACHE_DIR)
    except Exception as e:  # noqa: BLE001
        log.info("worker context: cache dir unavailable (%s: %s)", type(e).__name__, e)
    return context


def _child_main(module_name, class_name, args, kwargs, connection, cancel, context):
    """Entry point of the worker process."""
    crash = None
    out = _Outbox(connection)
    try:
        root = logging.getLogger()
        for handler in list(root.handlers):
            root.removeHandler(handler)
        root.addHandler(_ForwardLogs(out))
        root.setLevel(int(context.get("log_level", logging.INFO)))
        _watch_parent()
        site = context.get("site")
        if site:
            from common import set_site
            set_site(*site)
        cache_dir = context.get("cache_dir")
        if cache_dir:
            import consolidation_meta
            consolidation_meta.enable_persistent_caches(cache_dir)
        worker_cls = getattr(importlib.import_module(module_name), class_name)
        child_queue = _ChildQueue(out)

        def restore(value):
            if isinstance(value, str) and value == _QUEUE_TOKEN:
                return child_queue
            if isinstance(value, str) and value == _CANCEL_TOKEN:
                return cancel
            return value

        worker = worker_cls(*[restore(a) for a in args],
                            **{k: restore(v) for k, v in kwargs.items()})
        out.send(_READY, os.getpid())
        worker.run()
    except BaseException as e:  # noqa: BLE001 — reported to the relay, which ends the job
        crash = "".join(traceback.format_exception(type(e), e, e.__traceback__))
    finally:
        try:
            if context.get("cache_dir"):
                import consolidation_meta
                consolidation_meta.flush_persistent_caches(force=True)
        except Exception:  # noqa: BLE001  # silent-ok: a cache write must not mask the job's outcome
            pass
        try:
            out.send(_EXIT, crash)
            out.close()
        except Exception:  # noqa: BLE001  # silent-ok: the relay notices the exit either way
            pass
