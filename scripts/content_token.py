"""Change tokens and the token-validated digest memo (CMP-AUD-080).

A file's SHA-256 may be reused only while a CHANGE TOKEN proves the bytes are the
ones that were hashed. On Windows the token carries the filesystem's own
ChangeTime (``FILE_BASIC_INFO.ChangeTime``), which the OS advances on any write to
the file's data or metadata and which ``SetFileTime`` cannot restore: an in-place
same-size rewrite with the mtime put back leaves (size, mtime_ns, file id)
identical while ChangeTime moves, and a whole-file replacement moves the file id
as well. Where no change token can be read the memo is refused and the file is
hashed again.

This is a leaf module (stdlib only) so every reader that proves content identity
shares ONE memo: ``artifact_store.content_digest`` (source folders, evidence),
``consolidation_meta._bound_file_digest`` (published comparison members) and
``tsn_library.normalized_workbook_identity`` (the TSN library). A matrix snapshot
asks for the same workbook's digest several times per cell; before they shared a
memo, each of those was a full read of a file that can be hundreds of megabytes.

The memo may also be kept across launches (``enable_persistence``). That adds no
new trust: an entry is served only when the file's current token equals the
stored one, exactly as in memory, and NTFS keeps ChangeTime and the file id
across restarts.
"""
from __future__ import annotations

import json
import logging
import os
import threading
import time
from pathlib import Path

log = logging.getLogger("tsmis.identity")

HASH_BLOCK = 1 << 20
MEMO = {}                       # normcased path -> (token, sha256 hex)
MEMO_MAX = 50_000               # ~6 statewide environments' worth of routes
# FILETIME (100 ns ticks since 1601-01-01) -> unix nanoseconds.
FILETIME_EPOCH_TICKS = 116_444_736_000_000_000
# A change token can only be as fine as the system clock that stamps it: two
# writes inside one tick get the SAME ChangeTime (measured ~10 ms here, coarser
# on some virtualized hosts — a CI runner produced an unmoved token). So a digest
# is memoized only once its file's change stamp is comfortably in the PAST: a
# file changed within this window is re-hashed next time instead of trusted.
# (The same "racily clean" rule Git applies to its stat cache.) One second is far
# above any plausible tick, and the cost is only re-hashing a just-written file.
RACY_WINDOW_NS = 1_000_000_000

_PERSIST_SCHEMA = 1
_PERSIST_MIN_INTERVAL_S = 20.0
_persist_lock = threading.Lock()
_persist_path = None
_persist_dirty = False
_persist_saved_at = 0.0


def change_token(path, st):
    """A tamper-visible identity for one file, or None when none is available."""
    base = (int(st.st_size), int(getattr(st, "st_mtime_ns", 0)),
            int(getattr(st, "st_dev", 0)), int(getattr(st, "st_ino", 0)))
    if os.name != "nt":
        return None                # no verified change signal -> never memoize
    change_time = windows_change_time(path)
    return None if change_time is None else base + (change_time,)


_WIN32_METADATA = None


def _win32_metadata_api():
    """``(ctypes, kernel32, FILE_BASIC_INFO, INVALID_HANDLE)`` or None. Built once,
    lazily: ctypes stays off every non-Windows and non-fingerprint path."""
    global _WIN32_METADATA
    if _WIN32_METADATA is None:
        _WIN32_METADATA = False
        if os.name == "nt":
            try:
                import ctypes
                from ctypes import wintypes
                k32 = ctypes.WinDLL("kernel32", use_last_error=True)
                k32.CreateFileW.restype = wintypes.HANDLE
                k32.CreateFileW.argtypes = [
                    wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                    ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                    wintypes.HANDLE]
                k32.GetFileInformationByHandleEx.restype = wintypes.BOOL
                k32.GetFileInformationByHandleEx.argtypes = [
                    wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD]
                k32.GetFileInformationByHandle.restype = wintypes.BOOL
                k32.GetFileInformationByHandle.argtypes = [
                    wintypes.HANDLE, ctypes.c_void_p]
                k32.CloseHandle.argtypes = [wintypes.HANDLE]

                class FILE_BASIC_INFO(ctypes.Structure):
                    _fields_ = [("CreationTime", ctypes.c_longlong),
                                ("LastAccessTime", ctypes.c_longlong),
                                ("LastWriteTime", ctypes.c_longlong),
                                ("ChangeTime", ctypes.c_longlong),
                                ("FileAttributes", wintypes.DWORD)]

                class BY_HANDLE_FILE_INFORMATION(ctypes.Structure):
                    # FILETIME is two DWORDs, so the 64-bit views below must
                    # not be 8-byte aligned. (_pack_ has to precede _fields_.)
                    _pack_ = 4
                    _fields_ = [("dwFileAttributes", wintypes.DWORD),
                                ("ftCreationTime", ctypes.c_ulonglong),
                                ("ftLastAccessTime", ctypes.c_ulonglong),
                                ("ftLastWriteTime", ctypes.c_ulonglong),
                                ("dwVolumeSerialNumber", wintypes.DWORD),
                                ("nFileSizeHigh", wintypes.DWORD),
                                ("nFileSizeLow", wintypes.DWORD),
                                ("nNumberOfLinks", wintypes.DWORD),
                                ("nFileIndexHigh", wintypes.DWORD),
                                ("nFileIndexLow", wintypes.DWORD)]

                invalid = ctypes.cast(wintypes.HANDLE(-1), ctypes.c_void_p).value
                _WIN32_METADATA = (ctypes, k32, FILE_BASIC_INFO, invalid,
                                   BY_HANDLE_FILE_INFORMATION)
            except (OSError, AttributeError, ValueError) as e:
                log.warning("content identity: no Windows change token available "
                            "(%s: %s); source digests will be recomputed",
                            type(e).__name__, e)
    return _WIN32_METADATA or None


def windows_change_time(path):
    """FILE_BASIC_INFO.ChangeTime for `path` (None when it can't be read).

    Measured cheaper than ``os.stat`` (one metadata handle, no path
    re-resolution), so validating the memo costs nothing next to the v1 stat walk
    it replaces."""
    api = _win32_metadata_api()
    if api is None:
        return None
    ctypes, k32, basic_info, invalid, _by_handle = api
    # 0 access + share-all + FILE_FLAG_BACKUP_SEMANTICS: metadata only, never
    # blocking a workbook another process has open.
    handle = k32.CreateFileW(str(path), 0, 7, None, 3, 0x02000000, None)
    if not handle or handle == invalid:
        return None
    try:
        info = basic_info()
        if not k32.GetFileInformationByHandleEx(
                handle, 0, ctypes.byref(info), ctypes.sizeof(info)):
            return None
        return int(info.ChangeTime)
    finally:
        k32.CloseHandle(handle)


# None = not yet decided, True = the one-handle read agrees with os.stat on this
# interpreter, False = it does not (or cannot run) and is never used.
_ONE_HANDLE_OK = None


def _one_handle_token(path):
    """The five-part token for `path` from ONE metadata handle, or None."""
    api = _win32_metadata_api()
    if api is None:
        return None
    ctypes, k32, basic_info, invalid, by_handle = api
    handle = k32.CreateFileW(str(path), 0, 7, None, 3, 0x02000000, None)
    if not handle or handle == invalid:
        return None
    try:
        info = by_handle()
        basic = basic_info()
        if (not k32.GetFileInformationByHandle(handle, ctypes.byref(info))
                or not k32.GetFileInformationByHandleEx(
                    handle, 0, ctypes.byref(basic), ctypes.sizeof(basic))):
            return None
        return ((int(info.nFileSizeHigh) << 32) | int(info.nFileSizeLow),
                (int(info.ftLastWriteTime) - FILETIME_EPOCH_TICKS) * 100,
                int(info.dwVolumeSerialNumber),
                (int(info.nFileIndexHigh) << 32) | int(info.nFileIndexLow),
                int(basic.ChangeTime))
    finally:
        k32.CloseHandle(handle)


def path_token(path):
    """`change_token(path, os.stat(path))`, or None when the file has no token.

    A folder fingerprint needs the token of every file in it, and the two-call
    form opens each file twice (once inside ``os.stat``, once for ChangeTime) —
    most of what a warm statewide snapshot spent. This reads the same five facts
    through one handle. The two forms must agree EXACTLY, because tokens from
    both are compared with each other in the memo, so the one-handle form is
    used only after it has reproduced ``os.stat``'s token on this interpreter;
    until then, and on any disagreement, the two-call form is the answer. Never
    raises: an unreadable file has no token, and the caller's own read reports
    why."""
    global _ONE_HANDLE_OK
    if _ONE_HANDLE_OK:
        return _one_handle_token(path)
    try:
        reference = change_token(path, os.stat(path))
    except (OSError, ValueError):  # silent-ok: no token; the caller's read reports the cause
        return None
    if _ONE_HANDLE_OK is None and reference is not None:
        try:
            fast = _one_handle_token(path)
        except Exception as e:      # noqa: BLE001 — any ctypes surprise disables it
            fast = None
            log.info("content identity: one-handle token unavailable (%s: %s)",
                     type(e).__name__, e)
        _ONE_HANDLE_OK = fast == reference
        if not _ONE_HANDLE_OK:
            log.info("content identity: one-handle token differs from os.stat's; "
                     "using the two-call form")
    return reference


def memoizable(token):
    """True when `token`'s change stamp is old enough to be trusted (see
    ``RACY_WINDOW_NS``). A token with no change stamp is never memoizable."""
    if token is None:
        return False
    changed_ns = (token[-1] - FILETIME_EPOCH_TICKS) * 100
    return (time.time_ns() - changed_ns) > RACY_WINDOW_NS


def memo_key(path):
    return os.path.normcase(str(path))


def lookup(path, token):
    """The memoized SHA-256 for `path` when its stored token equals `token`."""
    if token is None:
        return None
    cached = MEMO.get(memo_key(path))
    if cached is not None and cached[0] == token:
        return cached[1]
    return None


def remember(path, token, digest):
    """Memoize `digest` for `path` under `token`.

    The CALLER has already proved the token did not move across its read; this
    adds the racy-window rule, so a just-written file is never trusted."""
    global _persist_dirty
    if token is None or not memoizable(token):
        return False
    if len(MEMO) >= MEMO_MAX:
        MEMO.clear()               # bounded: a cleared memo only costs a re-hash
    MEMO[memo_key(path)] = (token, digest)
    _persist_dirty = True
    return True


def _valid_entry(key, value):
    if not isinstance(key, str) or not isinstance(value, list) or len(value) != 6:
        return None
    *token, digest = value
    if not all(isinstance(part, int) and not isinstance(part, bool) for part in token):
        return None
    # A zero change stamp means the filesystem did not report one; such a token
    # is no stronger than stat, so it is never carried across launches.
    if token[-1] <= 0:
        return None
    if (not isinstance(digest, str) or len(digest) != 64
            or any(ch not in "0123456789abcdef" for ch in digest)):
        return None
    return tuple(token), digest


def reload():
    """Merge the memo file into memory again — after another process (a hosted
    job) wrote what it learned there. Unlike the first load, a file entry
    replaces a memory entry whose token differs: the file is the newer view of
    a file the other process just rewrote, and either way an entry is only ever
    served against the file's current token."""
    with _persist_lock:
        path = _persist_path
    if path is None:
        return 0
    return enable_persistence(path, replace=True)


def enable_persistence(path, replace=False):
    """Keep the memo in `path` across launches. Loads what is there now; entries
    already in memory win unless `replace`. Never raises: a missing or unreadable
    file is an empty memo, which only costs hashing."""
    global _persist_path
    path = Path(path)
    loaded = 0
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:  # silent-ok: no memo yet is the first-launch state
        data = None
    except (OSError, ValueError) as e:
        log.info("digest memo %s unreadable (%s: %s); starting empty",
                 path.name, type(e).__name__, e)
        data = None
    if (isinstance(data, dict) and data.get("schema_version") == _PERSIST_SCHEMA
            and isinstance(data.get("entries"), dict)):
        for key, value in data["entries"].items():
            entry = _valid_entry(key, value)
            if entry is None:
                continue
            known = MEMO.get(key)
            if known is None:
                if len(MEMO) >= MEMO_MAX:
                    continue
            elif not replace or known == entry:
                continue
            MEMO[key] = entry
            loaded += 1
    with _persist_lock:
        _persist_path = path
    log.info("digest memo: %d entries loaded from %s", loaded, path)
    return loaded


def flush(force=False):
    """Write the memo to its file when it changed. Throttled unless `force`.
    Returns True when a file was written. Never raises."""
    global _persist_dirty, _persist_saved_at
    with _persist_lock:
        path = _persist_path
        if path is None or not _persist_dirty:
            return False
        now = time.monotonic()
        if not force and now - _persist_saved_at < _PERSIST_MIN_INTERVAL_S:
            return False
        entries = {key: [*token, digest]
                   for key, (token, digest) in dict(MEMO).items()
                   if isinstance(token, tuple) and len(token) == 5 and token[-1] > 0
                   and isinstance(digest, str) and len(digest) == 64}
        tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump({"schema_version": _PERSIST_SCHEMA, "entries": entries}, f,
                          separators=(",", ":"))
            os.replace(tmp, path)
        except OSError as e:
            log.info("digest memo could not be saved (%s: %s)", type(e).__name__, e)
            try:
                tmp.unlink()
            except OSError:  # silent-ok: a leftover temp is harmless and named by pid
                pass
            return False
        _persist_dirty = False
        _persist_saved_at = now
        return True
