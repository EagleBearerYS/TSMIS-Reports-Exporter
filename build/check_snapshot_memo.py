"""The read-side memos behind a fast matrix snapshot.

Every matrix snapshot used to re-hash every built comparison workbook three
times and decode its whole typed payload — seconds per tab open once a few day
columns were built. Three things now make a repeat snapshot cheap, and each is
only allowed to skip work that a change token proves unnecessary:

  * ONE change-token digest memo (`content_token`) shared by the source-folder
    fingerprint, the published-comparison member digest and the TSN library
    identity;
  * `consolidation_meta.read_comparison_headline` — the strict reader's verdict,
    reduced to what a cell renders, kept against the token of every file the
    strict reader consulted;
  * `artifact_store.read_scope` — one read of each folder per snapshot.

This holds, for each: it is really used; it is never served across the audit's
tamper (same size, same mtime, same file id, different bytes); and it survives a
restart without being trusted any further than it is in memory.

Run from the repo root:  build\\.venv\\Scripts\\python.exe build\\check_snapshot_memo.py
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT / "scripts"), str(ROOT / "build"), str(ROOT)]

import artifact_store as a  # noqa: E402
import compare_core  # noqa: E402
import consolidation_meta as cm  # noqa: E402
import content_token as ct  # noqa: E402
import matrix_state  # noqa: E402
import tsn_library  # noqa: E402
from _checklib import publish_bound_comparison  # noqa: E402

_failures: list[str] = []
WINDOWS = os.name == "nt"


def check(name, cond, detail=""):
    print(f"  [{'OK ' if cond else 'FAIL'}] {name}")
    if not cond:
        if detail:
            print(f"       {detail}")
        _failures.append(name)


def _tamper_in_place(path, filler=b"Z"):
    """Overwrite with different bytes of the SAME length and put the times back."""
    path = Path(path)
    before = os.stat(path)
    data = path.read_bytes()
    with open(path, "r+b") as f:
        f.write(filler * len(data))
    os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns))
    after = os.stat(path)
    return ((before.st_size, before.st_mtime_ns, before.st_ino)
            == (after.st_size, after.st_mtime_ns, after.st_ino))


def _settle():
    time.sleep(1.15)                     # past content_token.RACY_WINDOW_NS


SCHEMA = compare_core.CompareSchema(
    report_name="Snapshot Memo", header=["Loc", "Text"], side_a="TSMIS", side_b="TSN",
    id_noun="row", id_noun_plural="rows")


def _publish(folder, name, right_text="changed"):
    folder.mkdir(parents=True, exist_ok=True)
    side_a, side_b = folder / f"{name}-a.xlsx", folder / f"{name}-b.xlsx"
    side_a.write_bytes(b"side a")
    side_b.write_bytes(b"side b")
    book = folder / f"{name}.xlsx"
    publish_bound_comparison(book, SCHEMA, [["1", "a"], ["2", "b"]],
                             [["1", "a"], ["2", right_text]], (side_a, side_b),
                             has_route=False)
    return book


def tokens(tmp):
    print("one token, read either way:")
    probe = tmp / "probe.bin"
    probe.write_bytes(b"A" * 4096)
    if not WINDOWS:
        check("no change token off Windows, so nothing is ever memoized",
              ct.path_token(probe) is None and ct.change_token(probe, os.stat(probe)) is None)
        return
    reference = ct.change_token(probe, os.stat(probe))
    check("the one-handle token equals the os.stat token",
          ct.path_token(probe) == reference and ct._one_handle_token(probe) == reference,
          f"{ct._one_handle_token(probe)} vs {reference}")
    check("...and the one-handle form is in use after agreeing once",
          ct._ONE_HANDLE_OK is True)
    check("a missing file has no token (and does not raise)",
          ct.path_token(tmp / "absent.bin") is None)
    # A change stamp is only as fine as the clock behind it: a rewrite inside the
    # same tick as the last write leaves the token unmoved (a CI runner did exactly
    # that to this check). That is the case the racy window exists for, so the two
    # claims are held apart: a fresh stamp is never trusted, and a SETTLED file --
    # the only kind a memo holds -- cannot be rewritten without its token moving.
    now_ticks = ct.FILETIME_EPOCH_TICKS + time.time_ns() // 100
    check("a change stamp inside the racy window is not trusted, an older one is",
          not ct.memoizable(reference[:-1] + (now_ticks,))
          and ct.memoizable(reference[:-1] + (now_ticks - 50_000_000,)))
    _settle()
    settled = ct.path_token(probe)
    check("a file left alone keeps its token, and it is trusted once settled",
          settled == reference and ct.memoizable(settled))
    _tamper_in_place(probe)
    check("the token moves on a same-size, timestamp-restored rewrite of a settled file",
          ct.path_token(probe) != settled)


def bound_digest(tmp):
    print("the published-member digest is memoized against the token:")
    member = tmp / "member.bin"
    member.write_bytes(b"M" * 8192)
    _settle()
    facts = cm._bound_file_digest(member)
    key = ct.memo_key(member)
    if not WINDOWS:
        check("off Windows the digest is simply recomputed", key not in ct.MEMO)
        return
    check("a settled member's digest is stored", key in ct.MEMO)
    token, real = ct.MEMO[key]
    ct.MEMO[key] = (token, "0" * 64)
    served = cm._bound_file_digest(member)
    check("a matching token serves the memo (no re-read)",
          served and served["sha256"] == "0" * 64 and served["size"] == 8192
          and served["mtime_ns"] == facts["mtime_ns"])
    ct.MEMO[key] = (token, real)
    check("the source-folder digest reads the SAME memo",
          a.content_digest(member) == real == facts["sha256"])
    identical = _tamper_in_place(member, b"Q")
    after = cm._bound_file_digest(member)
    check("a same-metadata rewrite is never served from it",
          identical and after and after["sha256"] != real)

    workbook = tmp / "tsn.xlsx"
    workbook.write_bytes(b"T" * 4096)
    _settle()
    identity = tsn_library.normalized_workbook_identity(workbook)
    key = ct.memo_key(workbook)
    token, real = ct.MEMO[key]
    ct.MEMO[key] = (token, "1" * 64)
    check("the TSN workbook identity is served from the memo too",
          tsn_library.normalized_workbook_identity(workbook)["sha256"] == "1" * 64)
    ct.MEMO[key] = (token, real)
    _tamper_in_place(workbook, b"U")
    check("...and re-read after a same-metadata rewrite",
          tsn_library.normalized_workbook_identity(workbook)["sha256"]
          not in (real, "1" * 64, identity["sha256"]))


def headline(tmp):
    print("the comparison headline is the strict reader's answer, kept by token:")
    book = _publish(tmp / "cmp", "clean")
    strict = cm.read_comparison_outcome(book)
    first = cm.read_comparison_headline(book)
    check("a published comparison reads trusted",
          strict is not None and strict.trusted and first is not None and first.trusted)
    check("the headline carries exactly the strict record's rendered values",
          first == cm._headline_of(strict)
          and first.differing_cells == strict.comparison_outcome.counts.differing_cells == 1
          and first.generation_id == strict.artifact_generation.generation_id
          and first.verdict == strict.comparison_outcome.verdict)

    calls = []
    real_reader = cm.read_comparison_outcome

    def counted(path):
        calls.append(str(path))
        return real_reader(path)

    cm.read_comparison_outcome = counted
    try:
        cm._HEADLINE_MEMO.clear()
        cm.read_comparison_headline(book)             # just written: inside the racy window
        cm.read_comparison_headline(book)
        fresh_calls = len(calls)
        if WINDOWS:
            check("a just-written generation is NOT memoized (the racy window)",
                  fresh_calls == 2, f"strict reads: {fresh_calls}")
        _settle()
        calls.clear()
        cm._HEADLINE_MEMO.clear()
        settled = cm.read_comparison_headline(book)
        again = cm.read_comparison_headline(book)
        if WINDOWS:
            check("once settled, a repeat read does not run the strict reader",
                  len(calls) == 1 and again == settled == first, f"strict reads: {len(calls)}")
            dependencies = cm._HEADLINE_MEMO[os.path.normcase(str(book))][0]
            names = {Path(text).name for text, _state in dependencies}
            check("its dependencies include the workbook, its sidecar and a sentinel slot",
                  book.name in names and any(n.endswith(".outcome.json") for n in names)
                  and any(state == cm._ABSENT_STATE for _t, state in dependencies),
                  repr(sorted(names)))
        else:
            check("off Windows every read is the strict reader", len(calls) == 2)

        # --- the audit's tamper, then every other dependency --------------
        calls.clear()
        identical = _tamper_in_place(book)
        tampered = cm.read_comparison_headline(book)
        check("a same-metadata rewrite of the workbook is NOT served: it reads untrusted",
              identical and tampered is not None and not tampered.trusted and len(calls) == 1,
              repr(tampered))

        book2 = _publish(tmp / "cmp2", "second")
        _settle()
        cm.read_comparison_headline(book2)
        sidecar = cm.meta_path(book2)
        calls.clear()
        sidecar.write_bytes(sidecar.read_bytes() + b" ")     # still valid JSON
        check("a rewritten sidecar is re-read by the strict reader",
              cm.read_comparison_headline(book2).trusted and len(calls) == 1,
              f"strict reads: {len(calls)}")
        _settle()
        cm.read_comparison_headline(book2)
        calls.clear()
        sidecar.write_bytes(b"{not json")
        changed = cm.read_comparison_headline(book2)
        check("a corrupted sidecar is never hidden behind the entry",
              len(calls) == 1 and changed is not None and not changed.trusted,
              repr(changed))

        book3 = _publish(tmp / "cmp3", "third")
        _settle()
        good = cm.read_comparison_headline(book3)
        sentinel = cm.meta_path(book3)
        sentinel = sentinel.with_name(sentinel.name + ".tmp")
        calls.clear()
        sentinel.write_text("{not json", encoding="utf-8")
        with_sentinel = cm.read_comparison_headline(book3)
        check("a publication sentinel that APPEARS invalidates the entry",
              good.trusted and len(calls) == 1 and with_sentinel is not None
              and not with_sentinel.trusted, repr(with_sentinel))
        sentinel.unlink()
        calls.clear()
        check("...and removing it reads trusted again",
              cm.read_comparison_headline(book3).trusted and len(calls) == 1)

        book4 = _publish(tmp / "cmp4", "fourth")
        _settle()
        cm.read_comparison_headline(book4)
        calls.clear()
        cm.meta_path(book4).unlink()
        check("a removed sidecar reads as no record",
              cm.read_comparison_headline(book4) is None and len(calls) == 1)
    finally:
        cm.read_comparison_outcome = real_reader

    print("the matrix freshness reader uses it:")
    book5 = _publish(tmp / "cmp5", "fifth")
    record = cm.read_comparison_outcome(book5)
    stale = matrix_state._staleness(
        book5.stat().st_mtime, [], None, (), None, comparison_output=book5)
    check("a cell reports the published verdict and counts",
          stale["verdict"] == record.comparison_outcome.verdict
          and stale["diff_cells"] == 1 and stale["one_sided"] == 0
          and stale["completion"] == record.comparison_outcome.completion, repr(stale))


def persistence(tmp):
    print("the memos survive a restart, trusted no further than in memory:")
    if not WINDOWS:
        check("nothing is memoized off Windows, so nothing is persisted", True)
        return
    cache = tmp / "cache"
    book = _publish(tmp / "persist", "kept")
    export = tmp / "exports"
    export.mkdir()
    (export / "route_001.xlsx").write_bytes(b"R" * 4096)
    _settle()
    ct.MEMO.clear()
    cm._HEADLINE_MEMO.clear()
    cm.enable_persistent_caches(cache)
    head = cm.read_comparison_headline(book)
    fingerprint = a.fingerprint(export)
    cm.flush_persistent_caches(force=True)
    files = sorted(p.name for p in cache.iterdir())
    check("both memo files are written", files == ["comparison_headlines.json", "digests.json"],
          repr(files))

    ct.MEMO.clear()
    cm._HEADLINE_MEMO.clear()
    loaded = cm.enable_persistent_caches(cache)
    check("a new session loads them", loaded == 1 and len(ct.MEMO) >= 2,
          f"headlines {loaded}, digests {len(ct.MEMO)}")
    calls = []
    real_reader = cm.read_comparison_outcome
    cm.read_comparison_outcome = lambda p: (calls.append(p), real_reader(p))[1]
    try:
        check("the loaded headline is served without the strict reader",
              cm.read_comparison_headline(book) == head and not calls)
        check("the loaded digest serves the fingerprint", a.fingerprint(export) == fingerprint)
        _tamper_in_place(book)
        check("a file changed while the app was closed is re-read, not trusted",
              not cm.read_comparison_headline(book).trusted and len(calls) == 1)
        _tamper_in_place(export / "route_001.xlsx")
        check("...and so is a changed export", a.fingerprint(export) != fingerprint)
    finally:
        cm.read_comparison_outcome = real_reader

    ct.MEMO.clear()
    cm._HEADLINE_MEMO.clear()
    real_version = cm.producer_app_version
    cm.producer_app_version = lambda: "0.0.0-other"
    try:
        check("headlines written by another app version are not loaded",
              cm.enable_persistent_caches(cache) == 0)
    finally:
        cm.producer_app_version = real_version

    digests = cache / "digests.json"
    data = json.loads(digests.read_text(encoding="utf-8"))
    key = next(iter(data["entries"]))
    data["entries"][key][4] = 0                      # no change stamp
    data["entries"]["bogus"] = ["not", "an", "entry"]
    digests.write_text(json.dumps(data), encoding="utf-8")
    ct.MEMO.clear()
    ct.enable_persistence(digests)
    check("an entry without a change stamp, or a malformed one, is not loaded",
          key not in ct.MEMO and "bogus" not in ct.MEMO)
    digests.write_text("{broken", encoding="utf-8")
    ct.MEMO.clear()
    check("an unreadable memo file is an empty memo, not an error",
          ct.enable_persistence(digests) == 0)


def scope(tmp):
    print("one read of each folder per snapshot:")
    store = tmp / "store"
    store.mkdir()
    (store / "route_001.xlsx").write_bytes(b"A" * 1024)
    before = a.fingerprint(store)
    newest = a.newest_report_file_mtime(store)
    with a.read_scope():
        inside = a.fingerprint(store)
        a.newest_report_file_mtime(store)
        (store / "route_002.xlsx").write_bytes(b"B" * 1024)
        os.utime(store / "route_002.xlsx", (newest + 100, newest + 100))
        check("inside a scope the folder is read once",
              a.fingerprint(store) == inside == before
              and a.newest_report_file_mtime(store) == newest)
        with a.read_scope():
            check("a nested scope shares the outer one", a.fingerprint(store) == inside)
        check("...and leaving the nested scope does not end the outer one",
              a.fingerprint(store) == inside)
    check("outside a scope every call reads the disk",
          a.fingerprint(store) != before and a.newest_report_file_mtime(store) == newest + 100)
    check("the scope is gone after its block", getattr(a._READ_SCOPE, "cache", None) is None)


if __name__ == "__main__":
    print("read-side memos for the matrix snapshot:")
    root = Path(tempfile.mkdtemp(prefix="tsmis_snapshot_memo_"))
    tokens(root)
    bound_digest(root)
    headline(root)
    persistence(root)
    scope(root)
    if _failures:
        print(f"\n{len(_failures)} check(s) FAILED")
        raise SystemExit(1)
    print("all good")
