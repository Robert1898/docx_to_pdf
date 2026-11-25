from __future__ import annotations

import json
import logging
import os
import threading
from pathlib import Path

import pytest

from convertor import watch
from convertor.engine import Result, Status

from .helpers import wait_until


def _doc(folder: Path, name: str = "a.docx", data: bytes = b"x") -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / name
    path.write_bytes(data)
    return path


def _fp(path: Path) -> watch.Fingerprint:
    fingerprint = watch._fingerprint(path)
    assert fingerprint is not None
    return fingerprint


def _record(state: watch.WatchState, path: Path, status: Status | None, message: str = "",
            now: float = 0.0, error: str = "") -> list[tuple[Path, str]]:
    """Ruleaza record_results pentru un singur fisier (status None = lotul a crapat)."""
    jobs = watch.plan_watch_jobs([path], path.parent / "Convertite", "docx")
    results = None if status is None else [Result(jobs[0], status, message)]
    return watch.record_results(state, jobs, results, {watch._key(path): _fp(path)},
                                error=error, now=now)


def test_file_must_be_stable_before_conversion(tmp_path: Path) -> None:
    folder_in, folder_out = tmp_path / "in", tmp_path / "in" / "Convertite"
    doc = _doc(folder_in)
    state = watch.WatchState(tmp_path / "state.json")

    ready, seen = watch.ready_files(folder_in, folder_out, state, {})
    assert ready == []  # vazut prima data: poate inca se copiaza
    ready, seen = watch.ready_files(folder_in, folder_out, state, seen)
    assert ready == [doc]

    doc.write_bytes(b"xy")  # inca se scrie
    ready, seen = watch.ready_files(folder_in, folder_out, state, seen)
    assert ready == []


def test_done_files_and_output_folder_are_ignored(tmp_path: Path) -> None:
    folder_in = tmp_path / "in"
    folder_out = watch.default_output(folder_in)
    doc = _doc(folder_in)
    _doc(folder_out, "a.pdf", b"%PDF")
    state_path = tmp_path / "state.json"
    state = watch.WatchState(state_path)
    _, seen = watch.ready_files(folder_in, folder_out, state, {})
    assert list(seen) == [os.path.normcase(str(doc))]  # rezultatul nu e reconvertit

    state.mark_done(doc, _fp(doc))
    state.save()
    reloaded = watch.WatchState(state_path)
    ready, _ = watch.ready_files(folder_in, folder_out, reloaded, seen)
    assert ready == []


def test_plan_watch_jobs_never_overwrites(tmp_path: Path) -> None:
    out = tmp_path / "out"
    out.mkdir()
    (out / "a.pdf").write_bytes(b"old")
    src = tmp_path / "a.docx"
    src.write_bytes(b"x")
    pdf = tmp_path / "r.pdf"
    pdf.write_bytes(b"%PDF")
    jobs = watch.plan_watch_jobs([src, pdf], out, "xlsx")
    assert jobs[0].dst.name == "a (2).pdf"
    assert jobs[1].dst.suffix == ".xlsx"


def test_only_successful_files_are_marked_done(tmp_path: Path) -> None:
    good, bad = _doc(tmp_path, "good.docx"), _doc(tmp_path, "bad.docx")
    state = watch.WatchState(tmp_path / "state.json")
    jobs = watch.plan_watch_jobs([good, bad], tmp_path / "out", "docx")
    results = [Result(jobs[0], Status.DONE), Result(jobs[1], Status.FAILED, "parola")]
    fingerprints = {watch._key(p): _fp(p) for p in (good, bad)}
    assert watch.record_results(state, jobs, results, fingerprints, now=0) == []
    assert state.status(good, _fp(good)) == watch.DONE
    assert state.status(bad, _fp(bad)) == watch.RETRY
    assert state.entries[watch._key(bad)].error == "parola"


def test_failed_file_is_retried_then_given_up(tmp_path: Path,
                                              caplog: pytest.LogCaptureFixture) -> None:
    doc = _doc(tmp_path)
    state = watch.WatchState(tmp_path / "state.json", retry_delay=10)
    for attempt in range(1, watch.MAX_RETRIES + 1):
        assert _record(state, doc, Status.FAILED, "eroare", now=attempt * 100) == []
        assert state.status(doc, _fp(doc)) == watch.RETRY
        assert state.entries[watch._key(doc)].attempts == attempt
        assert state.should_convert(doc, _fp(doc), now=attempt * 100 + 10 * attempt)

    with caplog.at_level(logging.WARNING, logger="convertor.watch"):
        gave_up = _record(state, doc, Status.FAILED, "eroare finala", now=1000)
    assert gave_up == [(doc, "eroare finala")]
    assert state.status(doc, _fp(doc)) == watch.FAILED
    assert not state.should_convert(doc, _fp(doc), now=10**9)  # nu mai e reincercat
    assert "renunt" in caplog.text
    assert "eroare finala" in caplog.text


def test_retry_waits_longer_after_each_attempt(tmp_path: Path) -> None:
    doc = _doc(tmp_path)
    state = watch.WatchState(tmp_path / "state.json", retry_delay=60)
    _record(state, doc, Status.FAILED, now=1000)
    assert not state.should_convert(doc, _fp(doc), now=1059)
    assert state.should_convert(doc, _fp(doc), now=1060)
    _record(state, doc, Status.FAILED, now=2000)
    assert not state.should_convert(doc, _fp(doc), now=2119)  # a doua: 2 x 60 s
    assert state.should_convert(doc, _fp(doc), now=2120)


def test_cancelled_does_not_count_as_attempt(tmp_path: Path) -> None:
    doc = _doc(tmp_path)
    state = watch.WatchState(tmp_path / "state.json")
    assert _record(state, doc, Status.CANCELLED) == []
    assert state.status(doc, _fp(doc)) is None
    assert state.should_convert(doc, _fp(doc))


def test_crashed_batch_counts_as_failure(tmp_path: Path) -> None:
    doc = _doc(tmp_path)
    state = watch.WatchState(tmp_path / "state.json")
    _record(state, doc, None, error="Office a murit")
    entry = state.entries[watch._key(doc)]
    assert (entry.status, entry.attempts, entry.error) == (watch.RETRY, 1, "Office a murit")


def test_changed_file_starts_over(tmp_path: Path) -> None:
    doc = _doc(tmp_path)
    state = watch.WatchState(tmp_path / "state.json")
    for _ in range(watch.MAX_RETRIES + 1):
        _record(state, doc, Status.FAILED)
    assert state.status(doc, _fp(doc)) == watch.FAILED
    doc.write_bytes(b"versiune noua")  # utilizatorul a pus alt fisier cu acelasi nume
    assert state.should_convert(doc, _fp(doc))
    _record(state, doc, Status.FAILED)
    assert state.entries[watch._key(doc)].attempts == 1


def test_old_state_format_is_migrated(tmp_path: Path) -> None:
    doc = _doc(tmp_path)
    other = _doc(tmp_path, "b.docx")
    state_path = tmp_path / "watch_done.json"
    old = {watch._key(doc): list(_fp(doc)), watch._key(other): "stricat"}
    state_path.write_text(json.dumps(old), encoding="utf-8")

    state = watch.WatchState(state_path)
    assert state.is_done(doc, _fp(doc))
    assert not state.should_convert(doc, _fp(doc))
    assert state.should_convert(other, _fp(other))  # intrarea ilizibila e ignorata

    state.save()
    data = json.loads(state_path.read_text(encoding="utf-8"))
    assert data["version"] == watch.STATE_VERSION
    assert data["files"][watch._key(doc)]["status"] == watch.DONE
    assert watch.WatchState(state_path).is_done(doc, _fp(doc))


def test_unreadable_state_starts_empty(tmp_path: Path) -> None:
    state_path = tmp_path / "watch_done.json"
    state_path.write_text("{nu e json", encoding="utf-8")
    assert watch.WatchState(state_path).entries == {}
    state_path.write_text(json.dumps({"version": 2, "files": {"x": {"fp": [1, 2],
                                                                     "status": "ciudat"}}}))
    assert watch.WatchState(state_path).entries == {}


def test_folder_watcher_converts_new_files(tmp_path: Path) -> None:
    folder_in = tmp_path / "in"
    _doc(folder_in)
    converted = threading.Event()
    received = []

    def convert(jobs):
        received.extend(jobs)
        converted.set()
        return [Result(job, Status.DONE) for job in jobs]

    watcher = watch.FolderWatcher(folder_in, watch.default_output(folder_in), convert,
                                  interval=0.05)
    watcher.state = watch.WatchState(tmp_path / "state.json")
    watcher.start()
    try:
        assert converted.wait(5)
    finally:
        watcher.stop()
        watcher.join(2)
    assert [j.src.name for j in received] == ["a.docx"]
    assert received[0].dst.parent == watch.default_output(folder_in)


def test_folder_watcher_retries_and_gives_up(tmp_path: Path,
                                             monkeypatch: pytest.MonkeyPatch) -> None:
    folder_in = tmp_path / "in"
    doc = _doc(folder_in)
    calls: list[str] = []
    gave_up: list[tuple[Path, str]] = []
    finished = threading.Event()

    def convert(jobs):
        calls.append(jobs[0].src.name)
        return [Result(job, Status.FAILED, "nu merge") for job in jobs]

    def on_give_up(path: Path, error: str) -> None:
        gave_up.append((path, error))
        finished.set()

    scans: list[int] = []
    real_ready_files = watch.ready_files

    def counted_ready_files(*args, **kwargs):
        scans.append(1)
        return real_ready_files(*args, **kwargs)

    monkeypatch.setattr(watch, "ready_files", counted_ready_files)
    watcher = watch.FolderWatcher(folder_in, watch.default_output(folder_in), convert,
                                  interval=0.02, on_give_up=on_give_up, retry_delay=0)
    watcher.state = watch.WatchState(tmp_path / "state.json", retry_delay=0)
    watcher.start()
    try:
        assert finished.wait(5)
        # cateva verificari ale folderului in plus (numarate, nu un timp fix): nu mai e reluat
        after_give_up = len(scans)
        assert wait_until(lambda: len(scans) >= after_give_up + 5)
    finally:
        watcher.stop()
        watcher.join(2)
    assert calls == ["a.docx"] * (watch.MAX_RETRIES + 1)
    assert gave_up == [(doc, "nu merge")]
    saved = json.loads((tmp_path / "state.json").read_text(encoding="utf-8"))
    assert saved["files"][watch._key(doc)]["status"] == watch.FAILED
    assert saved["files"][watch._key(doc)]["attempts"] == watch.MAX_RETRIES + 1
