from __future__ import annotations

import os
import threading
from pathlib import Path

from convertor import watch


def test_file_must_be_stable_before_conversion(tmp_path: Path) -> None:
    folder_in, folder_out = tmp_path / "in", tmp_path / "in" / "Convertite"
    folder_in.mkdir()
    doc = folder_in / "a.docx"
    doc.write_bytes(b"x")
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
    folder_out.mkdir(parents=True)
    doc = folder_in / "a.docx"
    doc.write_bytes(b"x")
    (folder_out / "a.pdf").write_bytes(b"%PDF")
    state_path = tmp_path / "state.json"
    state = watch.WatchState(state_path)
    _, seen = watch.ready_files(folder_in, folder_out, state, {})
    assert list(seen) == [os.path.normcase(str(doc))]  # rezultatul nu e reconvertit

    state.mark(doc, watch._fingerprint(doc))
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


def test_folder_watcher_converts_new_files(tmp_path: Path) -> None:
    folder_in = tmp_path / "in"
    folder_in.mkdir()
    (folder_in / "a.docx").write_bytes(b"x")
    converted = threading.Event()
    received = []

    def convert(jobs):
        received.extend(jobs)
        converted.set()
        return []

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
