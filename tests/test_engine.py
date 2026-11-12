from __future__ import annotations

import threading
import time
from pathlib import Path
from typing import ClassVar

import pytest
from PIL import Image
from reportlab.pdfgen import canvas

from convertor import engine
from convertor.core import ExportOptions, plan_jobs


class FakeSession:
    instances: ClassVar[list[FakeSession]] = []

    def __init__(self) -> None:
        self.closed = False
        self.killed = False
        FakeSession.instances.append(self)

    def __enter__(self) -> FakeSession:
        return self

    def __exit__(self, *exc: object) -> None:
        self.closed = True

    def kill_all(self) -> None:
        self.killed = True


@pytest.fixture
def fake_office(monkeypatch: pytest.MonkeyPatch) -> type[FakeSession]:
    FakeSession.instances = []
    monkeypatch.setattr(engine, "OfficeSession", FakeSession)
    return FakeSession


def test_pool_runs_tasks_and_reuses_sessions(fake_office: type[FakeSession]) -> None:
    pool = engine.OfficePool(2, idle_timeout=5)
    try:
        assert [pool.submit(lambda s, i=i: i * 2).result(5) for i in range(3)] == [0, 2, 4]
        assert len(fake_office.instances) == 1  # lot mic: o singura instanta Office
    finally:
        pool.close()
    assert all(s.closed for s in fake_office.instances)
    with pytest.raises(RuntimeError):
        pool.submit(lambda s: None)


def test_pool_grows_for_large_batches(fake_office: type[FakeSession]) -> None:
    pool = engine.OfficePool(3, idle_timeout=5)
    gate = threading.Event()
    try:
        futures = [pool.submit(lambda s: (gate.wait(5), threading.current_thread().name))
                   for _ in range(engine.JOBS_PER_WORKER * 3)]
        time.sleep(0.2)
        assert len(fake_office.instances) == 3
        gate.set()
        for future in futures:
            future.result(5)
    finally:
        pool.close()


def test_pool_propagates_errors(fake_office: type[FakeSession]) -> None:
    pool = engine.OfficePool(1, idle_timeout=5)
    try:
        def boom(session: FakeSession) -> None:
            raise ValueError("stricat")

        with pytest.raises(ValueError, match="stricat"):
            pool.submit(boom).result(5)
        assert pool.submit(lambda s: "inca merge").result(5) == "inca merge"
    finally:
        pool.close()


def test_idle_workers_close_office(fake_office: type[FakeSession]) -> None:
    pool = engine.OfficePool(1, idle_timeout=0.1)
    pool.submit(lambda s: None).result(5)
    deadline = time.time() + 5
    while not fake_office.instances[0].closed and time.time() < deadline:
        time.sleep(0.02)
    assert fake_office.instances[0].closed
    pool.close()


def test_engine_converts_images_and_pdf_tables_without_office(
        tmp_path: Path, fake_office: type[FakeSession]) -> None:
    img = tmp_path / "poza.png"
    Image.new("RGB", (300, 200), "green").save(img)
    pdf = tmp_path / "raport.pdf"
    c = canvas.Canvas(str(pdf))
    for i in range(5):
        c.drawString(60, 700 - i * 15, f"Rand {i}")
        c.drawRightString(400, 700 - i * 15, f"{i * 10:.2f}")
    c.save()
    jobs = plan_jobs([img, pdf], pdf_target="xlsx")
    seen = []
    results = engine.Engine(ExportOptions(pdf_target="xlsx"), listener=seen.append,
                            backends=engine.Backends(office={}, soffice=None,
                                                     tesseract=None)).run(jobs)
    assert [r.status for r in results] == [engine.Status.DONE, engine.Status.DONE]
    assert (tmp_path / "poza.pdf").exists()
    assert (tmp_path / "raport.xlsx").exists()
    assert any(r.status is engine.Status.RUNNING for r in seen)
