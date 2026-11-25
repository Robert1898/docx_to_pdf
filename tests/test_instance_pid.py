"""Identificarea instantei Office pornite de aplicatie: pid din fereastra, tasklist ca rezerva
(obiect COM simulat, fara Office)."""

from __future__ import annotations

import logging
from collections.abc import Callable
from types import SimpleNamespace

import pytest
import win32com.client
import win32gui

from convertor import office

WINDOW = 4242  # handle-ul ferestrei simulate


class FakeApp(office.OfficeApp):
    progid = "Test.Application"
    image = "TEST.EXE"
    name = "Test"

    def __init__(self, handle: int | Exception | None) -> None:
        super().__init__()
        self.handle = handle

    def _window_handle(self) -> int | None:
        if isinstance(self.handle, Exception):
            raise self.handle
        return self.handle


@pytest.fixture
def start(monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
          ) -> Callable[..., FakeApp]:
    """Porneste FakeApp cu listele tasklist si pid-ul ferestrei date."""

    def run(handle: int | Exception | None, before: set[int], after: set[int],
            window_pid: int | None = None) -> FakeApp:
        snapshots = iter([before, after])
        monkeypatch.setattr(office.base, "_pids", lambda image: next(snapshots))
        monkeypatch.setattr(office.base, "_pid_of_window",
                            lambda hwnd: window_pid if hwnd == WINDOW else None)
        monkeypatch.setattr(win32com.client, "DispatchEx", lambda progid: SimpleNamespace())
        app = FakeApp(handle)
        with caplog.at_level(logging.INFO, logger="convertor.office"):
            app.start()
        return app

    return run


def test_pid_from_window(start: Callable[..., FakeApp], caplog: pytest.LogCaptureFixture
                         ) -> None:
    app = start(WINDOW, before={1}, after={1, 50}, window_pid=50)
    assert (app.owned, app.pid) == (True, 50)
    assert "pid 50 din fereastra" in caplog.text
    assert "WARNING" not in caplog.text


def test_window_pid_wins_when_user_starts_the_same_app(start: Callable[..., FakeApp]) -> None:
    # utilizatorul a pornit aceeasi aplicatie in acelasi timp: doua procese noi
    app = start(WINDOW, before={1}, after={1, 50, 60}, window_pid=60)
    assert (app.owned, app.pid) == (True, 60)


@pytest.mark.parametrize("handle", [None, OSError("Member not found")])
def test_no_window_falls_back_to_tasklist(start: Callable[..., FakeApp],
                                         caplog: pytest.LogCaptureFixture,
                                         handle: Exception | None) -> None:
    app = start(handle, before={1}, after={1, 50})
    assert (app.owned, app.pid) == (True, 50)
    assert "pid 50 din tasklist" in caplog.text


def test_no_new_process_is_not_ours(start: Callable[..., FakeApp]) -> None:
    app = start(None, before={1}, after={1})  # PowerPoint deja deschis de utilizator
    assert (app.owned, app.pid) == (False, None)
    assert app.kill() is False


def test_two_new_processes_without_window_are_not_killed(start: Callable[..., FakeApp],
                                                        caplog: pytest.LogCaptureFixture,
                                                        monkeypatch: pytest.MonkeyPatch) -> None:
    app = start(None, before={1}, after={1, 50, 60})
    assert (app.owned, app.pid) == (True, None)  # Quit normal, dar fara oprire fortata
    killed: list[int] = []
    monkeypatch.setattr(office.base, "_kill", killed.append)
    assert app.kill() is False
    assert killed == []
    assert "2 procese noi [50, 60]" in caplog.text


def test_different_pids_warn_and_use_the_window(start: Callable[..., FakeApp],
                                                caplog: pytest.LogCaptureFixture) -> None:
    app = start(WINDOW, before={1}, after={1, 50}, window_pid=70)
    assert (app.owned, app.pid) == (True, 70)
    assert "pid-ul din fereastra (70) difera de cel din tasklist (50)" in caplog.text


def test_window_of_an_existing_instance_is_not_ours(start: Callable[..., FakeApp],
                                                    caplog: pytest.LogCaptureFixture) -> None:
    app = start(WINDOW, before={1, 30}, after={1, 30}, window_pid=30)
    assert app.owned is False
    assert app.kill() is False
    assert "instante deja pornite (pid 30)" in caplog.text


def test_excel_window_handle() -> None:
    app = office.ExcelApp()
    app.app = SimpleNamespace(Hwnd=197576)
    assert app._window_handle() == 197576


class FakeWord:
    Caption = "Word"


@pytest.mark.parametrize("found", [459866, 0])
def test_word_window_found_by_temporary_caption(monkeypatch: pytest.MonkeyPatch,
                                                found: int) -> None:
    app = office.WordApp()
    app.app = FakeWord()
    searched: list[tuple[str, str]] = []

    def find(cls: str, title: str) -> int:
        searched.append((cls, title))
        return found

    monkeypatch.setattr(win32gui, "FindWindow", find)
    assert app._window_handle() == (found or None)
    assert searched[0][0] == "OpusApp"
    assert searched[0][1].startswith("convertor-")  # titlu unic, nu „Word”
    assert app.app.Caption == "Word"  # refacut


def test_word_caption_restored_when_search_fails(monkeypatch: pytest.MonkeyPatch) -> None:
    app = office.WordApp()
    app.app = FakeWord()

    def fail(cls: str, title: str) -> int:
        raise win32gui.error(2, "FindWindow", "nu exista")

    monkeypatch.setattr(win32gui, "FindWindow", fail)
    with pytest.raises(win32gui.error):
        app._window_handle()
    assert app.app.Caption == "Word"
