"""PowerPoint are o singura instanta: prezentarile deschise de utilizator ajung in procesul
nostru si nu au voie sa fie inchise (obiect COM simulat, fara Office)."""

from __future__ import annotations

import logging
import threading
from pathlib import Path
from types import SimpleNamespace

import pytest

from convertor import office
from convertor.core import ExportOptions, Job, Kind


class FakePresentation:
    def __init__(self, ppt: FakePowerPoint) -> None:
        self.ppt = ppt

    def ExportAsFixedFormat(self, *args: object) -> None:
        self.ppt.calls.append("export")
        if self.ppt.block is not None:  # blocat (de ex. dialog invizibil) pana la anulare
            assert self.ppt.block.wait(5), "apelul nu a fost anulat"
            raise OSError("Call was canceled by the message filter.")

    def SaveAs(self, *args: object) -> None:
        self.ppt.calls.append("saveas")

    def Close(self) -> None:
        self.ppt.calls.append("close")
        self.ppt.count -= 1


class FakePowerPoint:
    """PowerPoint.Application simulat. `count` = toate prezentarile din proces."""

    def __init__(self, foreign: int = 0, count_fails: bool = False) -> None:
        self.count = foreign
        self.count_fails = count_fails
        self.calls: list[str] = []
        self.block: threading.Event | None = None

    @property
    def Presentations(self) -> SimpleNamespace:
        if self.count_fails:
            raise OSError("RPC server unavailable")
        return SimpleNamespace(Count=self.count, Open=self._open)

    def _open(self, **kwargs: object) -> FakePresentation:
        self.calls.append("open")
        self.count += 1
        return FakePresentation(self)

    def Quit(self) -> None:
        self.calls.append("quit")


@pytest.fixture
def env(monkeypatch: pytest.MonkeyPatch) -> SimpleNamespace:
    """Fara ferestre ale utilizatorului si fara procese reale; inregistreaza opririle."""
    state = SimpleNamespace(windows=[], killed=[])
    monkeypatch.setattr(office.PowerPointApp, "open_presentations", 0)
    monkeypatch.setattr(office.powerpoint, "_visible_windows",
                        lambda pid, classes: list(state.windows))
    monkeypatch.setattr(office.base, "_kill", state.killed.append)
    return state


def powerpoint(fake: FakePowerPoint, owned: bool = True) -> office.PowerPointApp:
    app = office.PowerPointApp()
    app.app, app.owned, app.pid = fake, owned, 4321 if owned else None
    app.configure()
    return app


def test_foreign_presentation_keeps_powerpoint_at_quit(env: SimpleNamespace,
                                                       caplog: pytest.LogCaptureFixture) -> None:
    fake = FakePowerPoint()
    app = powerpoint(fake)
    fake.count = 1  # utilizatorul a deschis o prezentare din Explorer, in procesul nostru
    with caplog.at_level(logging.WARNING, logger="convertor.office"):
        app.quit()
    assert "quit" not in fake.calls
    assert env.killed == []
    assert "prezentari deschise de altcineva (1" in caplog.text
    assert app.app is None


def test_quit_without_foreign_presentations(env: SimpleNamespace) -> None:
    fake = FakePowerPoint()
    app = powerpoint(fake)
    app.convert(Job(Path("a.pptx"), Path("a.pdf"), Kind.POWERPOINT), ExportOptions(),
                Path("a.pptx"))
    app.quit()
    assert fake.calls == ["open", "export", "close", "quit"]


def test_own_presentation_left_open_is_not_foreign(env: SimpleNamespace) -> None:
    fake = FakePowerPoint()
    app = powerpoint(fake)
    fake.count, office.PowerPointApp.open_presentations = 1, 1  # Close a esuat
    app.quit()
    assert fake.calls == ["quit"]


def test_unreadable_count_quits_unless_a_user_window_is_visible(env: SimpleNamespace) -> None:
    fake = FakePowerPoint()
    app = powerpoint(fake)
    fake.count_fails = True
    app.quit()
    assert fake.calls == ["quit"]

    fake = FakePowerPoint()
    app = powerpoint(fake)
    fake.count_fails = True
    env.windows = [7]  # fereastra PPTFrameClass a prezentarii utilizatorului
    app.quit()
    assert fake.calls == []


@pytest.mark.parametrize("signal", ["window", "count"])
def test_no_forced_kill_with_foreign_presentations(env: SimpleNamespace, signal: str) -> None:
    fake = FakePowerPoint(foreign=1 if signal == "count" else 0)
    app = powerpoint(fake)
    if signal == "window":
        env.windows = [7]
    else:  # numarat de firul nostru inainte de deschidere
        app.convert(Job(Path("a.pptx"), Path("a.pdf"), Kind.POWERPOINT), ExportOptions(),
                    Path("a.pptx"))
    assert app.kill() is False
    assert env.killed == []
    assert app.timed_out is False


def test_kill_without_foreign_presentations(env: SimpleNamespace) -> None:
    app = powerpoint(FakePowerPoint())
    assert app.kill() is True
    assert env.killed == [4321]


def session_with(app: office.PowerPointApp) -> office.OfficeSession:
    session = office.OfficeSession()
    session._apps[office.PowerPointApp] = app
    session._thread_id = 1  # fara CoInitialize: anularea e simulata
    return session


def test_timeout_with_foreign_presentation_fails_the_file_but_keeps_powerpoint(
        env: SimpleNamespace, monkeypatch: pytest.MonkeyPatch,
        caplog: pytest.LogCaptureFixture) -> None:
    fake = FakePowerPoint()
    fake.block = threading.Event()
    app = powerpoint(fake)
    env.windows = [7]
    cancelled: list[int] = []

    def cancel(thread_id: int) -> bool:
        cancelled.append(thread_id)
        fake.block.set()
        return True

    monkeypatch.setattr(office.session, "_cancel_call", cancel)
    session = session_with(app)
    job = Job(Path("a.pptx"), Path("a.pdf"), Kind.POWERPOINT)
    with caplog.at_level(logging.WARNING, logger="convertor.office"), \
            pytest.raises(office.ConversionTimeout, match="documente deschise de utilizator"):
        session.convert(job, ExportOptions(timeout=0.2))
    assert cancelled == [1]
    assert env.killed == []  # procesul ramane
    assert "saveas" not in fake.calls  # fara a doua incercare in PowerPoint-ul blocat
    assert "close" not in fake.calls
    assert app.app is None  # instanta abandonata nu mai e folosita si nu primeste Quit
    session.on_idle()
    assert "quit" not in fake.calls
    assert office.powerpoint.POWERPOINT_LOCK.acquire(blocking=False)
    office.powerpoint.POWERPOINT_LOCK.release()
    assert "nu il opresc, renunt doar la fisier" in caplog.text


def test_on_idle_closes_only_powerpoint(env: SimpleNamespace) -> None:
    fake = FakePowerPoint()
    app = powerpoint(fake)
    session = session_with(app)
    word = SimpleNamespace(quit=lambda: pytest.fail("Word ramane pentru lotul urmator"))
    session._apps[office.WordApp] = word  # type: ignore[assignment]
    session.on_idle()
    assert fake.calls == ["quit"]
    assert session._apps == {office.WordApp: word}
    session.on_idle()  # nimic de inchis
    assert fake.calls == ["quit"]


def test_shared_powerpoint_reference_is_dropped_after_each_file(env: SimpleNamespace) -> None:
    fake = FakePowerPoint()
    app = powerpoint(fake, owned=False)  # PowerPoint pornit de altcineva (alt fir, utilizator)
    session = session_with(app)
    session.convert(Job(Path("a.pptx"), Path("a.pdf"), Kind.POWERPOINT), ExportOptions())
    assert app.app is None
    assert "quit" not in fake.calls


def test_abort_with_foreign_presentation_cancels_the_blocked_call(
        env: SimpleNamespace, monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakePowerPoint()
    fake.block = threading.Event()
    app = powerpoint(fake)
    env.windows = [7]
    cancelled: list[int] = []

    def cancel(thread_id: int) -> bool:
        cancelled.append(thread_id)
        fake.block.set()
        return True

    monkeypatch.setattr(office.session, "_cancel_call", cancel)
    session = session_with(app)
    errors: list[BaseException] = []

    def run() -> None:
        try:
            session.convert(Job(Path("a.pptx"), Path("a.pdf"), Kind.POWERPOINT), ExportOptions())
        except BaseException as exc:
            errors.append(exc)

    worker = threading.Thread(target=run)
    worker.start()
    deadline = threading.Event()
    while "export" not in fake.calls and not deadline.wait(0.01):
        pass
    session.kill_all()  # Anulare / inchiderea aplicatiei, din alt fir
    worker.join(5)
    assert not worker.is_alive()  # firul nu mai e blocat
    assert cancelled
    assert env.killed == []  # PowerPoint cu prezentarea utilizatorului ramane
    assert [type(e) for e in errors] == [office.ConversionTimeout]
    assert "close" not in fake.calls and "saveas" not in fake.calls
    assert office.powerpoint.POWERPOINT_LOCK.acquire(blocking=False)
    office.powerpoint.POWERPOINT_LOCK.release()


def test_abort_without_foreign_presentation_still_kills(env: SimpleNamespace,
                                                         monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(office.session, "_cancel_call",
                        lambda tid: pytest.fail("nu se anuleaza cand procesul poate fi oprit"))
    fake = FakePowerPoint()
    app = powerpoint(fake)
    session = session_with(app)
    session._current = (app, Path("a.pptx"), threading.Event())
    session.kill_all()
    assert env.killed == [4321]


def test_abandoned_cleanup_has_a_deadline(monkeypatch: pytest.MonkeyPatch) -> None:
    """Ultimele apeluri catre o instanta abandonata pot bloca si ele: sunt anulate la limita."""
    unblock = threading.Event()
    monkeypatch.setattr(office.session, "ABANDON_CLEANUP_SECONDS", 0.1)
    monkeypatch.setattr(office.session, "_cancel_call", lambda tid: unblock.set() or True)
    app = office.PowerPointApp()
    app.app, app.abandoned = object(), True
    monkeypatch.setattr(app, "release_abandoned", lambda: unblock.wait(5))
    session = office.OfficeSession()
    session._thread_id = 1
    session._drop_abandoned(app, threading.Event())
    assert unblock.is_set()
    assert app.app is None
