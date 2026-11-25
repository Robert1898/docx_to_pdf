"""Inchiderea instantei Excel: registrele straine pastrate, fara oprire fortata peste ele
(obiect COM simulat, fara Office)."""

from __future__ import annotations

import logging
import threading
from pathlib import Path
from types import SimpleNamespace

import pytest

from convertor import office
from convertor.core import ExportOptions, Job, Kind


class FakeWorkbook:
    def __init__(self, excel: FakeExcel, close_fails: bool = False) -> None:
        self.excel = excel
        self.close_fails = close_fails

    def ExportAsFixedFormat(self, **kwargs) -> None:
        self.excel.seen_open.append(self.excel.Workbooks.Count)

    def Close(self, SaveChanges: bool) -> None:
        if self.close_fails:
            raise OSError("Close a esuat")
        self.excel.Workbooks.Count -= 1


class FakeWorkbooks(SimpleNamespace):
    def Open(self, **kwargs) -> FakeWorkbook:
        self.Count += 1
        return FakeWorkbook(self.excel, self.close_fails)


class FakeExcel:
    """Excel.Application simulat: inregistreaza Quit-urile."""

    def __init__(self, workbooks: int = 0, count_fails: bool = False,
                 close_fails: bool = False) -> None:
        self.Visible = self.UserControl = False
        self.Workbooks = FakeWorkbooks(Count=workbooks, excel=self, close_fails=close_fails)
        self.count_fails = count_fails
        self.quit_calls: list[str] = []
        self.seen_open: list[int] = []

    def __getattribute__(self, name: str):
        if name == "Workbooks" and object.__getattribute__(self, "count_fails"):
            raise OSError("RPC server unavailable")
        return object.__getattribute__(self, name)

    def __setattr__(self, name: str, value: object) -> None:
        # setare globala, salvata de Excel la iesire: aplicatia nu are voie sa o schimbe
        assert name != "IgnoreRemoteRequests", "IgnoreRemoteRequests nu se atinge"
        super().__setattr__(name, value)

    def Quit(self) -> None:
        self.quit_calls.append("quit")


@pytest.fixture
def process(monkeypatch: pytest.MonkeyPatch) -> list[str]:
    """Inregistreaza deschiderea handle-ului si asteptarea/oprirea procesului."""
    calls: list[str] = []
    monkeypatch.setattr(office.base, "_open_process", lambda pid: calls.append(f"open {pid}") or 99)

    def wait(handle: int, seconds: float) -> bool:
        calls.append(f"wait {handle}")
        return False

    monkeypatch.setattr(office.base, "_wait_or_terminate", wait)
    return calls


def excel_app(fake: FakeExcel, owned: bool = True) -> office.ExcelApp:
    app = office.ExcelApp()
    app.app, app.owned, app.pid = fake, owned, 1234 if owned else None
    app.configure()
    return app


def test_our_instance_is_quit_and_awaited(process: list[str]) -> None:
    fake = FakeExcel()
    app = excel_app(fake)
    app.quit()
    assert fake.quit_calls == ["quit"]
    assert process == ["open 1234", "wait 99"]
    assert app.app is None


def test_instance_not_started_by_us_is_left_alone(process: list[str]) -> None:
    fake = FakeExcel()
    app = excel_app(fake, owned=False)
    app.quit()
    assert fake.quit_calls == []
    assert process == []


def test_foreign_workbook_keeps_excel_alive(process: list[str],
                                            caplog: pytest.LogCaptureFixture) -> None:
    fake = FakeExcel()
    app = excel_app(fake)
    fake.Workbooks.Count = 1  # registru deschis de utilizator in instanta noastra
    with caplog.at_level(logging.WARNING, logger="convertor.office"):
        app.quit()
    assert fake.quit_calls == []  # Quit cu DisplayAlerts = False l-ar inchide nesalvat
    assert process == []  # nici oprit fortat
    assert (fake.Visible, fake.UserControl, fake.DisplayAlerts) == (True, True, True)
    assert "1 registru(e) deschis(e) de altcineva" in caplog.text
    assert app.app is None


class StubbornExcel(FakeExcel):
    """Excel care refuza unele setari (arunca la atribuire), de ex. cand e ocupat."""

    refuse: frozenset[str] = frozenset()

    def __setattr__(self, name: str, value: object) -> None:
        if name in self.refuse:
            raise OSError(f"{name} refuzat")
        super().__setattr__(name, value)


HANDOVER = ("DisplayAlerts", "ScreenUpdating", "EnableEvents", "Visible", "UserControl")


@pytest.mark.parametrize("refused", HANDOVER)
def test_refused_setting_does_not_keep_foreign_workbook_hidden(process: list[str],
                                                              refused: str) -> None:
    fake = StubbornExcel()
    app = excel_app(fake)
    fake.Workbooks.Count = 1
    fake.refuse = frozenset({refused})
    app.quit()
    assert fake.quit_calls == []
    assert process == []
    assert all(getattr(fake, name) is True for name in HANDOVER if name != refused)
    assert getattr(fake, refused) is False  # ramasa cum a pus-o configure


def test_own_open_workbook_is_not_foreign(process: list[str]) -> None:
    fake = FakeExcel()
    app = excel_app(fake)
    fake.Workbooks.Count = 1
    app.open_workbooks = 1  # Close a esuat: registrul e tot al nostru
    app.quit()
    assert fake.quit_calls == ["quit"]
    assert process == ["open 1234", "wait 99"]


def test_unreadable_workbook_count_still_quits(process: list[str]) -> None:
    fake = FakeExcel()
    app = excel_app(fake)
    fake.count_fails = True
    app.quit()
    assert fake.quit_calls == ["quit"]


def test_timed_out_instance_is_not_quit(process: list[str]) -> None:
    fake = FakeExcel()
    app = excel_app(fake)
    app.timed_out = True  # deja oprita fortat: Excel nu a apucat sa salveze setarea
    app.quit()
    assert fake.quit_calls == []
    assert process == []


def test_convert_counts_its_open_workbooks(monkeypatch: pytest.MonkeyPatch,
                                           tmp_path: Path) -> None:
    monkeypatch.setattr(office.ExcelApp, "_limit_to_content", lambda self, wb: None)
    fake = FakeExcel()
    app = excel_app(fake)
    job = Job(tmp_path / "a.xlsx", tmp_path / "a.pdf", Kind.EXCEL)
    app.convert(job, ExportOptions(), job.src)
    assert fake.seen_open == [1]
    assert app.open_workbooks == 0

    fake.Workbooks.close_fails = True
    with pytest.raises(OSError):
        app.convert(job, ExportOptions(), job.src)
    assert app.open_workbooks == 1  # inca deschis: nu e numarat ca strain


@pytest.fixture
def blocked(monkeypatch: pytest.MonkeyPatch, process: list[str]) -> SimpleNamespace:
    """Excel simulat al carui export ramane blocat pana la anulare; ferestre si opriri false."""
    state = SimpleNamespace(windows=[], killed=[], cancelled=[], release=threading.Event())
    monkeypatch.setattr(office.excel, "_visible_windows", lambda pid, classes: list(state.windows))
    monkeypatch.setattr(office.base, "_kill", state.killed.append)
    monkeypatch.setattr(office.ExcelApp, "_limit_to_content", lambda self, wb: None)

    def cancel(thread_id: int) -> bool:
        state.cancelled.append(thread_id)
        state.release.set()
        return True

    monkeypatch.setattr(office.session, "_cancel_call", cancel)
    return state


class BlockingWorkbook(FakeWorkbook):
    def ExportAsFixedFormat(self, **kwargs) -> None:
        assert self.excel.release.wait(5), "apelul nu a fost anulat"
        raise OSError("Call was canceled by the message filter.")

    def Close(self, SaveChanges: bool) -> None:
        self.excel.closed = True
        super().Close(SaveChanges)


class BlockingExcel(FakeExcel):
    def __init__(self, release: threading.Event, **kwargs) -> None:
        super().__init__(**kwargs)
        self.release, self.closed = release, False
        self.Workbooks.Open = lambda **kw: self._open()

    def _open(self) -> BlockingWorkbook:
        self.Workbooks.Count += 1
        return BlockingWorkbook(self)


def excel_session(app: office.ExcelApp) -> office.OfficeSession:
    session = office.OfficeSession()
    session._apps[office.ExcelApp] = app
    session._thread_id = 1
    return session


def test_timeout_with_foreign_workbook_fails_the_file_but_keeps_excel(
        blocked: SimpleNamespace, caplog: pytest.LogCaptureFixture) -> None:
    fake = BlockingExcel(blocked.release)
    app = excel_app(fake)
    blocked.windows = [9]  # fereastra XLMAIN vizibila: altcineva lucreaza in instanta
    session = excel_session(app)
    job = Job(Path("a.xlsx"), Path("a.pdf"), Kind.EXCEL)
    with caplog.at_level(logging.WARNING, logger="convertor.office"), \
            pytest.raises(office.ConversionTimeout, match="documente deschise de utilizator"):
        session.convert(job, ExportOptions(timeout=0.2))
    assert blocked.cancelled == [1]
    assert blocked.killed == []
    assert fake.closed is False  # fara Close in Excel-ul blocat
    assert fake.quit_calls == []
    assert app.app is None
    assert "are documente deschise de altcineva" in caplog.text


def test_abort_with_foreign_workbook_cancels_the_blocked_call(blocked: SimpleNamespace) -> None:
    fake = BlockingExcel(blocked.release, workbooks=1)  # registru strain deja deschis
    app = excel_app(fake)
    session = excel_session(app)
    errors: list[BaseException] = []

    def run() -> None:
        try:
            session.convert(Job(Path("a.xlsx"), Path("a.pdf"), Kind.EXCEL), ExportOptions())
        except BaseException as exc:
            errors.append(exc)

    worker = threading.Thread(target=run)
    worker.start()
    while session._current is None and worker.is_alive():
        threading.Event().wait(0.01)
    session.kill_all()
    worker.join(5)
    assert not worker.is_alive()
    assert [type(e) for e in errors] == [office.ConversionTimeout]
    assert blocked.killed == []
    # registrul strain ramane la vedere, in grija utilizatorului
    assert (fake.Visible, fake.UserControl) == (True, True)


def test_foreign_workbook_counted_before_open_blocks_kill(blocked: SimpleNamespace) -> None:
    fake = FakeExcel(workbooks=1)
    app = excel_app(fake)
    app.convert(Job(Path("a.xlsx"), Path("a.pdf"), Kind.EXCEL), ExportOptions(), Path("a.xlsx"))
    assert app.foreign_seen is True
    assert app.kill() is False
    assert blocked.killed == []


def test_kill_without_foreign_workbooks(blocked: SimpleNamespace) -> None:
    app = excel_app(FakeExcel())
    assert app.kill() is True
    assert blocked.killed == [1234]
