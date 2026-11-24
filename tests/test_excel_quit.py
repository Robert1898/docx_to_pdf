"""Inchiderea instantei Excel: IgnoreRemoteRequests refacut si registrele straine pastrate
(obiect COM simulat, fara Office)."""

from __future__ import annotations

import logging
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
    """Excel.Application simulat: tine minte valoarea IgnoreRemoteRequests la Quit."""

    def __init__(self, ignore_remote: bool = False, workbooks: int = 0,
                 count_fails: bool = False, close_fails: bool = False) -> None:
        self.IgnoreRemoteRequests = ignore_remote
        self.Visible = self.UserControl = False
        self.Workbooks = FakeWorkbooks(Count=workbooks, excel=self, close_fails=close_fails)
        self.count_fails = count_fails
        self.quit_calls: list[bool] = []  # IgnoreRemoteRequests in momentul fiecarui Quit
        self.seen_open: list[int] = []

    def __getattribute__(self, name: str):
        if name == "Workbooks" and object.__getattribute__(self, "count_fails"):
            raise OSError("RPC server unavailable")
        return object.__getattribute__(self, name)

    def Quit(self) -> None:
        self.quit_calls.append(self.IgnoreRemoteRequests)


@pytest.fixture
def process(monkeypatch: pytest.MonkeyPatch) -> list[str]:
    """Inregistreaza deschiderea handle-ului si asteptarea/oprirea procesului."""
    calls: list[str] = []
    monkeypatch.setattr(office, "_open_process", lambda pid: calls.append(f"open {pid}") or 99)

    def wait(handle: int, seconds: float) -> bool:
        calls.append(f"wait {handle}")
        return False

    monkeypatch.setattr(office, "_wait_or_terminate", wait)
    return calls


def excel_app(fake: FakeExcel, owned: bool = True) -> office.ExcelApp:
    app = office.ExcelApp()
    app.app, app.owned, app.pid = fake, owned, 1234 if owned else None
    app.configure()
    return app


@pytest.mark.parametrize("original", [False, True])
def test_remote_requests_ignored_while_running_and_restored_before_quit(
        process: list[str], original: bool) -> None:
    fake = FakeExcel(ignore_remote=original)
    app = excel_app(fake)
    assert fake.IgnoreRemoteRequests is True  # instanta ascunsa nu primeste fisiere din Explorer
    app.quit()
    assert fake.quit_calls == [original]  # refacut inainte de Quit, nu dupa
    assert process == ["open 1234", "wait 99"]
    assert app.app is None


def test_instance_not_started_by_us_is_left_alone(process: list[str]) -> None:
    fake = FakeExcel()
    app = excel_app(fake, owned=False)
    assert fake.IgnoreRemoteRequests is False
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
    assert fake.IgnoreRemoteRequests is False  # setarea globala e refacuta oricum
    assert (fake.Visible, fake.UserControl, fake.DisplayAlerts) == (True, True, True)
    assert "1 registru(e) deschis(e) de altcineva" in caplog.text
    assert app.app is None


def test_own_open_workbook_is_not_foreign(process: list[str]) -> None:
    fake = FakeExcel()
    app = excel_app(fake)
    fake.Workbooks.Count = 1
    app.open_workbooks = 1  # Close a esuat: registrul e tot al nostru
    app.quit()
    assert fake.quit_calls == [False]
    assert process == ["open 1234", "wait 99"]


def test_unreadable_workbook_count_still_quits(process: list[str]) -> None:
    fake = FakeExcel()
    app = excel_app(fake)
    fake.count_fails = True
    app.quit()
    assert fake.quit_calls == [False]


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
