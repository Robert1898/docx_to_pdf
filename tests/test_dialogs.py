"""Alegerea butonului din dialogurile Office (fara Office si fara PowerShell)."""

from __future__ import annotations

import logging

import pytest

from convertor import office


@pytest.mark.parametrize(("buttons", "expected"), [
    (["Da", "Nu"], "Da"),
    (["Yes", "No"], "Yes"),
    (["OK"], "OK"),
    (["Deschide", "Anulare"], "Deschide"),
    (["Open", "Cancel"], "Open"),
    # „Doar citire” castiga in fata lui „OK” (parola de scriere: OK cu parola goala ar esua)
    (["OK", "Doar citire", "Anulare"], "Doar citire"),
    (["Read-Only", "OK", "Cancel"], "Read-Only"),
    (["Da", "OK"], "Da"),
    # numele sunt comparate fara accelerator, spatii in plus si majuscule
    (["&Da", "&Nu"], "&Da"),
    (["  read-only ", "Cancel"], "  read-only "),
    (["DOAR  CITIRE"], "DOAR  CITIRE"),
])
def test_choose_allowed_button(buttons: list[str], expected: str) -> None:
    assert office.choose_button(buttons) == expected


@pytest.mark.parametrize("buttons", [
    [],
    ["Nu", "Anulare"],
    ["No", "Cancel", "Help"],
    ["Salvează", "Nu salva"],  # nu e in lista alba: nu apasam niciodata „Salvează”
    ["Notificare", "Anulare"],
    ["Data"],  # doar potrivire exacta, nu prefix
    ["Okay"],
])
def test_no_allowed_button(buttons: list[str]) -> None:
    assert office.choose_button(buttons) is None


def test_first_matching_button_in_dialog_order_is_kept() -> None:
    assert office.choose_button(["Da", "da"]) == "Da"


def test_custom_whitelist() -> None:
    assert office.choose_button(["OK", "Retry"], allowed=("Retry",)) == "Retry"


def test_whitelist_is_the_documented_one() -> None:
    assert set(office.ALLOWED_DIALOG_BUTTONS) == {
        "Da", "Yes", "OK", "Doar citire", "Read-Only", "Deschide", "Open"}


def _fake_scripts(monkeypatch: pytest.MonkeyPatch, read: dict, pressed: str = "") -> list:
    calls: list[tuple[str, tuple[str, ...]]] = []

    def run(script: str, *args: str) -> dict:
        kind = "read" if script is office._READ_DIALOG_PS else "press"
        calls.append((kind, args))
        return read if kind == "read" else {"button": pressed}

    monkeypatch.setattr(office, "_run_dialog_script", run)
    return calls


def test_answer_dialog_presses_only_the_chosen_button(monkeypatch: pytest.MonkeyPatch) -> None:
    calls = _fake_scripts(monkeypatch, {"text": "Îl deschideți oricum?",
                                        "buttons": ["Da", "Nu"]}, pressed="Da")
    assert office.answer_dialog(42) == ("Îl deschideți oricum?", ["Da", "Nu"], "Da")
    assert calls == [("read", ("42",)), ("press", ("42", "'Da'"))]


def test_answer_dialog_without_allowed_button_presses_nothing(
        monkeypatch: pytest.MonkeyPatch) -> None:
    calls = _fake_scripts(monkeypatch, {"text": "Salvați modificările?",
                                        "buttons": ["Salvează", "Nu salva", "Anulare"]})
    assert office.answer_dialog(7) == (
        "Salvați modificările?", ["Salvează", "Nu salva", "Anulare"], None)
    assert [kind for kind, _ in calls] == ["read"]


def test_answer_dialog_single_button_from_json(monkeypatch: pytest.MonkeyPatch) -> None:
    # ConvertTo-Json scrie o lista cu un element ca text simplu
    _fake_scripts(monkeypatch, {"text": "Eroare", "buttons": "OK"}, pressed="OK")
    assert office.answer_dialog(1) == ("Eroare", ["OK"], "OK")


def test_answer_dialog_closed_meanwhile(monkeypatch: pytest.MonkeyPatch) -> None:
    _fake_scripts(monkeypatch, {"text": "x", "buttons": ["OK"]}, pressed="")
    assert office.answer_dialog(1) == ("x", ["OK"], None)


def test_answer_dialog_unreadable(monkeypatch: pytest.MonkeyPatch) -> None:
    def fail(script: str, *args: str) -> dict:
        raise OSError("powershell lipseste")

    monkeypatch.setattr(office, "_run_dialog_script", fail)
    assert office.answer_dialog(1) is None


def test_watcher_logs_unanswered_dialog(monkeypatch: pytest.MonkeyPatch,
                                        caplog: pytest.LogCaptureFixture) -> None:
    answers = {1: ("Îl deschideți oricum?", ["Da", "Nu"], "Da"),
               2: ("Ceva neașteptat", ["Nu", "Anulare"], None)}
    monkeypatch.setattr(office, "_visible_dialogs", lambda pid: [1, 2])
    monkeypatch.setattr(office, "answer_dialog", answers.get)
    watcher = office.DialogWatcher(pid=123, interval=0.01)
    with caplog.at_level(logging.WARNING, logger="convertor.office"):
        watcher.start()
        watcher.join(0.2)
        watcher.stop()
        watcher.join(1)
    assert watcher.answers == ["Îl deschideți oricum?"]
    assert watcher.unanswered == ["Ceva neașteptat"]
    assert watcher.last_text == "Ceva neașteptat"
    assert "am apasat 'Da'" in caplog.text
    assert "niciun buton permis" in caplog.text
    assert caplog.text.count("Ceva neașteptat") == 1  # fiecare dialog e tratat o singura data
