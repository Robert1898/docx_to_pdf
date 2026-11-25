"""Versiunea lui Word din registry (fara Office: registry-ul e simulat)."""

from __future__ import annotations

import contextlib
import winreg

import pytest

from convertor import office
from convertor.office import registry


@pytest.fixture
def curver(monkeypatch: pytest.MonkeyPatch) -> list[str]:
    """Valoarea simulata a cheii Word.Application\\CurVer."""
    value = ["Word.Application.16"]
    monkeypatch.setattr(winreg, "OpenKey", lambda *args: contextlib.nullcontext())
    monkeypatch.setattr(winreg, "QueryValue", lambda key, name: value[0])
    return value


def test_word_version_and_the_keys_built_from_it(curver: list[str]) -> None:
    assert registry.word_version() == "16.0"
    assert office.PdfWarningGuard._options_key() == r"Software\Microsoft\Office\16.0\Word\Options"
    assert office.ResiliencyCleaner.root() == r"Software\Microsoft\Office\16.0\Word\Resiliency"
    curver[0] = "Word.Application.15"
    assert registry.word_version() == "15.0"


@pytest.mark.parametrize("value", ["Word.Application", "Word.Application.", "ceva ciudat"])
def test_unreadable_word_version_is_an_os_error(curver: list[str], value: str) -> None:
    # OSError: apelantii (de ex. curatarea marcajelor Word) o trateaza deja
    curver[0] = value
    with pytest.raises(OSError, match="Versiune Word necunoscuta"):
        registry.word_version()


def test_word_guard_flag_is_per_instance() -> None:
    first, second = office.WordApp(), office.WordApp()
    first._guarded = True
    assert second._guarded is False
    assert "_guarded" not in vars(office.WordApp)
