"""Integrarea cu Explorer scrie in HKCU real; testul ruleaza doar daca nu era deja activa."""

from __future__ import annotations

import winreg

import pytest

from convertor import shell


@pytest.mark.skipif(shell.context_menu_registered(), reason="meniul e deja activ pe acest PC")
def test_context_menu_roundtrip() -> None:
    try:
        shell.register_context_menu()
        assert shell.context_menu_registered()
        key = rf"{shell.CLASSES}\.pdf\shell\{shell.MENU_KEY}"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
            assert winreg.QueryValueEx(k, "MUIVerb")[0] == "Convertește în DOCX"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key + r"\command") as k:
            command = winreg.QueryValueEx(k, "")[0]
        assert command.endswith('--auto "%1"')
    finally:
        shell.unregister_context_menu()
    assert not shell.context_menu_registered()


def test_sendto_roundtrip() -> None:
    # APPDATA e redirectat de conftest, deci scurtatura ajunge intr-un folder temporar
    shell.sendto_path().parent.mkdir(parents=True, exist_ok=True)
    shell.create_sendto()
    assert shell.sendto_registered()
    shell.remove_sendto()
    assert not shell.sendto_registered()
