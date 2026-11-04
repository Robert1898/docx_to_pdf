"""Trecerea de la numele vechi („Convertor DOCX-PDF”) la „Document Converter”."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from convertor import APP_ID, LEGACY_APP_ID, LEGACY_APP_NAME, LEGACY_MENU_KEY, paths, shell


def test_legacy_settings_folder_is_taken_over() -> None:
    legacy = Path(os.environ["APPDATA"]) / LEGACY_APP_ID
    legacy.mkdir(parents=True)
    (legacy / "settings.json").write_text('{"theme": "dark"}', encoding="utf-8")

    folder = paths.config_dir()

    assert folder.name == APP_ID
    assert (folder / "settings.json").read_text(encoding="utf-8") == '{"theme": "dark"}'
    assert not legacy.exists()


def test_existing_new_folder_wins_over_legacy() -> None:
    base = Path(os.environ["LOCALAPPDATA"])
    (base / LEGACY_APP_ID).mkdir(parents=True)
    (base / APP_ID).mkdir(parents=True)
    (base / APP_ID / "keep.txt").write_text("nou")

    assert (paths.state_dir() / "keep.txt").read_text() == "nou"
    assert (base / LEGACY_APP_ID).exists()  # nu suprascriem nimic


@pytest.mark.skipif(shell.context_menu_registered(LEGACY_MENU_KEY),
                    reason="meniul vechi e activ pe acest PC")
def test_legacy_sendto_shortcut_is_replaced() -> None:
    legacy = shell.sendto_path(LEGACY_APP_NAME)
    legacy.parent.mkdir(parents=True, exist_ok=True)
    legacy.write_bytes(b"scurtatura veche")

    shell.migrate_legacy()

    assert not legacy.exists()
    assert shell.sendto_registered()
    shell.remove_sendto()
