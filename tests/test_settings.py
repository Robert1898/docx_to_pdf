from __future__ import annotations

import json
from pathlib import Path

from convertor import settings


def test_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / "s.json"
    s = settings.Settings(theme="dark", timeout=300, pdfa=True, output_dir="D:/out")
    settings.save(s, path)
    assert settings.load(path) == s


def test_missing_file_gives_defaults(tmp_path: Path) -> None:
    assert settings.load(tmp_path / "nope.json") == settings.Settings()


def test_corrupt_file_gives_defaults(tmp_path: Path) -> None:
    path = tmp_path / "s.json"
    path.write_text("{ nu e json", encoding="utf-8")
    assert settings.load(path) == settings.Settings()


def test_wrong_types_and_unknown_keys_are_ignored(tmp_path: Path) -> None:
    path = tmp_path / "s.json"
    path.write_text(json.dumps({"theme": 5, "pdfa": "da", "extra": 1, "same_folder": False}))
    loaded = settings.load(path)
    assert loaded.theme == "system"
    assert loaded.pdfa is False
    assert loaded.same_folder is False


def test_values_are_clamped(tmp_path: Path) -> None:
    path = tmp_path / "s.json"
    path.write_text(json.dumps({"theme": "neon", "timeout": 1, "watermark_opacity": 900}))
    loaded = settings.load(path)
    assert (loaded.theme, loaded.timeout, loaded.watermark_opacity) == ("system", 10, 100)


def test_export_options_follow_settings() -> None:
    opts = settings.Settings(bookmarks=False, for_screen=True, timeout=60).export_options()
    assert (opts.bookmarks, opts.for_screen, opts.timeout) == (False, True, 60)
