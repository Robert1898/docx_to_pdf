from __future__ import annotations

import pytest

from convertor import updater


def test_parse_version() -> None:
    assert updater.parse_version("v2.3.0") == (2, 3, 0)
    assert updater.parse_version("2.10") == (2, 10)
    assert updater.parse_version("nimic") == (0,)


def test_is_newer() -> None:
    assert updater.is_newer("2.10.0", "2.9.9")
    assert updater.is_newer("v3.0", "2.99")
    assert not updater.is_newer("2.3.0", "2.3.0")
    assert not updater.is_newer("2.2.9", "2.3.0")


def test_asset_pattern() -> None:
    assert updater.ASSET.search("DocumentConverter-Setup-2.3.0.exe")
    assert not updater.ASSET.search("DocumentConverter-2.3.0.zip")
    assert not updater.ASSET.search("DocumentConverter-Setup-2.3.0.exe.sha256")


def test_latest_release_picks_the_installer(monkeypatch: pytest.MonkeyPatch) -> None:
    data = {
        "tag_name": "v9.1.0",
        "html_url": "https://example.invalid/release",
        "body": "note",
        "assets": [
            {"name": "notes.txt", "browser_download_url": "https://x/notes.txt", "size": 1},
            {"name": "DocumentConverter-Setup-9.1.0.exe",
             "browser_download_url": "https://x/setup.exe", "size": 123},
        ],
    }
    monkeypatch.setattr(updater, "_get_json", lambda url, timeout: data)
    release = updater.latest_release()
    assert release is not None
    assert (release.version, release.asset_url, release.asset_size) == (
        "9.1.0", "https://x/setup.exe", 123)


def test_latest_release_offline(monkeypatch: pytest.MonkeyPatch) -> None:
    def fail(url, timeout):
        raise OSError("offline")

    monkeypatch.setattr(updater, "_get_json", fail)
    assert updater.latest_release() is None


def test_latest_release_without_installer(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(updater, "_get_json", lambda url, timeout: {"tag_name": "v9", "assets": []})
    assert updater.latest_release() is None
