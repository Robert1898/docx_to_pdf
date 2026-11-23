from __future__ import annotations

import hashlib
import io
import re
import tempfile
import urllib.request
from pathlib import Path

import pytest

from convertor import updater

INSTALLER = b"MZ installer simulat" * 1000
SETUP = "Document-Converter-Setup-9.1.0.exe"
SETUP_URL = f"https://github.com/o/r/releases/download/v9.1.0/{SETUP}"
SUMS_URL = "https://github.com/o/r/releases/download/v9.1.0/SHA256SUMS.txt"


class FakeResponse(io.BytesIO):
    """Raspuns HTTP simulat: continut + antete."""

    def __init__(self, body: bytes) -> None:
        super().__init__(body)
        self.headers = {"Content-Length": str(len(body))}


def release(**changes) -> updater.Release:
    values = dict(version="9.1.0", asset_url=SETUP_URL, asset_name=SETUP,
                  asset_size=len(INSTALLER), page_url="https://github.com/o/r/releases/v9.1.0",
                  notes="", checksums_url=SUMS_URL)
    values.update(changes)
    return updater.Release(**values)


@pytest.fixture
def temp_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Folderul temporar al descarcarilor, ca sa vedem ce ramane pe disc."""
    monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
    return tmp_path


def serve(monkeypatch: pytest.MonkeyPatch, files: dict[str, bytes]) -> list[str]:
    """Simuleaza descarcarile: url -> continut; intoarce adresele cerute."""
    requested: list[str] = []
    real_check = updater._require_allowed

    def fake_open(url: str, timeout: float) -> FakeResponse:
        real_check(url)  # verificarea domeniului ramane cea adevarata
        requested.append(url)
        if url not in files:
            raise OSError("404")
        return FakeResponse(files[url])

    monkeypatch.setattr(updater, "_open", fake_open)
    return requested


def sums(digest: str, name: str = SETUP) -> bytes:
    return f"{digest}  {name}\n".encode()


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


def test_allowed_urls() -> None:
    assert updater.is_allowed_url(SETUP_URL)
    assert updater.is_allowed_url("https://objects.githubusercontent.com/x")
    assert updater.is_allowed_url("https://release-assets.githubusercontent.com/x")
    assert not updater.is_allowed_url("http://github.com/x")  # fara https
    assert not updater.is_allowed_url("https://github.com.evil.example/x")
    assert not updater.is_allowed_url("https://evil.example/github.com/x")
    assert not updater.is_allowed_url("file:///C:/setup.exe")
    assert not updater.is_allowed_url("")


def test_parse_checksums() -> None:
    digest = "a" * 64
    text = f"{digest}  {SETUP}\r\n{'B' * 64} *alt.exe\nrand stricat\n"
    assert updater.parse_checksums(text) == {SETUP: digest, "alt.exe": "b" * 64}


def test_latest_release_picks_the_installer(monkeypatch: pytest.MonkeyPatch) -> None:
    data = {
        "tag_name": "v9.1.0",
        "html_url": "https://github.com/o/r/releases/v9.1.0",
        "body": "note",
        "assets": [
            {"name": "notes.txt", "browser_download_url": "https://x/notes.txt", "size": 1},
            {"name": SETUP, "browser_download_url": SETUP_URL, "size": 123},
            {"name": "SHA256SUMS.txt", "browser_download_url": SUMS_URL, "size": 100},
        ],
    }
    monkeypatch.setattr(updater, "_get_json", lambda url, timeout: data)
    found = updater.latest_release()
    assert found is not None
    assert (found.version, found.asset_url, found.asset_size, found.checksums_url) == (
        "9.1.0", SETUP_URL, 123, SUMS_URL)
    assert found.verifiable


def test_release_without_checksums_is_not_verifiable(monkeypatch: pytest.MonkeyPatch) -> None:
    data = {"tag_name": "v9.1.0", "assets": [
        {"name": SETUP, "browser_download_url": SETUP_URL, "size": 123}]}
    monkeypatch.setattr(updater, "_get_json", lambda url, timeout: data)
    found = updater.latest_release()
    assert found is not None
    assert found.checksums_url == ""
    assert not found.verifiable


def test_foreign_domain_is_not_verifiable() -> None:
    assert not release(asset_url="https://evil.example/setup.exe").verifiable
    assert not release(checksums_url="http://github.com/SHA256SUMS.txt").verifiable


def test_latest_release_offline(monkeypatch: pytest.MonkeyPatch) -> None:
    def fail(url, timeout):
        raise OSError("offline")

    monkeypatch.setattr(updater, "_get_json", fail)
    assert updater.latest_release() is None


def test_latest_release_without_installer(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(updater, "_get_json", lambda url, timeout: {"tag_name": "v9", "assets": []})
    assert updater.latest_release() is None


def test_download_with_correct_hash(monkeypatch: pytest.MonkeyPatch, temp_dir: Path) -> None:
    digest = hashlib.sha256(INSTALLER).hexdigest()
    serve(monkeypatch, {SUMS_URL: sums(digest.upper()), SETUP_URL: INSTALLER})
    seen: list[tuple[int, int]] = []
    path = updater.download(release(), lambda done, total: seen.append((done, total)))
    assert path.read_bytes() == INSTALLER
    assert path.name == SETUP
    assert seen[-1] == (len(INSTALLER), len(INSTALLER))


def test_download_with_wrong_hash_deletes_the_file(monkeypatch: pytest.MonkeyPatch,
                                                   temp_dir: Path) -> None:
    serve(monkeypatch, {SUMS_URL: sums("0" * 64), SETUP_URL: INSTALLER})
    launched: list[Path] = []
    monkeypatch.setattr(updater, "launch_installer", launched.append)
    with pytest.raises(updater.UpdateError, match="suma de control"):
        updater.download(release())
    assert list(temp_dir.rglob("*.exe")) == []  # installer-ul nu ramane pe disc
    assert list(temp_dir.iterdir()) == []
    assert launched == []


def test_download_refused_without_checksums(monkeypatch: pytest.MonkeyPatch,
                                            temp_dir: Path) -> None:
    requested = serve(monkeypatch, {SETUP_URL: INSTALLER})
    with pytest.raises(updater.UpdateError, match="SHA256SUMS"):
        updater.download(release(checksums_url=""))
    assert requested == []  # installer-ul nici nu e descarcat


def test_download_refused_when_checksums_missing_on_server(monkeypatch: pytest.MonkeyPatch,
                                                           temp_dir: Path) -> None:
    requested = serve(monkeypatch, {SETUP_URL: INSTALLER})
    with pytest.raises(updater.UpdateError, match="SHA256SUMS"):
        updater.download(release())
    assert requested == [SUMS_URL]


def test_download_refused_when_installer_not_listed(monkeypatch: pytest.MonkeyPatch,
                                                   temp_dir: Path) -> None:
    digest = hashlib.sha256(INSTALLER).hexdigest()
    requested = serve(monkeypatch, {SUMS_URL: sums(digest, "alt.exe"), SETUP_URL: INSTALLER})
    with pytest.raises(updater.UpdateError, match=re.escape(SETUP)):
        updater.download(release())
    assert requested == [SUMS_URL]


def test_download_refused_for_foreign_domain(monkeypatch: pytest.MonkeyPatch,
                                             temp_dir: Path) -> None:
    evil = "https://evil.example/Document-Converter-Setup-9.1.0.exe"
    digest = hashlib.sha256(INSTALLER).hexdigest()
    requested = serve(monkeypatch, {SUMS_URL: sums(digest), evil: INSTALLER})
    with pytest.raises(updater.UpdateError, match=r"evil\.example"):
        updater.download(release(asset_url=evil))
    assert requested == []
    with pytest.raises(updater.UpdateError, match="nepermis"):
        updater.download(release(asset_url=SETUP_URL.replace("https", "http")))


def test_download_incomplete(monkeypatch: pytest.MonkeyPatch, temp_dir: Path) -> None:
    digest = hashlib.sha256(INSTALLER).hexdigest()
    serve(monkeypatch, {SUMS_URL: sums(digest), SETUP_URL: INSTALLER[:-10]})
    with pytest.raises(updater.UpdateError, match="incompletă"):
        updater.download(release())
    assert list(temp_dir.iterdir()) == []


def test_redirect_to_foreign_domain_is_refused() -> None:
    handler = updater._CheckedRedirects()
    request = urllib.request.Request(SETUP_URL)
    allowed = handler.redirect_request(request, None, 302, "Found", {},
                                       "https://release-assets.githubusercontent.com/x")
    assert allowed is not None
    with pytest.raises(updater.UpdateError, match=r"evil\.example"):
        handler.redirect_request(request, None, 302, "Found", {}, "https://evil.example/x")
