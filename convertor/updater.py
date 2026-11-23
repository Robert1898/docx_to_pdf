"""Actualizari din GitHub Releases (publicate de workflow-ul de release al proiectului).

Se verifica cel mult o data pe zi. Daca exista o versiune mai noua, installer-ul ei
(„Document-Converter-Setup-X.Y.Z.exe”) e descarcat si pornit; aplicatia se inchide, iar
installer-ul inlocuieste versiunea veche.

Installer-ul e pornit doar daca e verificat: hash-ul SHA-256 trebuie sa fie cel din
SHA256SUMS.txt, publicat in acelasi release. Un release fara SHA256SUMS.txt nu se instaleaza
automat (aplicatia ofera doar linkul spre pagina lui). Toate descarcarile merg doar pe https,
la domeniile GitHub (si dupa redirectionari).
"""

from __future__ import annotations

import hashlib
import json
import logging
import re
import shutil
import subprocess
import tempfile
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

from . import GITHUB_REPO, __version__
from .i18n import _

log = logging.getLogger(__name__)

API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
ASSET = re.compile(r"-Setup-[\d.]+\.exe$", re.IGNORECASE)
CHECKSUMS = "SHA256SUMS.txt"
CHECK_INTERVAL = 24 * 3600
# github.com redirectioneaza descarcarile spre release-assets.githubusercontent.com
# (inainte spre objects.githubusercontent.com)
ALLOWED_HOSTS = frozenset({
    "github.com", "objects.githubusercontent.com", "release-assets.githubusercontent.com"})
MAX_CHECKSUMS_SIZE = 64 * 1024
CHUNK = 256 * 1024


class UpdateError(Exception):
    """Actualizare refuzata sau esuata; mesajul e afisabil direct utilizatorului."""


@dataclass
class Release:
    version: str
    asset_url: str
    asset_name: str
    asset_size: int
    page_url: str
    notes: str
    checksums_url: str = ""

    @property
    def verifiable(self) -> bool:
        """Poate fi instalat automat: are sume de control si adrese GitHub pe https."""
        return bool(self.checksums_url) and all(
            is_allowed_url(url) for url in (self.asset_url, self.checksums_url))


def parse_version(text: str) -> tuple[int, ...]:
    """„v2.3.0” -> (2, 3, 0); partile nenumerice sunt ignorate."""
    return tuple(int(p) for p in re.findall(r"\d+", text)[:4]) or (0,)


def is_newer(remote: str, local: str = __version__) -> bool:
    return parse_version(remote) > parse_version(local)


def is_allowed_url(url: str) -> bool:
    """Doar https, doar domeniile GitHub de unde vin release-urile."""
    try:
        parts = urlsplit(url)
    except ValueError:
        return False
    return parts.scheme == "https" and (parts.hostname or "").lower() in ALLOWED_HOSTS


def _require_allowed(url: str) -> None:
    if not is_allowed_url(url):
        host = urlsplit(url).hostname or url
        raise UpdateError(_("Adresă de descărcare nepermisă: {host}").format(host=host))


class _CheckedRedirects(urllib.request.HTTPRedirectHandler):
    """Urmeaza redirectionarile doar spre domeniile permise."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):  # type: ignore[override]
        _require_allowed(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _open(url: str, timeout: float) -> Any:
    """Raspunsul HTTP pentru o descarcare (separat, ca testele sa-l poata simula)."""
    _require_allowed(url)
    opener = urllib.request.build_opener(_CheckedRedirects)
    request = urllib.request.Request(url, headers={"User-Agent": "DocumentConverter"})
    return opener.open(request, timeout=timeout)


def _get_json(url: str, timeout: float) -> dict:
    request = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github+json", "User-Agent": "DocumentConverter"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def latest_release(timeout: float = 10) -> Release | None:
    """Ultimul release cu installer; None daca nu se poate afla (offline, depozit privat)."""
    try:
        data = _get_json(API, timeout)
    except (OSError, ValueError) as exc:
        log.info("Nu am putut verifica actualizarile: %s", exc)
        return None
    assets = data.get("assets", [])
    checksums = next((a.get("browser_download_url", "") for a in assets
                      if a.get("name") == CHECKSUMS), "")
    for asset in assets:
        if ASSET.search(asset.get("name", "")):
            return Release(
                version=str(data.get("tag_name", "")).lstrip("v"),
                asset_url=asset["browser_download_url"],
                asset_name=asset["name"],
                asset_size=int(asset.get("size", 0)),
                page_url=str(data.get("html_url", "")),
                notes=str(data.get("body") or ""),
                checksums_url=str(checksums),
            )
    return None


def parse_checksums(text: str) -> dict[str, str]:
    """Formatul lui sha256sum: „<hash>  <nume>” (sau „<hash> *<nume>”) pe fiecare rand."""
    sums = {}
    for line in text.splitlines():
        match = re.fullmatch(r"\s*([0-9a-fA-F]{64}) [ *](.+?)\s*", line)
        if match:
            sums[match.group(2)] = match.group(1).lower()
    return sums


def expected_hash(release: Release, timeout: float = 30) -> str:
    """Hash-ul installer-ului, din SHA256SUMS.txt publicat in release."""
    if not release.checksums_url:
        raise UpdateError(_("Release-ul nu are {file}; actualizarea automată nu e sigură.")
                          .format(file=CHECKSUMS))
    try:
        with _open(release.checksums_url, timeout) as response:
            text = response.read(MAX_CHECKSUMS_SIZE + 1)
    except OSError as exc:
        raise UpdateError(_("Nu am putut descărca {file}: {error}").format(
            file=CHECKSUMS, error=exc)) from exc
    if len(text) > MAX_CHECKSUMS_SIZE:
        raise UpdateError(_("{file} este prea mare.").format(file=CHECKSUMS))
    digest = parse_checksums(text.decode("utf-8", errors="replace")).get(release.asset_name)
    if digest is None:
        raise UpdateError(_("{file} nu conține installer-ul {name}.").format(
            file=CHECKSUMS, name=release.asset_name))
    return digest


def download(release: Release, progress: Callable[[int, int], None] | None = None) -> Path:
    """Descarca installer-ul intr-un folder temporar si il verifica (marime + SHA-256).

    Intoarce calea doar daca verificarea a reusit; altfel sterge fisierul si ridica
    UpdateError."""
    _require_allowed(release.asset_url)
    expected = expected_hash(release)
    folder = Path(tempfile.mkdtemp(prefix="docconv-update-"))
    dst = folder / Path(release.asset_name).name
    sha = hashlib.sha256()
    try:
        try:
            with _open(release.asset_url, 60) as response, open(dst, "wb") as fh:
                total = int(response.headers.get("Content-Length") or release.asset_size or 0)
                done = 0
                while chunk := response.read(CHUNK):
                    fh.write(chunk)
                    sha.update(chunk)
                    done += len(chunk)
                    if progress and total:
                        progress(done, total)
        except OSError as exc:
            raise UpdateError(_("Descărcarea a eșuat: {error}").format(error=exc)) from exc
        if release.asset_size and dst.stat().st_size != release.asset_size:
            raise UpdateError(_("Descărcare incompletă."))
        if sha.hexdigest() != expected:
            log.warning("Hash gresit pentru %s: %s, asteptam %s", release.asset_name,
                        sha.hexdigest(), expected)
            raise UpdateError(_("Installer-ul descărcat nu se potrivește cu suma de control "
                                "din release; l-am șters."))
    except BaseException:
        shutil.rmtree(folder, ignore_errors=True)
        raise
    log.info("Installer verificat: %s (sha256 %s)", dst.name, expected)
    return dst


def launch_installer(path: Path) -> None:
    """Porneste installer-ul; aplicatia trebuie sa se inchida imediat dupa (installer-ul o
    inchide oricum daca mai ruleaza). Primeste doar calea intoarsa de download()."""
    subprocess.Popen([str(path)], close_fds=True)
