"""Actualizari din GitHub Releases (publicate de workflow-ul de release al proiectului).

Se verifica cel mult o data pe zi. Daca exista o versiune mai noua, installer-ul ei
(„Document-Converter-Setup-X.Y.Z.exe”) e descarcat si pornit; aplicatia se inchide, iar
installer-ul inlocuieste versiunea veche.
"""

from __future__ import annotations

import json
import logging
import re
import subprocess
import tempfile
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from . import GITHUB_REPO, __version__

log = logging.getLogger(__name__)

API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
ASSET = re.compile(r"-Setup-[\d.]+\.exe$", re.IGNORECASE)
CHECK_INTERVAL = 24 * 3600


@dataclass
class Release:
    version: str
    asset_url: str
    asset_name: str
    asset_size: int
    page_url: str
    notes: str


def parse_version(text: str) -> tuple[int, ...]:
    """„v2.3.0” -> (2, 3, 0); partile nenumerice sunt ignorate."""
    return tuple(int(p) for p in re.findall(r"\d+", text)[:4]) or (0,)


def is_newer(remote: str, local: str = __version__) -> bool:
    return parse_version(remote) > parse_version(local)


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
    for asset in data.get("assets", []):
        if ASSET.search(asset.get("name", "")):
            return Release(
                version=str(data.get("tag_name", "")).lstrip("v"),
                asset_url=asset["browser_download_url"],
                asset_name=asset["name"],
                asset_size=int(asset.get("size", 0)),
                page_url=str(data.get("html_url", "")),
                notes=str(data.get("body") or ""),
            )
    return None


def download(release: Release, progress: Callable[[int, int], None] | None = None) -> Path:
    """Descarca installer-ul intr-un folder temporar; verifica marimea anuntata."""
    folder = Path(tempfile.mkdtemp(prefix="docconv-update-"))
    dst = folder / release.asset_name
    request = urllib.request.Request(release.asset_url,
                                     headers={"User-Agent": "DocumentConverter"})
    with urllib.request.urlopen(request, timeout=60) as response, open(dst, "wb") as fh:
        total = int(response.headers.get("Content-Length") or release.asset_size or 0)
        done = 0
        while chunk := response.read(256 * 1024):
            fh.write(chunk)
            done += len(chunk)
            if progress and total:
                progress(done, total)
    if release.asset_size and dst.stat().st_size != release.asset_size:
        dst.unlink(missing_ok=True)
        raise OSError("descarcare incompleta")
    return dst


def launch_installer(path: Path) -> None:
    """Porneste installer-ul; aplicatia trebuie sa se inchida imediat dupa (installer-ul o
    inchide oricum daca mai ruleaza)."""
    subprocess.Popen([str(path)], close_fds=True)
