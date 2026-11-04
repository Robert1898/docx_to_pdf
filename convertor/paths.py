"""Locatii pentru setari, jurnale si resurse incluse in executabil."""

from __future__ import annotations

import contextlib
import os
import sys
from pathlib import Path

from . import APP_ID, LEGACY_APP_ID


def _env_dir(var: str, fallback: Path) -> Path:
    base = os.environ.get(var)
    return Path(base) if base else fallback


def _app_dir(base: Path) -> Path:
    """Folderul aplicatiei; cel de dinainte de redenumire (pana la 2.0) e preluat ca atare."""
    path = base / APP_ID
    legacy = base / LEGACY_APP_ID
    if not path.exists() and legacy.is_dir():
        # folder in uz: pornim curat, setarile vechi raman acolo
        with contextlib.suppress(OSError):
            legacy.rename(path)
    path.mkdir(parents=True, exist_ok=True)
    return path


def config_dir() -> Path:
    """%APPDATA%\\DocumentConverter: setari care merg cu profilul utilizatorului."""
    return _app_dir(_env_dir("APPDATA", Path.home() / "AppData" / "Roaming"))


def state_dir() -> Path:
    """%LOCALAPPDATA%\\DocumentConverter: jurnale si stare locala."""
    return _app_dir(_env_dir("LOCALAPPDATA", Path.home() / "AppData" / "Local"))


def log_dir() -> Path:
    path = state_dir() / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def resource_path(name: str) -> Path:
    """Fisier inclus cu PyInstaller (sys._MEIPASS) sau din radacina proiectului."""
    base = getattr(sys, "_MEIPASS", None)
    root = Path(base) if base else Path(__file__).resolve().parent.parent
    return root / name
