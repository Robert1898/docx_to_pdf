"""Setari memorate intre porniri, in %APPDATA%\\DocumentConverter\\settings.json."""

from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass, fields
from pathlib import Path

from .core import PDF_MODES, ExportOptions
from .paths import config_dir

log = logging.getLogger(__name__)

THEMES = ("system", "light", "dark")
SPLIT_MODES = ("each", "every", "parts", "ranges")


@dataclass
class Settings:
    theme: str = "system"
    same_folder: bool = True
    output_dir: str = ""
    open_after: bool = False
    last_dir: str = ""
    bookmarks: bool = True
    pdfa: bool = False
    for_screen: bool = False
    ocr: bool = False
    timeout: int = 180
    pdf_mode: str = "auto"
    watermark_text: str = "CONFIDENȚIAL"
    watermark_opacity: int = 15
    renumber_pages: bool = True
    split_mode: str = "each"  # each / every / parts / ranges
    split_every: int = 200
    split_parts: int = 3
    split_ranges: str = ""
    geometry: str = ""

    def export_options(self) -> ExportOptions:
        return ExportOptions(
            bookmarks=self.bookmarks,
            pdfa=self.pdfa,
            for_screen=self.for_screen,
            ocr=self.ocr,
            timeout=self.timeout,
            pdf_mode=self.pdf_mode,
        )


def settings_path() -> Path:
    return config_dir() / "settings.json"


def load(path: Path | None = None) -> Settings:
    path = path or settings_path()
    settings = Settings()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return settings
    except (OSError, ValueError) as exc:
        log.warning("Setari ilizibile in %s, folosesc valorile implicite: %s", path, exc)
        return settings
    if not isinstance(data, dict):
        return settings

    for f in fields(Settings):
        value = data.get(f.name)
        default = getattr(settings, f.name)
        # acceptam doar valori de acelasi tip ca valoarea implicita
        if value is not None and type(value) is type(default):
            setattr(settings, f.name, value)
    if settings.theme not in THEMES:
        settings.theme = "system"
    settings.timeout = max(10, min(settings.timeout, 3600))
    settings.watermark_opacity = max(5, min(settings.watermark_opacity, 100))
    if settings.pdf_mode not in PDF_MODES:
        settings.pdf_mode = "auto"
    if settings.split_mode not in SPLIT_MODES:
        settings.split_mode = "each"
    settings.split_every = max(1, settings.split_every)
    settings.split_parts = max(1, settings.split_parts)
    return settings


def save(settings: Settings, path: Path | None = None) -> None:
    path = path or settings_path()
    tmp = path.with_suffix(".tmp")
    try:
        tmp.write_text(json.dumps(asdict(settings), indent=2, ensure_ascii=False), encoding="utf-8")
        tmp.replace(path)
    except OSError as exc:
        log.warning("Nu am putut salva setarile: %s", exc)
