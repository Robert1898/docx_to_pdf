"""Setari memorate intre porniri, in %APPDATA%\\DocumentConverter\\settings.json."""

from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass, fields
from pathlib import Path

from .core import PDF_MODES, PDF_TARGETS, ExportOptions
from .i18n import LANGUAGES
from .paths import config_dir

log = logging.getLogger(__name__)

THEMES = ("system", "light", "dark")
SPLIT_MODES = ("each", "every", "parts", "ranges")
COMPRESS_LEVELS = ("light", "balanced", "strong")
IMAGE_PAGES = ("a4", "image")
IMAGE_FORMATS = ("png", "jpg")
MAX_WORKERS = 4


@dataclass
class Settings:
    theme: str = "system"
    language: str = "ro"  # „system” = dupa limba Windows
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
    pdf_target: str = "docx"
    workers: int = 2
    image_page: str = "a4"
    watermark_text: str = "CONFIDENȚIAL"
    watermark_opacity: int = 15
    renumber_pages: bool = True
    split_mode: str = "each"  # each / every / parts / ranges
    split_every: int = 200
    split_parts: int = 3
    split_ranges: str = ""
    compress_level: str = "balanced"
    images_format: str = "png"
    images_dpi: int = 150
    watch_enabled: bool = False
    watch_in: str = ""
    watch_out: str = ""
    autostart: bool = False
    check_updates: bool = True
    last_update_check: float = 0.0
    skipped_version: str = ""
    geometry: str = ""

    def export_options(self) -> ExportOptions:
        return ExportOptions(
            bookmarks=self.bookmarks,
            pdfa=self.pdfa,
            for_screen=self.for_screen,
            ocr=self.ocr,
            timeout=self.timeout,
            pdf_mode=self.pdf_mode,
            pdf_target=self.pdf_target,
            workers=self.workers,
            image_page=self.image_page,
        )


def settings_path() -> Path:
    return config_dir() / "settings.json"


def _choice(value: str, allowed: tuple[str, ...] | dict, default: str) -> str:
    return value if value in allowed else default


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
        if isinstance(default, float) and isinstance(value, int):
            value = float(value)
        # acceptam doar valori de acelasi tip ca valoarea implicita
        if value is not None and type(value) is type(default):
            setattr(settings, f.name, value)
    s = settings
    s.theme = _choice(s.theme, THEMES, "system")
    s.language = _choice(s.language, LANGUAGES, "ro")
    s.pdf_mode = _choice(s.pdf_mode, PDF_MODES, "auto")
    s.pdf_target = _choice(s.pdf_target, PDF_TARGETS, "docx")
    s.split_mode = _choice(s.split_mode, SPLIT_MODES, "each")
    s.compress_level = _choice(s.compress_level, COMPRESS_LEVELS, "balanced")
    s.image_page = _choice(s.image_page, IMAGE_PAGES, "a4")
    s.images_format = _choice(s.images_format, IMAGE_FORMATS, "png")
    s.timeout = max(10, min(s.timeout, 3600))
    s.watermark_opacity = max(5, min(s.watermark_opacity, 100))
    s.split_every = max(1, s.split_every)
    s.split_parts = max(1, s.split_parts)
    s.workers = max(1, min(s.workers, MAX_WORKERS))
    s.images_dpi = max(72, min(s.images_dpi, 600))
    return s


def save(settings: Settings, path: Path | None = None) -> None:
    path = path or settings_path()
    tmp = path.with_suffix(".tmp")
    try:
        tmp.write_text(json.dumps(asdict(settings), indent=2, ensure_ascii=False), encoding="utf-8")
        tmp.replace(path)
    except OSError as exc:
        log.warning("Nu am putut salva setarile: %s", exc)
