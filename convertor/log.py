"""Jurnal rotativ in %LOCALAPPDATA%, util cand exe-ul da erori la altcineva."""

from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from . import __version__
from .paths import log_dir

LOG_FILE = "convertor.log"


def setup_logging(console: bool = False, level: int = logging.INFO) -> Path:
    path = log_dir() / LOG_FILE
    root = logging.getLogger()
    if any(getattr(h, "_convertor", False) for h in root.handlers):
        return path

    root.setLevel(level)
    file_handler = RotatingFileHandler(path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)-7s %(name)s: %(message)s")
    )
    file_handler._convertor = True  # type: ignore[attr-defined]
    root.addHandler(file_handler)

    if console:
        stream = logging.StreamHandler(sys.stderr)
        stream.setLevel(logging.WARNING)
        stream.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        root.addHandler(stream)

    logging.getLogger(__name__).info("Pornit v%s (%s)", __version__, sys.executable)
    return path
