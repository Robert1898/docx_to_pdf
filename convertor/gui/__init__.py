"""Interfata grafica. Argumente: fisiere de incarcat in lista si --auto (porneste imediat)."""

from __future__ import annotations

import argparse
import logging
from collections.abc import Sequence

from .. import APP_NAME, shell
from ..log import setup_logging
from ..office import PdfWarningGuard, ResiliencyCleaner


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog=APP_NAME, add_help=False)
    parser.add_argument("files", nargs="*")
    parser.add_argument("--auto", action="store_true")
    args, unknown = parser.parse_known_args(argv)

    setup_logging()
    if unknown:
        logging.getLogger(__name__).info("Argumente ignorate: %s", unknown)
    # curatenie dupa o rulare anterioara care s-a oprit brusc
    PdfWarningGuard.restore()
    ResiliencyCleaner.clean()
    shell.migrate_legacy()

    from .app import App

    App(files=args.files, auto=args.auto).run()
    return 0
