"""Interfata grafica. Argumente: fisiere de incarcat in lista, --auto (porneste conversia
imediat) si --minimized (pornire cu Windows, de exemplu pentru folderul urmarit)."""

from __future__ import annotations

import argparse
import logging
import os
from collections.abc import Sequence

from .. import APP_NAME, i18n, settings, shell
from ..log import setup_logging
from ..office import PdfWarningGuard, ResiliencyCleaner


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog=APP_NAME, add_help=False)
    parser.add_argument("files", nargs="*")
    parser.add_argument("--auto", action="store_true")
    parser.add_argument("--minimized", action="store_true")
    args, unknown = parser.parse_known_args(argv)

    setup_logging()
    log = logging.getLogger(__name__)
    if unknown:
        log.info("Argumente ignorate: %s", unknown)

    from .instance import claim

    # caile relative ar insemna altceva in instanta principala (alt folder curent)
    files = [os.path.abspath(f) for f in args.files]
    listener = claim({"files": files, "auto": args.auto})
    if listener is None:
        log.info("Am predat %d fisier(e) ferestrei deja deschise", len(files))
        return 0

    # curatenie dupa o rulare anterioara care s-a oprit brusc
    PdfWarningGuard.restore()
    ResiliencyCleaner.clean()
    shell.migrate_legacy()
    i18n.set_language(settings.load().language)

    from .app import App

    App(files=files, auto=args.auto, listener=listener, minimized=args.minimized).run()
    return 0
