"""Linia de comanda: docconv.exe fisier.docx folder\\ --out D:\\pdf"""

from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import Sequence
from pathlib import Path

from . import APP_TITLE, CLI_NAME, __version__, shell
from .core import PDF_MODES, ExportOptions, collect_files, plan_jobs
from .engine import Engine, Result, Status
from .log import setup_logging
from .office import PdfWarningGuard

log = logging.getLogger(__name__)

MARKS = {Status.DONE: "OK", Status.FAILED: "EROARE", Status.SKIPPED: "OMIS",
         Status.CANCELLED: "ANULAT"}


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog=CLI_NAME,
        description=f"{APP_TITLE}: documente Word, Excel și PowerPoint în PDF, PDF în DOCX.",
    )
    p.add_argument("inputs", nargs="*", type=Path, help="fișiere sau foldere (recursiv)")
    p.add_argument("-o", "--out", type=Path, help="folder destinație (implicit: lângă original)")
    p.add_argument("--overwrite", action="store_true",
                   help="suprascrie rezultatele existente (implicit sunt omise)")
    p.add_argument("--pdfa", action="store_true", help="PDF/A pentru arhivare")
    p.add_argument("--no-bookmarks", action="store_true", help="fără semne de carte din titluri")
    p.add_argument("--screen", action="store_true", help="optimizat pentru ecran (fișier mai mic)")
    p.add_argument("--ocr", action="store_true", help="OCR pe PDF-urile scanate înainte de DOCX")
    p.add_argument("--pdf-mode", choices=PDF_MODES, default="auto",
                   help="PDF → DOCX: auto (implicit), layout = aspect fidel (rapoarte, "
                        "tabele), word = text curgător prin Word")
    p.add_argument("--timeout", type=int, default=180, metavar="SEC",
                   help="limită de timp per fișier (implicit 180)")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    integration = p.add_argument_group("integrare Windows")
    integration.add_argument("--register-shell", action="store_true",
                             help="adaugă „Convertește” în meniul click-dreapta")
    integration.add_argument("--register-sendto", action="store_true",
                             help="adaugă aplicația în „Trimite către”")
    integration.add_argument("--unregister-shell", action="store_true",
                             help="elimină ambele integrări")
    return p


def _integration(args: argparse.Namespace) -> bool:
    if args.unregister_shell:
        shell.unregister_context_menu()
        shell.remove_sendto()
    if args.register_shell:
        shell.register_context_menu()
    if args.register_sendto:
        shell.create_sendto()
    return bool(args.unregister_shell or args.register_shell or args.register_sendto)


def main(argv: Sequence[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")

    args = build_parser().parse_args(argv)
    setup_logging(console=True)
    PdfWarningGuard.restore()  # daca o rulare anterioara a cazut

    if _integration(args):
        return 0
    if not args.inputs:
        build_parser().print_help()
        return 2
    if args.timeout < 10:
        print("--timeout trebuie să fie cel puțin 10 secunde.", file=sys.stderr)
        return 2

    files = collect_files(args.inputs)
    if not files:
        print("Nu am găsit fișiere suportate.", file=sys.stderr)
        return 2
    if args.out:
        args.out.mkdir(parents=True, exist_ok=True)
    jobs = plan_jobs(files, args.out.resolve() if args.out else None)
    skip = [] if args.overwrite else [j for j in jobs if j.dst.exists()]

    options = ExportOptions(
        bookmarks=not args.no_bookmarks,
        pdfa=args.pdfa,
        for_screen=args.screen,
        ocr=args.ocr,
        timeout=args.timeout,
        pdf_mode=args.pdf_mode,
    )

    def report(result: Result) -> None:
        if result.status is Status.RUNNING:
            return
        line = f"[{MARKS[result.status]}] {result.job.src}"
        if result.status is Status.DONE:
            line += f" -> {result.job.dst}"
        if result.message:
            line += f" ({result.message})"
        print(line, flush=True)

    engine = Engine(options, listener=report)
    try:
        results = engine.run(jobs, skip=skip)
    except KeyboardInterrupt:
        engine.abort()
        print("Întrerupt.", file=sys.stderr)
        return 130

    failed = sum(r.status is Status.FAILED for r in results)
    done = sum(r.status is Status.DONE for r in results)
    print(f"\n{done} convertite, {len(skip)} omise, {failed} cu erori.")
    return 1 if failed else 0
