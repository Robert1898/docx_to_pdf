"""Linia de comanda: docconv.exe fisier.docx folder\\ --out D:\\pdf

Acelasi motor ca interfata grafica: documente Office si imagini in PDF, PDF in DOCX sau
Excel, conversii in paralel (--jobs) si urmarirea unui folder (--watch)."""

from __future__ import annotations

import argparse
import logging
import sys
import time
from collections.abc import Sequence
from pathlib import Path

from . import APP_TITLE, CLI_NAME, __version__, i18n, settings, shell
from .core import PDF_MODES, PDF_TARGETS, ExportOptions, Job, collect_files, plan_jobs
from .engine import Engine, OfficePool, Result, Status
from .i18n import _
from .log import setup_logging
from .office import PdfWarningGuard

log = logging.getLogger(__name__)


def _marks() -> dict[Status, str]:
    return {Status.DONE: "OK", Status.FAILED: _("EROARE"), Status.SKIPPED: _("OMIS"),
            Status.CANCELLED: _("ANULAT")}


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog=CLI_NAME,
        description=_("{app}: documente Word, Excel, PowerPoint și imagini în PDF; PDF în "
                      "DOCX sau Excel.").format(app=APP_TITLE),
    )
    p.add_argument("inputs", nargs="*", type=Path, help=_("fișiere sau foldere (recursiv)"))
    p.add_argument("-o", "--out", type=Path,
                   help=_("folder destinație (implicit: lângă original)"))
    p.add_argument("--overwrite", action="store_true",
                   help=_("suprascrie rezultatele existente (implicit sunt omise)"))
    p.add_argument("--pdfa", action="store_true", help=_("PDF/A pentru arhivare"))
    p.add_argument("--no-bookmarks", action="store_true",
                   help=_("fără semne de carte din titluri"))
    p.add_argument("--screen", action="store_true",
                   help=_("optimizat pentru ecran (fișier mai mic)"))
    p.add_argument("--ocr", action="store_true",
                   help=_("OCR pe PDF-urile scanate înainte de DOCX"))
    p.add_argument("--pdf-to", choices=tuple(PDF_TARGETS), default="docx",
                   help=_("în ce se convertesc PDF-urile: docx (implicit) sau xlsx (tabele)"))
    p.add_argument("--pdf-mode", choices=PDF_MODES, default="auto",
                   help=_("PDF → DOCX: auto (implicit), layout = aspect fidel (rapoarte, "
                          "tabele), word = text curgător prin Word"))
    p.add_argument("--image-page", choices=settings.IMAGE_PAGES, default="a4",
                   help=_("imagine → PDF: a4 = încadrată pe A4 (implicit), image = pagina "
                          "are mărimea imaginii"))
    p.add_argument("-j", "--jobs", type=int, default=2, metavar="N",
                   help=_("conversii Office simultane, 1-{max} (implicit 2)").format(
                       max=settings.MAX_WORKERS))
    p.add_argument("--timeout", type=int, default=180, metavar="SEC",
                   help=_("limită de timp per fișier (implicit 180)"))
    p.add_argument("--watch", type=Path, metavar="FOLDER",
                   help=_("urmărește folderul și convertește fișierele noi (Ctrl+C oprește); "
                          "rezultatele merg în --out sau în FOLDER\\Convertite"))
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    integration = p.add_argument_group(_("integrare Windows"))
    integration.add_argument("--register-shell", action="store_true",
                             help=_("adaugă „Convertește” în meniul click-dreapta"))
    integration.add_argument("--register-sendto", action="store_true",
                             help=_("adaugă aplicația în „Trimite către”"))
    integration.add_argument("--unregister-shell", action="store_true",
                             help=_("elimină integrările (meniu, Trimite către, pornire "
                                    "cu Windows)"))
    return p


def _integration(args: argparse.Namespace) -> bool:
    if args.unregister_shell:
        shell.unregister_context_menu()
        shell.remove_sendto()
        shell.set_autostart(False)
    if args.register_shell:
        shell.register_context_menu()
    if args.register_sendto:
        shell.create_sendto()
    return bool(args.unregister_shell or args.register_shell or args.register_sendto)


def _report(result: Result) -> None:
    if result.status is Status.RUNNING:
        return
    line = f"[{_marks()[result.status]}] {result.job.src}"
    if result.status is Status.DONE:
        line += f" -> {result.job.dst}"
    if result.message:
        line += f" ({result.message})"
    print(line, flush=True)


def _watch(args: argparse.Namespace, options: ExportOptions) -> int:
    from .watch import FolderWatcher, default_output

    folder_in = args.watch.resolve()
    if not folder_in.is_dir():
        print(_("Folderul urmărit nu există: {folder}").format(folder=folder_in),
              file=sys.stderr)
        return 2
    folder_out = args.out.resolve() if args.out else default_output(folder_in)
    pool = OfficePool(options.workers)

    def convert(jobs: list[Job]) -> list[Result]:
        return Engine(options, listener=_report, pool=pool).run(jobs)

    watcher = FolderWatcher(folder_in, folder_out, convert, interval=3,
                            pdf_target=lambda: options.pdf_target)
    watcher.start()
    print(_("Urmăresc {folder} → {out}. Ctrl+C oprește.").format(folder=folder_in,
                                                                 out=folder_out), flush=True)
    try:
        while watcher.is_alive():
            time.sleep(0.5)
    except KeyboardInterrupt:
        print(_("Oprit."), file=sys.stderr)
    finally:
        watcher.stop()
        watcher.join(timeout=30)
        pool.close()
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")

    i18n.set_language(settings.load().language)
    args = build_parser().parse_args(argv)
    setup_logging(console=True)
    PdfWarningGuard.restore()  # daca o rulare anterioara a cazut

    if _integration(args):
        return 0
    if not args.inputs and args.watch is None:
        build_parser().print_help()
        return 2
    if args.timeout < 10:
        print(_("--timeout trebuie să fie cel puțin 10 secunde."), file=sys.stderr)
        return 2
    if not 1 <= args.jobs <= settings.MAX_WORKERS:
        print(_("--jobs trebuie să fie între 1 și {max}.").format(max=settings.MAX_WORKERS),
              file=sys.stderr)
        return 2

    options = ExportOptions(
        bookmarks=not args.no_bookmarks,
        pdfa=args.pdfa,
        for_screen=args.screen,
        ocr=args.ocr,
        timeout=args.timeout,
        pdf_mode=args.pdf_mode,
        pdf_target=args.pdf_to,
        workers=args.jobs,
        image_page=args.image_page,
    )
    if args.watch is not None:
        return _watch(args, options)

    files = collect_files(args.inputs)
    if not files:
        print(_("Nu am găsit fișiere suportate."), file=sys.stderr)
        return 2
    if args.out:
        args.out.mkdir(parents=True, exist_ok=True)
    jobs = plan_jobs(files, args.out.resolve() if args.out else None, pdf_target=args.pdf_to)
    skip = [] if args.overwrite else [j for j in jobs if j.dst.exists()]

    engine = Engine(options, listener=_report)
    try:
        results = engine.run(jobs, skip=skip)
    except KeyboardInterrupt:
        engine.abort()
        print(_("Întrerupt."), file=sys.stderr)
        return 130

    failed = sum(r.status is Status.FAILED for r in results)
    done = sum(r.status is Status.DONE for r in results)
    print("\n" + _("{done} convertite, {skipped} omise, {failed} cu erori.").format(
        done=done, skipped=len(skip), failed=failed))
    return 1 if failed else 0
