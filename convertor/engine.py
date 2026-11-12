"""Ruleaza un lot de conversii; folosit atat de interfata grafica, cat si de linia de comanda.

Conversiile merg in paralel pe un pool de fire de lucru. Fiecare fir are propriile aplicatii
Office (COM tine obiectele legate de firul care le-a creat), pe care le pastreaza intre
loturi: urmatoarea conversie nu mai asteapta pornirea lui Word. Dupa cateva minute fara
activitate, firul inchide Office si se opreste.
"""

from __future__ import annotations

import logging
import queue
import tempfile
import threading
from collections.abc import Callable, Sequence
from concurrent.futures import Future
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from . import libreoffice, pdflayout, pdftools
from .core import ExportOptions, Job, Kind
from .i18n import _
from .office import OFFICE_FOR_KIND, ConversionError, OfficeSession, describe_error

log = logging.getLogger(__name__)

IDLE_TIMEOUT = 300  # secunde dupa care un fir inchide Office
# un fir nou (cu Office-ul lui, ~2 s de pornire) merita doar daca are destule fisiere de lucrat
JOBS_PER_WORKER = 4


class Status(Enum):
    PENDING = "În așteptare"
    RUNNING = "Se convertește…"
    DONE = "Gata"
    FAILED = "Eroare"
    SKIPPED = "Omis (există deja)"
    CANCELLED = "Anulat"

    @property
    def text(self) -> str:
        return _(self.value)


@dataclass
class Result:
    job: Job
    status: Status
    message: str = ""
    seconds: float = 0.0


@dataclass
class Backends:
    """Ce poate folosi aplicatia pe calculatorul curent."""

    office: dict[Kind, bool] = field(default_factory=dict)
    soffice: Path | None = None
    tesseract: Path | None = None

    @classmethod
    def detect(cls) -> Backends:
        return cls(
            office={kind: app.available() for kind, app in OFFICE_FOR_KIND.items()},
            soffice=libreoffice.find_soffice(),
            tesseract=pdftools.find_tesseract(),
        )

    def can_convert(self, kind: Kind) -> bool:
        if kind in (Kind.PDF, Kind.IMAGE):
            return True  # „aspect fidel”, Excel si imaginile nu au nevoie de Office
        return self.office.get(kind, False) or self.soffice is not None


# --- pool -----------------------------------------------------------------------------------


class OfficePool:
    """Fire de lucru permanente, fiecare cu sesiunea lui Office."""

    def __init__(self, size: int, idle_timeout: float = IDLE_TIMEOUT) -> None:
        self.size = max(1, size)
        self.idle_timeout = idle_timeout
        self._tasks: queue.Queue[tuple[Future, Callable[[OfficeSession], object]] | None] = (
            queue.Queue())
        self._threads: list[threading.Thread] = []
        self._sessions: set[OfficeSession] = set()
        self._lock = threading.Lock()
        self._closed = False

    def _ensure_workers(self) -> None:
        """Porneste fire dupa cat de mare e coada: un lot mic ramane pe o singura instanta
        Office, unul mare se imparte pe pana la `size` instante."""
        wanted = min(self.size, max(1, -(-self._tasks.qsize() // JOBS_PER_WORKER)))
        with self._lock:
            self._threads = [t for t in self._threads if t.is_alive()]
            while len(self._threads) < wanted:
                thread = threading.Thread(target=self._worker, name="office-worker", daemon=True)
                thread.start()
                self._threads.append(thread)

    def submit(self, fn: Callable[[OfficeSession], object]) -> Future:
        if self._closed:
            raise RuntimeError("pool inchis")
        future: Future = Future()
        self._tasks.put((future, fn))
        self._ensure_workers()
        return future

    def _worker(self) -> None:
        with OfficeSession() as session:
            with self._lock:
                self._sessions.add(session)
            try:
                while True:
                    try:
                        item = self._tasks.get(timeout=self.idle_timeout)
                    except queue.Empty:
                        return  # fara activitate: inchidem Office (la iesirea din „with”)
                    if item is None:
                        return
                    future, fn = item
                    if not future.set_running_or_notify_cancel():
                        continue
                    try:
                        future.set_result(fn(session))
                    except BaseException as exc:
                        future.set_exception(exc)
            finally:
                with self._lock:
                    self._sessions.discard(session)

    def kill_all(self) -> None:
        """Opreste imediat aplicatiile Office pornite de pool (anulare urgenta)."""
        with self._lock:
            sessions = list(self._sessions)
        for session in sessions:
            session.kill_all()

    def close(self, wait: float = 15) -> None:
        """Inchide Office in toate firele (la iesirea din aplicatie)."""
        self._closed = True
        with self._lock:
            threads = [t for t in self._threads if t.is_alive()]
        for _thread in threads:
            self._tasks.put(None)
        for thread in threads:
            thread.join(timeout=wait)


# --- motor ----------------------------------------------------------------------------------

Listener = Callable[[Result], None]


class Engine:
    def __init__(
        self,
        options: ExportOptions,
        listener: Listener | None = None,
        backends: Backends | None = None,
        pool: OfficePool | None = None,
    ) -> None:
        self.options = options
        self.listener = listener
        self.backends = backends or Backends.detect()
        self.cancel_event = threading.Event()
        self.pool = pool
        self._active_pool: OfficePool | None = None

    def cancel(self) -> None:
        """Opreste dupa fisierele care se convertesc acum."""
        self.cancel_event.set()

    def abort(self) -> None:
        """Opreste imediat, inclusiv aplicatiile Office care lucreaza acum."""
        self.cancel_event.set()
        if self._active_pool is not None:
            self._active_pool.kill_all()

    def _emit(self, result: Result) -> None:
        if self.listener is not None:
            try:
                self.listener(result)
            except Exception:  # un listener stricat nu are voie sa opreasca lotul
                log.exception("Listener esuat")

    def run(self, jobs: Sequence[Job], skip: Sequence[Job] = ()) -> list[Result]:
        skipped = {id(job) for job in skip}
        own_pool = self.pool is None
        pool = self.pool or OfficePool(self.options.workers, idle_timeout=1)
        self._active_pool = pool
        results: dict[int, Result] = {}
        futures: list[tuple[Job, Future]] = []
        try:
            for job in jobs:
                if id(job) in skipped:
                    results[id(job)] = Result(job, Status.SKIPPED)
                    self._emit(results[id(job)])
                else:
                    futures.append((job, pool.submit(
                        lambda session, job=job: self._process(session, job))))
            for job, future in futures:
                try:
                    results[id(job)] = future.result()
                except Exception as exc:  # nu ar trebui sa ajunga aici; _process prinde tot
                    log.exception("Fir de lucru esuat")
                    results[id(job)] = Result(job, Status.FAILED, describe_error(exc))
                    self._emit(results[id(job)])
        finally:
            self._active_pool = None
            if own_pool:
                pool.close()
        return [results[id(job)] for job in jobs]

    def _process(self, session: OfficeSession, job: Job) -> Result:
        import time

        if self.cancel_event.is_set():
            result = Result(job, Status.CANCELLED)
        else:
            self._emit(Result(job, Status.RUNNING))
            start = time.monotonic()
            result = self._convert_one(session, job)
            result.seconds = time.monotonic() - start
        self._emit(result)
        return result

    def _convert_one(self, session: OfficeSession, job: Job) -> Result:
        log.info("Convertesc %s -> %s", job.src, job.dst)
        try:
            note = self._convert(session, job)
        except Exception as exc:
            if self.cancel_event.is_set():
                return Result(job, Status.CANCELLED)
            friendly = isinstance(exc, ConversionError | pdftools.PdfError)
            message = str(exc) if friendly else describe_error(exc)
            log.warning("Eroare la %s: %s", job.src, message, exc_info=not friendly)
            return Result(job, Status.FAILED, message)
        return Result(job, Status.DONE, note)

    def _convert(self, session: OfficeSession, job: Job) -> str:
        if not self.backends.can_convert(job.kind):
            raise ConversionError(_("Nu am găsit Microsoft Office sau LibreOffice instalat."))

        job.dst.parent.mkdir(parents=True, exist_ok=True)

        if job.kind is Kind.IMAGE:
            pages = pdftools.images_to_pdf([job.src], job.dst, page=self.options.image_page)
            return _("{count} pagini").format(count=pages) if pages > 1 else ""
        if job.kind is Kind.PDF:
            if job.dst.suffix.lower() == ".xlsx":
                return self._pdf_to_excel(job)
            return self._pdf_to_docx(session, job)
        if not self.backends.office.get(job.kind, False):
            assert self.backends.soffice is not None
            libreoffice.convert(self.backends.soffice, job, self.options)
            return _("prin LibreOffice")
        return session.convert(job, self.options)

    def _searchable(self, job: Job, tmp: str, notes: list[str]) -> tuple[Path, bool]:
        """PDF-ul de lucrat (cu OCR daca e scanat si se cere) si daca a ramas doar imagine."""
        if not pdftools.is_scanned(job.src):
            return job.src, False
        if self.options.ocr and self.backends.tesseract is not None:
            src = Path(tmp) / job.src.name
            pdftools.ocr(job.src, src, cancel=self.cancel_event)
            notes.append(_("cu OCR"))
            return src, False
        if self.options.ocr:
            notes.append(_("PDF scanat, dar Tesseract nu e instalat: fără OCR"))
        return job.src, True

    def _pdf_to_excel(self, job: Job) -> str:
        from . import pdftable

        notes: list[str] = []
        with tempfile.TemporaryDirectory(prefix="convertor-") as tmp:
            src, image_only = self._searchable(job, tmp, notes)
            if image_only:
                raise ConversionError(_("PDF-ul e scanat (conține doar imagini). Activează OCR "
                                        "(necesită Tesseract) ca să obții tabele."))
            rows = pdftable.convert(src, job.dst)
        notes.insert(0, _("{rows} rânduri").format(rows=rows))
        return "; ".join(notes)

    def _pdf_to_docx(self, session: OfficeSession, job: Job) -> str:
        """Alege intre „aspect fidel” (rapoarte pe coloane) si Word (text curgator).

        * PDF scanat: Word, care pastreaza imaginile; cu OCR inainte, daca e cerut;
        * modul ales explicit e respectat (Word doar daca e instalat);
        * automat: rapoartele pe coloane merg pe aspect fidel, restul prin Word.
        """
        has_word = self.backends.office.get(Kind.PDF, False)
        mode = self.options.pdf_mode
        notes: list[str] = []
        with tempfile.TemporaryDirectory(prefix="convertor-") as tmp:
            src, image_only = self._searchable(job, tmp, notes)
            if mode == "auto":
                mode = ("word" if has_word and (image_only or not pdflayout.looks_tabular(src))
                        else "layout")
            if mode == "layout" and image_only:
                # fara text de asezat, aspectul fidel ar da pagini goale
                if not has_word:
                    raise ConversionError(_(
                        "PDF-ul e scanat (conține doar imagini). Activează OCR (necesită "
                        "Tesseract) ca să obții text, sau instalează Microsoft Word."))
                mode = "word"
                notes.append(_("PDF scanat: prin Word, care păstrează imaginile"))
            if mode == "word" and not has_word:
                mode = "layout"
                notes.append(_("Word lipsește"))

            if mode == "layout":
                pdflayout.convert(src, job.dst)
                notes.insert(0, _("aspect fidel"))
            else:
                notes.insert(0, _("prin Word"))
                notes.append(session.convert(job, self.options, src=src))
        return "; ".join(filter(None, notes))
