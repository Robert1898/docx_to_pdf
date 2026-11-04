"""Ruleaza un lot de conversii; folosit atat de interfata grafica, cat si de linia de comanda."""

from __future__ import annotations

import logging
import tempfile
import threading
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from . import libreoffice, pdflayout, pdftools
from .core import ExportOptions, Job, Kind
from .office import OFFICE_FOR_KIND, ConversionError, OfficeSession, describe_error

log = logging.getLogger(__name__)


class Status(Enum):
    PENDING = "În așteptare"
    RUNNING = "Se convertește…"
    DONE = "Gata"
    FAILED = "Eroare"
    SKIPPED = "Omis (există deja)"
    CANCELLED = "Anulat"


@dataclass
class Result:
    job: Job
    status: Status
    message: str = ""


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
        if kind is Kind.PDF:
            return True  # „aspect fidel” nu are nevoie de Word
        return self.office.get(kind, False) or self.soffice is not None


Listener = Callable[[Result], None]


class Engine:
    def __init__(
        self,
        options: ExportOptions,
        listener: Listener | None = None,
        backends: Backends | None = None,
    ) -> None:
        self.options = options
        self.listener = listener
        self.backends = backends or Backends.detect()
        self.cancel_event = threading.Event()
        self._session: OfficeSession | None = None

    def cancel(self) -> None:
        """Opreste dupa fisierul curent."""
        self.cancel_event.set()

    def abort(self) -> None:
        """Opreste imediat, inclusiv aplicatia Office care lucreaza acum."""
        self.cancel_event.set()
        if self._session is not None:
            self._session.kill_all()

    def _emit(self, result: Result) -> None:
        if self.listener is not None:
            try:
                self.listener(result)
            except Exception:  # un listener stricat nu are voie sa opreasca lotul
                log.exception("Listener esuat")

    def run(self, jobs: Sequence[Job], skip: Sequence[Job] = ()) -> list[Result]:
        skipped = {id(job) for job in skip}
        results: list[Result] = []
        with OfficeSession() as session:
            self._session = session
            try:
                for job in jobs:
                    if id(job) in skipped:
                        result = Result(job, Status.SKIPPED)
                    elif self.cancel_event.is_set():
                        result = Result(job, Status.CANCELLED)
                    else:
                        self._emit(Result(job, Status.RUNNING))
                        result = self._convert_one(session, job)
                    results.append(result)
                    self._emit(result)
            finally:
                self._session = None
        return results

    def _convert_one(self, session: OfficeSession, job: Job) -> Result:
        log.info("Convertesc %s -> %s", job.src, job.dst)
        try:
            note = self._convert(session, job)
        except Exception as exc:
            if self.cancel_event.is_set():
                return Result(job, Status.CANCELLED)
            message = str(exc) if isinstance(exc, ConversionError | pdftools.PdfError) else (
                describe_error(exc)
            )
            log.warning("Eroare la %s: %s", job.src, message, exc_info=not isinstance(
                exc, ConversionError | pdftools.PdfError))
            return Result(job, Status.FAILED, message)
        return Result(job, Status.DONE, note)

    def _convert(self, session: OfficeSession, job: Job) -> str:
        if not self.backends.can_convert(job.kind):
            raise ConversionError("Nu am găsit Microsoft Office sau LibreOffice instalat.")

        job.dst.parent.mkdir(parents=True, exist_ok=True)

        if job.kind is Kind.PDF:
            return self._pdf_to_docx(session, job)
        if not self.backends.office.get(job.kind, False):
            assert self.backends.soffice is not None
            libreoffice.convert(self.backends.soffice, job, self.options)
            return "prin LibreOffice"
        return session.convert(job, self.options)

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
            src = job.src
            image_only = False
            if pdftools.is_scanned(src):
                if self.options.ocr and self.backends.tesseract is not None:
                    src = Path(tmp) / job.src.name
                    pdftools.ocr(job.src, src, cancel=self.cancel_event)
                    notes.append("cu OCR")
                else:
                    image_only = True
                    if self.options.ocr:
                        notes.append("PDF scanat, dar Tesseract nu e instalat: fără OCR")
            if mode == "auto":
                mode = ("word" if has_word and (image_only or not pdflayout.looks_tabular(src))
                        else "layout")
            if mode == "layout" and image_only:
                # fara text de asezat, aspectul fidel ar da pagini goale
                if not has_word:
                    raise ConversionError(
                        "PDF-ul e scanat (conține doar imagini). Activează OCR (necesită "
                        "Tesseract) ca să obții text, sau instalează Microsoft Word.")
                mode = "word"
                notes.append("PDF scanat: prin Word, care păstrează imaginile")
            if mode == "word" and not has_word:
                mode = "layout"
                notes.append("Word lipsește")

            if mode == "layout":
                pdflayout.convert(src, job.dst)
                notes.insert(0, "aspect fidel")
            else:
                notes.insert(0, "prin Word")
                notes.append(session.convert(job, self.options, src=src))
        return "; ".join(filter(None, notes))
