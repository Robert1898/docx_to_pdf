"""Sesiunea de conversie: aplicatiile Office pornite pe durata unui lot, limita de timp
per fisier si raspunsul la dialoguri."""

from __future__ import annotations

import contextlib
import logging
import threading
from pathlib import Path

from ..core import ExportOptions, Job, Kind
from ..i18n import _
from .base import OFFICE_DIED, ConversionError, ConversionTimeout, OfficeApp, describe_error
from .dialogs import DialogWatcher
from .excel import ExcelApp
from .powerpoint import POWERPOINT_LOCK, PowerPointApp
from .process import _cancel_call, _enable_call_cancellation
from .registry import ResiliencyCleaner
from .word import WordApp

log = logging.getLogger(__package__)  # un singur jurnal pentru tot pachetul

ABANDON_CLEANUP_SECONDS = 3.0  # cat asteptam o instanta abandonata (poate fi inca blocata)

OFFICE_FOR_KIND: dict[Kind, type[OfficeApp]] = {
    Kind.WORD: WordApp,
    Kind.PDF: WordApp,
    Kind.EXCEL: ExcelApp,
    Kind.POWERPOINT: PowerPointApp,
}


class OfficeSession:
    """Tine aplicatiile Office pornite pe durata unui lot. Se foloseste dintr-un singur fir."""

    def __init__(self) -> None:
        self._apps: dict[type[OfficeApp], OfficeApp] = {}
        self._lock = threading.Lock()
        self._com_ready = False
        # aplicatia si fisierul la care lucreaza firul, plus evenimentul „apel terminat”
        self._current: tuple[OfficeApp, Path, threading.Event] | None = None
        self._cancel_lock = threading.Lock()  # nicio anulare dupa ce apelul s-a terminat
        self._thread_id: int | None = None  # firul care face apelurile COM

    def __enter__(self) -> OfficeSession:
        import pythoncom

        pythoncom.CoInitialize()
        self._com_ready = True
        _enable_call_cancellation()
        self._thread_id = threading.get_native_id()
        return self

    def __exit__(self, *exc: object) -> None:
        import pythoncom

        for app in list(self._apps.values()):
            app.quit()
        self._apps.clear()
        if self._com_ready:
            pythoncom.CoUninitialize()
            self._com_ready = False

    def _app(self, cls: type[OfficeApp]) -> OfficeApp:
        with self._lock:
            if cls not in self._apps:
                self._apps[cls] = cls()
            return self._apps[cls]

    def on_idle(self) -> None:
        """Chemat cand nu mai e nimic de convertit (la sfarsitul unui lot). Inchide PowerPoint:
        are o singura instanta, deci orice prezentare deschisa de utilizator cat timp il tinem
        pornit ar ajunge in procesul nostru. Word si Excel raman pentru lotul urmator."""
        with self._lock:
            app = self._apps.pop(PowerPointApp, None)
        if app is not None:
            app.quit()

    def kill_all(self) -> None:
        """Oprire de urgenta (de exemplu la inchiderea ferestrei). Sigur din orice fir."""
        with self._lock:
            apps = list(self._apps.values())
            current = self._current
        if current is not None and isinstance(current[0], WordApp):
            ResiliencyCleaner.remember(current[1])
        for app in apps:
            if app.kill() or current is None or app is not current[0]:
                continue
            # Nu poate fi oprit (documente straine sau instanta altcuiva): anulam doar apelul
            # in curs, ca firul (si POWERPOINT_LOCK) sa fie eliberate. Fara asteptare aici.
            log.warning("Anulare: %s nu poate fi oprit; anulez doar apelul in curs", app.name)
            threading.Thread(target=self._abandon, args=(app, current[2]), daemon=True,
                             name="cancel-call").start()
        ResiliencyCleaner.clean()

    def convert(self, job: Job, options: ExportOptions, src: Path | None = None) -> str:
        """Converteste un fisier. Intoarce o nota (de ex. dialogurile la care am raspuns)."""
        cls = OFFICE_FOR_KIND[job.kind]
        # PowerPoint are o singura instanta pe calculator: firele paralele il folosesc pe rand
        with POWERPOINT_LOCK if cls is PowerPointApp else contextlib.nullcontext():
            return self._convert(cls, job, options, src)

    def _convert(self, cls: type[OfficeApp], job: Job, options: ExportOptions,
                 src: Path | None) -> str:
        app = self._app(cls)
        opened = src or job.src
        timer: threading.Timer | None = None
        watcher: DialogWatcher | None = None
        done = threading.Event()
        try:
            app.ensure()
            timer = threading.Timer(options.timeout, self._on_timeout, args=(app, opened, done))
            timer.daemon = True
            if app.owned and app.pid:
                watcher = DialogWatcher(app.pid)
                watcher.start()
            with self._lock:
                self._current = (app, opened, done)
            timer.start()
            app.convert(job, options, opened)
        except Exception as exc:
            dialog = (" " + _("Dialog: „{text}”").format(text=watcher.last_text)
                      if watcher and watcher.last_text else "")
            if app.timed_out:
                app.reset()
                raise ConversionTimeout(
                    _("A depășit limita de {seconds} s; {app} a fost oprit și va fi repornit.")
                    .format(seconds=options.timeout, app=app.name) + dialog
                ) from exc
            if app.abandoned:
                self._drop_abandoned(app, done)
                raise ConversionTimeout(
                    _("A depășit limita de {seconds} s; {app} are documente deschise de "
                      "utilizator, așa că nu l-am oprit.")
                    .format(seconds=options.timeout, app=app.name) + dialog
                ) from exc
            message = describe_error(exc)
            if message == _(OFFICE_DIED):
                app.reset()
            raise ConversionError(message + dialog) from exc
        finally:
            self._stop_cancelling(done)
            if timer is not None:
                timer.cancel()
            if watcher is not None:
                watcher.stop()
            with self._lock:
                self._current = None
            if isinstance(app, PowerPointApp) and not app.owned:
                app.reset()  # instanta altcuiva: nu tinem referinta intre fisiere
        if app.abandoned:  # anularea a venit prea tarziu: fisierul e gata, instanta nu
            self._drop_abandoned(app, done)
        if watcher and watcher.answers:
            return _("am răspuns automat la {count} dialog(uri) {app}").format(
                count=len(watcher.answers), app=app.name)
        return ""

    def _on_timeout(self, app: OfficeApp, opened: Path, done: threading.Event) -> None:
        if isinstance(app, WordApp):
            ResiliencyCleaner.remember(opened)
        if app.kill():
            log.warning("Timeout la %s: am oprit %s (pid %s)", opened, app.name, app.pid)
            if isinstance(app, WordApp):
                ResiliencyCleaner.clean()  # inainte ca Word sa fie repornit
            return
        # Procesul nu e al nostru sau are documentele utilizatorului: nu il oprim. Anulam doar
        # apelul nostru blocat, ca fisierul sa fie marcat esuat si firul sa poata continua.
        log.warning("Timeout la %s, dar %s e deschis de utilizator sau are documentele lui; nu il "
                    "opresc, renunt doar la fisier", opened, app.name)
        self._abandon(app, done)

    def _abandon(self, app: OfficeApp, done: threading.Event) -> None:
        """Renunta la instanta fara s-o opreasca: anuleaza apelul COM blocat al firului nostru
        (din alt fir), pana cand firul iese din conversie."""
        app.abandoned = True
        self._cancel_until(done)

    def _cancel_until(self, done: threading.Event) -> None:
        while self._thread_id is not None:
            with self._cancel_lock:
                if done.is_set():
                    return
                _cancel_call(self._thread_id)  # apelul blocat poate incepe abia dupa cerere
            done.wait(0.5)

    def _stop_cancelling(self, done: threading.Event) -> None:
        with self._cancel_lock:  # dupa asta nu mai pleaca nicio anulare pentru acest apel
            done.set()

    def _drop_abandoned(self, app: OfficeApp, done: threading.Event) -> None:
        """Ultimele apeluri catre o instanta abandonata (de ex. setarile globale refacute), cu
        limita de timp, apoi renuntam la ea. Procesul ramane pornit."""
        self._stop_cancelling(done)
        finished = threading.Event()
        canceller = threading.Timer(ABANDON_CLEANUP_SECONDS, self._cancel_until, args=(finished,))
        canceller.daemon = True
        canceller.start()
        try:
            app.release_abandoned()
        except Exception as exc:  # intentionat larg: instanta poate fi inca blocata
            log.debug("Eliberare %s: %s", app.name, describe_error(exc))
        finally:
            self._stop_cancelling(finished)
            canceller.cancel()
        app.reset()
