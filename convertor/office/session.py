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
from .registry import ResiliencyCleaner
from .word import WordApp

log = logging.getLogger(__package__)  # un singur jurnal pentru tot pachetul

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
        self._current: tuple[OfficeApp, Path] | None = None

    def __enter__(self) -> OfficeSession:
        import pythoncom

        pythoncom.CoInitialize()
        self._com_ready = True
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

    def kill_all(self) -> None:
        """Oprire de urgenta (de exemplu la inchiderea ferestrei). Sigur din orice fir."""
        with self._lock:
            apps = list(self._apps.values())
            current = self._current
        if current is not None and isinstance(current[0], WordApp):
            ResiliencyCleaner.remember(current[1])
        for app in apps:
            app.kill()
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
        try:
            app.ensure()
            timer = threading.Timer(options.timeout, self._on_timeout, args=(app, opened))
            timer.daemon = True
            if app.owned and app.pid:
                watcher = DialogWatcher(app.pid)
                watcher.start()
            with self._lock:
                self._current = (app, opened)
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
            message = describe_error(exc)
            if message == _(OFFICE_DIED):
                app.reset()
            raise ConversionError(message + dialog) from exc
        finally:
            if timer is not None:
                timer.cancel()
            if watcher is not None:
                watcher.stop()
            with self._lock:
                self._current = None
        if watcher and watcher.answers:
            return _("am răspuns automat la {count} dialog(uri) {app}").format(
                count=len(watcher.answers), app=app.name)
        return ""

    @staticmethod
    def _on_timeout(app: OfficeApp, opened: Path) -> None:
        if isinstance(app, WordApp):
            ResiliencyCleaner.remember(opened)
        if app.kill():
            log.warning("Timeout la %s: am oprit %s (pid %s)", opened, app.name, app.pid)
            if isinstance(app, WordApp):
                ResiliencyCleaner.clean()  # inainte ca Word sa fie repornit
        else:
            log.warning(
                "Timeout la %s, dar %s era deschis de utilizator si nu il opresc", opened, app.name
            )
