"""Partea comuna: erorile afisate utilizatorului si aplicatia Office generica (pornire,
Quit, oprire fortata)."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ..core import ExportOptions, Job
from ..i18n import _
from .process import _kill, _open_process, _pid_of_window, _pids, _wait_or_terminate
from .registry import progid_available

log = logging.getLogger(__package__)  # un singur jurnal pentru tot pachetul

# Parola gresita intentionat: un document neprotejat o ignora, unul protejat da eroare
# imediat in loc sa deschida un dialog invizibil de parola.
NO_PASSWORD = "\u0001convertor-no-password"

MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3  # macro-urile nu ruleaza la deschidere

# coduri HRESULT care inseamna ca procesul Office a murit sau nu mai raspunde
OFFICE_DIED = "Aplicația Office s-a închis neașteptat."

DEAD_SERVER_CODES = {
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2147023170,  # RPC_S_CALL_FAILED
    -2147417848,  # RPC_E_DISCONNECTED
    -2147417851,  # RPC_E_SERVERFAULT
    -2147418111,  # RPC_E_CALL_REJECTED
}


class ConversionError(Exception):
    """Eroare cu mesaj prietenos, afisabil direct utilizatorului."""


class ConversionTimeout(ConversionError):
    pass


def describe_error(exc: BaseException) -> str:
    try:
        import pywintypes
    except ImportError:  # pragma: no cover - doar pe sisteme fara pywin32
        return str(exc)
    if isinstance(exc, pywintypes.com_error):
        hresult, message, excepinfo, _arg = exc.args
        detail = (excepinfo[2] if excepinfo and excepinfo[2] else message) or str(exc)
        low = detail.lower()
        if "password" in low or "parol" in low:
            return _("Documentul este protejat cu parolă.")
        if hresult in DEAD_SERVER_CODES or (excepinfo and excepinfo[5] in DEAD_SERVER_CODES):
            return _(OFFICE_DIED)
        return " ".join(detail.split())
    return str(exc) or exc.__class__.__name__


class OfficeApp:
    progid = ""
    image = ""
    name = ""
    # Cat asteptam dupa Quit ca procesul pornit de noi sa dispara, inainte sa-l oprim.
    # None: nu asteptam (Word poate inca salva Normal.dotm, nu are voie sa fie oprit).
    quit_grace: float | None = None

    def __init__(self) -> None:
        self.app: Any = None
        self.pid: int | None = None
        self.owned = False
        self.timed_out = False
        # La timeout nu am putut opri procesul: apelul blocat a fost anulat, procesul ramane
        self.abandoned = False

    @classmethod
    def available(cls) -> bool:
        return progid_available(cls.progid)

    def ensure(self) -> Any:
        if self.app is None:
            self.start()
        return self.app

    def start(self) -> None:
        import win32com.client

        before = _pids(self.image)
        self.app = win32com.client.DispatchEx(self.progid)
        self.owned, self.pid = self._identify(before, _pids(self.image))
        self.timed_out = self.abandoned = False
        log.info("%s pornit (pid=%s, propriu=%s)", self.name, self.pid, self.owned)
        self.configure()

    def _window_handle(self) -> int | None:
        """Handle-ul ferestrei principale (ascunse) a instantei, din modelul COM. None daca
        aplicatia nu il expune; atunci pid-ul vine din tasklist."""
        return None

    def _window_pid(self) -> int | None:
        try:
            hwnd = self._window_handle()
            return _pid_of_window(hwnd) if hwnd else None
        except Exception as exc:  # intentionat larg: COM si Win32 pot refuza in multe feluri
            log.debug("Fereastra %s: %s", self.name, describe_error(exc))
            return None

    def _identify(self, before: set[int], after: set[int]) -> tuple[bool, int | None]:
        """Decide daca instanta e pornita de noi (Quit la final) si pid-ul ei (oprire fortata).

        Pid-ul din fereastra aplicatiei e sigur. Diferenta dintre doua liste tasklist nu e: daca
        utilizatorul porneste aceeasi aplicatie in acelasi timp, apar doua procese noi. De aceea,
        fara fereastra, oprirea fortata e permisa doar cand a aparut exact un proces nou.
        PowerPoint permite o singura instanta: daca era deja deschis, il impartim (nu e al nostru).
        """
        new = after - before
        window_pid = self._window_pid()
        if window_pid is not None:
            if len(new) == 1 and window_pid not in new:
                log.warning("%s: pid-ul din fereastra (%s) difera de cel din tasklist (%s); il "
                            "folosesc pe cel din fereastra", self.name, window_pid, next(iter(new)))
            if window_pid in before:
                log.warning("%s: fereastra e a unei instante deja pornite (pid %s); nu e a noastra",
                            self.name, window_pid)
                return False, window_pid
            log.info("%s: pid %s din fereastra aplicatiei", self.name, window_pid)
            return True, window_pid
        if len(new) == 1:
            pid = next(iter(new))
            log.info("%s: pid %s din tasklist (fereastra indisponibila)", self.name, pid)
            return True, pid
        if new:
            log.warning("%s: %d procese noi %s, nu stiu care e al nostru; fara oprire fortata, "
                        "doar Quit", self.name, len(new), sorted(new))
            return True, None
        log.info("%s: niciun proces nou in tasklist; instanta exista deja, nu e a noastra",
                 self.name)
        return False, None

    def configure(self) -> None:
        pass

    def prepare_quit(self) -> bool:
        """Chemat inainte de Quit, doar pentru o instanta pornita de noi. False: instanta nu
        trebuie inchisa (si nici oprita fortat)."""
        return True

    def quit(self) -> None:
        if self.app is None:
            return
        handle = None
        try:
            if self.owned and not (self.timed_out or self.abandoned) and self.prepare_quit():
                if self.quit_grace is not None and self.pid:
                    handle = _open_process(self.pid)
                self.app.Quit()
        except Exception as exc:  # aplicatia poate fi deja moarta
            log.debug("Quit %s: %s", self.name, exc)
        finally:
            self.app = None  # eliberam referinta COM: altfel procesul nu se poate inchide
        grace = self.quit_grace
        if handle is not None and grace is not None and _wait_or_terminate(handle, grace):
            log.info("%s (pid %s) inca rula la %.0f s dupa Quit; l-am oprit",
                     self.name, self.pid, grace)

    def release_abandoned(self) -> None:
        """Ultimele apeluri inainte sa renuntam la o instanta pe care nu o putem opri (de ex.
        setari globale de refacut). Instanta poate fi inca blocata: apelantul pune o limita."""

    def has_foreign_documents(self) -> bool:
        """Instanta are documente deschise de altcineva? Apelat din alt fir, cat timp aplicatia
        poate fi blocata, deci fara COM."""
        return False

    def kill(self) -> bool:
        """Opreste fortat procesul (apelat din alt fir). False daca nu e al nostru."""
        if not (self.owned and self.pid):
            return False
        if self.has_foreign_documents():
            log.warning("%s (pid %s) are documente deschise de altcineva; nu il opresc",
                        self.name, self.pid)
            return False
        self.timed_out = True
        _kill(self.pid)
        return True

    def reset(self) -> None:
        """Dupa un timeout sau un crash: uitam instanta, se reporneste la nevoie."""
        self.app = None
        self.pid = None

    def convert(self, job: Job, options: ExportOptions, src: Path) -> None:
        raise NotImplementedError
