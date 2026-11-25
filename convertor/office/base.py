"""Partea comuna: erorile afisate utilizatorului si aplicatia Office generica (pornire,
Quit, oprire fortata)."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ..core import ExportOptions, Job
from ..i18n import _
from .process import _kill, _open_process, _pids, _wait_or_terminate
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
        new = _pids(self.image) - before
        # PowerPoint permite o singura instanta: daca era deja deschis, il impartim
        self.owned = bool(new)
        self.pid = next(iter(new)) if len(new) == 1 else None
        self.timed_out = False
        log.info("%s pornit (pid=%s, propriu=%s)", self.name, self.pid, self.owned)
        self.configure()

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
            if self.owned and not self.timed_out and self.prepare_quit():
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

    def kill(self) -> bool:
        """Opreste fortat procesul (apelat din alt fir). False daca nu e al nostru."""
        if not (self.owned and self.pid):
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
