"""O singura fereastra a aplicatiei per utilizator.

La click-dreapta pe mai multe fisiere, Explorer porneste aplicatia o data pentru fiecare
fisier. Prima instanta asculta pe un „named pipe”; urmatoarele ii trimit fisierele si se
inchid, iar fereastra existenta le primeste pe toate in aceeasi lista.
"""

from __future__ import annotations

import contextlib
import logging
import os
import re
import threading
import time
from collections.abc import Callable
from multiprocessing.connection import Client, Listener

from .. import APP_ID

log = logging.getLogger(__name__)

AUTHKEY = APP_ID.encode()
Message = dict  # {"files": [...], "auto": bool}


def pipe_name() -> str:
    user = re.sub(r"[^A-Za-z0-9_.-]", "_", os.environ.get("USERNAME", "user"))
    return rf"\\.\pipe\{APP_ID}-{user}"


def send_to_running(message: Message, address: str | None = None) -> bool:
    """Trimite fisierele instantei deja pornite. False daca nu exista una."""
    try:
        with Client(address or pipe_name(), family="AF_PIPE", authkey=AUTHKEY) as conn:
            conn.send(message)
        return True
    except OSError:  # nu exista nicio instanta (sau tocmai se inchide)
        return False


def try_listen(address: str | None = None) -> Listener | None:
    """Devine instanta principala. None daca alta instanta asculta deja."""
    try:
        return Listener(address or pipe_name(), family="AF_PIPE", authkey=AUTHKEY)
    except OSError:  # pipe-ul exista deja (FILE_FLAG_FIRST_PIPE_INSTANCE)
        return None


def claim(message: Message, attempts: int = 25, address: str | None = None) -> Listener | None:
    """Trimite catre instanta existenta sau devine instanta principala.

    Intoarce Listener-ul daca aceasta instanta trebuie sa porneasca fereastra, altfel None
    (fisierele au fost predate). La pornirea simultana a mai multor instante, una castiga
    pipe-ul, iar celelalte reincearca pana o gasesc."""
    for _attempt in range(attempts):
        if send_to_running(message, address):
            return None
        listener = try_listen(address)
        if listener is not None:
            return listener
        time.sleep(0.2)
    log.warning("Nu am putut contacta instanta principala; pornesc o fereastra noua")
    return try_listen(address) or _NullListener()  # type: ignore[return-value]


class _NullListener:
    """Folosit doar daca pipe-ul nu poate fi nici creat, nici contactat."""

    def accept(self) -> None:
        threading.Event().wait()

    def close(self) -> None:
        pass


class InstanceServer(threading.Thread):
    """Primeste mesajele instantelor pornite ulterior (in fundal)."""

    def __init__(self, listener: Listener, on_message: Callable[[Message], None]) -> None:
        super().__init__(name="instance-server", daemon=True)
        self.listener = listener
        self.on_message = on_message
        self._stopped = False

    def run(self) -> None:
        while not self._stopped:
            try:
                with self.listener.accept() as conn:
                    message = conn.recv()
            except (OSError, EOFError) as exc:
                if self._stopped:
                    return
                log.debug("Conexiune instanta: %s", exc)
                continue
            except Exception as exc:  # mesaj stricat sau cheie gresita
                log.warning("Mesaj respins de la alta instanta: %s", exc)
                continue
            if self._stopped:
                return
            if isinstance(message, dict):
                self.on_message(message)

    def stop(self) -> None:
        self._stopped = True
        # accept() asteapta blocant: o conexiune proprie il trezeste ca pipe-ul sa fie eliberat
        address = getattr(self.listener, "address", None)
        if address and self.is_alive():
            send_to_running({}, address)
        self.join(timeout=2)
        with contextlib.suppress(OSError):
            self.listener.close()
