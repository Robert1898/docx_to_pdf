"""Procesele Office pornite de noi: gasire, asteptare si oprire fortata."""

from __future__ import annotations

import logging
import os
import signal
import subprocess
from typing import Any

log = logging.getLogger(__package__)  # un singur jurnal pentru tot pachetul

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _pids(image: str) -> set[int]:
    out = subprocess.run(
        ["tasklist", "/FI", f"IMAGENAME eq {image}", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        creationflags=CREATE_NO_WINDOW,
    ).stdout
    pids = set()
    for line in out.splitlines():
        parts = line.strip().strip('"').split('","')
        if len(parts) > 1 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def _pid_of_window(hwnd: int) -> int | None:
    """Procesul care detine fereastra; None daca fereastra nu (mai) exista."""
    import win32process

    return win32process.GetWindowThreadProcessId(hwnd)[1] or None


def _visible_windows(pid: int, classes: set[str]) -> list[int]:
    """Ferestrele vizibile de nivel superior ale procesului, cu una dintre clasele date."""
    import win32gui
    import win32process

    found: list[int] = []

    def visit(hwnd: int, _: object) -> bool:
        if (
            win32gui.IsWindowVisible(hwnd)
            and win32gui.GetClassName(hwnd) in classes
            and win32process.GetWindowThreadProcessId(hwnd)[1] == pid
        ):
            found.append(hwnd)
        return True

    win32gui.EnumWindows(visit, None)
    return found


def _enable_call_cancellation() -> None:
    """Permite ca un apel COM blocat al firului curent sa fie anulat din alt fir."""
    import ctypes

    ctypes.WinDLL("ole32").CoEnableCallCancellation(None)


def _cancel_call(thread_id: int) -> bool:
    """Anuleaza apelul COM in curs al firului (intoarce acolo RPC_E_CALL_CANCELED); procesul
    Office nu e atins. False daca firul nu are acum niciun apel in curs."""
    import ctypes

    return ctypes.WinDLL("ole32").CoCancelCall(thread_id, 0) == 0


def _kill(pid: int) -> None:
    try:
        os.kill(pid, signal.SIGTERM)  # pe Windows: TerminateProcess
    except OSError as exc:
        log.debug("Procesul %s nu mai exista: %s", pid, exc)


_SYNCHRONIZE = 0x00100000
_PROCESS_TERMINATE = 0x0001
_WAIT_TIMEOUT = 0x00000102


def _kernel32() -> Any:
    import ctypes
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)  # instanta proprie: argtypes locale
    k32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    k32.OpenProcess.restype = wintypes.HANDLE
    k32.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
    k32.WaitForSingleObject.restype = wintypes.DWORD
    k32.TerminateProcess.argtypes = (wintypes.HANDLE, wintypes.UINT)
    k32.CloseHandle.argtypes = (wintypes.HANDLE,)
    return k32


def _open_process(pid: int) -> int | None:
    """Handle pentru asteptat/oprit procesul. Luat inainte de Quit, ca pid-ul sa nu poata fi
    refolosit de alt proces intre timp."""
    return _kernel32().OpenProcess(_SYNCHRONIZE | _PROCESS_TERMINATE, False, pid) or None


def _wait_or_terminate(handle: int, seconds: float) -> bool:
    """Asteapta sa se termine procesul; daca nu se termina la timp, il opreste.
    True daca a trebuit oprit. Inchide handle-ul."""
    k32 = _kernel32()
    try:
        if k32.WaitForSingleObject(handle, int(seconds * 1000)) != _WAIT_TIMEOUT:
            return False
        k32.TerminateProcess(handle, 1)
        k32.WaitForSingleObject(handle, 5000)
        return True
    finally:
        k32.CloseHandle(handle)
