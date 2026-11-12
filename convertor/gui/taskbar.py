"""Progresul conversiei pe iconita din taskbar (ITaskbarList3), ca la copierea fisierelor.

Apelurile COM sunt facute direct prin ctypes (pywin32 nu expune aceasta interfata).
Daca ceva nu merge (Windows foarte vechi), progresul pur si simplu nu se afiseaza.
"""

from __future__ import annotations

import ctypes
import logging
import tkinter as tk
import uuid
from ctypes import wintypes

log = logging.getLogger(__name__)

NOPROGRESS, INDETERMINATE, NORMAL, ERROR, PAUSED = 0, 1, 2, 4, 8
_CLSID = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
_IID = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"


class _GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort), ("Data4", ctypes.c_ubyte * 8)]

    @classmethod
    def parse(cls, text: str) -> _GUID:
        return cls.from_buffer_copy(uuid.UUID(text).bytes_le)


class TaskbarProgress:
    def __init__(self, window: tk.Misc) -> None:
        self._window = window
        self._obj = ctypes.c_void_p()
        self._ok = False
        try:
            ctypes.windll.ole32.CoInitialize(None)
            hr = ctypes.windll.ole32.CoCreateInstance(
                ctypes.byref(_GUID.parse(_CLSID)), None, 1,  # CLSCTX_INPROC_SERVER
                ctypes.byref(_GUID.parse(_IID)), ctypes.byref(self._obj))
            self._ok = hr == 0 and bool(self._obj) and self._call(3) == 0  # HrInit
        except (OSError, AttributeError) as exc:
            log.debug("Progres in taskbar indisponibil: %s", exc)

    def _method(self, index: int, *argtypes: type) -> ctypes._CFuncPtr:  # type: ignore[name-defined]
        vtable = ctypes.cast(self._obj, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
        prototype = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *argtypes)
        return prototype(vtable[index])

    def _call(self, index: int, *args: object, types: tuple = ()) -> int:
        return self._method(index, *types)(self._obj, *args)

    def _hwnd(self) -> int:
        return ctypes.windll.user32.GetParent(self._window.winfo_id())

    def set(self, done: float, total: float) -> None:
        """Bara verde cu procentul dat."""
        if not self._ok:
            return
        try:
            self._call(10, self._hwnd(), NORMAL, types=(wintypes.HWND, ctypes.c_int))
            self._call(9, self._hwnd(), ctypes.c_ulonglong(int(done)),
                       ctypes.c_ulonglong(max(1, int(total))),
                       types=(wintypes.HWND, ctypes.c_ulonglong, ctypes.c_ulonglong))
        except OSError as exc:
            log.debug("SetProgressValue: %s", exc)

    def state(self, value: int) -> None:
        """NOPROGRESS (ascuns), INDETERMINATE, ERROR (rosu), PAUSED (galben)."""
        if not self._ok:
            return
        try:
            self._call(10, self._hwnd(), value, types=(wintypes.HWND, ctypes.c_int))
        except OSError as exc:
            log.debug("SetProgressState: %s", exc)

    def clear(self) -> None:
        self.state(NOPROGRESS)
