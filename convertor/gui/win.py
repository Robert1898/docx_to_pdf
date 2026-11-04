"""Detalii Windows: DPI, tema sistemului, bara de titlu intunecata."""

from __future__ import annotations

import ctypes
import logging
import tkinter as tk
import winreg

log = logging.getLogger(__name__)

LIGHT_THEME = "bootstrap-light"
DARK_THEME = "bootstrap-dark"


def enable_dpi_awareness() -> None:
    """Fara asta, Windows mareste fereastra ca pe o imagine si textul iese incetosat."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError) as exc:
            log.debug("DPI awareness indisponibil: %s", exc)


def system_prefers_dark() -> bool:
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def resolve_theme(choice: str) -> str:
    if choice == "dark" or (choice == "system" and system_prefers_dark()):
        return DARK_THEME
    return LIGHT_THEME


def set_dark_titlebar(window: tk.Misc, dark: bool) -> None:
    """Bara de titlu Windows 10/11 in ton cu tema (DWMWA_USE_IMMERSIVE_DARK_MODE = 20)."""
    try:
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), 4)
    except (AttributeError, OSError) as exc:
        log.debug("Bara de titlu intunecata indisponibila: %s", exc)
