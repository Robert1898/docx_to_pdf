"""Fereastra principala: tema, tab-uri, setari, inchidere sigura."""

from __future__ import annotations

import logging
import sys
import tkinter as tk
import traceback
from collections.abc import Sequence
from tkinter import messagebox

import ttkbootstrap as ttk

from .. import APP_TITLE, settings
from ..engine import Backends
from ..log import setup_logging
from ..paths import resource_path
from .convert_tab import ConvertTab
from .pdf_tab import PdfToolsTab
from .settings_tab import SettingsTab
from .win import DARK_THEME, enable_dpi_awareness, resolve_theme, set_dark_titlebar

log = logging.getLogger(__name__)

DEFAULT_SIZE = (1000, 760)


def _make_root() -> tk.Tk:
    """Fereastra cu suport drag & drop; fara tkdnd, una obisnuita."""
    try:
        from tkinterdnd2 import TkinterDnD

        return TkinterDnD.Tk()
    except Exception as exc:
        log.warning("tkdnd indisponibil, fara drag & drop: %s", exc)
        return tk.Tk()


class App:
    def __init__(self, files: Sequence[str] = (), auto: bool = False) -> None:
        enable_dpi_awareness()
        self.settings = settings.load()
        self.backends = Backends.detect()

        self.root = _make_root()
        self.root.withdraw()
        self.root.title(APP_TITLE)
        self.root.report_callback_exception = self._on_tk_error
        icon = resource_path("icon.ico")
        if icon.exists():
            self.root.iconbitmap(default=str(icon))
        self.style = ttk.Style()
        self.apply_theme()
        self._restore_geometry()

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.convert_tab = ConvertTab(notebook, self)
        self.pdf_tab = PdfToolsTab(notebook, self)
        self.settings_tab = SettingsTab(notebook, self)
        notebook.add(self.convert_tab, text="  Conversie  ")
        notebook.add(self.pdf_tab, text="  Unelte PDF  ")
        notebook.add(self.settings_tab, text="  Setări  ")

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.deiconify()

        if files:
            self.convert_tab.add_paths(files)
            if auto:
                self.root.after(400, self.convert_tab.start)

    def run(self) -> None:
        self.root.mainloop()

    # --- aspect ---------------------------------------------------------------------------

    def apply_theme(self) -> None:
        theme = resolve_theme(self.settings.theme)
        self.style.theme_use(theme)
        set_dark_titlebar(self.root, theme == DARK_THEME)
        # randuri mai inalte, proportional cu DPI-ul (Treeview le are in pixeli)
        scale = self.root.winfo_fpixels("1i") / 96
        self.style.configure("Treeview", rowheight=int(26 * scale))

    def _restore_geometry(self) -> None:
        scale = self.root.winfo_fpixels("1i") / 96
        width, height = (int(v * scale) for v in DEFAULT_SIZE)
        self.root.minsize(int(760 * scale), int(600 * scale))
        geometry = self.settings.geometry
        if geometry:
            try:
                self.root.geometry(geometry)
                return
            except tk.TclError:
                pass
        x = (self.root.winfo_screenwidth() - width) // 2
        y = max(0, (self.root.winfo_screenheight() - height) // 3)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def notify(self, title: str, message: str) -> None:
        """Notificare discreta in coltul ecranului, fara sa blocheze fereastra."""
        try:
            ttk.ToastNotification(title=title, message=message, duration=4000).show_toast()
        except Exception as exc:
            log.debug("Toast indisponibil: %s", exc)

    # --- setari si inchidere --------------------------------------------------------------

    def save_settings(self) -> None:
        self.convert_tab.store_settings(self.settings)
        self.pdf_tab.store_settings(self.settings)
        if self.root.state() == "normal":
            self.settings.geometry = self.root.geometry()
        settings.save(self.settings)

    def on_close(self) -> None:
        busy = self.convert_tab.busy or self.pdf_tab.busy
        if busy and not messagebox.askyesno(
            "Lucru în desfășurare",
            "O operație e încă în curs. Închizi oricum? Fișierul curent nu va fi terminat.",
            parent=self.root,
        ):
            return
        self.save_settings()
        if busy:
            self.convert_tab.abort()
            self.pdf_tab.cancel()
            for thread in (self.convert_tab.thread, self.pdf_tab.thread):
                if thread is not None:
                    thread.join(timeout=8)  # lasa Office sa se inchida si registry-ul refacut
        self.root.destroy()

    def _on_tk_error(self, exc: type[BaseException], value: BaseException, tb: object) -> None:
        log.error("Eroare neasteptata:\n%s", "".join(traceback.format_exception(exc, value, tb)))
        messagebox.showerror(
            "Eroare neașteptată",
            f"{value}\n\nDetaliile sunt în jurnal (Setări → Jurnal).",
            parent=self.root,
        )


if __name__ == "__main__":  # pragma: no cover
    setup_logging()
    App(sys.argv[1:]).run()
