"""Fereastra principala: tema, tab-uri, scurtaturi, instanta unica, folder urmarit,
actualizari si inchidere sigura."""

from __future__ import annotations

import logging
import subprocess
import sys
import threading
import time
import tkinter as tk
import traceback
from collections.abc import Sequence
from pathlib import Path
from tkinter import messagebox
from typing import Any

import ttkbootstrap as ttk

from .. import APP_TITLE, settings, shell, updater
from ..core import Job
from ..engine import Backends, Engine, OfficePool, Result, Status
from ..i18n import _
from ..log import setup_logging
from ..paths import resource_path
from ..watch import FolderWatcher, default_output
from .convert_tab import ConvertTab
from .instance import InstanceServer
from .pages_tab import PagesTab
from .pdf_tab import PdfToolsTab
from .settings_tab import SettingsTab
from .widgets import UiQueue
from .win import DARK_THEME, enable_dpi_awareness, resolve_theme, set_dark_titlebar

log = logging.getLogger(__name__)

DEFAULT_SIZE = (1040, 780)
UPDATE_DELAY_MS = 8000  # verificarea actualizarilor nu incetineste pornirea


def _make_root() -> tk.Tk:
    """Fereastra cu suport drag & drop; fara tkdnd, una obisnuita."""
    try:
        from tkinterdnd2 import TkinterDnD

        return TkinterDnD.Tk()
    except Exception as exc:
        log.warning("tkdnd indisponibil, fara drag & drop: %s", exc)
        return tk.Tk()


class App:
    def __init__(self, files: Sequence[str] = (), auto: bool = False,
                 listener: Any = None, minimized: bool = False) -> None:
        enable_dpi_awareness()
        self.settings = settings.load()
        self.backends = Backends.detect()
        self._pool: OfficePool | None = None
        self._pool_lock = threading.Lock()
        self.watcher: FolderWatcher | None = None
        self._closing = False

        self.root = _make_root()
        self.root.withdraw()
        self.root.title(APP_TITLE)
        self.root.report_callback_exception = self._on_tk_error
        icon = resource_path("icon.ico")
        if icon.exists():
            self.root.iconbitmap(default=str(icon))
        self.style = ttk.Style()
        self.ui = UiQueue(self.root)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.convert_tab = ConvertTab(self.notebook, self)
        self.pdf_tab = PdfToolsTab(self.notebook, self)
        self.pages_tab = PagesTab(self.notebook, self)
        self.settings_tab = SettingsTab(self.notebook, self)
        for tab, text in ((self.convert_tab, _("Conversie")), (self.pdf_tab, _("Unelte PDF")),
                          (self.pages_tab, _("Pagini")), (self.settings_tab, _("Setări"))):
            self.notebook.add(tab, text=f"  {text}  ")
        self.apply_theme()
        self._restore_geometry()
        self._bind_shortcuts()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.deiconify()
        if minimized:
            self.root.iconify()

        self.server: InstanceServer | None = None
        if listener is not None and hasattr(listener, "address"):
            self.server = InstanceServer(listener, lambda m: self.ui.post(lambda: self._message(m)))
            self.server.start()

        self.update_watch()
        if self.settings.check_updates and (
                time.time() - self.settings.last_update_check > updater.CHECK_INTERVAL):
            self.root.after(UPDATE_DELAY_MS, lambda: self.settings_tab.check_now(quiet=True))

        if files:
            self.convert_tab.receive(files, auto)

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
        for tab in (getattr(self, "pages_tab", None), getattr(self, "settings_tab", None)):
            if tab is not None:
                tab.apply_theme()

    def _restore_geometry(self) -> None:
        scale = self.root.winfo_fpixels("1i") / 96
        width, height = (int(v * scale) for v in DEFAULT_SIZE)
        self.root.minsize(int(780 * scale), int(600 * scale))
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

    def bring_to_front(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(200, lambda: self.root.attributes("-topmost", False))
        self.root.focus_force()

    # --- scurtaturi -----------------------------------------------------------------------

    def _current(self) -> tk.Misc:
        return self.notebook.nametowidget(self.notebook.select())

    def _bind_shortcuts(self) -> None:
        def handler(fn: Any) -> Any:
            def run(_event: tk.Event) -> str:
                fn()
                return "break"
            return run

        bindings = {
            "<Control-o>": self._shortcut_open,
            "<Control-O>": self._shortcut_open_folder,  # Ctrl+Shift+O
            "<Control-Return>": self._shortcut_start,
            "<Escape>": self._shortcut_cancel,
            "<Control-s>": self._shortcut_save,
        }
        for sequence, fn in bindings.items():
            self.root.bind(sequence, handler(fn))
        for number in range(1, 5):
            self.root.bind(f"<Control-Key-{number}>",
                           handler(lambda n=number: self.notebook.select(n - 1)))

    def _shortcut_open(self) -> None:
        tab = self._current()
        if tab is self.convert_tab:
            self.convert_tab.add_files_dialog()
        elif tab is self.pdf_tab:
            self.pdf_tab.add_dialog()
        elif tab is self.pages_tab:
            self.pages_tab.open_dialog()

    def _shortcut_open_folder(self) -> None:
        if self._current() is self.convert_tab:
            self.convert_tab.add_folder_dialog()

    def _shortcut_start(self) -> None:
        tab = self._current()
        if tab is self.convert_tab:
            self.convert_tab.start()
        elif tab is self.pdf_tab:
            self.pdf_tab.start()

    def _shortcut_cancel(self) -> None:
        tab = self._current()
        if tab is self.convert_tab and self.convert_tab.busy:
            self.convert_tab.cancel()
        elif tab is self.pdf_tab and self.pdf_tab.busy:
            self.pdf_tab.cancel()

    def _shortcut_save(self) -> None:
        if self._current() is self.pages_tab:
            self.pages_tab.save_as()

    # --- alte instante --------------------------------------------------------------------

    def _message(self, message: dict) -> None:
        files = [f for f in message.get("files", []) if isinstance(f, str)]
        self.bring_to_front()
        if files:
            self.notebook.select(self.convert_tab)
            self.convert_tab.receive(files, bool(message.get("auto")))

    # --- Office -----------------------------------------------------------------------------

    def pool(self, workers: int) -> OfficePool:
        """Pool-ul ramane pornit intre conversii: al doilea lot gaseste Word deja deschis."""
        with self._pool_lock:
            if self._pool is not None and self._pool.size != workers:
                old = self._pool
                threading.Thread(target=old.close, name="pool-close", daemon=True).start()
                self._pool = None
            if self._pool is None:
                self._pool = OfficePool(workers)
            return self._pool

    # --- folder urmarit ---------------------------------------------------------------------

    def update_watch(self) -> None:
        """Porneste, opreste sau reporneste urmarirea dupa setarile curente."""
        if self.watcher is not None:
            self.watcher.stop()
            self.watcher = None
        s = self.settings
        if not s.watch_enabled or not s.watch_in:
            self.settings_tab.watch_info("")
            return
        folder_in = Path(s.watch_in)
        folder_out = Path(s.watch_out) if s.watch_out else default_output(folder_in)
        self.watcher = FolderWatcher(folder_in, folder_out, self._watch_convert,
                                     pdf_target=lambda: self.settings.pdf_target)
        self.watcher.start()
        self.settings_tab.watch_info(_("Urmăresc „{folder}”.").format(folder=folder_in))

    def _watch_convert(self, jobs: list[Job]) -> list[Result]:
        """Ruleaza in firul de urmarire."""
        options = self.settings.export_options()
        engine = Engine(options, backends=self.backends, pool=self.pool(options.workers))
        results = engine.run(jobs)
        done = sum(r.status is Status.DONE for r in results)
        failed = sum(r.status is Status.FAILED for r in results)
        text = _("{done} convertite, {failed} eșuate din folderul urmărit.").format(
            done=done, failed=failed)
        self.ui.post(lambda: (self.settings_tab.watch_info(text),
                              self.notify(_("Folder urmărit"), text)))
        return results

    # --- setari si inchidere --------------------------------------------------------------

    def save_settings(self) -> None:
        self.convert_tab.store_settings(self.settings)
        self.pdf_tab.store_settings(self.settings)
        self.settings_tab.store_settings(self.settings)
        if self.root.state() == "normal":
            self.settings.geometry = self.root.geometry()
        settings.save(self.settings)

    def restart(self) -> None:
        if self.on_close():
            subprocess.Popen(shell.gui_command(), close_fds=True)

    def on_close(self, force: bool = False) -> bool:
        if self._closing:
            return False
        tabs = (self.convert_tab, self.pdf_tab, self.pages_tab)
        busy = any(tab.busy for tab in tabs)
        if not force:
            if busy and not messagebox.askyesno(
                _("Lucru în desfășurare"),
                _("O operație e încă în curs. Închizi oricum? Fișierul curent nu va fi "
                  "terminat."),
                parent=self.root,
            ):
                return False
            if not self.pages_tab.confirm_discard():
                return False
        self._closing = True
        self.save_settings()
        self.root.withdraw()
        if self.watcher is not None:
            self.watcher.stop()
        if busy:
            self.convert_tab.abort()
            self.pdf_tab.cancel()
            for tab in tabs:
                if tab.thread is not None:
                    tab.thread.join(timeout=8)  # lasa Office sa se inchida, registry refacut
        if self.server is not None:
            self.server.stop()
        if self._pool is not None:
            self._pool.close(wait=10)
        self.root.destroy()
        return True

    def _on_tk_error(self, exc: type[BaseException], value: BaseException, tb: object) -> None:
        log.error("Eroare neasteptata:\n%s", "".join(traceback.format_exception(exc, value, tb)))
        messagebox.showerror(
            _("Eroare neașteptată"),
            _("{error}\n\nDetaliile sunt în jurnal (Setări → Jurnal).").format(error=value),
            parent=self.root,
        )


if __name__ == "__main__":  # pragma: no cover
    setup_logging()
    App(sys.argv[1:]).run()
