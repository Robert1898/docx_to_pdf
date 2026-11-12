"""Tab-ul de setari: aspect si limba, folder urmarit, actualizari, OCR, integrare Windows,
componente detectate, jurnal, despre."""

from __future__ import annotations

import logging
import os
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING

import ttkbootstrap as ttk

from .. import APP_TITLE, __version__, pdftools, shell, updater
from ..core import Kind
from ..i18n import _
from ..log import LOG_FILE
from ..paths import is_frozen, log_dir
from ..watch import default_output
from .widgets import UiQueue, icon_button

if TYPE_CHECKING:
    from .app import App

log = logging.getLogger(__name__)


class Scrollable(ttk.Frame):
    """Continut care se poate derula cu rotita mouse-ului."""

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0)
        scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        self.inner = ttk.Frame(self.canvas, padding=16)
        window = self.canvas.create_window(0, 0, window=self.inner, anchor=tk.NW)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(window,
                                                                            width=e.width))
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.bind_all("<MouseWheel>", self._wheel, add="+")

    def _wheel(self, event: tk.Event) -> None:
        widget = self.winfo_containing(event.x_root, event.y_root)
        while widget is not None:
            if widget is self:
                self.canvas.yview_scroll(-1 * (event.delta // 120), "units")
                return
            widget = widget.master

    def apply_theme(self) -> None:
        self.canvas.configure(background=ttk.Style().lookup("TFrame", "background") or "white")


class SettingsTab(Scrollable):
    def __init__(self, master: tk.Misc, app: App) -> None:
        super().__init__(master)
        self.app = app
        s = app.settings
        self.ui = UiQueue(self)
        self.theme = tk.StringVar(value=s.theme)
        self.language = tk.StringVar(value=s.language)
        self.context_menu = tk.BooleanVar(value=shell.context_menu_registered())
        self.sendto = tk.BooleanVar(value=shell.sendto_registered())
        self.autostart = tk.BooleanVar(value=shell.autostart_enabled())
        self.watch_enabled = tk.BooleanVar(value=s.watch_enabled)
        self.watch_in = tk.StringVar(value=s.watch_in)
        self.watch_out = tk.StringVar(value=s.watch_out)
        self.watch_status = tk.StringVar()
        self.check_updates = tk.BooleanVar(value=s.check_updates)
        self.update_status = tk.StringVar(value=_("Versiunea instalată: {v}.").format(
            v=__version__))
        self.ocr_status = tk.StringVar()
        self.release: updater.Release | None = None
        self._build()

    def _section(self, title: str) -> ttk.Labelframe:
        frame = ttk.Labelframe(self.inner, text=title, padding=(12, 8))
        frame.pack(fill=tk.X, pady=(0, 12))
        return frame

    def _build(self) -> None:
        self._build_look()
        self._build_watch()
        self._build_updates()
        self._build_ocr()
        self._build_integration()
        self._build_components()
        self._build_log_and_about()

    def _build_look(self) -> None:
        look = self._section(_("Aspect"))
        row = ttk.Frame(look)
        row.pack(fill=tk.X)
        ttk.Label(row, text=_("Temă:"), width=11).pack(side=tk.LEFT)
        for value, text in (("system", _("Ca sistemul")), ("light", _("Luminoasă")),
                            ("dark", _("Întunecată"))):
            ttk.Radiobutton(row, text=text, value=value, variable=self.theme,
                            command=self._theme_changed).pack(side=tk.LEFT, padx=(0, 18))
        row = ttk.Frame(look)
        row.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(row, text=_("Limbă:"), width=11).pack(side=tk.LEFT)
        for value, text in (("system", _("Ca sistemul")), ("ro", "Română"), ("en", "English")):
            ttk.Radiobutton(row, text=text, value=value, variable=self.language,
                            command=self._language_changed).pack(side=tk.LEFT, padx=(0, 18))

    def _build_watch(self) -> None:
        watch = self._section(_("Folder urmărit"))
        ttk.Checkbutton(watch, text=_("Convertește automat fișierele puse în folderul de mai jos"),
                        variable=self.watch_enabled, command=self._watch_changed,
                        bootstyle="round-toggle").pack(anchor=tk.W)
        for label_text, var, choose in ((_("Folder urmărit:"), self.watch_in, self._choose_in),
                                        (_("Rezultate în:"), self.watch_out, self._choose_out)):
            row = ttk.Frame(watch)
            row.pack(fill=tk.X, pady=(6, 0))
            ttk.Label(row, text=label_text, width=16).pack(side=tk.LEFT)
            ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True)
            icon_button(row, _("Alege…"), choose, "secondary-outline", "folder").pack(
                side=tk.LEFT, padx=(6, 0))
        ttk.Checkbutton(watch, text=_("Pornește aplicația (minimizată) odată cu Windows"),
                        variable=self.autostart, command=self._toggle_autostart,
                        bootstyle="round-toggle").pack(anchor=tk.W, pady=(8, 0))
        ttk.Label(watch, text=_("Folosește opțiunile din tab-ul Conversie. Funcționează cât timp "
                                "aplicația e deschisă (și minimizată)."),
                  bootstyle="secondary", wraplength=760).pack(anchor=tk.W, pady=(4, 0))
        ttk.Label(watch, textvariable=self.watch_status, bootstyle="info").pack(anchor=tk.W)

    def _build_updates(self) -> None:
        updates = self._section(_("Actualizări"))
        ttk.Checkbutton(updates, text=_("Verifică automat (cel mult o dată pe zi)"),
                        variable=self.check_updates, command=self._updates_changed,
                        bootstyle="round-toggle").pack(anchor=tk.W)
        row = ttk.Frame(updates)
        row.pack(fill=tk.X, pady=(6, 0))
        self.check_button = icon_button(row, _("Verifică acum"), self.check_now,
                                        "secondary-outline", "sync")
        self.check_button.pack(side=tk.LEFT)
        self.install_button = icon_button(row, _("Descarcă și instalează"), self.install,
                                          "success", "download")
        self.update_label = ttk.Label(row, textvariable=self.update_status,
                                      bootstyle="secondary")
        self.update_label.pack(side=tk.LEFT, padx=12)

    def _build_ocr(self) -> None:
        ocr = self._section(_("OCR (text din PDF-uri scanate)"))
        row = ttk.Frame(ocr)
        row.pack(fill=tk.X)
        self.ocr_button = icon_button(row, _("Descarcă limba română"), self.download_romanian,
                                      "secondary-outline", "download")
        self.ocr_button.pack(side=tk.LEFT)
        ttk.Label(row, textvariable=self.ocr_status, bootstyle="secondary").pack(
            side=tk.LEFT, padx=12)
        self._refresh_ocr()

    def _build_integration(self) -> None:
        integration = self._section(_("Integrare cu Windows Explorer"))
        ttk.Checkbutton(integration, text=_("„Convertește în PDF / DOCX” în meniul click-dreapta"),
                        variable=self.context_menu, command=self._toggle_context_menu,
                        bootstyle="round-toggle").pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(integration, text=_("Aplicația în „Trimite către”"),
                        variable=self.sendto, command=self._toggle_sendto,
                        bootstyle="round-toggle").pack(anchor=tk.W, pady=2)
        note = _("Mai multe fișiere selectate deodată ajung în aceeași fereastră. Pe Windows 11, "
                 "opțiunile apar la „Afișați mai multe opțiuni” (Shift + F10).")
        if not is_frozen():
            note += " " + _("Rulezi din sursă: comanda va porni scriptul Python curent.")
        ttk.Label(integration, text=note, bootstyle="secondary", wraplength=760).pack(
            anchor=tk.W, pady=(4, 0))

    def _build_components(self) -> None:
        components = self._section(_("Componente detectate"))
        backends = self.app.backends
        rows = [
            ("Microsoft Word", backends.office.get(Kind.WORD, False),
             _("Word → PDF și PDF → DOCX prin Word")),
            ("Microsoft Excel", backends.office.get(Kind.EXCEL, False), _("Excel → PDF")),
            ("Microsoft PowerPoint", backends.office.get(Kind.POWERPOINT, False),
             _("PowerPoint → PDF")),
            ("LibreOffice", backends.soffice is not None,
             _("alternativă pentru documente → PDF când lipsește Office")),
            ("Tesseract OCR", backends.tesseract is not None, _("necesar pentru OCR")),
        ]
        for i, (name, ok, what) in enumerate(rows):
            ttk.Label(components, text="✓" if ok else "✗",
                      bootstyle="success" if ok else "danger").grid(row=i, column=0, padx=(0, 8))
            ttk.Label(components, text=name).grid(row=i, column=1, sticky=tk.W)
            ttk.Label(components, text=what, bootstyle="secondary").grid(
                row=i, column=2, sticky=tk.W, padx=(16, 0))
        ttk.Label(components, text=_("„Aspect fidel”, PDF → Excel, imaginile și uneltele PDF nu "
                                     "au nevoie de Office."), bootstyle="secondary").grid(
            row=len(rows), column=0, columnspan=3, sticky=tk.W, pady=(6, 0))

    def _build_log_and_about(self) -> None:
        logs = self._section(_("Jurnal"))
        icon_button(logs, _("Deschide jurnalul"), self._open_log, "secondary-outline",
                    "open").pack(side=tk.LEFT)
        icon_button(logs, _("Deschide folderul"), lambda: os.startfile(log_dir()),
                    "secondary-outline", "folder").pack(side=tk.LEFT, padx=6)
        ttk.Label(logs, text=_("Util când ceva nu merge: conține detaliile fiecărei erori."),
                  bootstyle="secondary").pack(side=tk.LEFT, padx=12)

        about = self._section(_("Despre"))
        ttk.Label(about, text=f"{APP_TITLE}  v{__version__}", font="-weight bold").pack(
            anchor=tk.W)
        ttk.Label(about, text=_("© 2026 Robert-Raul Mîrza · Licență MIT"),
                  bootstyle="secondary").pack(anchor=tk.W)
        ttk.Label(about, text=_("Scurtături: Ctrl+O adaugă fișiere, Ctrl+Shift+O adaugă un "
                                "folder, Ctrl+Enter pornește, Esc anulează, Delete elimină, "
                                "Ctrl+1…4 schimbă tab-ul, Ctrl+S salvează paginile."),
                  bootstyle="secondary", wraplength=760).pack(anchor=tk.W, pady=(6, 0))

    def store_settings(self, s: object) -> None:
        s.watch_in = self.watch_in.get().strip()  # type: ignore[attr-defined]
        s.watch_out = self.watch_out.get().strip()  # type: ignore[attr-defined]
        s.watch_enabled = self.watch_enabled.get()  # type: ignore[attr-defined]
        s.check_updates = self.check_updates.get()  # type: ignore[attr-defined]

    # --- aspect si limba ------------------------------------------------------------------

    def _theme_changed(self) -> None:
        self.app.settings.theme = self.theme.get()
        self.app.apply_theme()
        self.app.save_settings()

    def _language_changed(self) -> None:
        self.app.settings.language = self.language.get()
        self.app.save_settings()
        if messagebox.askyesno(_("Limbă"), _("Limba se schimbă după repornire. Repornesc "
                                             "aplicația acum?"), parent=self):
            self.app.restart()

    # --- folder urmarit -------------------------------------------------------------------

    def _choose_in(self) -> None:
        folder = filedialog.askdirectory(parent=self, title=_("Alege folderul urmărit"))
        if folder:
            self.watch_in.set(os.path.normpath(folder))
            if not self.watch_out.get():
                self.watch_out.set(str(default_output(Path(folder))))
            self._watch_changed()

    def _choose_out(self) -> None:
        folder = filedialog.askdirectory(parent=self, title=_("Alege folderul pentru rezultate"))
        if folder:
            self.watch_out.set(os.path.normpath(folder))
            self._watch_changed()

    def _watch_changed(self) -> None:
        if self.watch_enabled.get() and not self.watch_in.get().strip():
            messagebox.showwarning(_("Folder urmărit"), _("Alege întâi folderul urmărit."),
                                   parent=self)
            self.watch_enabled.set(False)
        self.store_settings(self.app.settings)
        self.app.save_settings()
        self.app.update_watch()

    def watch_info(self, text: str) -> None:
        self.watch_status.set(text)

    def _toggle_autostart(self) -> None:
        try:
            shell.set_autostart(self.autostart.get())
        except OSError as exc:
            messagebox.showerror(_("Pornire cu Windows"), str(exc), parent=self)
        self.autostart.set(shell.autostart_enabled())
        self.app.settings.autostart = self.autostart.get()
        self.app.save_settings()

    # --- actualizari ----------------------------------------------------------------------

    def _updates_changed(self) -> None:
        self.app.settings.check_updates = self.check_updates.get()
        self.app.save_settings()

    def check_now(self, quiet: bool = False) -> None:
        self.check_button.configure(state=tk.DISABLED)
        self.update_status.set(_("Verific…"))

        def work() -> None:
            release = updater.latest_release()
            self.ui.post(lambda: self._checked(release, quiet))

        threading.Thread(target=work, name="update-check", daemon=True).start()

    def _checked(self, release: updater.Release | None, quiet: bool) -> None:
        self.check_button.configure(state=tk.NORMAL)
        self.app.settings.last_update_check = time.time()
        self.app.save_settings()
        if release is None:
            self.update_status.set(_("Nu am putut verifica (fără internet sau depozit privat)."))
            return
        if not updater.is_newer(release.version):
            self.update_status.set(_("Ai cea mai nouă versiune ({v}).").format(v=__version__))
            return
        self.release = release
        self.update_status.set(_("Versiune nouă: {v} (ai {current}).").format(
            v=release.version, current=__version__))
        self.install_button.pack(side=tk.LEFT, padx=(6, 0), before=self.update_label)
        if quiet and release.version != self.app.settings.skipped_version:
            self.app.notify(_("Actualizare disponibilă"),
                            _("{app} {v} poate fi instalat din Setări.").format(
                                app=APP_TITLE, v=release.version))

    def install(self) -> None:
        release = self.release
        if release is None:
            return
        if not is_frozen():
            messagebox.showinfo(_("Actualizare"), _("Rulezi din sursă; actualizează cu git."),
                                parent=self)
            return
        if not messagebox.askyesno(_("Actualizare"), _(
                "Descarc versiunea {v} și pornesc installer-ul? Aplicația se va închide."
                ).format(v=release.version), parent=self):
            return
        self.install_button.configure(state=tk.DISABLED)

        def progress(done: int, total: int) -> None:
            self.ui.post(lambda: self.update_status.set(_("Descarc… {p}%").format(
                p=done * 100 // max(total, 1))))

        def work() -> None:
            try:
                path = updater.download(release, progress)
            except OSError as exc:
                message = str(exc)
                self.ui.post(lambda: (self.update_status.set(
                    _("Descărcarea a eșuat: {error}").format(error=message)),
                    self.install_button.configure(state=tk.NORMAL)))
                return
            self.ui.post(lambda: self._launch(path))

        threading.Thread(target=work, name="update-download", daemon=True).start()

    def _launch(self, path: Path) -> None:
        updater.launch_installer(path)
        self.app.on_close(force=True)

    # --- OCR ------------------------------------------------------------------------------

    def _refresh_ocr(self) -> None:
        exe = self.app.backends.tesseract
        if exe is None:
            self.ocr_status.set(_("Tesseract OCR nu este instalat."))
            self.ocr_button.configure(state=tk.DISABLED)
            return
        langs = pdftools.tesseract_languages(exe)
        self.ocr_status.set(_("Limbi disponibile: {langs}.").format(langs=", ".join(langs)
                                                                    or "-"))
        self.ocr_button.configure(state=tk.DISABLED if "ron" in langs else tk.NORMAL)

    def download_romanian(self) -> None:
        self.ocr_button.configure(state=tk.DISABLED)

        def progress(done: int, total: int) -> None:
            self.ui.post(lambda: self.ocr_status.set(_("Descarc… {p}%").format(
                p=done * 100 // max(total, 1))))

        def work() -> None:
            try:
                pdftools.download_ocr_language("ron", progress)
            except pdftools.PdfError as exc:
                message = str(exc)
                self.ui.post(lambda: (self.ocr_status.set(message),
                                      self.ocr_button.configure(state=tk.NORMAL)))
                return
            self.ui.post(self._refresh_ocr)

        threading.Thread(target=work, name="ocr-download", daemon=True).start()

    # --- integrare ------------------------------------------------------------------------

    def _toggle_context_menu(self) -> None:
        try:
            if self.context_menu.get():
                shell.register_context_menu()
            else:
                shell.unregister_context_menu()
        except OSError as exc:
            log.exception("Meniu click-dreapta")
            messagebox.showerror(_("Integrare Windows"), str(exc), parent=self)
        self.context_menu.set(shell.context_menu_registered())

    def _toggle_sendto(self) -> None:
        try:
            if self.sendto.get():
                shell.create_sendto()
            else:
                shell.remove_sendto()
        except Exception as exc:
            log.exception("Trimite catre")
            messagebox.showerror(_("Integrare Windows"), str(exc), parent=self)
        self.sendto.set(shell.sendto_registered())

    def _open_log(self) -> None:
        path = log_dir() / LOG_FILE
        if path.exists():
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            messagebox.showinfo(_("Jurnal"), _("Jurnalul este încă gol."), parent=self)
