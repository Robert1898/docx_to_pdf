"""Tab-ul de setari: tema, integrare Windows, componente detectate, jurnal, despre."""

from __future__ import annotations

import logging
import os
import tkinter as tk
from tkinter import messagebox
from typing import TYPE_CHECKING

import ttkbootstrap as ttk

from .. import APP_TITLE, __version__, pdftools, shell
from ..core import Kind
from ..log import LOG_FILE
from ..paths import is_frozen, log_dir

if TYPE_CHECKING:
    from .app import App

log = logging.getLogger(__name__)


class SettingsTab(ttk.Frame):
    def __init__(self, master: tk.Misc, app: App) -> None:
        super().__init__(master, padding=16)
        self.app = app
        self.theme = tk.StringVar(value=app.settings.theme)
        self.context_menu = tk.BooleanVar(value=shell.context_menu_registered())
        self.sendto = tk.BooleanVar(value=shell.sendto_registered())
        self._build()

    def _section(self, title: str) -> ttk.Labelframe:
        frame = ttk.Labelframe(self, text=title, padding=(12, 8))
        frame.pack(fill=tk.X, pady=(0, 12))
        return frame

    def _build(self) -> None:
        look = self._section("Aspect")
        for value, label in (("system", "Ca sistemul"), ("light", "Luminoasă"),
                             ("dark", "Întunecată")):
            ttk.Radiobutton(look, text=label, value=value, variable=self.theme,
                            command=self._theme_changed).pack(side=tk.LEFT, padx=(0, 18))

        integration = self._section("Integrare cu Windows Explorer")
        ttk.Checkbutton(integration, text="„Convertește în PDF / DOCX” în meniul click-dreapta",
                        variable=self.context_menu, command=self._toggle_context_menu,
                        bootstyle="round-toggle").pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(integration, text="Aplicația în „Trimite către” (merge cu mai multe "
                                          "fișiere selectate deodată)",
                        variable=self.sendto, command=self._toggle_sendto,
                        bootstyle="round-toggle").pack(anchor=tk.W, pady=2)
        note = ("Pe Windows 11, opțiunile apar la „Afișați mai multe opțiuni” (Shift + F10).")
        if not is_frozen():
            note += " Rulezi din sursă: comanda va porni scriptul Python curent."
        ttk.Label(integration, text=note, bootstyle="secondary", wraplength=760).pack(
            anchor=tk.W, pady=(4, 0))

        components = self._section("Componente detectate")
        backends = self.app.backends
        rows = [
            ("Microsoft Word", backends.office.get(Kind.WORD, False),
             "Word → PDF și PDF → DOCX"),
            ("Microsoft Excel", backends.office.get(Kind.EXCEL, False), "Excel → PDF"),
            ("Microsoft PowerPoint", backends.office.get(Kind.POWERPOINT, False),
             "PowerPoint → PDF"),
            ("LibreOffice", backends.soffice is not None,
             "alternativă pentru documente → PDF când lipsește Office"),
        ]
        tess = backends.tesseract
        langs = pdftools.ocr_languages(tess) if tess else ""
        rows.append(("Tesseract OCR", tess is not None,
                     f"OCR, limbi: {langs}" if tess else "necesar pentru OCR"))
        for i, (name, ok, what) in enumerate(rows):
            ttk.Label(components, text="✓" if ok else "✗",
                      bootstyle="success" if ok else "danger").grid(row=i, column=0, padx=(0, 8))
            ttk.Label(components, text=name).grid(row=i, column=1, sticky=tk.W)
            ttk.Label(components, text=what, bootstyle="secondary").grid(
                row=i, column=2, sticky=tk.W, padx=(16, 0))

        logs = self._section("Jurnal")
        ttk.Button(logs, text="Deschide jurnalul", command=self._open_log,
                   bootstyle="secondary-outline").pack(side=tk.LEFT)
        ttk.Button(logs, text="Deschide folderul", command=lambda: os.startfile(log_dir()),
                   bootstyle="secondary-outline").pack(side=tk.LEFT, padx=6)
        ttk.Label(logs, text="Util când ceva nu merge: conține detaliile fiecărei erori.",
                  bootstyle="secondary").pack(side=tk.LEFT, padx=12)

        about = self._section("Despre")
        ttk.Label(about, text=f"{APP_TITLE}  v{__version__}", font="-weight bold").pack(
            anchor=tk.W)
        ttk.Label(about, text="© 2026 Robert-Raul Mîrza · Licență MIT",
                  bootstyle="secondary").pack(anchor=tk.W)

    def _theme_changed(self) -> None:
        self.app.settings.theme = self.theme.get()
        self.app.apply_theme()
        self.app.save_settings()

    def _toggle_context_menu(self) -> None:
        try:
            if self.context_menu.get():
                shell.register_context_menu()
            else:
                shell.unregister_context_menu()
        except OSError as exc:
            log.exception("Meniu click-dreapta")
            messagebox.showerror("Integrare Windows", str(exc), parent=self)
        self.context_menu.set(shell.context_menu_registered())

    def _toggle_sendto(self) -> None:
        try:
            if self.sendto.get():
                shell.create_sendto()
            else:
                shell.remove_sendto()
        except Exception as exc:
            log.exception("Trimite catre")
            messagebox.showerror("Integrare Windows", str(exc), parent=self)
        self.sendto.set(shell.sendto_registered())

    def _open_log(self) -> None:
        path = log_dir() / LOG_FILE
        if path.exists():
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            messagebox.showinfo("Jurnal", "Jurnalul este încă gol.", parent=self)
