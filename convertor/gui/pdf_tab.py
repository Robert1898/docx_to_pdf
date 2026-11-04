"""Tab-ul de unelte PDF: unire, impartire, parola, watermark, OCR."""

from __future__ import annotations

import contextlib
import logging
import os
import threading
import tkinter as tk
from collections.abc import Callable, Sequence
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING

import ttkbootstrap as ttk

from .. import pdftools
from ..core import free_path
from ..settings import Settings
from .widgets import Column, FileTable, UiQueue, tooltip

if TYPE_CHECKING:
    from .app import App

log = logging.getLogger(__name__)

COLUMNS = (
    Column("name", "Fișier", 260),
    Column("pages", "Pagini", 70, stretch=False),
    Column("folder", "Folder", 250),
    Column("status", "Status", 220),
)

OPERATIONS = (
    ("merge", "Unește"),
    ("split", "Împarte pe pagini"),
    ("protect", "Parolă"),
    ("watermark", "Watermark"),
    ("ocr", "OCR"),
)

HINTS = {
    "merge": "Toate PDF-urile din listă devin unul singur, în ordinea din tabel (folosește ↑ ↓).",
    "split": "Fiecare parte devine un PDF separat, într-un folder nou lângă original.",
    "protect": "PDF-ul nu se poate deschide fără parolă (criptare AES-256). Parola nu se salvează.",
    "watermark": "Textul apare în diagonală, semitransparent, pe fiecare pagină.",
    "ocr": "PDF-urile scanate primesc un strat de text: devin căutabile și se pot copia.",
}


class PdfToolsTab(ttk.Frame):
    def __init__(self, master: tk.Misc, app: App) -> None:
        super().__init__(master, padding=12)
        self.app = app
        self.files: dict[str, Path] = {}
        self.ui = UiQueue(self)
        self.thread: threading.Thread | None = None
        self.cancel_event = threading.Event()

        self.operation = tk.StringVar(value="merge")
        self.password = tk.StringVar()
        self.password_confirm = tk.StringVar()
        self.watermark_text = tk.StringVar(value=app.settings.watermark_text)
        self.opacity = tk.IntVar(value=app.settings.watermark_opacity)
        self.renumber = tk.BooleanVar(value=app.settings.renumber_pages)
        self.split_mode = tk.StringVar(value=app.settings.split_mode)
        self.split_every = tk.StringVar(value=str(app.settings.split_every))
        self.split_parts = tk.StringVar(value=str(app.settings.split_parts))
        self.split_ranges = tk.StringVar(value=app.settings.split_ranges)
        self.split_preview = tk.StringVar()
        self.pages: dict[str, int | None] = {}
        self.hint = tk.StringVar()
        self.status = tk.StringVar(value="Adaugă PDF-uri pentru a începe.")
        self._build()
        self._on_operation()
        for var in (self.split_mode, self.split_every, self.split_parts, self.split_ranges):
            var.trace_add("write", lambda *_: self._update_split_preview())

    @property
    def busy(self) -> bool:
        return self.thread is not None and self.thread.is_alive()

    def _build(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X)
        self.buttons = [
            ttk.Button(bar, text="Adaugă PDF-uri…", command=self.add_dialog, bootstyle="primary"),
            ttk.Button(bar, text="Elimină", command=self.remove_selected,
                       bootstyle="secondary-outline"),
            ttk.Button(bar, text="Golește lista", command=self.clear,
                       bootstyle="secondary-outline"),
            ttk.Button(bar, text="↑", width=3, command=lambda: self._move(-1),
                       bootstyle="secondary-outline"),
            ttk.Button(bar, text="↓", width=3, command=lambda: self._move(1),
                       bootstyle="secondary-outline"),
        ]
        for i, button in enumerate(self.buttons):
            button.pack(side=tk.LEFT, padx=(0 if i == 0 else 6, 0))

        self.table = FileTable(self, COLUMNS, "Trage aici fișiere PDF.", on_drop=self.add_paths)
        self.table.pack(fill=tk.BOTH, expand=True, pady=10)
        self.table.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.table.tree.bind("<Double-1>", self._open_row)
        self.table.tree.bind("<<TreeviewSelect>>", lambda e: self._update_split_preview())

        box = ttk.Labelframe(self, text="Operație", padding=(10, 8))
        box.pack(fill=tk.X)
        radios = ttk.Frame(box)
        radios.pack(fill=tk.X)
        self.radios: dict[str, ttk.Radiobutton] = {}
        for key, label in OPERATIONS:
            radio = ttk.Radiobutton(radios, text=label, value=key, variable=self.operation,
                                    command=self._on_operation, bootstyle="toolbutton-outline")
            radio.pack(side=tk.LEFT, padx=(0, 6))
            self.radios[key] = radio
        if self.app.backends.tesseract is None:
            self.radios["ocr"].configure(state=tk.DISABLED)

        ttk.Label(box, textvariable=self.hint, bootstyle="secondary", wraplength=760).pack(
            anchor=tk.W, pady=(8, 4)
        )

        self.params = ttk.Frame(box)
        self.params.pack(fill=tk.X)
        self.param_frames: dict[str, ttk.Frame] = {}

        merge = ttk.Frame(self.params)
        renumber = ttk.Checkbutton(
            merge, text="Renumerotează paginile (ex. 8/8 și 1/2 devin 8/10 și 9/10)",
            variable=self.renumber)
        renumber.pack(anchor=tk.W)
        tooltip(renumber, "Numerele de pagină tipărite în antet sau subsol („8/8”, "
                          "„Pagina 1 din 2”) sunt înlocuite cu numerotarea documentului unit, "
                          "în același font și loc.")
        self.param_frames["merge"] = merge

        split = ttk.Frame(self.params)

        def mode_row(value: str, text: str) -> ttk.Frame:
            row = ttk.Frame(split)
            row.pack(anchor=tk.W, pady=2)
            ttk.Radiobutton(row, text=text, value=value, variable=self.split_mode).pack(
                side=tk.LEFT)
            return row

        def field(widget: tk.Widget, value: str) -> None:
            """Campul se poate completa direct: alege si modul lui."""
            widget.pack(side=tk.LEFT, padx=6)
            widget.bind("<FocusIn>", lambda e: self.split_mode.set(value))

        mode_row("each", "Fiecare pagină separat")
        row = mode_row("every", "Câte")
        field(ttk.Spinbox(row, from_=1, to=100000, width=7, textvariable=self.split_every),
              "every")
        ttk.Label(row, text="pagini într-un fișier").pack(side=tk.LEFT)
        row = mode_row("parts", "În")
        field(ttk.Spinbox(row, from_=1, to=10000, width=7, textvariable=self.split_parts),
              "parts")
        ttk.Label(row, text="părți egale").pack(side=tk.LEFT)
        row = mode_row("ranges", "Intervale:")
        field(ttk.Entry(row, textvariable=self.split_ranges, width=30), "ranges")
        ttk.Label(row, text="ex. 1-200, 201-450, 451-  (gol după cratimă = până la sfârșit)",
                  bootstyle="secondary").pack(side=tk.LEFT)
        self.preview_label = ttk.Label(split, textvariable=self.split_preview, bootstyle="info")
        self.preview_label.pack(anchor=tk.W, pady=(6, 0))
        self.param_frames["split"] = split

        protect = ttk.Frame(self.params)
        ttk.Label(protect, text="Parolă:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(protect, textvariable=self.password, show="•", width=28).grid(
            row=0, column=1, padx=6, pady=2)
        ttk.Label(protect, text="Confirmă:").grid(row=0, column=2, sticky=tk.W, padx=(12, 0))
        ttk.Entry(protect, textvariable=self.password_confirm, show="•", width=28).grid(
            row=0, column=3, padx=6, pady=2)
        self.param_frames["protect"] = protect

        mark = ttk.Frame(self.params)
        ttk.Label(mark, text="Text:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(mark, textvariable=self.watermark_text, width=34).grid(
            row=0, column=1, padx=6, pady=2)
        ttk.Label(mark, text="Opacitate:").grid(row=0, column=2, sticky=tk.W, padx=(12, 0))
        ttk.Scale(mark, from_=5, to=100, variable=self.opacity, length=160,
                  command=lambda v: self.opacity.set(round(float(v)))).grid(
            row=0, column=3, padx=6)
        ttk.Label(mark, textvariable=self.opacity, width=4).grid(row=0, column=4, sticky=tk.W)
        ttk.Label(mark, text="%").grid(row=0, column=5, sticky=tk.W)
        self.param_frames["watermark"] = mark

        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, pady=(12, 0))
        self.progress = ttk.Progressbar(bottom, mode="determinate", bootstyle="info")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cancel_button = ttk.Button(bottom, text="Anulează", command=self.cancel,
                                        bootstyle="danger-outline", state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT, padx=(10, 0))
        self.run_button = ttk.Button(bottom, text="Execută", command=self.start,
                                     bootstyle="info")
        self.run_button.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(self, textvariable=self.status, bootstyle="secondary").pack(
            fill=tk.X, pady=(6, 0))

    def _on_operation(self) -> None:
        op = self.operation.get()
        self.hint.set(HINTS[op])
        if op == "ocr" and self.app.backends.tesseract is not None:
            langs = pdftools.ocr_languages(self.app.backends.tesseract)
            self.hint.set(f"{HINTS[op]} Limbi folosite: {langs}.")
        for key, frame in self.param_frames.items():
            if key == op:
                frame.pack(anchor=tk.W, pady=(4, 0))
            else:
                frame.pack_forget()
        self._update_split_preview()

    # --- impartire ------------------------------------------------------------------------

    def _split_rule(self) -> Callable[[int], list[pdftools.PageRange]]:
        """Regula de impartire aleasa acum, ca functie de numarul de pagini. Citeste
        interfata o singura data, deci poate fi folosita apoi si din firul de lucru."""
        mode = self.split_mode.get()
        every, parts = self.split_every.get().strip(), self.split_parts.get().strip()
        text = self.split_ranges.get()

        def number(raw: str, what: str) -> int:
            try:
                return int(raw)
            except ValueError:
                raise pdftools.PdfError(f"Scrie un număr întreg pentru {what}.") from None

        def rule(total: int) -> list[pdftools.PageRange]:
            if mode == "every":
                return pdftools.ranges_every(total, number(every, "paginile pe fișier"))
            if mode == "parts":
                return pdftools.ranges_parts(total, number(parts, "numărul de părți"))
            if mode == "ranges":
                return pdftools.parse_ranges(text, total)
            return pdftools.ranges_every(total, 1)

        return rule

    def _update_split_preview(self) -> None:
        if self.operation.get() != "split":
            return
        selected = [iid for iid in self.table.selection() if iid in self.files]
        iid = (selected or self.table.order() or [None])[0]
        if iid is None:
            self.split_preview.set("Adaugă un PDF ca să vezi cum va fi împărțit.")
            self.preview_label.configure(bootstyle="secondary")
            return
        total = self.pages.get(iid)
        name = self.files[iid].name
        if not total:
            self.split_preview.set(f"„{name}” nu poate fi citit (protejat sau deteriorat).")
            self.preview_label.configure(bootstyle="danger")
            return
        try:
            ranges = self._split_rule()(total)
        except pdftools.PdfError as exc:
            self.split_preview.set(str(exc))
            self.preview_label.configure(bootstyle="danger")
            return
        count = f"{len(ranges)} fișier" + ("e" if len(ranges) != 1 else "")
        self.split_preview.set(f"„{name}” ({total} pagini) → {count}: "
                               f"{pdftools.describe_ranges(ranges)}")
        self.preview_label.configure(bootstyle="info")

    # --- lista ----------------------------------------------------------------------------

    def add_paths(self, paths: Sequence[str | os.PathLike[str]]) -> int:
        if self.busy:
            return 0
        added = 0
        for raw in paths:
            path = Path(raw)
            candidates = sorted(path.rglob("*.pdf")) if path.is_dir() else [path]
            for pdf in candidates:
                if pdf.suffix.lower() != ".pdf" or not pdf.is_file():
                    continue
                iid = os.path.normcase(str(pdf.resolve()))
                if iid in self.files:
                    continue
                try:
                    count: int | None = pdftools.page_count(pdf)
                    pages = str(count)
                except pdftools.PdfError as exc:
                    count, pages = None, "🔒" if "parolă" in str(exc) else "?"
                self.files[iid] = pdf.resolve()
                self.pages[iid] = count
                self.table.insert(iid, (pdf.name, pages, str(pdf.parent), ""))
                added += 1
        n = len(self.files)
        self.status.set(f"{n} PDF-uri în listă." if n else "Adaugă PDF-uri pentru a începe.")
        self._update_split_preview()
        return added

    def add_dialog(self) -> None:
        paths = filedialog.askopenfilenames(parent=self, title="Alege PDF-uri",
                                            filetypes=[("PDF", "*.pdf")],
                                            initialdir=self.app.settings.last_dir or None)
        if paths:
            self.app.settings.last_dir = str(Path(paths[0]).parent)
            self.add_paths(paths)

    def remove_selected(self) -> None:
        if self.busy:
            return
        for iid in self.table.selection():
            self.files.pop(iid, None)
            self.pages.pop(iid, None)
        self.table.delete(self.table.selection())
        self._update_split_preview()

    def clear(self) -> None:
        if not self.busy:
            self.files.clear()
            self.pages.clear()
            self.table.clear()
            self._update_split_preview()

    def _move(self, delta: int) -> None:
        if not self.busy:
            self.table.move(self.table.selection(), delta)

    def _open_row(self, event: tk.Event) -> None:
        iid = self.table.tree.identify_row(event.y)
        if iid in self.files:
            os.startfile(self.files[iid])  # type: ignore[attr-defined]

    def store_settings(self, s: Settings) -> None:
        s.watermark_text = self.watermark_text.get().strip() or s.watermark_text
        s.watermark_opacity = int(self.opacity.get())
        s.renumber_pages = self.renumber.get()
        s.split_mode = self.split_mode.get()
        s.split_ranges = self.split_ranges.get().strip()
        for name, var in (("split_every", self.split_every), ("split_parts", self.split_parts)):
            with contextlib.suppress(ValueError):
                setattr(s, name, max(1, int(var.get())))

    # --- executie -------------------------------------------------------------------------

    def start(self) -> None:
        if self.busy:
            return
        rows = self.table.order()
        files = [self.files[iid] for iid in rows]
        op = self.operation.get()
        if not files:
            messagebox.showwarning("Lipsă fișiere", "Adaugă cel puțin un PDF.", parent=self)
            return

        work: Callable[[], None]
        if op == "merge":
            if len(files) < 2:
                messagebox.showwarning("Unire", "Ai nevoie de cel puțin două PDF-uri.", parent=self)
                return
            target = filedialog.asksaveasfilename(
                parent=self, title="Salvează PDF-ul unit", defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")], initialdir=str(files[0].parent),
                initialfile=f"{files[0].stem} (unit).pdf",
            )
            if not target:
                return
            dst = Path(target)
            if any(os.path.normcase(str(dst.resolve())) == iid for iid in rows):
                messagebox.showwarning("Unire", "Alege un nume diferit de fișierele din listă.",
                                       parent=self)
                return

            renumber = self.renumber.get()

            def work() -> None:
                self._merge(rows, files, dst, renumber)
        else:
            per_file = self._per_file_action(op)
            if per_file is None:
                return

            def work() -> None:
                self._each(rows, files, per_file)

        self.app.settings.watermark_text = self.watermark_text.get().strip()
        self.cancel_event.clear()
        for iid in rows:
            self.table.set(iid, "status", "În așteptare", tag="")
        self.progress.configure(maximum=len(files), value=0)
        self._set_busy(True)
        self.thread = threading.Thread(target=self._guard, args=(work,), name="pdf-tools",
                                       daemon=True)
        self.thread.start()

    def _per_file_action(self, op: str) -> Callable[[Path, str], str] | None:
        if op == "protect":
            password = self.password.get()
            if not password:
                messagebox.showwarning("Parolă", "Introdu o parolă.", parent=self)
                return None
            if password != self.password_confirm.get():
                messagebox.showwarning("Parolă", "Parolele nu coincid.", parent=self)
                return None

            def protect(src: Path, iid: str) -> str:
                dst = free_path(src.with_name(f"{src.stem} (protejat).pdf"))
                pdftools.protect(src, dst, password)
                return dst.name
            return protect

        if op == "watermark":
            text = self.watermark_text.get().strip()
            if not text:
                messagebox.showwarning("Watermark", "Scrie textul pentru watermark.", parent=self)
                return None
            opacity = int(self.opacity.get()) / 100

            def stamp(src: Path, iid: str) -> str:
                dst = free_path(src.with_name(f"{src.stem} (watermark).pdf"))
                pdftools.watermark(src, dst, text, opacity)
                return dst.name
            return stamp

        if op == "split":
            # verificam alegerea pe primul PDF inainte de a porni, ca greselile de scriere
            # („1-2OO”) sa apara imediat, nu ca eroare pe fiecare fisier
            rule = self._split_rule()
            first = next((self.pages[i] for i in self.table.order() if self.pages.get(i)), None)
            if first:
                try:
                    rule(first)
                except pdftools.PdfError as exc:
                    messagebox.showwarning("Împărțire", str(exc), parent=self)
                    return None

            def split(src: Path, iid: str) -> str:
                ranges = rule(pdftools.page_count(src))
                out_dir = free_path(src.parent / f"{src.stem} - pagini")
                written = pdftools.split(src, out_dir, ranges)
                return f"{len(written)} fișiere în „{out_dir.name}”"
            return split

        def ocr(src: Path, iid: str) -> str:
            dst = free_path(src.with_name(f"{src.stem} (OCR).pdf"))

            def progress(page: int, total: int) -> None:
                self.ui.post(lambda: self.table.set(iid, "status", f"OCR pagina {page}/{total}…",
                                                    tag="running"))

            pdftools.ocr(src, dst, progress=progress, cancel=self.cancel_event)
            return dst.name
        return ocr

    def _guard(self, work: Callable[[], None]) -> None:
        try:
            work()
        except Exception as exc:
            log.exception("Unealta PDF a esuat")
            message = str(exc)
            self.ui.post(lambda: messagebox.showerror("Eroare", message, parent=self))
        finally:
            self.ui.post(self._finished)

    def _merge(self, rows: list[str], files: list[Path], dst: Path, renumber: bool) -> None:
        for iid in rows:
            self.ui.post(lambda iid=iid: self.table.set(iid, "status", "Se unește…", tag="running"))
        try:
            replaced = pdftools.merge(files, dst, renumber=renumber)
        except Exception as exc:
            message = str(exc)
            for iid in rows:
                self.ui.post(lambda iid=iid: self.table.set(iid, "status", f"Eroare: {message}",
                                                            tag="failed"))
            raise
        for iid in rows:
            self.ui.post(lambda iid=iid: self.table.set(iid, "status", f"Gata → {dst.name}",
                                                        tag="done"))
        self.ui.post(lambda: self.progress.configure(value=len(files)))
        note = ""
        if renumber:
            note = (f" Am renumerotat {replaced} numere de pagină." if replaced else
                    " Nu am găsit numere de pagină tipărite de renumerotat.")
        self.ui.post(lambda: self.status.set(f"Am creat {dst}.{note}"))

    def _each(self, rows: list[str], files: list[Path],
              action: Callable[[Path, str], str]) -> None:
        failures = 0
        for iid, src in zip(rows, files, strict=True):
            if self.cancel_event.is_set():
                self.ui.post(lambda iid=iid: self.table.set(iid, "status", "Anulat", tag="muted"))
                continue
            self.ui.post(lambda iid=iid: self.table.set(iid, "status", "Se procesează…",
                                                        tag="running"))
            try:
                note = action(src, iid)
                self.ui.post(lambda iid=iid, note=note: self.table.set(
                    iid, "status", f"Gata → {note}", tag="done"))
            except pdftools.Cancelled:
                self.ui.post(lambda iid=iid: self.table.set(iid, "status", "Anulat", tag="muted"))
            except Exception as exc:
                failures += 1
                log.warning("Unealta PDF pe %s: %s", src, exc)
                message = str(exc)
                self.ui.post(lambda iid=iid, m=message: self.table.set(
                    iid, "status", f"Eroare: {m}", tag="failed"))
            self.ui.post(lambda: self.progress.step(1))
        ok = len(files) - failures
        self.ui.post(lambda: self.status.set(
            f"Gata: {ok} reușite, {failures} cu erori." if failures else f"Gata: {ok} reușite."))

    def _finished(self) -> None:
        self._set_busy(False)
        self.app.notify("Unelte PDF", self.status.get())

    def _set_busy(self, busy: bool) -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        for widget in [*self.buttons, self.run_button]:
            widget.configure(state=state)
        self.cancel_button.configure(state=tk.NORMAL if busy else tk.DISABLED)
        if busy:
            self.status.set("Se lucrează…")

    def cancel(self) -> None:
        self.cancel_event.set()
        self.cancel_button.configure(state=tk.DISABLED)
        self.status.set("Se oprește după fișierul curent…")
