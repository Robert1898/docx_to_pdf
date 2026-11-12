"""Tab-ul de unelte PDF: unire, impartire, comprimare, imagini, parola, watermark, OCR."""

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
from ..i18n import _
from ..settings import Settings
from .convert_tab import IMAGES, choice_box
from .widgets import Column, FileTable, UiQueue, icon_button, tooltip

if TYPE_CHECKING:
    from .app import App

log = logging.getLogger(__name__)

PREVIEW_WIDTH = 150


def operations() -> tuple[tuple[str, str], ...]:
    return (("merge", _("Unește")), ("split", _("Împarte")), ("compress", _("Comprimă")),
            ("images", _("În imagini")), ("protect", _("Parolă")),
            ("watermark", _("Watermark")), ("ocr", _("OCR")))


def hints() -> dict[str, str]:
    return {
        "merge": _("Toate fișierele din listă devin un singur PDF, în ordinea din tabel "
                   "(trage rândurile sau folosește ↑ ↓). Imaginile devin pagini."),
        "split": _("Fiecare parte devine un PDF separat, într-un folder nou lângă original."),
        "compress": _("Imaginile mari sunt micșorate și recomprimate; textul rămâne intact. "
                      "Dacă rezultatul n-ar fi mai mic, se păstrează originalul."),
        "images": _("Fiecare pagină devine o imagine, într-un folder nou lângă original."),
        "protect": _("PDF-ul nu se poate deschide fără parolă (criptare AES-256). "
                     "Parola nu se salvează."),
        "watermark": _("Textul apare în diagonală, semitransparent, pe fiecare pagină."),
        "ocr": _("PDF-urile scanate primesc un strat de text: devin căutabile și se pot copia."),
    }


def compress_labels() -> tuple[tuple[str, str], ...]:
    return (("light", _("Ușor (calitate maximă)")), ("balanced", _("Echilibrat")),
            ("strong", _("Puternic (fișier minim)")))


def _size(n: int) -> str:
    return f"{n / 1024 / 1024:.1f} MB" if n >= 1024 * 1024 else f"{n / 1024:.0f} KB"


class PdfToolsTab(ttk.Frame):
    def __init__(self, master: tk.Misc, app: App) -> None:
        super().__init__(master, padding=12)
        self.app = app
        s = app.settings
        self.files: dict[str, Path] = {}
        self.pages: dict[str, int | None] = {}
        self.ui = UiQueue(self)
        self.thread: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self._preview_image: object = None

        self.operation = tk.StringVar(value="merge")
        self.password = tk.StringVar()
        self.password_confirm = tk.StringVar()
        self.watermark_text = tk.StringVar(value=s.watermark_text)
        self.opacity = tk.IntVar(value=s.watermark_opacity)
        self.renumber = tk.BooleanVar(value=s.renumber_pages)
        self.split_mode = tk.StringVar(value=s.split_mode)
        self.split_every = tk.StringVar(value=str(s.split_every))
        self.split_parts = tk.StringVar(value=str(s.split_parts))
        self.split_ranges = tk.StringVar(value=s.split_ranges)
        self.split_preview = tk.StringVar()
        self.compress_level = tk.StringVar(value=s.compress_level)
        self.images_format = tk.StringVar(value=s.images_format)
        self.images_dpi = tk.StringVar(value=str(s.images_dpi))
        self.hint = tk.StringVar()
        self.status = tk.StringVar(value=_("Adaugă PDF-uri pentru a începe."))
        self._build()
        self._on_operation()
        for var in (self.split_mode, self.split_every, self.split_parts, self.split_ranges):
            var.trace_add("write", lambda *_a: self._update_split_preview())

    @property
    def busy(self) -> bool:
        return self.thread is not None and self.thread.is_alive()

    # --- construire interfata -------------------------------------------------------------

    def _build(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X)
        self.buttons = [
            icon_button(bar, _("Adaugă fișiere…"), self.add_dialog, "primary", "add"),
            icon_button(bar, _("Elimină"), self.remove_selected, "secondary-outline", "remove"),
            icon_button(bar, _("Golește lista"), self.clear, "secondary-outline", "clear"),
            icon_button(bar, "", lambda: self._move(-1), "secondary-outline", "up", width=3),
            icon_button(bar, "", lambda: self._move(1), "secondary-outline", "down", width=3),
        ]
        for i, button in enumerate(self.buttons):
            button.pack(side=tk.LEFT, padx=(0 if i == 0 else 6, 0))
        tooltip(self.buttons[3], _("Mută în sus"))
        tooltip(self.buttons[4], _("Mută în jos"))

        middle = ttk.Frame(self)
        middle.pack(fill=tk.BOTH, expand=True, pady=10)
        columns = (Column("name", _("Fișier"), 250), Column("pages", _("Pagini"), 60, False),
                   Column("folder", _("Folder"), 200), Column("status", _("Status"), 220))
        self.table = FileTable(middle, columns, _("Trage aici fișiere PDF (sau imagini, pentru "
                                                  "unire)."), on_drop=self.add_paths,
                               reorderable=True)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.table.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.table.tree.bind("<Double-1>", self._open_row, add="+")
        self.table.tree.bind("<<TreeviewSelect>>", lambda e: self._on_select())

        preview = ttk.Frame(middle, width=PREVIEW_WIDTH + 20)
        preview.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 0))
        preview.pack_propagate(False)
        ttk.Label(preview, text=_("Previzualizare"), bootstyle="secondary").pack(anchor=tk.W)
        self.preview_label = ttk.Label(preview, anchor=tk.N)
        self.preview_label.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        box = ttk.Labelframe(self, text=_("Operație"), padding=(10, 8))
        box.pack(fill=tk.X)
        radios = ttk.Frame(box)
        radios.pack(fill=tk.X)
        self.radios: dict[str, ttk.Radiobutton] = {}
        for key, text in operations():
            radio = ttk.Radiobutton(radios, text=text, value=key, variable=self.operation,
                                    command=self._on_operation,
                                    bootstyle="primary-outline-toolbutton")
            radio.pack(side=tk.LEFT, padx=(0, 6))
            self.radios[key] = radio
        if self.app.backends.tesseract is None:
            self.radios["ocr"].configure(state=tk.DISABLED)

        ttk.Label(box, textvariable=self.hint, bootstyle="secondary", wraplength=820).pack(
            anchor=tk.W, pady=(8, 4))

        self.params = ttk.Frame(box)
        self.params.pack(fill=tk.X)
        self.param_frames: dict[str, ttk.Frame] = {}
        self._build_merge()
        self._build_split()
        self._build_compress()
        self._build_images()
        self._build_protect()
        self._build_watermark()

        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, pady=(12, 0))
        self.progress = ttk.Progressbar(bottom, mode="determinate", bootstyle="info")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cancel_button = icon_button(bottom, _("Anulează"), self.cancel, "danger-outline",
                                         "stop", state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT, padx=(10, 0))
        self.run_button = icon_button(bottom, _("Execută"), self.start, "info", "play")
        self.run_button.pack(side=tk.LEFT, padx=(6, 0))
        tooltip(self.run_button, "Ctrl+Enter")
        ttk.Label(self, textvariable=self.status, bootstyle="secondary").pack(
            fill=tk.X, pady=(6, 0))

    def _build_merge(self) -> None:
        merge = ttk.Frame(self.params)
        renumber = ttk.Checkbutton(
            merge, text=_("Renumerotează paginile (ex. 8/8 și 1/2 devin 8/10 și 9/10)"),
            variable=self.renumber)
        renumber.pack(anchor=tk.W)
        tooltip(renumber, _("Numerele de pagină tipărite în antet sau subsol („8/8”, "
                            "„Pagina 1 din 2”) sunt înlocuite cu numerotarea documentului unit, "
                            "în același font și loc."))
        self.param_frames["merge"] = merge

    def _build_split(self) -> None:
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

        mode_row("each", _("Fiecare pagină separat"))
        row = mode_row("every", _("Câte"))
        field(ttk.Spinbox(row, from_=1, to=100000, width=7, textvariable=self.split_every),
              "every")
        ttk.Label(row, text=_("pagini într-un fișier")).pack(side=tk.LEFT)
        row = mode_row("parts", _("În"))
        field(ttk.Spinbox(row, from_=1, to=10000, width=7, textvariable=self.split_parts),
              "parts")
        ttk.Label(row, text=_("părți egale")).pack(side=tk.LEFT)
        row = mode_row("ranges", _("Intervale:"))
        field(ttk.Entry(row, textvariable=self.split_ranges, width=30), "ranges")
        ttk.Label(row, text=_("ex. 1-200, 201-450, 451-  (gol după cratimă = până la sfârșit)"),
                  bootstyle="secondary").pack(side=tk.LEFT)
        self.split_preview_label = ttk.Label(split, textvariable=self.split_preview,
                                             bootstyle="info")
        self.split_preview_label.pack(anchor=tk.W, pady=(6, 0))
        self.param_frames["split"] = split

    def _build_compress(self) -> None:
        frame = ttk.Frame(self.params)
        ttk.Label(frame, text=_("Nivel:")).pack(side=tk.LEFT)
        choice_box(frame, compress_labels(), self.compress_level, 26).pack(side=tk.LEFT, padx=6)
        self.param_frames["compress"] = frame

    def _build_images(self) -> None:
        frame = ttk.Frame(self.params)
        ttk.Label(frame, text=_("Format:")).pack(side=tk.LEFT)
        choice_box(frame, (("png", "PNG"), ("jpg", "JPG")), self.images_format, 6).pack(
            side=tk.LEFT, padx=6)
        ttk.Label(frame, text=_("Rezoluție:")).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Spinbox(frame, from_=72, to=600, increment=50, width=5,
                    textvariable=self.images_dpi).pack(side=tk.LEFT, padx=6)
        ttk.Label(frame, text="DPI").pack(side=tk.LEFT)
        self.param_frames["images"] = frame

    def _build_protect(self) -> None:
        protect = ttk.Frame(self.params)
        ttk.Label(protect, text=_("Parolă:")).grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(protect, textvariable=self.password, show="•", width=28).grid(
            row=0, column=1, padx=6, pady=2)
        ttk.Label(protect, text=_("Confirmă:")).grid(row=0, column=2, sticky=tk.W, padx=(12, 0))
        ttk.Entry(protect, textvariable=self.password_confirm, show="•", width=28).grid(
            row=0, column=3, padx=6, pady=2)
        self.param_frames["protect"] = protect

    def _build_watermark(self) -> None:
        mark = ttk.Frame(self.params)
        ttk.Label(mark, text=_("Text:")).grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(mark, textvariable=self.watermark_text, width=34).grid(
            row=0, column=1, padx=6, pady=2)
        ttk.Label(mark, text=_("Opacitate:")).grid(row=0, column=2, sticky=tk.W, padx=(12, 0))
        ttk.Scale(mark, from_=5, to=100, variable=self.opacity, length=160,
                  command=lambda v: self.opacity.set(round(float(v)))).grid(
            row=0, column=3, padx=6)
        ttk.Label(mark, textvariable=self.opacity, width=4).grid(row=0, column=4, sticky=tk.W)
        ttk.Label(mark, text="%").grid(row=0, column=5, sticky=tk.W)
        self.param_frames["watermark"] = mark

    def _on_operation(self) -> None:
        op = self.operation.get()
        self.hint.set(hints()[op])
        if op == "ocr" and self.app.backends.tesseract is not None:
            langs = pdftools.ocr_languages(self.app.backends.tesseract)
            self.hint.set(hints()[op] + " " + _("Limbi folosite: {langs}.").format(langs=langs))
        for key, frame in self.param_frames.items():
            if key == op:
                frame.pack(anchor=tk.W, pady=(4, 0))
            else:
                frame.pack_forget()
        self._update_split_preview()

    # --- previzualizare -------------------------------------------------------------------

    def _on_select(self) -> None:
        self._update_split_preview()
        selected = self.table.selection()
        if not selected or selected[0] not in self.files:
            return
        path = self.files[selected[0]]

        def render() -> None:
            try:
                image = _thumbnail(path, PREVIEW_WIDTH)
            except Exception as exc:  # PDF protejat, imagine stricata
                log.debug("Previzualizare %s: %s", path, exc)
                image = None
            self.ui.post(lambda: self._show_preview(image))

        threading.Thread(target=render, name="preview", daemon=True).start()

    def _show_preview(self, image: object) -> None:
        if image is None:
            self.preview_label.configure(image="", text=_("fără previzualizare"))
            return
        from PIL import ImageTk

        self._preview_image = ImageTk.PhotoImage(image)
        self.preview_label.configure(image=self._preview_image, text="")

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
                raise pdftools.PdfError(_("Scrie un număr întreg pentru {what}.").format(
                    what=what)) from None

        def rule(total: int) -> list[pdftools.PageRange]:
            if mode == "every":
                return pdftools.ranges_every(total, number(every, _("paginile pe fișier")))
            if mode == "parts":
                return pdftools.ranges_parts(total, number(parts, _("numărul de părți")))
            if mode == "ranges":
                return pdftools.parse_ranges(text, total)
            return pdftools.ranges_every(total, 1)

        return rule

    def _update_split_preview(self) -> None:
        if self.operation.get() != "split":
            return
        selected = [iid for iid in self.table.selection() if iid in self.files]
        iid = (selected or [i for i in self.table.order() if self.pages.get(i)] or [None])[0]
        if iid is None:
            self.split_preview.set(_("Adaugă un PDF ca să vezi cum va fi împărțit."))
            self.split_preview_label.configure(bootstyle="secondary")
            return
        total = self.pages.get(iid)
        name = self.files[iid].name
        if not total:
            self.split_preview.set(_("„{name}” nu poate fi împărțit (nu e PDF, e protejat sau "
                                     "deteriorat).").format(name=name))
            self.split_preview_label.configure(bootstyle="danger")
            return
        try:
            ranges = self._split_rule()(total)
        except pdftools.PdfError as exc:
            self.split_preview.set(str(exc))
            self.split_preview_label.configure(bootstyle="danger")
            return
        self.split_preview.set(_("„{name}” ({total} pagini) → {count} fișier(e): {ranges}")
                               .format(name=name, total=total, count=len(ranges),
                                       ranges=pdftools.describe_ranges(ranges)))
        self.split_preview_label.configure(bootstyle="info")

    # --- lista ----------------------------------------------------------------------------

    def add_paths(self, paths: Sequence[str | os.PathLike[str]]) -> int:
        if self.busy:
            return 0
        added = 0
        for raw in paths:
            path = Path(raw)
            candidates = (sorted(p for p in path.rglob("*") if p.suffix.lower() == ".pdf"
                                 or pdftools.is_image(p)) if path.is_dir() else [path])
            for item in candidates:
                if not item.is_file() or not (item.suffix.lower() == ".pdf"
                                              or pdftools.is_image(item)):
                    continue
                iid = os.path.normcase(str(item.resolve()))
                if iid in self.files:
                    continue
                count: int | None = None
                if pdftools.is_image(item):
                    pages = _("imagine")
                else:
                    try:
                        count = pdftools.page_count(item)
                        pages = str(count)
                    except pdftools.PdfError as exc:
                        pages = "🔒" if _("parolă") in str(exc) else "?"
                self.files[iid] = item.resolve()
                self.pages[iid] = count
                self.table.insert(iid, (item.name, pages, str(item.parent), ""))
                added += 1
        n = len(self.files)
        self.status.set(_("{n} fișier(e) în listă.").format(n=n) if n else
                        _("Adaugă PDF-uri pentru a începe."))
        self._update_split_preview()
        if added and not self.table.selection():
            first = self.table.order()[0]
            self.table.tree.selection_set(first)  # previzualizarea arata ceva imediat
            self.table.tree.focus(first)
        return added

    def add_dialog(self) -> None:
        paths = filedialog.askopenfilenames(
            parent=self, title=_("Alege PDF-uri sau imagini"),
            filetypes=[("PDF", "*.pdf"), (_("Imagini"), IMAGES),
                       (_("PDF și imagini"), f"*.pdf {IMAGES}")],
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
            self.preview_label.configure(image="", text="")

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
        s.compress_level = self.compress_level.get()
        s.images_format = self.images_format.get()
        for name, var in (("split_every", self.split_every), ("split_parts", self.split_parts),
                          ("images_dpi", self.images_dpi)):
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
            messagebox.showwarning(_("Lipsă fișiere"), _("Adaugă cel puțin un PDF."), parent=self)
            return
        if op != "merge" and any(pdftools.is_image(f) for f in files):
            messagebox.showwarning(_("Imagini în listă"), _(
                "Imaginile pot fi doar unite cu PDF-uri. Pentru a le transforma în PDF-uri "
                "separate, folosește tab-ul Conversie."), parent=self)
            return

        work: Callable[[], None]
        if op == "merge":
            if len(files) < 2:
                messagebox.showwarning(_("Unire"), _("Ai nevoie de cel puțin două fișiere."),
                                       parent=self)
                return
            target = filedialog.asksaveasfilename(
                parent=self, title=_("Salvează PDF-ul unit"), defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")], initialdir=str(files[0].parent),
                initialfile=_("{name} (unit).pdf").format(name=files[0].stem))
            if not target:
                return
            dst = Path(target)
            if any(os.path.normcase(str(dst.resolve())) == iid for iid in rows):
                messagebox.showwarning(_("Unire"), _("Alege un nume diferit de fișierele din "
                                                     "listă."), parent=self)
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

        self.store_settings(self.app.settings)
        self.cancel_event.clear()
        for iid in rows:
            self.table.set(iid, "status", _("În așteptare"), tag="")
        self.progress.configure(maximum=len(files), value=0)
        self._set_busy(True)
        self.thread = threading.Thread(target=self._guard, args=(work,), name="pdf-tools",
                                       daemon=True)
        self.thread.start()

    def _per_file_action(self, op: str) -> Callable[[Path, str], str] | None:
        if op == "protect":
            password = self.password.get()
            if not password:
                messagebox.showwarning(_("Parolă"), _("Introdu o parolă."), parent=self)
                return None
            if password != self.password_confirm.get():
                messagebox.showwarning(_("Parolă"), _("Parolele nu coincid."), parent=self)
                return None

            def protect(src: Path, iid: str) -> str:
                dst = free_path(src.with_name(_("{name} (protejat).pdf").format(name=src.stem)))
                pdftools.protect(src, dst, password)
                return dst.name
            return protect

        if op == "watermark":
            text = self.watermark_text.get().strip()
            if not text:
                messagebox.showwarning(_("Watermark"), _("Scrie textul pentru watermark."),
                                       parent=self)
                return None
            opacity = int(self.opacity.get()) / 100

            def stamp(src: Path, iid: str) -> str:
                dst = free_path(src.with_name(_("{name} (watermark).pdf").format(name=src.stem)))
                pdftools.watermark(src, dst, text, opacity)
                return dst.name
            return stamp

        if op == "compress":
            level = self.compress_level.get()

            def compress(src: Path, iid: str) -> str:
                dst = free_path(src.with_name(_("{name} (comprimat).pdf").format(
                    name=src.stem)))
                before, after = pdftools.compress(src, dst, level, cancel=self.cancel_event)
                saved = 100 - after * 100 // max(before, 1)
                return _("{before} → {after} (−{saved}%)").format(
                    before=_size(before), after=_size(after), saved=saved)
            return compress

        if op == "images":
            fmt = self.images_format.get()
            try:
                dpi = max(72, min(int(self.images_dpi.get()), 600))
            except ValueError:
                dpi = 150

            def to_images(src: Path, iid: str) -> str:
                out_dir = free_path(src.parent / _("{name} - imagini").format(name=src.stem))

                def progress(page: int, total: int) -> None:
                    self.ui.post(lambda: self.table.set(
                        iid, "status", _("Pagina {page}/{total}…").format(page=page,
                                                                          total=total),
                        tag="running"))

                written = pdftools.pdf_to_images(src, out_dir, fmt, dpi, progress=progress,
                                                 cancel=self.cancel_event)
                return _("{n} imagini în „{folder}”").format(n=len(written), folder=out_dir.name)
            return to_images

        if op == "split":
            # verificam alegerea pe primul PDF inainte de a porni, ca greselile de scriere
            # („1-2OO”) sa apara imediat, nu ca eroare pe fiecare fisier
            rule = self._split_rule()
            first = next((self.pages[i] for i in self.table.order() if self.pages.get(i)), None)
            if first:
                try:
                    rule(first)
                except pdftools.PdfError as exc:
                    messagebox.showwarning(_("Împărțire"), str(exc), parent=self)
                    return None

            def split(src: Path, iid: str) -> str:
                ranges = rule(pdftools.page_count(src))
                out_dir = free_path(src.parent / _("{name} - pagini").format(name=src.stem))
                written = pdftools.split(src, out_dir, ranges)
                return _("{n} fișiere în „{folder}”").format(n=len(written), folder=out_dir.name)
            return split

        def ocr(src: Path, iid: str) -> str:
            dst = free_path(src.with_name(_("{name} (OCR).pdf").format(name=src.stem)))

            def progress(page: int, total: int) -> None:
                self.ui.post(lambda: self.table.set(
                    iid, "status", _("OCR pagina {page}/{total}…").format(page=page,
                                                                          total=total),
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
            self.ui.post(lambda: messagebox.showerror(_("Eroare"), message, parent=self))
        finally:
            self.ui.post(self._finished)

    def _merge(self, rows: list[str], files: list[Path], dst: Path, renumber: bool) -> None:
        for iid in rows:
            self.ui.post(lambda iid=iid: self.table.set(iid, "status", _("Se unește…"),
                                                        tag="running"))
        try:
            replaced = pdftools.merge(files, dst, renumber=renumber)
        except Exception as exc:
            message = str(exc)
            for iid in rows:
                self.ui.post(lambda iid=iid: self.table.set(
                    iid, "status", _("Eroare: {message}").format(message=message),
                    tag="failed"))
            raise
        for iid in rows:
            self.ui.post(lambda iid=iid: self.table.set(
                iid, "status", _("Gata → {name}").format(name=dst.name), tag="done"))
        self.ui.post(lambda: self.progress.configure(value=len(files)))
        note = ""
        if renumber:
            note = (" " + _("Am renumerotat {n} numere de pagină.").format(n=replaced)
                    if replaced else " " + _("Nu am găsit numere de pagină tipărite de "
                                             "renumerotat."))
        self.ui.post(lambda: self.status.set(_("Am creat {path}.").format(path=dst) + note))

    def _each(self, rows: list[str], files: list[Path],
              action: Callable[[Path, str], str]) -> None:
        failures = 0
        for iid, src in zip(rows, files, strict=True):
            if self.cancel_event.is_set():
                self.ui.post(lambda iid=iid: self.table.set(iid, "status", _("Anulat"),
                                                            tag="muted"))
                continue
            self.ui.post(lambda iid=iid: self.table.set(iid, "status", _("Se procesează…"),
                                                        tag="running"))
            try:
                note = action(src, iid)
                self.ui.post(lambda iid=iid, note=note: self.table.set(
                    iid, "status", _("Gata → {name}").format(name=note), tag="done"))
            except pdftools.Cancelled:
                self.ui.post(lambda iid=iid: self.table.set(iid, "status", _("Anulat"),
                                                            tag="muted"))
            except Exception as exc:
                failures += 1
                log.warning("Unealta PDF pe %s: %s", src, exc)
                message = str(exc)
                self.ui.post(lambda iid=iid, m=message: self.table.set(
                    iid, "status", _("Eroare: {message}").format(message=m), tag="failed"))
            self.ui.post(lambda: self.progress.step(1))
        ok = len(files) - failures
        self.ui.post(lambda: self.status.set(
            _("Gata: {ok} reușite, {failed} cu erori.").format(ok=ok, failed=failures)
            if failures else _("Gata: {ok} reușite.").format(ok=ok)))

    def _finished(self) -> None:
        self._set_busy(False)
        self.app.notify(_("Unelte PDF"), self.status.get())

    def _set_busy(self, busy: bool) -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        for widget in [*self.buttons, self.run_button]:
            widget.configure(state=state)
        self.cancel_button.configure(state=tk.NORMAL if busy else tk.DISABLED)
        if busy:
            self.status.set(_("Se lucrează…"))

    def cancel(self) -> None:
        self.cancel_event.set()
        self.cancel_button.configure(state=tk.DISABLED)
        self.status.set(_("Se oprește după fișierul curent…"))


def _thumbnail(path: Path, width: int) -> object:
    """Prima pagina a unui PDF (sau imaginea) micsorata la latimea data."""
    from PIL import Image, ImageOps

    if pdftools.is_image(path):
        image = Image.open(path)
        image.thumbnail((width, width * 2))
        image = image.convert("RGB")
    else:
        import pypdfium2 as pdfium

        document = pdfium.PdfDocument(str(path))
        try:
            page = document[0]
            image = page.render(scale=width / page.get_width()).to_pil().convert("RGB")
        finally:
            document.close()
    return ImageOps.expand(image, border=1, fill="#9aa0a6")  # pagina alba pe fundal alb
