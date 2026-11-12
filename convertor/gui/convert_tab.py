"""Tab-ul principal: documente si imagini -> PDF, PDF -> DOCX sau Excel."""

from __future__ import annotations

import contextlib
import logging
import os
import threading
import time
import tkinter as tk
from collections.abc import Sequence
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING

import ttkbootstrap as ttk

from ..core import Job, collect_files, kind_of, label, plan_jobs
from ..engine import Engine, Result, Status
from ..i18n import _
from ..settings import MAX_WORKERS, Settings
from .taskbar import ERROR, TaskbarProgress
from .widgets import Column, FileTable, UiQueue, icon_button, tooltip

if TYPE_CHECKING:
    from .app import App

log = logging.getLogger(__name__)

IMAGES = "*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff *.webp"


def filetypes() -> list[tuple[str, str]]:
    return [
        (_("Toate documentele"), "*.docx *.doc *.docm *.rtf *.odt *.xlsx *.xls *.xlsm *.ods "
                                 f"*.pptx *.ppt *.pptm *.odp *.pdf {IMAGES}"),
        ("Word", "*.docx *.doc *.docm *.rtf *.odt"),
        ("Excel", "*.xlsx *.xls *.xlsm *.ods"),
        ("PowerPoint", "*.pptx *.ppt *.pptm *.odp"),
        ("PDF", "*.pdf"),
        (_("Imagini"), IMAGES),
    ]


def pdf_mode_labels() -> tuple[tuple[str, str], ...]:
    return (("auto", _("Automat (recomandat)")),
            ("layout", _("Aspect fidel (rapoarte, tabele)")),
            ("word", _("Text editabil (Word)")))


def pdf_target_labels() -> tuple[tuple[str, str], ...]:
    return (("docx", _("Word (DOCX)")), ("xlsx", _("Excel (XLSX)")))


def image_page_labels() -> tuple[tuple[str, str], ...]:
    return (("a4", _("încadrată pe A4")), ("image", _("la mărimea imaginii")))


STATUS_TAGS = {
    Status.PENDING: "",
    Status.RUNNING: "running",
    Status.DONE: "done",
    Status.FAILED: "failed",
    Status.SKIPPED: "muted",
    Status.CANCELLED: "muted",
}


def _duration(seconds: float) -> str:
    seconds = round(seconds)
    if seconds < 60:
        return _("{s} s").format(s=max(seconds, 1))
    return _("{m} min {s:02d} s").format(m=seconds // 60, s=seconds % 60)


def choice_box(master: tk.Misc, labels: Sequence[tuple[str, str]], variable: tk.StringVar,
               width: int, on_change: object = None) -> ttk.Combobox:
    """Combobox care afiseaza textul tradus, dar memoreaza cheia in `variable`."""
    box = ttk.Combobox(master, state="readonly", width=width, values=[t for _k, t in labels])
    box.set(dict(labels).get(variable.get(), labels[0][1]))

    def changed(_event: tk.Event) -> None:
        variable.set(next(k for k, text in labels if text == box.get()))
        if callable(on_change):
            on_change()

    box.bind("<<ComboboxSelected>>", changed)
    return box


class ConvertTab(ttk.Frame):
    def __init__(self, master: tk.Misc, app: App) -> None:
        super().__init__(master, padding=12)
        self.app = app
        s = app.settings
        self.files: dict[str, Path] = {}
        self.results: dict[str, Result] = {}
        self.engine: Engine | None = None
        self.thread: threading.Thread | None = None
        self.job_rows: dict[int, str] = {}
        self.pending: list[str] = []  # fisiere primite in timpul unei conversii
        self.pending_auto = False
        self._auto_after: str | None = None
        self._started_at = 0.0
        self._durations: list[float] = []
        self.ui = UiQueue(self)
        self.taskbar = TaskbarProgress(app.root)

        self.same_folder = tk.BooleanVar(value=s.same_folder)
        self.output_dir = tk.StringVar(value=s.output_dir)
        self.open_after = tk.BooleanVar(value=s.open_after)
        self.bookmarks = tk.BooleanVar(value=s.bookmarks)
        self.pdfa = tk.BooleanVar(value=s.pdfa)
        self.for_screen = tk.BooleanVar(value=s.for_screen)
        self.ocr = tk.BooleanVar(value=s.ocr and app.backends.tesseract is not None)
        self.timeout = tk.StringVar(value=str(s.timeout))
        self.pdf_mode = tk.StringVar(value=s.pdf_mode)
        self.pdf_target = tk.StringVar(value=s.pdf_target)
        self.workers = tk.StringVar(value=str(s.workers))
        self.image_page = tk.StringVar(value=s.image_page)
        self.status = tk.StringVar()

        self._build()
        self._toggle_output()
        self._update_count()

    @property
    def busy(self) -> bool:
        return self.thread is not None and self.thread.is_alive()

    # --- construire interfata -------------------------------------------------------------

    def _build(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X)
        self.add_buttons = [
            icon_button(bar, _("Adaugă fișiere…"), self.add_files_dialog, "primary", "add"),
            icon_button(bar, _("Adaugă folder…"), self.add_folder_dialog, "primary-outline",
                        "folder"),
            icon_button(bar, _("Elimină"), self.remove_selected, "secondary-outline", "remove"),
            icon_button(bar, _("Golește lista"), self.clear, "secondary-outline", "clear"),
        ]
        for i, button in enumerate(self.add_buttons):
            button.pack(side=tk.LEFT, padx=(0 if i == 0 else 6, 0))
        tooltip(self.add_buttons[0], "Ctrl+O")
        tooltip(self.add_buttons[1], "Ctrl+Shift+O")
        tooltip(self.add_buttons[2], "Delete")
        self.count_label = ttk.Label(bar, bootstyle="secondary")
        self.count_label.pack(side=tk.RIGHT)

        placeholder = _("Trage aici documente Word, Excel, PowerPoint, PDF sau imagini\n"
                        "(sau foldere întregi), ori folosește butoanele de mai sus.")
        columns = (Column("name", _("Fișier"), 250), Column("folder", _("Folder"), 220),
                   Column("kind", _("Conversie"), 140, stretch=False),
                   Column("status", _("Status"), 240))
        self.table = FileTable(self, columns, placeholder, on_drop=self.add_paths,
                               on_activate=self._activate, reorderable=True)
        self.table.pack(fill=tk.BOTH, expand=True, pady=10)
        self.table.bind_context_menu(self._context_menu)
        self.table.tree.bind("<Delete>", lambda e: self.remove_selected())

        options = ttk.Labelframe(self, text=_("Opțiuni"), padding=(10, 6))
        options.pack(fill=tk.X)
        bookmarks = ttk.Checkbutton(options, text=_("Semne de carte din titluri"),
                                    variable=self.bookmarks)
        pdfa = ttk.Checkbutton(options, text=_("PDF/A (arhivare)"), variable=self.pdfa)
        screen = ttk.Checkbutton(options, text=_("Optimizat pentru ecran (fișier mai mic)"),
                                 variable=self.for_screen)
        ocr = ttk.Checkbutton(options, text=_("OCR pentru PDF-uri scanate"), variable=self.ocr)
        bookmarks.grid(row=0, column=0, sticky=tk.W, padx=(0, 18), pady=3)
        pdfa.grid(row=0, column=1, sticky=tk.W, padx=(0, 18), pady=3)
        screen.grid(row=0, column=2, sticky=tk.W, pady=3)
        ocr.grid(row=1, column=0, sticky=tk.W, padx=(0, 18), pady=3)
        tooltip(bookmarks, _("PDF-ul primește un cuprins navigabil din titlurile documentului "
                             "Word."))
        tooltip(pdfa, _("Format standard pentru arhivare pe termen lung, cerut uneori de "
                        "instituții."))
        tooltip(screen, _("Imagini la rezoluție mai mică: fișier mai mic, potrivit pentru "
                          "e-mail."))
        if self.app.backends.tesseract is None:
            ocr.configure(state=tk.DISABLED)
            tooltip(ocr, _("Necesită Tesseract OCR instalat (vezi README)."))
        else:
            tooltip(ocr, _("PDF-urile scanate (imagini) sunt trecute prin OCR înainte de "
                           "conversia în DOCX sau Excel, ca să rezulte text editabil."))

        limit = ttk.Frame(options)
        limit.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=3)
        ttk.Label(limit, text=_("Limită per fișier:")).pack(side=tk.LEFT)
        spin = ttk.Spinbox(limit, from_=30, to=3600, increment=30, width=6,
                           textvariable=self.timeout)
        spin.pack(side=tk.LEFT, padx=6)
        ttk.Label(limit, text=_("secunde")).pack(side=tk.LEFT, padx=(0, 18))
        tooltip(spin, _("Dacă un fișier durează mai mult (de obicei un dialog ascuns), "
                        "aplicația Office pornită de program e oprită și se trece la următorul "
                        "fișier."))
        ttk.Label(limit, text=_("Conversii simultane:")).pack(side=tk.LEFT)
        workers = ttk.Spinbox(limit, from_=1, to=MAX_WORKERS, width=3, textvariable=self.workers)
        workers.pack(side=tk.LEFT, padx=6)
        tooltip(workers, _("Câte fișiere se convertesc în același timp, fiecare cu propria "
                           "instanță Office. Ajută la loturi mari; folosește mai multă memorie."))

        pdf_row = ttk.Frame(options)
        pdf_row.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=3)
        ttk.Label(pdf_row, text=_("PDF în:")).pack(side=tk.LEFT)
        choice_box(pdf_row, pdf_target_labels(), self.pdf_target, 14,
                   self._refresh_kinds).pack(side=tk.LEFT, padx=6)
        ttk.Label(pdf_row, text=_("Mod DOCX:")).pack(side=tk.LEFT, padx=(12, 0))
        mode = choice_box(pdf_row, pdf_mode_labels(), self.pdf_mode, 30)
        mode.pack(side=tk.LEFT, padx=6)
        tooltip(mode, _("Aspect fidel: fiecare pagină arată ca în PDF, cu coloanele aliniate "
                        "exact; potrivit pentru rapoarte, balanțe, extrase, tabele.\n"
                        "Text editabil (Word): textul curge liber, potrivit pentru scrisori și "
                        "contracte; Word ghicește tabelele și poate muta valori.\n"
                        "Automat: rapoartele pe coloane merg pe aspect fidel, restul prin Word."))
        ttk.Label(pdf_row, text=_("Imagine → PDF:")).pack(side=tk.LEFT, padx=(12, 0))
        choice_box(pdf_row, image_page_labels(), self.image_page, 20).pack(side=tk.LEFT, padx=6)

        dest = ttk.Labelframe(self, text=_("Destinație"), padding=(10, 6))
        dest.pack(fill=tk.X, pady=(10, 0))
        ttk.Checkbutton(dest, text=_("Lângă fișierul original"), variable=self.same_folder,
                        command=self._toggle_output).pack(anchor=tk.W)
        row = ttk.Frame(dest)
        row.pack(fill=tk.X, pady=6)
        self.out_entry = ttk.Entry(row, textvariable=self.output_dir)
        self.out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.out_button = icon_button(row, _("Alege…"), self.choose_output, "secondary-outline",
                                      "folder")
        self.out_button.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Checkbutton(dest, text=_("Deschide folderul la final"),
                        variable=self.open_after).pack(anchor=tk.W)

        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, pady=(12, 0))
        self.progress = ttk.Progressbar(bottom, mode="determinate", bootstyle="success")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cancel_button = icon_button(bottom, _("Anulează"), self.cancel, "danger-outline",
                                         "stop", state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT, padx=(10, 0))
        tooltip(self.cancel_button, "Esc")
        self.start_button = icon_button(bottom, _("Convertește"), self.start, "success", "play")
        self.start_button.pack(side=tk.LEFT, padx=(6, 0))
        tooltip(self.start_button, "Ctrl+Enter")
        ttk.Label(self, textvariable=self.status, bootstyle="secondary").pack(
            fill=tk.X, pady=(6, 0))

    def _toggle_output(self) -> None:
        state = tk.DISABLED if self.same_folder.get() else tk.NORMAL
        self.out_entry.configure(state=state)
        self.out_button.configure(state=state)

    def _update_count(self) -> None:
        n = len(self.files)
        self.count_label.configure(text=_("{n} fișier(e)").format(n=n) if n else "")
        if not self.busy:
            self.status.set(_("Adaugă documente pentru a începe.") if not n else
                            _("Gata de conversie. Dublu-click pe un rând deschide rezultatul."))

    def _refresh_kinds(self) -> None:
        """PDF -> DOCX sau Excel: actualizeaza coloana „Conversie”."""
        for iid, path in self.files.items():
            kind = kind_of(path)
            if kind is not None:
                self.table.set(iid, "kind", label(kind, self.pdf_target.get()))

    # --- lista de fisiere -----------------------------------------------------------------

    def add_paths(self, paths: Sequence[str | os.PathLike[str]]) -> int:
        if self.busy:
            # vin de la alta instanta sau din Explorer in timpul conversiei: le punem deoparte
            self.pending.extend(str(p) for p in paths)
            self.status.set(_("Am primit {n} fișier(e); le adaug după conversia curentă.")
                            .format(n=len(self.pending)))
            return 0
        added = 0
        for path in collect_files(paths):
            iid = os.path.normcase(str(path))
            if iid in self.files:
                continue
            kind = kind_of(path)
            assert kind is not None
            self.files[iid] = path
            self.table.insert(iid, (path.name, str(path.parent),
                                    label(kind, self.pdf_target.get()), Status.PENDING.text))
            added += 1
        self._update_count()
        if paths and not added:
            self.status.set(_("Nu am găsit fișiere noi suportate."))
        return added

    def receive(self, paths: Sequence[str], auto: bool) -> None:
        """Fisiere primite de la alta instanta (click-dreapta pe mai multe fisiere)."""
        self.add_paths(paths)
        if not auto:
            return
        if self.busy:
            self.pending_auto = True
            return
        # Explorer porneste cate o instanta per fisier: asteptam sa soseasca toate
        if self._auto_after is not None:
            self.after_cancel(self._auto_after)
        self._auto_after = self.after(900, self._auto_start)

    def _auto_start(self) -> None:
        self._auto_after = None
        if not self.busy and self.files:
            self.start(ask_existing=False)

    def add_files_dialog(self) -> None:
        paths = filedialog.askopenfilenames(
            parent=self, title=_("Alege documente"), filetypes=filetypes(),
            initialdir=self.app.settings.last_dir or None)
        if paths:
            self.app.settings.last_dir = str(Path(paths[0]).parent)
            self.add_paths(paths)

    def add_folder_dialog(self) -> None:
        folder = filedialog.askdirectory(parent=self, title=_("Alege un folder"),
                                         initialdir=self.app.settings.last_dir or None)
        if folder:
            self.app.settings.last_dir = folder
            if not self.add_paths([folder]):
                messagebox.showinfo(_("Niciun fișier"),
                                    _("Nu am găsit documente noi în folderul ales."), parent=self)

    def remove_selected(self) -> None:
        if self.busy:
            return
        for iid in self.table.selection():
            self.files.pop(iid, None)
            self.results.pop(iid, None)
        self.table.delete(self.table.selection())
        self._update_count()

    def clear(self) -> None:
        if self.busy:
            return
        self.files.clear()
        self.results.clear()
        self.table.clear()
        self._update_count()

    def choose_output(self) -> None:
        folder = filedialog.askdirectory(parent=self, title=_("Alege folderul de destinație"))
        if folder:
            self.output_dir.set(os.path.normpath(folder))

    # --- interactiuni cu randurile --------------------------------------------------------

    def _activate(self, iid: str) -> None:
        result = self.results.get(iid)
        if result is None:
            return
        if result.status is Status.DONE:
            self._open(result.job.dst)
        elif result.status is Status.FAILED:
            messagebox.showerror(result.job.src.name, result.message, parent=self)

    def _open(self, path: Path) -> None:
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except OSError as exc:
            messagebox.showerror(_("Nu pot deschide"), str(exc), parent=self)

    def _context_menu(self, iid: str | None) -> tk.Menu | None:
        if iid is None:
            return None
        src = self.files[iid]
        result = self.results.get(iid)
        menu = tk.Menu(self, tearoff=False)
        if result is not None and result.status is Status.DONE:
            menu.add_command(label=_("Deschide rezultatul"),
                             command=lambda: self._open(result.job.dst))
            menu.add_command(label=_("Deschide folderul rezultatului"),
                             command=lambda: self._open(result.job.dst.parent))
            menu.add_separator()
        if result is not None and result.status is Status.FAILED:
            menu.add_command(label=_("Vezi eroarea"), command=lambda: self._activate(iid))
            menu.add_separator()
        menu.add_command(label=_("Deschide originalul"), command=lambda: self._open(src))
        menu.add_command(label=_("Deschide folderul originalului"),
                         command=lambda: self._open(src.parent))
        if not self.busy:
            menu.add_separator()
            menu.add_command(label=_("Elimină din listă"), command=self.remove_selected)
        return menu

    # --- conversie ------------------------------------------------------------------------

    def _read_int(self, var: tk.StringVar, low: int, high: int, message: str) -> int | None:
        try:
            value = int(var.get())
        except ValueError:
            value = low - 1
        if not low <= value <= high:
            messagebox.showwarning(_("Valoare invalidă"), message, parent=self)
            return None
        return value

    def store_settings(self, s: Settings) -> None:
        s.same_folder = self.same_folder.get()
        s.output_dir = self.output_dir.get().strip()
        s.open_after = self.open_after.get()
        s.bookmarks = self.bookmarks.get()
        s.pdfa = self.pdfa.get()
        s.for_screen = self.for_screen.get()
        s.ocr = self.ocr.get()
        s.pdf_mode = self.pdf_mode.get()
        s.pdf_target = self.pdf_target.get()
        s.image_page = self.image_page.get()
        with contextlib.suppress(ValueError):
            s.timeout = max(10, min(int(self.timeout.get()), 3600))
        with contextlib.suppress(ValueError):
            s.workers = max(1, min(int(self.workers.get()), MAX_WORKERS))

    def start(self, ask_existing: bool = True) -> None:
        if self.busy:
            return
        if not self.files:
            messagebox.showwarning(_("Lipsă fișiere"), _("Adaugă cel puțin un document."),
                                   parent=self)
            return
        if self._read_int(self.timeout, 10, 3600, _("Limita per fișier trebuie să fie între "
                                                    "10 și 3600 de secunde.")) is None:
            return
        if self._read_int(self.workers, 1, MAX_WORKERS, _("Conversiile simultane trebuie să fie "
                                                          "între 1 și {n}.").format(
                                                              n=MAX_WORKERS)) is None:
            return

        out_dir: Path | None = None
        if not self.same_folder.get():
            text = self.output_dir.get().strip()
            if not text:
                messagebox.showwarning(_("Lipsă destinație"), _("Alege folderul de destinație."),
                                       parent=self)
                return
            out_dir = Path(text)
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                messagebox.showerror(_("Destinație invalidă"), str(exc), parent=self)
                return

        rows = self.table.order()
        jobs = plan_jobs([self.files[iid] for iid in rows], out_dir, self.pdf_target.get())
        existing = [job for job in jobs if job.dst.exists()]
        skip: list[Job] = []
        if existing:
            if ask_existing:
                answer = messagebox.askyesnocancel(
                    _("Fișiere existente"),
                    _("{n} fișier(e) rezultat există deja.\n\nDa – le suprascriu\nNu – le omit\n"
                      "Anulează – nu pornesc conversia").format(n=len(existing)),
                    parent=self)
                if answer is None:
                    return
                if answer is False:
                    skip = existing
            else:
                skip = existing  # pornire automata: nu suprascriem nimic fara sa intrebam

        self.store_settings(self.app.settings)
        self.app.save_settings()
        options = self.app.settings.export_options()

        self.job_rows = {id(job): os.path.normcase(str(job.src)) for job in jobs}
        self.results.clear()
        for iid in rows:
            self.table.set(iid, "status", Status.PENDING.text, tag="")
        self.progress.configure(maximum=len(jobs), value=0)
        self._durations = []
        self._started_at = time.monotonic()
        self._set_busy(True)
        self.status.set(_("Pornesc aplicațiile necesare…"))
        self.taskbar.set(0, len(jobs))

        self.engine = Engine(options, listener=lambda r: self.ui.post(lambda: self._on_result(r)),
                             backends=self.app.backends, pool=self.app.pool(options.workers))
        self.thread = threading.Thread(target=self._run, args=(self.engine, jobs, skip),
                                       name="convert", daemon=True)
        self.thread.start()

    def _run(self, engine: Engine, jobs: list[Job], skip: list[Job]) -> None:
        results: list[Result] = []
        try:
            results = engine.run(jobs, skip=skip)
        except Exception as exc:
            log.exception("Lotul de conversie a esuat")
            message = str(exc)
            self.ui.post(lambda: messagebox.showerror(_("Eroare"), message, parent=self))
        finally:
            self.ui.post(lambda: self._finished(results))

    def _eta(self) -> str:
        done = int(self.progress["value"])
        total = int(self.progress["maximum"])
        if done == 0 or done >= total:
            return ""
        elapsed = time.monotonic() - self._started_at
        return " · " + _("încă ~{time}").format(time=_duration(elapsed / done * (total - done)))

    def _on_result(self, result: Result) -> None:
        iid = self.job_rows.get(id(result.job))
        if iid is None:
            return
        text = result.status.text
        if result.status is Status.FAILED:
            text = _("Eroare: {message}").format(message=result.message)
        elif result.message:
            text = f"{text} ({result.message})"
        self.table.set(iid, "status", text, tag=STATUS_TAGS[result.status])
        total = int(self.progress["maximum"])
        if result.status is Status.RUNNING:
            self.table.tree.see(iid)
            self.status.set(_("{done}/{total}: {name}").format(
                done=int(self.progress["value"]) + 1, total=total, name=result.job.src.name)
                + self._eta())
        else:
            self.results[iid] = result
            self.progress.step(1)
            done = int(self.progress["value"])
            self.taskbar.set(done, total)
            self.status.set(_("{done}/{total} terminate").format(done=done, total=total)
                            + self._eta())

    def _finished(self, results: list[Result]) -> None:
        self._set_busy(False)
        self.engine = None
        counts = {status: sum(r.status is status for r in results) for status in Status}
        parts = [_("{n} convertite").format(n=counts[Status.DONE])]
        if counts[Status.SKIPPED]:
            parts.append(_("{n} omise").format(n=counts[Status.SKIPPED]))
        if counts[Status.CANCELLED]:
            parts.append(_("{n} anulate").format(n=counts[Status.CANCELLED]))
        if counts[Status.FAILED]:
            parts.append(_("{n} cu erori").format(n=counts[Status.FAILED]))
        took = _duration(time.monotonic() - self._started_at)
        summary = ", ".join(parts) + _(" în {time}.").format(time=took)
        self.status.set(_("Gata: {summary} Dublu-click pe un rând deschide rezultatul.")
                        .format(summary=summary))

        failed = [r for r in results if r.status is Status.FAILED]
        if failed:
            self.taskbar.state(ERROR)
            lines = "\n".join(f"• {r.job.src.name}: {r.message}" for r in failed[:10])
            more = ("\n" + _("… și încă {n}").format(n=len(failed) - 10)) if len(failed) > 10 \
                else ""
            messagebox.showwarning(_("Terminat cu erori"), f"{summary}\n\n{lines}{more}",
                                   parent=self)
        elif results:
            self.app.notify(_("Conversie terminată"), summary)
        self.taskbar.clear()

        done = [r for r in results if r.status is Status.DONE]
        if self.open_after.get() and done:
            self._open(done[0].job.dst.parent)

        if self.pending:  # fisiere primite in timpul conversiei
            pending, auto = self.pending, self.pending_auto
            self.pending, self.pending_auto = [], False
            self.receive(pending, auto)

    def _set_busy(self, busy: bool) -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        for button in [*self.add_buttons, self.start_button]:
            button.configure(state=state)
        self.cancel_button.configure(state=tk.NORMAL if busy else tk.DISABLED)

    def cancel(self) -> None:
        if self.engine is not None:
            self.engine.cancel()
            self.cancel_button.configure(state=tk.DISABLED)
            self.status.set(_("Se oprește după fișierele care se convertesc acum…"))

    def abort(self) -> None:
        """La inchiderea ferestrei: opreste imediat si aplicatiile Office care lucreaza."""
        if self.engine is not None:
            self.engine.abort()
