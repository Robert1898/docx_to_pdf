"""Tab-ul principal: conversie Word/Excel/PowerPoint -> PDF si PDF -> DOCX."""

from __future__ import annotations

import contextlib
import logging
import os
import threading
import tkinter as tk
from collections.abc import Sequence
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING

import ttkbootstrap as ttk

from ..core import LABELS, Job, collect_files, kind_of, plan_jobs
from ..engine import Engine, Result, Status
from ..settings import Settings
from .widgets import Column, FileTable, UiQueue, tooltip

if TYPE_CHECKING:
    from .app import App

log = logging.getLogger(__name__)

FILETYPES = [
    ("Toate documentele", "*.docx *.doc *.docm *.rtf *.odt *.xlsx *.xls *.xlsm *.ods "
                          "*.pptx *.ppt *.pptm *.odp *.pdf"),
    ("Word", "*.docx *.doc *.docm *.rtf *.odt"),
    ("Excel", "*.xlsx *.xls *.xlsm *.ods"),
    ("PowerPoint", "*.pptx *.ppt *.pptm *.odp"),
    ("PDF", "*.pdf"),
]

COLUMNS = (
    Column("name", "Fișier", 250),
    Column("folder", "Folder", 230),
    Column("kind", "Conversie", 140, stretch=False),
    Column("status", "Status", 230),
)

PDF_MODE_LABELS = (
    ("auto", "Automat (recomandat)"),
    ("layout", "Aspect fidel (rapoarte, tabele)"),
    ("word", "Text editabil (Word)"),
)

STATUS_TAGS = {
    Status.PENDING: "",
    Status.RUNNING: "running",
    Status.DONE: "done",
    Status.FAILED: "failed",
    Status.SKIPPED: "muted",
    Status.CANCELLED: "muted",
}


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
        self.ui = UiQueue(self)

        self.same_folder = tk.BooleanVar(value=s.same_folder)
        self.output_dir = tk.StringVar(value=s.output_dir)
        self.open_after = tk.BooleanVar(value=s.open_after)
        self.bookmarks = tk.BooleanVar(value=s.bookmarks)
        self.pdfa = tk.BooleanVar(value=s.pdfa)
        self.for_screen = tk.BooleanVar(value=s.for_screen)
        self.ocr = tk.BooleanVar(value=s.ocr and app.backends.tesseract is not None)
        self.timeout = tk.StringVar(value=str(s.timeout))
        self.pdf_mode = tk.StringVar(value=s.pdf_mode)
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
            ttk.Button(bar, text="Adaugă fișiere…", command=self.add_files_dialog,
                       bootstyle="primary"),
            ttk.Button(bar, text="Adaugă folder…", command=self.add_folder_dialog,
                       bootstyle="primary-outline"),
            ttk.Button(bar, text="Elimină", command=self.remove_selected,
                       bootstyle="secondary-outline"),
            ttk.Button(bar, text="Golește lista", command=self.clear,
                       bootstyle="secondary-outline"),
        ]
        for i, button in enumerate(self.add_buttons):
            button.pack(side=tk.LEFT, padx=(0 if i == 0 else 6, 0))
        self.count_label = ttk.Label(bar, bootstyle="secondary")
        self.count_label.pack(side=tk.RIGHT)

        placeholder = (
            "Trage aici documente Word, Excel, PowerPoint sau PDF\n"
            "(sau foldere întregi), ori folosește butoanele de mai sus."
        )
        self.table = FileTable(self, COLUMNS, placeholder, on_drop=self.add_paths,
                               on_activate=self._activate)
        self.table.pack(fill=tk.BOTH, expand=True, pady=10)
        self.table.bind_context_menu(self._context_menu)
        self.table.tree.bind("<Delete>", lambda e: self.remove_selected())

        options = ttk.Labelframe(self, text="Opțiuni", padding=(10, 6))
        options.pack(fill=tk.X)
        bookmarks = ttk.Checkbutton(options, text="Semne de carte din titluri",
                                    variable=self.bookmarks)
        pdfa = ttk.Checkbutton(options, text="PDF/A (arhivare)", variable=self.pdfa)
        screen = ttk.Checkbutton(options, text="Optimizat pentru ecran (fișier mai mic)",
                                 variable=self.for_screen)
        ocr = ttk.Checkbutton(options, text="OCR pentru PDF-uri scanate", variable=self.ocr)
        bookmarks.grid(row=0, column=0, sticky=tk.W, padx=(0, 18), pady=3)
        pdfa.grid(row=0, column=1, sticky=tk.W, padx=(0, 18), pady=3)
        screen.grid(row=0, column=2, sticky=tk.W, pady=3)
        ocr.grid(row=1, column=0, sticky=tk.W, padx=(0, 18), pady=3)
        tooltip(bookmarks, "PDF-ul primește un cuprins navigabil din titlurile documentului Word.")
        tooltip(pdfa, "Format standard pentru arhivare pe termen lung, cerut uneori de instituții.")
        tooltip(screen, "Imagini la rezoluție mai mică: fișier mai mic, potrivit pentru e-mail.")
        if self.app.backends.tesseract is None:
            ocr.configure(state=tk.DISABLED)
            tooltip(ocr, "Necesită Tesseract OCR instalat (vezi README).")
        else:
            tooltip(ocr, "PDF-urile scanate (imagini) sunt trecute prin OCR înainte de conversia "
                         "în DOCX, ca să rezulte text editabil.")

        limit = ttk.Frame(options)
        limit.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=3)
        ttk.Label(limit, text="Limită per fișier:").pack(side=tk.LEFT)
        spin = ttk.Spinbox(limit, from_=30, to=3600, increment=30, width=6,
                           textvariable=self.timeout)
        spin.pack(side=tk.LEFT, padx=6)
        ttk.Label(limit, text="secunde").pack(side=tk.LEFT)
        tooltip(spin, "Dacă un fișier durează mai mult (de obicei un dialog ascuns), aplicația "
                      "Office pornită de program e oprită și se trece la următorul fișier.")

        pdf_row = ttk.Frame(options)
        pdf_row.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=3)
        ttk.Label(pdf_row, text="PDF → DOCX:").pack(side=tk.LEFT)
        mode = ttk.Combobox(pdf_row, state="readonly", width=34,
                            values=[label for _key, label in PDF_MODE_LABELS])
        mode.set(dict(PDF_MODE_LABELS)[self.app.settings.pdf_mode])
        mode.bind("<<ComboboxSelected>>", lambda e: self.pdf_mode.set(
            next(k for k, label in PDF_MODE_LABELS if label == mode.get())))
        mode.pack(side=tk.LEFT, padx=6)
        tooltip(mode, "Aspect fidel: fiecare pagină arată ca în PDF, cu coloanele aliniate exact; "
                      "potrivit pentru rapoarte, balanțe, extrase, tabele.\n"
                      "Text editabil (Word): textul curge liber, potrivit pentru scrisori și "
                      "contracte; Word ghicește tabelele și poate muta valori.\n"
                      "Automat: rapoartele pe coloane merg pe aspect fidel, restul prin Word.")

        dest = ttk.Labelframe(self, text="Destinație", padding=(10, 6))
        dest.pack(fill=tk.X, pady=(10, 0))
        ttk.Checkbutton(dest, text="Lângă fișierul original", variable=self.same_folder,
                        command=self._toggle_output).pack(anchor=tk.W)
        row = ttk.Frame(dest)
        row.pack(fill=tk.X, pady=6)
        self.out_entry = ttk.Entry(row, textvariable=self.output_dir)
        self.out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.out_button = ttk.Button(row, text="Alege…", command=self.choose_output,
                                     bootstyle="secondary-outline")
        self.out_button.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Checkbutton(dest, text="Deschide folderul la final",
                        variable=self.open_after).pack(anchor=tk.W)

        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, pady=(12, 0))
        self.progress = ttk.Progressbar(bottom, mode="determinate", bootstyle="success")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cancel_button = ttk.Button(bottom, text="Anulează", command=self.cancel,
                                        bootstyle="danger-outline", state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT, padx=(10, 0))
        self.start_button = ttk.Button(bottom, text="Convertește", command=self.start,
                                       bootstyle="success")
        self.start_button.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(self, textvariable=self.status, bootstyle="secondary").pack(
            fill=tk.X, pady=(6, 0)
        )

    def _toggle_output(self) -> None:
        state = tk.DISABLED if self.same_folder.get() else tk.NORMAL
        self.out_entry.configure(state=state)
        self.out_button.configure(state=state)

    def _update_count(self) -> None:
        n = len(self.files)
        self.count_label.configure(text=f"{n} fișier{'e' if n != 1 else ''}" if n else "")
        if not self.busy:
            self.status.set("Adaugă documente pentru a începe." if not n else
                            "Gata de conversie. Dublu-click pe un rând deschide rezultatul.")

    # --- lista de fisiere -----------------------------------------------------------------

    def add_paths(self, paths: Sequence[str | os.PathLike[str]]) -> int:
        if self.busy:
            self.status.set("Așteaptă terminarea conversiei înainte să adaugi fișiere.")
            return 0
        added = 0
        for path in collect_files(paths):
            iid = os.path.normcase(str(path))
            if iid in self.files:
                continue
            kind = kind_of(path)
            assert kind is not None
            self.files[iid] = path
            self.table.insert(iid, (path.name, str(path.parent), LABELS[kind],
                                    Status.PENDING.value))
            added += 1
        self._update_count()
        if paths and not added:
            self.status.set("Nu am găsit fișiere noi suportate.")
        return added

    def add_files_dialog(self) -> None:
        paths = filedialog.askopenfilenames(
            parent=self, title="Alege documente", filetypes=FILETYPES,
            initialdir=self.app.settings.last_dir or None,
        )
        if paths:
            self.app.settings.last_dir = str(Path(paths[0]).parent)
            self.add_paths(paths)

    def add_folder_dialog(self) -> None:
        folder = filedialog.askdirectory(parent=self, title="Alege un folder",
                                         initialdir=self.app.settings.last_dir or None)
        if folder:
            self.app.settings.last_dir = folder
            if not self.add_paths([folder]):
                messagebox.showinfo("Niciun fișier", "Nu am găsit documente noi în folderul ales.",
                                    parent=self)

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
        folder = filedialog.askdirectory(parent=self, title="Alege folderul de destinație")
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
            messagebox.showerror("Nu pot deschide", str(exc), parent=self)

    def _context_menu(self, iid: str | None) -> tk.Menu | None:
        if iid is None:
            return None
        src = self.files[iid]
        result = self.results.get(iid)
        menu = tk.Menu(self, tearoff=False)
        if result is not None and result.status is Status.DONE:
            menu.add_command(label="Deschide rezultatul",
                             command=lambda: self._open(result.job.dst))
            menu.add_command(label="Deschide folderul rezultatului",
                             command=lambda: self._open(result.job.dst.parent))
            menu.add_separator()
        if result is not None and result.status is Status.FAILED:
            menu.add_command(label="Vezi eroarea", command=lambda: self._activate(iid))
            menu.add_separator()
        menu.add_command(label="Deschide originalul", command=lambda: self._open(src))
        menu.add_command(label="Deschide folderul originalului",
                         command=lambda: self._open(src.parent))
        if not self.busy:
            menu.add_separator()
            menu.add_command(label="Elimină din listă", command=self.remove_selected)
        return menu

    # --- conversie ------------------------------------------------------------------------

    def _read_timeout(self) -> int | None:
        try:
            value = int(self.timeout.get())
        except ValueError:
            value = 0
        if not 10 <= value <= 3600:
            messagebox.showwarning("Limită invalidă",
                                   "Limita per fișier trebuie să fie între 10 și 3600 de secunde.",
                                   parent=self)
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
        with contextlib.suppress(ValueError):
            s.timeout = max(10, min(int(self.timeout.get()), 3600))

    def start(self) -> None:
        if self.busy:
            return
        if not self.files:
            messagebox.showwarning("Lipsă fișiere", "Adaugă cel puțin un document.", parent=self)
            return
        timeout = self._read_timeout()
        if timeout is None:
            return

        out_dir: Path | None = None
        if not self.same_folder.get():
            text = self.output_dir.get().strip()
            if not text:
                messagebox.showwarning("Lipsă destinație", "Alege folderul de destinație.",
                                       parent=self)
                return
            out_dir = Path(text)
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("Destinație invalidă", str(exc), parent=self)
                return

        rows = self.table.order()
        jobs = plan_jobs([self.files[iid] for iid in rows], out_dir)
        existing = [job for job in jobs if job.dst.exists()]
        skip: list[Job] = []
        if existing:
            answer = messagebox.askyesnocancel(
                "Fișiere existente",
                f"{len(existing)} fișier(e) rezultat există deja.\n\n"
                "Da – le suprascriu\nNu – le omit\nAnulează – nu pornesc conversia",
                parent=self,
            )
            if answer is None:
                return
            if answer is False:
                skip = existing

        self.store_settings(self.app.settings)
        self.app.save_settings()
        options = self.app.settings.export_options()

        self.job_rows = {id(job): os.path.normcase(str(job.src)) for job in jobs}
        self.results.clear()
        for iid in rows:
            self.table.set(iid, "status", Status.PENDING.value, tag="")
        self.progress.configure(maximum=len(jobs), value=0)
        self._set_busy(True)
        self.status.set("Pornesc aplicațiile necesare…")

        self.engine = Engine(options, listener=lambda r: self.ui.post(lambda: self._on_result(r)),
                             backends=self.app.backends)
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
            self.ui.post(lambda: messagebox.showerror("Eroare", message, parent=self))
        finally:
            self.ui.post(lambda: self._finished(results))

    def _on_result(self, result: Result) -> None:
        iid = self.job_rows.get(id(result.job))
        if iid is None:
            return
        text = result.status.value
        if result.status is Status.FAILED:
            text = f"Eroare: {result.message}"
        elif result.message:
            text = f"{text} ({result.message})"
        self.table.set(iid, "status", text, tag=STATUS_TAGS[result.status])
        if result.status is Status.RUNNING:
            self.table.tree.see(iid)
            done = int(self.progress["value"])
            self.status.set(f"{done + 1}/{int(self.progress['maximum'])}: {result.job.src.name}")
        else:
            self.results[iid] = result
            self.progress.step(1)

    def _finished(self, results: list[Result]) -> None:
        self._set_busy(False)
        self.engine = None
        counts = {status: sum(r.status is status for r in results) for status in Status}
        parts = [f"{counts[Status.DONE]} convertite"]
        if counts[Status.SKIPPED]:
            parts.append(f"{counts[Status.SKIPPED]} omise")
        if counts[Status.CANCELLED]:
            parts.append(f"{counts[Status.CANCELLED]} anulate")
        if counts[Status.FAILED]:
            parts.append(f"{counts[Status.FAILED]} cu erori")
        summary = ", ".join(parts) + "."
        self.status.set(f"Gata: {summary} Dublu-click pe un rând deschide rezultatul.")

        failed = [r for r in results if r.status is Status.FAILED]
        if failed:
            lines = "\n".join(f"• {r.job.src.name}: {r.message}" for r in failed[:10])
            more = f"\n… și încă {len(failed) - 10}" if len(failed) > 10 else ""
            messagebox.showwarning("Terminat cu erori", f"{summary}\n\n{lines}{more}", parent=self)
        elif results:
            self.app.notify("Conversie terminată", summary)

        done = [r for r in results if r.status is Status.DONE]
        if self.open_after.get() and done:
            self._open(done[0].job.dst.parent)

    def _set_busy(self, busy: bool) -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        for button in [*self.add_buttons, self.start_button]:
            button.configure(state=state)
        self.cancel_button.configure(state=tk.NORMAL if busy else tk.DISABLED)

    def cancel(self) -> None:
        if self.engine is not None:
            self.engine.cancel()
            self.cancel_button.configure(state=tk.DISABLED)
            self.status.set("Se oprește după fișierul curent…")

    def abort(self) -> None:
        """La inchiderea ferestrei: opreste imediat si aplicatia Office care lucreaza."""
        if self.engine is not None:
            self.engine.abort()
