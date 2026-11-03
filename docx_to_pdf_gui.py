"""Convertor DOCX <-> PDF cu interfata grafica (tkinter).

Documentele Word devin PDF, iar PDF-urile devin DOCX.
Foloseste Microsoft Word (prin pywin32) daca este instalat,
altfel LibreOffice (soffice) daca il gaseste (doar Word -> PDF).
"""

import os
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

WORD_PDF_FORMAT = 17
WORD_DOCX_FORMAT = 16
WORD_EXTENSIONS = (".docx", ".doc", ".rtf", ".odt")
EXTENSIONS = WORD_EXTENSIONS + (".pdf",)
PDF_WARNING_VALUE = "DisableConvertPdfWarning"


def is_pdf(path):
    return path.lower().endswith(".pdf")


def target_extension(src):
    return ".docx" if is_pdf(src) else ".pdf"


def find_soffice():
    candidates = [
        shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def word_available():
    try:
        import win32com.client  # noqa: F401
        import winreg

        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"))
        return True
    except Exception:
        return False


def word_options_key():
    import winreg

    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer") as key:
        major = winreg.QueryValue(key, None).rsplit(".", 1)[1]  # "Word.Application.16"
    return rf"Software\Microsoft\Office\{major}.0\Word\Options"


class WordConverter:
    name = "Microsoft Word"

    def __enter__(self):
        import pythoncom
        import win32com.client

        # Cand deschide un PDF, Word afiseaza un mesaj de confirmare chiar si invizibil,
        # ceea ce ar bloca conversia. Il dezactivam temporar si refacem valoarea la final.
        self._saved_warning = self._set_pdf_warning(1)
        pythoncom.CoInitialize()
        try:
            self.word = win32com.client.DispatchEx("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = 0
        except Exception:
            pythoncom.CoUninitialize()
            self._restore_pdf_warning()
            raise
        return self

    def _restore_pdf_warning(self):
        if self._saved_warning != 1:
            self._set_pdf_warning(self._saved_warning)

    @staticmethod
    def _set_pdf_warning(value):
        """Seteaza optiunea si intoarce valoarea anterioara (None daca lipsea)."""
        import winreg

        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, word_options_key()) as key:
                try:
                    old = winreg.QueryValueEx(key, PDF_WARNING_VALUE)[0]
                except FileNotFoundError:
                    old = None
                if value is None:
                    winreg.DeleteValue(key, PDF_WARNING_VALUE)
                else:
                    winreg.SetValueEx(key, PDF_WARNING_VALUE, 0, winreg.REG_DWORD, value)
                return old
        except OSError:
            return None

    def convert(self, src, dst):
        doc = self.word.Documents.Open(
            os.path.abspath(src),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        try:
            if is_pdf(src):
                doc.SaveAs2(os.path.abspath(dst), FileFormat=WORD_DOCX_FORMAT)
            else:
                doc.ExportAsFixedFormat(os.path.abspath(dst), WORD_PDF_FORMAT)
        finally:
            doc.Close(False)

    def __exit__(self, *exc):
        import pythoncom

        try:
            self.word.Quit()
        finally:
            pythoncom.CoUninitialize()
            self._restore_pdf_warning()


class LibreOfficeConverter:
    name = "LibreOffice"

    def __init__(self, soffice):
        self.soffice = soffice

    def __enter__(self):
        return self

    def convert(self, src, dst):
        if is_pdf(src):
            raise RuntimeError("conversia PDF → DOCX necesită Microsoft Word")
        out_dir = os.path.dirname(os.path.abspath(dst))
        subprocess.run(
            [self.soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, src],
            check=True,
            capture_output=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        produced = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
        if os.path.abspath(produced) != os.path.abspath(dst):
            os.replace(produced, dst)

    def __exit__(self, *exc):
        pass


def make_converter():
    if word_available():
        return WordConverter()
    soffice = find_soffice()
    if soffice:
        return LibreOfficeConverter(soffice)
    return None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Convertor DOCX ⇄ PDF")
        self.geometry("680x480")
        self.minsize(560, 400)
        self._set_icon()

        self.files = []
        self.output_dir = tk.StringVar(value="")
        self.same_folder = tk.BooleanVar(value=True)
        self.open_after = tk.BooleanVar(value=False)
        self.events = queue.Queue()
        self.worker = None

        self._build_ui()
        self._update_output_state()
        self.after(100, self._poll_events)

    def _set_icon(self):
        # in exe-ul PyInstaller fisierele incluse stau in sys._MEIPASS
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        icon = os.path.join(base, "icon.ico")
        if os.path.isfile(icon):
            self.iconbitmap(default=icon)

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)
        ttk.Button(top, text="Adaugă fișiere…", command=self.add_files).pack(side="left")
        ttk.Button(top, text="Adaugă folder…", command=self.add_folder).pack(side="left", padx=5)
        ttk.Button(top, text="Șterge selectate", command=self.remove_selected).pack(side="left")
        ttk.Button(top, text="Golește lista", command=self.clear_files).pack(side="left", padx=5)

        list_frame = ttk.LabelFrame(
            self, text="Fișiere de convertit (Word → PDF, PDF → DOCX, automat după tip)"
        )
        list_frame.pack(fill="both", expand=True, **pad)
        self.listbox = tk.Listbox(list_frame, selectmode="extended")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
        scroll.pack(side="right", fill="y", pady=5)

        out = ttk.LabelFrame(self, text="Destinație")
        out.pack(fill="x", **pad)
        ttk.Checkbutton(
            out,
            text="Salvează rezultatul în același folder cu fișierul original",
            variable=self.same_folder,
            command=self._update_output_state,
        ).pack(anchor="w", padx=5, pady=(5, 0))
        row = ttk.Frame(out)
        row.pack(fill="x", padx=5, pady=5)
        self.out_entry = ttk.Entry(row, textvariable=self.output_dir)
        self.out_entry.pack(side="left", fill="x", expand=True)
        self.out_button = ttk.Button(row, text="Alege…", command=self.choose_output)
        self.out_button.pack(side="left", padx=(5, 0))
        ttk.Checkbutton(
            out, text="Deschide folderul la final", variable=self.open_after
        ).pack(anchor="w", padx=5, pady=(0, 5))

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True)
        self.convert_button = ttk.Button(bottom, text="Convertește", command=self.start)
        self.convert_button.pack(side="left", padx=(10, 0))

        self.status = tk.StringVar(value="Adaugă documente Word sau PDF pentru a începe.")
        ttk.Label(self, textvariable=self.status, anchor="w").pack(fill="x", padx=10, pady=(0, 10))

    def _update_output_state(self):
        state = "disabled" if self.same_folder.get() else "normal"
        self.out_entry.configure(state=state)
        self.out_button.configure(state=state)

    def _add(self, paths):
        added = 0
        for path in paths:
            path = os.path.normpath(path)
            name = os.path.basename(path)
            if name.startswith("~$") or not name.lower().endswith(EXTENSIONS):
                continue
            if path not in self.files:
                self.files.append(path)
                arrow = "→ DOCX" if is_pdf(path) else "→ PDF "
                self.listbox.insert("end", f"[{arrow}]  {path}")
                added += 1
        self.status.set(f"{len(self.files)} fișier(e) în listă.")
        return added

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Alege documente Word sau PDF",
            filetypes=[
                ("Word și PDF", "*.docx *.doc *.rtf *.odt *.pdf"),
                ("Documente Word", "*.docx *.doc *.rtf *.odt"),
                ("PDF", "*.pdf"),
                ("Toate fișierele", "*.*"),
            ],
        )
        self._add(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Alege un folder cu documente")
        if not folder:
            return
        paths = []
        for root, _dirs, names in os.walk(folder):
            paths.extend(os.path.join(root, n) for n in names)
        if self._add(sorted(paths)) == 0:
            messagebox.showinfo("Niciun fișier", "Nu am găsit documente Word sau PDF noi în folderul ales.")

    def remove_selected(self):
        for index in reversed(self.listbox.curselection()):
            self.listbox.delete(index)
            del self.files[index]
        self.status.set(f"{len(self.files)} fișier(e) în listă.")

    def clear_files(self):
        self.files.clear()
        self.listbox.delete(0, "end")
        self.status.set("Lista e goală.")

    def choose_output(self):
        folder = filedialog.askdirectory(title="Alege folderul de destinație")
        if folder:
            self.output_dir.set(os.path.normpath(folder))

    def _target_for(self, src):
        folder = os.path.dirname(src) if self.same_folder.get() else self.output_dir.get()
        name = os.path.splitext(os.path.basename(src))[0] + target_extension(src)
        return os.path.join(folder, name)

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.files:
            messagebox.showwarning("Lipsă fișiere", "Adaugă cel puțin un document Word sau PDF.")
            return
        if not self.same_folder.get():
            out = self.output_dir.get().strip()
            if not out:
                messagebox.showwarning("Lipsă destinație", "Alege folderul de destinație.")
                return
            os.makedirs(out, exist_ok=True)

        # raport.docx si raport.pdf in aceeasi lista s-ar suprascrie reciproc
        sources = {os.path.normcase(src) for src in self.files}
        jobs = []
        for src in self.files:
            dst = self._target_for(src)
            if os.path.normcase(dst) in sources:
                base, ext = os.path.splitext(dst)
                dst = f"{base} (convertit){ext}"
            jobs.append((src, dst))
        existing = [dst for _src, dst in jobs if os.path.exists(dst)]
        if existing and not messagebox.askyesno(
            "Fișiere existente",
            f"{len(existing)} fișier(e) rezultat există deja și vor fi suprascrise. Continui?",
        ):
            return

        self.convert_button.configure(state="disabled")
        self.progress.configure(maximum=len(jobs), value=0)
        self.worker = threading.Thread(target=self._run, args=(jobs,), daemon=True)
        self.worker.start()

    def _run(self, jobs):
        converter = make_converter()
        if converter is None:
            self.events.put(("fatal", "Nu am găsit nici Microsoft Word, nici LibreOffice instalat."))
            return
        errors = []
        try:
            with converter:
                for i, (src, dst) in enumerate(jobs, 1):
                    self.events.put(("status", f"[{converter.name}] {i}/{len(jobs)}: {os.path.basename(src)}"))
                    try:
                        converter.convert(src, dst)
                    except Exception as exc:
                        errors.append(f"{os.path.basename(src)}: {exc}")
                    self.events.put(("progress", i))
        except Exception as exc:
            self.events.put(("fatal", f"Nu am putut porni {converter.name}:\n{exc}"))
            return
        self.events.put(("done", (len(jobs), errors, jobs)))

    def _poll_events(self):
        try:
            while True:
                kind, data = self.events.get_nowait()
                if kind == "status":
                    self.status.set(data)
                elif kind == "progress":
                    self.progress.configure(value=data)
                elif kind == "fatal":
                    self.convert_button.configure(state="normal")
                    self.status.set("Eroare.")
                    messagebox.showerror("Eroare", data)
                elif kind == "done":
                    self._finish(*data)
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _finish(self, total, errors, jobs):
        self.convert_button.configure(state="normal")
        ok = total - len(errors)
        self.status.set(f"Gata: {ok}/{total} convertite.")
        if errors:
            messagebox.showwarning(
                "Terminat cu erori",
                f"{ok} din {total} convertite.\n\nErori:\n" + "\n".join(errors[:15]),
            )
        else:
            messagebox.showinfo("Gata", f"Toate cele {total} fișier(e) au fost convertite.")
        if self.open_after.get() and ok:
            os.startfile(os.path.dirname(jobs[0][1]))


if __name__ == "__main__":
    App().mainloop()
