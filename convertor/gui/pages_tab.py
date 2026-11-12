"""Tab-ul „Pagini”: organizatorul unui PDF, cu miniaturi.

Paginile se selecteaza (click, Ctrl+click, Shift+click), se rotesc, se sterg, se muta (tragand
cu mouse-ul sau cu sagetile), se extrag intr-un PDF nou si pot primi o semnatura sau o
stampila (imagine). Nimic nu se scrie pana la „Salvează ca…”.
"""

from __future__ import annotations

import logging
import os
import threading
import tkinter as tk
from collections.abc import Sequence
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING

import ttkbootstrap as ttk

from .. import pageedit, pdftools
from ..core import free_path
from ..i18n import _
from .convert_tab import IMAGES
from .widgets import UiQueue, enable_drop, icon_button, tooltip

if TYPE_CHECKING:
    from PIL import Image

    from .app import App

log = logging.getLogger(__name__)

THUMB = 120  # latimea miniaturii, px (inainte de scalarea DPI)
GAP = 18


class PagesTab(ttk.Frame):
    def __init__(self, master: tk.Misc, app: App) -> None:
        super().__init__(master, padding=12)
        self.app = app
        self.ui = UiQueue(self)
        self.src: Path | None = None
        self.edits: list[pageedit.PageEdit] = []
        self.thumbs: dict[int, Image.Image] = {}  # pagina sursa -> miniatura (nerotita)
        self.photos: dict[tuple[int, int], object] = {}  # (sursa, rotire) -> PhotoImage
        self.selected: set[int] = set()  # pozitii in self.edits
        self.anchor: int | None = None
        self.dirty = False
        self._drag_from: int | None = None
        self._dragged = False
        self._generation = 0  # miniaturile unui document vechi nu mai sunt afisate
        self.thread: threading.Thread | None = None
        scale = app.root.winfo_fpixels("1i") / 96
        self.thumb_width = int(THUMB * scale)
        self.status = tk.StringVar(value=_("Deschide un PDF ca să-i organizezi paginile."))
        self._build()

    @property
    def busy(self) -> bool:
        return self.thread is not None and self.thread.is_alive()

    # --- construire interfata -------------------------------------------------------------

    def _build(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X)
        self.open_button = icon_button(bar, _("Deschide PDF…"), self.open_dialog, "primary",
                                       "open")
        self.open_button.pack(side=tk.LEFT)
        self.tools = [
            (icon_button(bar, "", lambda: self.rotate(-90), "secondary-outline", "rotate",
                         mirror=True, width=3), _("Rotește spre stânga")),
            (icon_button(bar, "", lambda: self.rotate(90), "secondary-outline", "rotate",
                         width=3), _("Rotește spre dreapta")),
            (icon_button(bar, "", self.delete_selected, "secondary-outline", "remove", width=3),
             _("Șterge paginile selectate (Delete)")),
            (icon_button(bar, "", lambda: self.move(-1), "secondary-outline", "left", width=3),
             _("Mută mai în față")),
            (icon_button(bar, "", lambda: self.move(1), "secondary-outline", "right", width=3),
             _("Mută mai în spate")),
            (icon_button(bar, _("Semnătură…"), self.add_stamp, "secondary-outline", "pen"),
             _("Pune o imagine (semnătură, ștampilă) pe pagina selectată")),
            (icon_button(bar, _("Extrage…"), self.extract, "secondary-outline", "copy"),
             _("Salvează doar paginile selectate într-un PDF nou")),
            (icon_button(bar, _("Renunță"), self.reset, "secondary-outline", "undo"),
             _("Revino la documentul original")),
        ]
        for button, tip in self.tools:
            button.pack(side=tk.LEFT, padx=(6, 0))
            tooltip(button, tip)
        self.save_button = icon_button(bar, _("Salvează ca…"), self.save_as, "success", "save")
        self.save_button.pack(side=tk.RIGHT)
        tooltip(self.save_button, "Ctrl+S")

        area = ttk.Frame(self)
        area.pack(fill=tk.BOTH, expand=True, pady=10)
        self.canvas = tk.Canvas(area, highlightthickness=0, bd=0)
        scroll = ttk.Scrollbar(area, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.canvas.bind("<Configure>", lambda e: self.redraw())
        self.canvas.bind("<ButtonPress-1>", self._press)
        self.canvas.bind("<B1-Motion>", self._motion)
        self.canvas.bind("<ButtonRelease-1>", self._release)
        self.canvas.bind("<Double-1>", lambda e: self.add_stamp())
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(
            -1 * (e.delta // 120), "units"))
        self.canvas.bind("<Delete>", lambda e: self.delete_selected())
        self.canvas.bind("<Control-a>", lambda e: self.select_all())
        enable_drop(self.canvas, self._dropped)

        ttk.Label(self, textvariable=self.status, bootstyle="secondary").pack(fill=tk.X)
        self._update_buttons()

    def apply_theme(self) -> None:
        background = ttk.Style().lookup("TFrame", "background") or "white"
        self.canvas.configure(background=background)
        self.redraw()

    # --- document -------------------------------------------------------------------------

    def _dropped(self, paths: Sequence[str]) -> None:
        pdfs = [p for p in paths if p.lower().endswith(".pdf")]
        if pdfs:
            self.open(Path(pdfs[0]))

    def open_dialog(self) -> None:
        path = filedialog.askopenfilename(parent=self, title=_("Alege un PDF"),
                                          filetypes=[("PDF", "*.pdf")],
                                          initialdir=self.app.settings.last_dir or None)
        if path:
            self.open(Path(path))

    def confirm_discard(self) -> bool:
        return not self.dirty or messagebox.askyesno(
            _("Modificări nesalvate"), _("Renunți la modificările făcute paginilor?"),
            parent=self)

    def open(self, path: Path) -> None:
        if self.busy or not self.confirm_discard():
            return
        try:
            edits = pageedit.initial(path)
        except pdftools.PdfError as exc:
            messagebox.showerror(_("Nu pot deschide"), str(exc), parent=self)
            return
        self.src, self.edits = path, edits
        self.thumbs.clear()
        self.photos.clear()
        self.selected.clear()
        self.dirty = False
        self._generation += 1
        self.app.settings.last_dir = str(path.parent)
        self.status.set(_("„{name}”: {n} pagini. Se generează miniaturile…").format(
            name=path.name, n=len(edits)))
        self.redraw()
        self._update_buttons()
        generation = self._generation
        threading.Thread(target=self._render_thumbs, args=(path, generation), name="thumbs",
                         daemon=True).start()

    def _render_thumbs(self, path: Path, generation: int) -> None:
        import pypdfium2 as pdfium

        document = pdfium.PdfDocument(str(path))
        try:
            for index in range(len(document)):
                if generation != self._generation:
                    return
                page = document[index]
                image = page.render(scale=self.thumb_width / page.get_width()).to_pil()
                page.close()
                self.ui.post(lambda i=index, img=image, g=generation: self._thumb_ready(
                    i, img, g))
        except Exception as exc:
            log.warning("Miniaturi %s: %s", path, exc)
        finally:
            document.close()
        self.ui.post(lambda: self._update_status() if generation == self._generation else None)

    def _thumb_ready(self, index: int, image: Image.Image, generation: int) -> None:
        if generation != self._generation:
            return
        self.thumbs[index] = image
        self.redraw()

    # --- desenare -------------------------------------------------------------------------

    def _layout(self) -> tuple[int, int, int]:
        width = max(self.canvas.winfo_width(), self.thumb_width + 2 * GAP)
        cell_w = self.thumb_width + GAP
        cell_h = int(self.thumb_width * 1.45) + GAP + 18
        columns = max(1, (width - GAP) // cell_w)
        return columns, cell_w, cell_h

    def _photo(self, source: int, rotation: int) -> object | None:
        image = self.thumbs.get(source)
        if image is None:
            return None
        key = (source, rotation % 360)
        if key not in self.photos:
            from PIL import ImageTk

            rotated = image.rotate(-rotation, expand=True)
            # paginile intoarse sau late trebuie sa incapa tot in caseta miniaturii
            rotated.thumbnail((self.thumb_width, int(self.thumb_width * 1.45)))
            self.photos[key] = ImageTk.PhotoImage(rotated)
        return self.photos[key]

    def redraw(self) -> None:
        canvas = self.canvas
        canvas.delete("all")
        if not self.edits:
            canvas.create_text(canvas.winfo_width() // 2, 80, text=_(
                "Deschide un PDF sau trage-l aici.\nPaginile se pot apoi roti, șterge, muta "
                "(trăgându-le), extrage sau semna."), justify=tk.CENTER, fill="#8a8f98")
            return
        columns, cell_w, cell_h = self._layout()
        accent = "#3b82f6"
        for pos, edit in enumerate(self.edits):
            col, row = pos % columns, pos // columns
            x = GAP + col * cell_w + self.thumb_width // 2
            y = GAP + row * cell_h
            photo = self._photo(edit.source, edit.rotation)
            box_h = int(self.thumb_width * 1.45)
            if pos in self.selected:
                canvas.create_rectangle(x - self.thumb_width // 2 - 5, y - 5,
                                        x + self.thumb_width // 2 + 5, y + box_h + 22,
                                        outline=accent, width=3, fill="")
            if photo is not None:
                half_w, half_h = photo.width() // 2, photo.height() // 2  # type: ignore[attr-defined]
                cy = y + box_h // 2
                # umbra si chenar: o pagina alba nu se vede pe fundal deschis
                canvas.create_rectangle(x - half_w + 3, cy - half_h + 3, x + half_w + 3,
                                        cy + half_h + 3, outline="", fill="#9aa0a6")
                canvas.create_image(x, cy, image=photo)
                canvas.create_rectangle(x - half_w, cy - half_h, x + half_w, cy + half_h,
                                        outline="#9aa0a6")
            else:
                canvas.create_rectangle(x - self.thumb_width // 2, y, x + self.thumb_width // 2,
                                        y + box_h, outline="#c0c4cc", dash=(3, 3))
            marks = " ✎" if edit.stamps else ""
            canvas.create_text(x, y + box_h + 10, text=f"{pos + 1}{marks}", fill="#8a8f98")
        rows = (len(self.edits) + columns - 1) // columns
        canvas.configure(scrollregion=(0, 0, columns * cell_w + GAP, rows * cell_h + GAP))

    def _position_at(self, event: tk.Event) -> int | None:
        columns, cell_w, cell_h = self._layout()
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        col, row = int((x - GAP / 2) // cell_w), int((y - GAP / 2) // cell_h)
        if col < 0 or col >= columns or row < 0:
            return None
        pos = row * columns + col
        return pos if pos < len(self.edits) else None

    # --- selectie si tragere --------------------------------------------------------------

    def _press(self, event: tk.Event) -> None:
        self.canvas.focus_set()
        pos = self._position_at(event)
        self._dragged = False
        if pos is None:
            self.selected.clear()
            self._drag_from = None
        elif event.state & 0x0001 and self.anchor is not None:  # Shift: interval
            low, high = sorted((self.anchor, pos))
            self.selected = set(range(low, high + 1))
            self._drag_from = None
        elif event.state & 0x0004:  # Ctrl: adauga / scoate
            self.selected ^= {pos}
            self.anchor = pos
            self._drag_from = None
        else:
            if pos not in self.selected:
                self.selected = {pos}
            self.anchor = pos
            self._drag_from = pos
        self.redraw()
        self._update_buttons()

    def _motion(self, event: tk.Event) -> None:
        if self._drag_from is None:
            return
        self._dragged = True
        self.canvas.configure(cursor="fleur")

    def _release(self, event: tk.Event) -> None:
        self.canvas.configure(cursor="")
        if self._drag_from is None or not self._dragged:
            return
        target = self._position_at(event)
        if target is None:
            target = len(self.edits) - 1
        if target not in self.selected:
            self._move_selection_to(target)
        self._drag_from = None

    def _move_selection_to(self, target: int) -> None:
        moving = [self.edits[i] for i in sorted(self.selected)]
        rest = [e for i, e in enumerate(self.edits) if i not in self.selected]
        before = sum(1 for i in range(target) if i not in self.selected)
        if target > max(self.selected):
            before += 1  # mutare in spate: dupa pagina tinta
        self.edits = rest[:before] + moving + rest[before:]
        self.selected = set(range(before, before + len(moving)))
        self._changed()

    def select_all(self) -> None:
        self.selected = set(range(len(self.edits)))
        self.redraw()
        self._update_buttons()

    # --- operatii -------------------------------------------------------------------------

    def _changed(self) -> None:
        self.dirty = True
        self.redraw()
        self._update_buttons()
        self._update_status()

    def _update_status(self) -> None:
        if self.src is None:
            return
        changed = _(" (modificat, nesalvat)") if self.dirty else ""
        self.status.set(_("„{name}”: {n} pagini{changed}. Selectate: {s}.").format(
            name=self.src.name, n=len(self.edits), changed=changed, s=len(self.selected)))

    def _update_buttons(self) -> None:
        has_doc = bool(self.edits) and not self.busy
        for button, _tip in self.tools:
            button.configure(state=tk.NORMAL if has_doc else tk.DISABLED)
        self.save_button.configure(state=tk.NORMAL if has_doc else tk.DISABLED)
        self._update_status()

    def rotate(self, degrees: int) -> None:
        if not self.selected:
            return
        for pos in self.selected:
            self.edits[pos].rotation = (self.edits[pos].rotation + degrees) % 360
        self._changed()

    def delete_selected(self) -> None:
        if not self.selected:
            return
        if len(self.selected) == len(self.edits):
            messagebox.showwarning(_("Ștergere"), _("Documentul trebuie să păstreze cel puțin "
                                                    "o pagină."), parent=self)
            return
        self.edits = [e for i, e in enumerate(self.edits) if i not in self.selected]
        self.selected.clear()
        self._changed()

    def move(self, delta: int) -> None:
        if not self.selected:
            return
        low, high = min(self.selected), max(self.selected)
        target = low - 1 if delta < 0 else high + 1
        if 0 <= target < len(self.edits):
            self._move_selection_to(target)

    def reset(self) -> None:
        if self.src is not None and self.confirm_discard():
            self.dirty = False
            self.open(self.src)

    def add_stamp(self) -> None:
        if self.src is None or len(self.selected) != 1:
            messagebox.showinfo(_("Semnătură"), _("Selectează o singură pagină."), parent=self)
            return
        image = filedialog.askopenfilename(
            parent=self, title=_("Alege imaginea semnăturii sau a ștampilei"),
            filetypes=[(_("Imagini"), IMAGES)])
        if not image:
            return
        pos = next(iter(self.selected))
        stamp = StampDialog(self, self.src, self.edits[pos], Path(image)).result
        if stamp is not None:
            self.edits[pos].stamps.append(stamp)
            self._changed()

    def _save(self, edits: list[pageedit.PageEdit], dst: Path, after: str) -> None:
        assert self.src is not None
        src = self.src

        def work() -> None:
            try:
                pageedit.save(src, edits, dst)
            except Exception as exc:
                log.exception("Salvare pagini")
                message = str(exc)
                self.ui.post(lambda: messagebox.showerror(_("Eroare"), message, parent=self))
                return
            self.ui.post(lambda: self._saved(dst, after))

        self.status.set(_("Se salvează…"))
        self.thread = threading.Thread(target=work, name="pages-save", daemon=True)
        self.thread.start()
        self._update_buttons()

    def _saved(self, dst: Path, after: str) -> None:
        if after == "save":
            self.dirty = False
        self._update_buttons()
        self.status.set(_("Am salvat {path}.").format(path=dst))
        self.app.notify(_("Pagini"), _("Am salvat {name}.").format(name=dst.name))

    def save_as(self) -> None:
        if self.src is None or self.busy:
            return
        target = filedialog.asksaveasfilename(
            parent=self, title=_("Salvează PDF-ul"), defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")], initialdir=str(self.src.parent),
            initialfile=free_path(self.src.with_name(_("{name} (editat).pdf").format(
                name=self.src.stem))).name)
        if not target:
            return
        if os.path.normcase(str(Path(target).resolve())) == os.path.normcase(
                str(self.src.resolve())):
            messagebox.showwarning(_("Salvare"), _("Salvează cu alt nume; originalul rămâne "
                                                   "neatins."), parent=self)
            return
        self._save(list(self.edits), Path(target), "save")

    def extract(self) -> None:
        if self.src is None or not self.selected or self.busy:
            return
        target = filedialog.asksaveasfilename(
            parent=self, title=_("Salvează paginile selectate"), defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")], initialdir=str(self.src.parent),
            initialfile=free_path(self.src.with_name(_("{name} (extras).pdf").format(
                name=self.src.stem))).name)
        if target:
            self._save([self.edits[i] for i in sorted(self.selected)], Path(target), "extract")


class StampDialog(tk.Toplevel):
    """Pagina marita; utilizatorul trage dreptunghiul in care apare imaginea."""

    def __init__(self, parent: PagesTab, src: Path, edit: pageedit.PageEdit,
                 image: Path) -> None:
        super().__init__(parent)
        self.title(_("Semnătură sau ștampilă"))
        self.transient(parent.winfo_toplevel())
        self.result: pageedit.Stamp | None = None
        self.src, self.edit, self.image_path = src, edit, image
        from PIL import Image, ImageTk

        try:
            self.stamp_image = Image.open(image)
            self.stamp_image.load()
        except OSError as exc:
            messagebox.showerror(_("Imagine"), str(exc), parent=parent)
            self.destroy()
            return

        import pypdfium2 as pdfium

        self.document = pdfium.PdfDocument(str(src))
        self.page = self.document[edit.source]
        screen_h = self.winfo_screenheight()
        height = int(screen_h * 0.72)
        width_pt, height_pt = self.page.get_width(), self.page.get_height()
        if edit.rotation % 180:
            width_pt, height_pt = height_pt, width_pt
        scale = height / height_pt
        rendered = self.page.render(scale=scale, rotation=edit.rotation % 360).to_pil()
        self.size = rendered.size
        self.page_photo = ImageTk.PhotoImage(rendered)

        ttk.Label(self, text=_("Trage cu mouse-ul dreptunghiul în care să apară imaginea."),
                  padding=8).pack(anchor=tk.W)
        self.canvas = tk.Canvas(self, width=self.size[0], height=self.size[1],
                                highlightthickness=0, cursor="crosshair")
        self.canvas.pack(padx=8)
        self.canvas.create_image(0, 0, image=self.page_photo, anchor=tk.NW)
        buttons = ttk.Frame(self, padding=8)
        buttons.pack(fill=tk.X)
        self.ok = ttk.Button(buttons, text=_("Aplică"), bootstyle="success", command=self._ok,
                             state=tk.DISABLED)
        self.ok.pack(side=tk.RIGHT)
        ttk.Button(buttons, text=_("Anulează"), bootstyle="secondary-outline",
                   command=self._close).pack(side=tk.RIGHT, padx=6)
        self.rect: tuple[int, int, int, int] | None = None
        self.start: tuple[int, int] | None = None
        self.canvas.bind("<ButtonPress-1>", self._press)
        self.canvas.bind("<B1-Motion>", self._drag)
        self.bind("<Escape>", lambda e: self._close())
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.grab_set()
        self.wait_window()

    def _press(self, event: tk.Event) -> None:
        self.start = (event.x, event.y)

    def _drag(self, event: tk.Event) -> None:
        from PIL import ImageTk

        if self.start is None:
            return
        x0, y0 = self.start
        # pastram proportiile imaginii: latimea data de mouse, inaltimea calculata
        w = max(10, abs(event.x - x0))
        ratio = self.stamp_image.height / max(self.stamp_image.width, 1)
        h = max(10, int(w * ratio))
        x1 = x0 + w if event.x >= x0 else x0 - w
        y1 = y0 + h if event.y >= y0 else y0 - h
        left, right = sorted((x0, x1))
        top, bottom = sorted((y0, y1))
        self.rect = (left, top, right, bottom)
        preview = self.stamp_image.convert("RGBA").resize((right - left, bottom - top))
        self.preview_photo = ImageTk.PhotoImage(preview)
        self.canvas.delete("stamp")
        self.canvas.create_image(left, top, image=self.preview_photo, anchor=tk.NW, tags="stamp")
        self.canvas.create_rectangle(left, top, right, bottom, outline="#3b82f6", dash=(4, 2),
                                     tags="stamp")
        self.ok.configure(state=tk.NORMAL)

    def _ok(self) -> None:
        assert self.rect is not None
        left, top, right, bottom = self.rect
        a = pageedit.device_to_page(self.page, self.edit.rotation, self.size, (left, top))
        b = pageedit.device_to_page(self.page, self.edit.rotation, self.size, (right, bottom))
        angle = (pageedit.page_rotation(self.page) + self.edit.rotation) % 360
        self.result = pageedit.Stamp(self.image_path, min(a[0], b[0]), min(a[1], b[1]),
                                     max(a[0], b[0]), max(a[1], b[1]), angle)
        self._close()

    def _close(self) -> None:
        try:
            self.page.close()
            self.document.close()
        finally:
            self.grab_release()
            self.destroy()
