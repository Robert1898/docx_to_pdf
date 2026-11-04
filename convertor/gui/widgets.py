"""Componente refolosite de tab-uri: tabelul de fisiere, drag & drop, coada catre UI."""

from __future__ import annotations

import contextlib
import logging
import queue
import tkinter as tk
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass

import ttkbootstrap as ttk

log = logging.getLogger(__name__)

STATUS_COLORS = {"done": "#2e9e5b", "failed": "#d9534f", "muted": "#8a8f98", "running": "#3b82f6"}


class UiQueue:
    """Firele de lucru nu au voie sa atinga widget-urile; trimit functii aici."""

    def __init__(self, widget: tk.Misc, interval: int = 80) -> None:
        self._queue: queue.Queue[Callable[[], None]] = queue.Queue()
        self._widget = widget
        self._interval = interval
        widget.after(interval, self._poll)

    def post(self, fn: Callable[[], None]) -> None:
        self._queue.put(fn)

    def _poll(self) -> None:
        try:
            while True:
                fn = self._queue.get_nowait()
                try:
                    fn()
                except Exception:
                    log.exception("Eroare in actualizarea interfetei")
        except queue.Empty:
            pass
        with contextlib.suppress(tk.TclError):  # fereastra a fost inchisa
            self._widget.after(self._interval, self._poll)


def enable_drop(widget: tk.Misc, callback: Callable[[Sequence[str]], None]) -> bool:
    """Accepta fisiere trase din Explorer. False daca tkdnd nu e disponibil."""
    try:
        from tkinterdnd2 import DND_FILES

        widget.drop_target_register(DND_FILES)  # type: ignore[attr-defined]

        def on_drop(event: tk.Event) -> str:
            callback(widget.tk.splitlist(event.data))  # type: ignore[attr-defined]
            return event.action  # type: ignore[attr-defined]

        widget.dnd_bind("<<Drop>>", on_drop)  # type: ignore[attr-defined]
        return True
    except Exception as exc:
        log.debug("Drag & drop indisponibil: %s", exc)
        return False


def tooltip(widget: tk.Misc, text: str) -> None:
    try:
        ttk.ToolTip(widget, text=text)
    except Exception as exc:  # doar decorativ
        log.debug("ToolTip: %s", exc)


@dataclass(frozen=True)
class Column:
    key: str
    heading: str
    width: int
    stretch: bool = True


class FileTable(ttk.Frame):
    """Treeview cu bara de defilare, mesaj cand e gol si suport pentru drag & drop."""

    def __init__(
        self,
        master: tk.Misc,
        columns: Sequence[Column],
        placeholder: str,
        on_drop: Callable[[Sequence[str]], object] | None = None,
        on_activate: Callable[[str], None] | None = None,
    ) -> None:
        super().__init__(master)
        self.columns = columns
        self.tree = ttk.Treeview(
            self, columns=[c.key for c in columns], show="headings", selectmode="extended"
        )
        for col in columns:
            self.tree.heading(col.key, text=col.heading, anchor=tk.W)
            self.tree.column(col.key, width=col.width, stretch=col.stretch, anchor=tk.W)
        for tag, color in STATUS_COLORS.items():
            self.tree.tag_configure(tag, foreground=color)

        scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.placeholder = ttk.Label(
            self, text=placeholder, justify=tk.CENTER, anchor=tk.CENTER, bootstyle="secondary"
        )
        self.drop_enabled = False
        if on_drop is not None:
            self.drop_enabled = enable_drop(self.tree, on_drop)
            enable_drop(self.placeholder, on_drop)
        if on_activate is not None:
            self.tree.bind("<Double-1>", lambda e: self._activate(e, on_activate))
            self.tree.bind("<Return>", lambda e: self._activate(e, on_activate))
        self.tree.bind("<Control-a>", lambda e: self.tree.selection_set(self.tree.get_children()))
        self._refresh_placeholder()

    def _activate(self, event: tk.Event, callback: Callable[[str], None]) -> None:
        iid = self.tree.identify_row(event.y) if event.type == tk.EventType.ButtonPress else (
            self.tree.focus()
        )
        if iid:
            callback(iid)

    def _refresh_placeholder(self) -> None:
        if self.tree.get_children():
            self.placeholder.place_forget()
        else:
            self.placeholder.place(relx=0.5, rely=0.55, anchor=tk.CENTER)

    # --- operatii pe randuri --------------------------------------------------------------

    def insert(self, iid: str, values: Sequence[str], tags: Iterable[str] = ()) -> None:
        self.tree.insert("", tk.END, iid=iid, values=list(values), tags=tuple(tags))
        self._refresh_placeholder()

    def set(self, iid: str, key: str, value: str, tag: str | None = None) -> None:
        if not self.tree.exists(iid):
            return
        self.tree.set(iid, key, value)
        if tag is not None:
            self.tree.item(iid, tags=(tag,) if tag else ())

    def order(self) -> list[str]:
        return list(self.tree.get_children())

    def selection(self) -> list[str]:
        return list(self.tree.selection())

    def delete(self, iids: Iterable[str]) -> None:
        for iid in iids:
            if self.tree.exists(iid):
                self.tree.delete(iid)
        self._refresh_placeholder()

    def clear(self) -> None:
        self.delete(self.order())

    def move(self, iids: Sequence[str], delta: int) -> None:
        """Muta randurile selectate cu o pozitie in sus (-1) sau in jos (+1)."""
        rows = self.order()
        chosen = sorted((rows.index(i) for i in iids if i in rows), reverse=delta > 0)
        for index in chosen:
            target = index + delta
            if 0 <= target < len(rows) and rows[target] not in iids:
                self.tree.move(rows[index], "", target)
                rows.insert(target, rows.pop(index))
        self.tree.selection_set(list(iids))
        if iids:
            self.tree.see(iids[0])

    def bind_context_menu(self, build: Callable[[str | None], tk.Menu | None]) -> None:
        def show(event: tk.Event) -> None:
            iid = self.tree.identify_row(event.y) or None
            if iid and iid not in self.tree.selection():
                self.tree.selection_set(iid)
            menu = build(iid)
            if menu is not None:
                menu.tk_popup(event.x_root, event.y_root)

        self.tree.bind("<Button-3>", show)
