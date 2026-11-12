"""Iconite pentru butoane, desenate din fontul de iconite al Windows (Segoe Fluent / MDL2).

Nu adaugam fisiere grafice, iar iconitele arata ca in restul sistemului. Daca fontul lipseste,
butoanele raman doar cu text.
"""

from __future__ import annotations

import logging
import tkinter as tk
from functools import cache
from pathlib import Path

log = logging.getLogger(__name__)

FONT_FILES = (r"C:\Windows\Fonts\SegoeIcons.ttf", r"C:\Windows\Fonts\segmdl2.ttf")
GLYPHS = {
    "add": "\ue710",
    "folder": "\ue8b7",
    "remove": "\ue74d",
    "clear": "\ue894",
    "play": "\ue768",
    "stop": "\ue71a",
    "up": "\ue70e",
    "down": "\ue70d",
    "left": "\ue72b",
    "right": "\ue72a",
    "rotate": "\ue7ad",
    "save": "\ue74e",
    "open": "\ue8e5",
    "pen": "\ue70f",
    "copy": "\ue8c8",
    "undo": "\ue7a7",
    "download": "\ue896",
    "sync": "\ue895",
    "zoom": "\ue71e",
}
_images: dict[tuple, tk.PhotoImage] = {}  # referinte pastrate, altfel Tk sterge imaginea


@cache
def _font_path() -> str | None:
    for path in FONT_FILES:
        if Path(path).exists():
            return path
    return None


def icon(name: str, color: str = "#5f6b7a", size: int = 16, mirror: bool = False
         ) -> tk.PhotoImage | None:
    key = (name, color, size, mirror)
    if key in _images:
        return _images[key]
    path = _font_path()
    if path is None or name not in GLYPHS:
        return None
    try:
        from PIL import Image, ImageDraw, ImageFont, ImageOps, ImageTk

        scale = 4  # desenam mai mare si micsoram, pentru margini fine
        font = ImageFont.truetype(path, size * scale)
        image = Image.new("RGBA", (size * scale, size * scale), (0, 0, 0, 0))
        ImageDraw.Draw(image).text((size * scale / 2, size * scale / 2), GLYPHS[name],
                                   font=font, fill=color, anchor="mm")
        image = image.resize((size, size), Image.Resampling.LANCZOS)
        if mirror:
            image = ImageOps.mirror(image)
        photo = ImageTk.PhotoImage(image)
    except Exception as exc:  # doar decorativ
        log.debug("Iconita %s: %s", name, exc)
        return None
    _images[key] = photo
    return photo
