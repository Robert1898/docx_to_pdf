"""Editarea paginilor unui PDF: ordine, rotire, stergere, extragere, semnatura/stampila.

Modelul e o lista de pagini editate (`PageEdit`), fiecare cu pagina sursa, rotirea
suplimentara si imaginile aplicate pe ea; salvarea scrie un PDF nou.
"""

from __future__ import annotations

import ctypes
from collections.abc import Sequence
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from .i18n import _
from .pdftools import PdfError, _write, open_pdf


@dataclass
class Stamp:
    """Imagine (semnatura, stampila) pe pagina, in coordonatele PDF ale paginii sursa."""

    image: Path
    x0: float
    y0: float
    x1: float
    y1: float
    angle: int = 0  # cat trebuie rotita imaginea ca sa apara dreapta pe pagina afisata


@dataclass
class PageEdit:
    source: int  # indexul paginii in PDF-ul original
    rotation: int = 0  # rotire suplimentara, multiplu de 90
    stamps: list[Stamp] = field(default_factory=list)


def initial(path: Path) -> list[PageEdit]:
    return [PageEdit(i) for i in range(len(open_pdf(path).pages))]


def _stamp_page(width: float, height: float, stamps: Sequence[Stamp]) -> object:
    from PIL import Image
    from pypdf import PdfReader
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    for stamp in stamps:
        try:
            image = Image.open(stamp.image)
            image.load()
        except OSError as exc:
            raise PdfError(_("Nu pot citi imaginea {name}: {error}").format(
                name=stamp.image.name, error=exc)) from exc
        if image.mode not in ("RGB", "RGBA", "L"):
            image = image.convert("RGBA")
        if stamp.angle:
            image = image.rotate(stamp.angle, expand=True)
        png = BytesIO()
        image.save(png, "PNG")
        png.seek(0)
        c.drawImage(ImageReader(png), stamp.x0, stamp.y0, stamp.x1 - stamp.x0,
                    stamp.y1 - stamp.y0, mask="auto", preserveAspectRatio=True,
                    anchor="c")
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def save(src: Path, edits: Sequence[PageEdit], dst: Path) -> None:
    """Scrie un PDF nou cu paginile in ordinea si forma din `edits`."""
    from pypdf import PdfWriter

    if not edits:
        raise PdfError(_("Documentul nu mai are nicio pagină."))
    reader = open_pdf(src)
    writer = PdfWriter()
    for edit in edits:
        page = writer.add_page(reader.pages[edit.source])
        if edit.stamps:
            box = page.mediabox
            overlay = _stamp_page(float(box.right), float(box.top), edit.stamps)
            page.merge_page(overlay)
        if edit.rotation % 360:
            page.rotate(edit.rotation % 360)
    _write(writer, dst)


def device_to_page(pdf_page: object, rotation: int, size: tuple[int, int],
                   point: tuple[float, float]) -> tuple[float, float]:
    """Punctul dintr-o imagine a paginii (randata cu `rotation` suplimentara) -> coordonate
    PDF ale paginii. Tine cont si de rotirea proprie a paginii (/Rotate)."""
    import pypdfium2.raw as pdfium_c

    px, py = ctypes.c_double(), ctypes.c_double()
    pdfium_c.FPDF_DeviceToPage(pdf_page.raw, 0, 0, size[0], size[1],  # type: ignore[attr-defined]
                               (rotation // 90) % 4, int(point[0]), int(point[1]),
                               ctypes.byref(px), ctypes.byref(py))
    return px.value, py.value


def page_rotation(pdf_page: object) -> int:
    """Rotirea proprie a paginii (/Rotate), in grade."""
    return int(pdf_page.get_rotation()) % 360  # type: ignore[attr-defined]
