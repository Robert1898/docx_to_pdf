"""Renumerotarea paginilor dupa unirea PDF-urilor.

Numerele de pagina sunt text tiparit in pagina („8/8”, „Pagina 1 din 2”), nu metadate, deci nu
pot fi modificate direct. Le gasim in fiecare PDF sursa (pozitie, marime, culoare, font),
acoperim numarul vechi cu un petic de culoarea fundalului si scriem numarul nou in acelasi loc.

Ca sa nu atingem alte cifre din text, un numar e considerat numar de pagina doar daca:
  * corespunde paginii: „8/8” pe pagina 8 dintr-un document de 8 pagini;
  * e in antet sau subsol (zona de sus/jos a paginii);
  * e la inceputul sau la sfarsitul randului lui, nu in mijlocul unei fraze.
"""

from __future__ import annotations

import ctypes
import logging
import re
import winreg
from collections import Counter
from collections.abc import Iterator
from dataclasses import dataclass
from functools import cache
from io import BytesIO
from pathlib import Path

from pypdf import PageObject, PdfReader

log = logging.getLogger(__name__)

EDGE_ZONE = 0.12  # cat din inaltimea paginii conteaza ca antet / subsol
LINE_EDGE = re.compile(r"^[\s\-–—|·•.,:()\[\]]*$")  # ce poate sta intre numar si capatul randului

# grupuri numite: n = pagina curenta, total = numarul de pagini (optional)
PATTERNS = (
    re.compile(r"(?i)\b(?:pagina|pag\.|page|p\.)\s*(?P<n>\d{1,4})\s*(?:din|of|de|/)\s*"
               r"(?P<total>\d{1,4})\b"),
    re.compile(r"(?<![\d/])(?P<n>\d{1,4})\s*/\s*(?P<total>\d{1,4})(?![\d/])"),
    re.compile(r"(?i)\b(?:pagina|pag\.|page)\s*(?P<n>\d{1,4})\b"),
    re.compile(r"(?<![\d.,/])(?P<n>\d{1,4})(?![\d.,/])"),
)


@dataclass
class PageNumber:
    """Un numar de pagina gasit intr-un PDF sursa."""

    template: str  # textul original cu {page} si {total} in locul cifrelor
    box: tuple[float, float, float, float]  # stanga, jos, dreapta, sus (coordonate PDF)
    baseline: float
    font_size: float
    color: tuple[float, float, float]
    background: tuple[float, float, float]
    font_name: str
    bold: bool
    align: str  # "left" / "center" / "right"

    def text(self, page: int, total: int) -> str:
        return self.template.format(page=page, total=total)


# --- detectie -----------------------------------------------------------------------------


def _template(match: re.Match[str]) -> str:
    """„Pagina 8 din 8” -> „Pagina {page} din {total}”."""
    text, start = match.group(0), match.start()
    spans = sorted(
        (match.span(name), "{" + ("page" if name == "n" else "total") + "}")
        for name in ("n", "total")
        if name in match.re.groupindex and match.group(name) is not None
    )
    out, pos = [], 0
    for (a, b), placeholder in spans:
        out.append(text[pos:a - start].replace("{", "{{").replace("}", "}}"))
        out.append(placeholder)
        pos = b - start
    out.append(text[pos:].replace("{", "{{").replace("}", "}}"))
    return "".join(out)


def _alone_on_line(text: str, start: int, end: int) -> bool:
    line_start = max(text.rfind("\n", 0, start), text.rfind("\r", 0, start)) + 1
    ends = [i for i in (text.find("\r", end), text.find("\n", end)) if i != -1]
    line_end = min(ends) if ends else len(text)
    before, after = text[line_start:start], text[end:line_end]
    return bool(LINE_EDGE.match(before) or LINE_EDGE.match(after))


def _font(textpage_raw: object, index: int) -> tuple[str, bool]:
    import pypdfium2.raw as pdfium_c

    flags = ctypes.c_int()
    size = pdfium_c.FPDFText_GetFontInfo(textpage_raw, index, None, 0, ctypes.byref(flags))
    name = ""
    if size:
        buf = ctypes.create_string_buffer(size)
        pdfium_c.FPDFText_GetFontInfo(textpage_raw, index, buf, size, ctypes.byref(flags))
        name = buf.value.decode("utf-8", errors="replace")
    weight = pdfium_c.FPDFText_GetFontWeight(textpage_raw, index)
    return name, weight >= 600 or "bold" in name.lower()


def _color(textpage_raw: object, index: int) -> tuple[float, float, float]:
    import pypdfium2.raw as pdfium_c

    r, g, b, a = (ctypes.c_uint() for _ in range(4))
    refs = (ctypes.byref(v) for v in (r, g, b, a))
    if pdfium_c.FPDFText_GetFillColor(textpage_raw, index, *refs):
        return r.value / 255, g.value / 255, b.value / 255
    return 0.0, 0.0, 0.0


def _background(page: object, box: tuple[float, float, float, float]) -> tuple[float, ...]:
    """Culoarea dominanta din jurul numarului (alb pentru majoritatea documentelor)."""
    import pypdfium2.raw as pdfium_c

    scale = 2
    image = page.render(scale=scale).to_pil().convert("RGB")  # type: ignore[attr-defined]
    width, height = image.size

    def device(x: float, y: float) -> tuple[int, int]:
        dx, dy = ctypes.c_int(), ctypes.c_int()
        pdfium_c.FPDF_PageToDevice(page.raw, 0, 0, width, height, 0, x, y,  # type: ignore[attr-defined]
                                   ctypes.byref(dx), ctypes.byref(dy))
        return dx.value, dy.value

    pad = 3
    (x0, y0), (x1, y1) = device(box[0] - pad, box[1] - pad), device(box[2] + pad, box[3] + pad)
    left, right = sorted((x0, x1))
    top, bottom = sorted((y0, y1))
    ring = []
    for x in range(max(left, 0), min(right, width - 1) + 1):
        for y in (top, bottom):
            if 0 <= y < height:
                ring.append(image.getpixel((x, y)))
    for y in range(max(top, 0), min(bottom, height - 1) + 1):
        for x in (left, right):
            if 0 <= x < width:
                ring.append(image.getpixel((x, y)))
    if not ring:
        return 1.0, 1.0, 1.0
    r, g, b = Counter(ring).most_common(1)[0][0]
    return r / 255, g / 255, b / 255


def _align(box: tuple[float, float, float, float], mediabox: tuple[float, ...]) -> str:
    left, _bottom, right, _top = mediabox
    width = right - left
    center = (box[0] + box[2]) / 2
    if abs(center - (left + width / 2)) < 0.06 * width:
        return "center"
    return "right" if box[2] > left + 0.6 * width else "left"


def _page_text(textpage: object) -> str:
    """Textul paginii caracter cu caracter, ca indicii sa corespunda exact cu pdfium."""
    import pypdfium2.raw as pdfium_c

    raw = textpage.raw  # type: ignore[attr-defined]
    count = textpage.count_chars()  # type: ignore[attr-defined]
    return "".join(chr(pdfium_c.FPDFText_GetUnicode(raw, i) or 0xFFFD) for i in range(count))


def _find_on_page(page: object, index: int, count: int) -> Iterator[PageNumber]:
    import pypdfium2.raw as pdfium_c

    textpage = page.get_textpage()  # type: ignore[attr-defined]
    try:
        text = _page_text(textpage)
        mediabox = page.get_mediabox()  # type: ignore[attr-defined]
        bottom_edge, top_edge = mediabox[1], mediabox[3]
        zone = (top_edge - bottom_edge) * EDGE_ZONE
        taken: list[tuple[int, int]] = []
        for pattern in PATTERNS:
            for match in pattern.finditer(text):
                start, end = match.span()
                if any(a < end and start < b for a, b in taken):
                    continue
                n = int(match.group("n"))
                total = match.groupdict().get("total")
                if n != index + 1 or (total is not None and int(total) != count):
                    continue
                if not _alone_on_line(text, start, end):
                    continue
                boxes = [textpage.get_charbox(i) for i in range(start, end)
                         if not text[i].isspace()]
                boxes = [b for b in boxes if b[2] > b[0] and b[3] > b[1]]
                if not boxes:
                    continue
                box = (min(b[0] for b in boxes), min(b[1] for b in boxes),
                       max(b[2] for b in boxes), max(b[3] for b in boxes))
                if not (box[1] < bottom_edge + zone or box[3] > top_edge - zone):
                    continue

                digit = match.start("n")
                ox, oy = ctypes.c_double(), ctypes.c_double()
                pdfium_c.FPDFText_GetCharOrigin(textpage.raw, digit, ctypes.byref(ox),
                                                ctypes.byref(oy))
                font_name, bold = _font(textpage.raw, digit)
                taken.append((start, end))
                yield PageNumber(
                    template=_template(match),
                    box=box,
                    baseline=oy.value,
                    font_size=pdfium_c.FPDFText_GetFontSize(textpage.raw, digit) or 10.0,
                    color=_color(textpage.raw, digit),
                    background=_background(page, box),  # type: ignore[arg-type]
                    font_name=font_name,
                    bold=bold,
                    align=_align(box, mediabox),
                )
    finally:
        textpage.close()  # type: ignore[attr-defined]


def find_page_numbers(path: Path) -> list[list[PageNumber]]:
    """Pentru fiecare pagina din PDF: numerele de pagina tiparite pe ea."""
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(path))
    try:
        count = len(document)
        result = []
        for index in range(count):
            page = document[index]
            try:
                result.append(list(_find_on_page(page, index, count)))
            finally:
                page.close()
        return result
    finally:
        document.close()


# --- fonturi ------------------------------------------------------------------------------


def _normalize(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


@cache
def _system_fonts() -> dict[str, Path]:
    """Fonturile TrueType instalate: nume normalizat -> fisier."""
    fonts: dict[str, Path] = {}
    windows_fonts = Path(r"C:\Windows\Fonts")
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
        except OSError:
            continue
        with key:
            index = 0
            while True:
                try:
                    name, file, _type = winreg.EnumValue(key, index)
                except OSError:
                    break
                index += 1
                path = Path(file) if Path(file).is_absolute() else windows_fonts / file
                if path.suffix.lower() == ".ttf" and path.exists():
                    label = re.sub(r"\s*\((TrueType|OpenType)\)\s*$", "", name)
                    fonts.setdefault(_normalize(label), path)

    # fonturile descarcate de Office (Aptos etc.) au nume numerice; citim numele din fisier
    import os

    from PIL import ImageFont

    cloud = Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts"))
    if cloud.is_dir():
        for path in cloud.rglob("*.ttf"):
            try:
                family, style = ImageFont.truetype(str(path), 10).getname()
            except OSError:
                continue
            style = "" if (style or "").lower() == "regular" else (style or "")
            fonts.setdefault(_normalize(f"{family}{style}"), path)
    return fonts


def _font_file(pdf_font: str, bold: bool) -> Path | None:
    """„ABCDEF+Calibri,Bold” / „TimesNewRomanPS-BoldMT” -> fisierul fontului instalat."""
    name = pdf_font.split("+", 1)[-1]
    family = re.split(r"[,\-]", name, maxsplit=1)[0]
    family = re.sub(r"(PS|MT|PSMT)$", "", family)
    base = _normalize(family)
    fonts = _system_fonts()
    candidates = [base + "bold", base] if bold else [base, base + "regular"]
    candidates += ["arialbold", "arial"] if bold else ["arial"]
    for candidate in candidates:
        if candidate in fonts:
            return fonts[candidate]
    return None


@cache
def _register(path: Path) -> str:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    name = f"PageNumber-{_normalize(path.stem)}"
    try:
        pdfmetrics.registerFont(TTFont(name, str(path)))
    except Exception as exc:  # font cu restrictii sau format neasteptat
        log.debug("Nu pot folosi fontul %s: %s", path, exc)
        return "Helvetica-Bold" if "bold" in path.stem.lower() else "Helvetica"
    return name


# --- aplicare -----------------------------------------------------------------------------


def stamp(page: PageObject, numbers: list[PageNumber], page_no: int, total: int) -> None:
    """Acopera numerele vechi si scrie numerele noi pe pagina (deja unita)."""
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas

    box = page.mediabox
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(float(box.right), float(box.top)))
    for number in numbers:
        file = _font_file(number.font_name, number.bold)
        font = _register(file) if file else ("Helvetica-Bold" if number.bold else "Helvetica")
        text = number.text(page_no, total)
        width = stringWidth(text, font, number.font_size)
        left, bottom, right, top = number.box

        pad = 1.0
        c.setFillColorRGB(*number.background)
        c.rect(left - pad, bottom - pad, right - left + 2 * pad, top - bottom + 2 * pad,
               stroke=0, fill=1)

        if number.align == "center":
            x = (left + right) / 2 - width / 2
        elif number.align == "right":
            x = right - width
        else:
            x = left
        c.setFillColorRGB(*number.color)
        c.setFont(font, number.font_size)
        c.drawString(x, number.baseline, text)
    c.save()
    buf.seek(0)
    page.merge_page(PdfReader(buf).pages[0])
