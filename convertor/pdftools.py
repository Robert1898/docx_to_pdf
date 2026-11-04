"""Unelte PDF: unire, impartire, parola, watermark si OCR (Tesseract)."""

from __future__ import annotations

import logging
import math
import re
import shutil
import subprocess
import tempfile
import threading
from collections.abc import Callable, Sequence
from functools import cache
from io import BytesIO
from pathlib import Path

from pypdf import PageObject, PdfReader, PdfWriter, Transformation

from .office import CREATE_NO_WINDOW

log = logging.getLogger(__name__)

Progress = Callable[[int, int], None]

OCR_DPI = 300
PREFERRED_OCR_LANGS = ("ron", "eng")


class PdfError(Exception):
    """Eroare cu mesaj prietenos."""


class Cancelled(Exception):
    pass


def open_pdf(path: Path) -> PdfReader:
    try:
        reader = PdfReader(path)
    except Exception as exc:
        raise PdfError(f"Nu pot citi PDF-ul: {exc}") from exc
    if reader.is_encrypted and not reader.decrypt(""):
        raise PdfError("PDF-ul este protejat cu parolă.")
    return reader


def page_count(path: Path) -> int:
    return len(open_pdf(path).pages)


def is_scanned(path: Path, sample: int = 3) -> bool:
    """PDF fara strat de text (de obicei scanat): primele pagini nu au aproape deloc text."""
    reader = open_pdf(path)
    pages = reader.pages[:sample]
    if not pages:
        return False
    chars = sum(len((page.extract_text() or "").strip()) for page in pages)
    return chars < 20 * len(pages)


def _write(writer: PdfWriter, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_name(dst.name + ".part")
    with open(tmp, "wb") as fh:
        writer.write(fh)
    tmp.replace(dst)


def merge(sources: Sequence[Path], dst: Path, renumber: bool = False) -> int:
    """Uneste PDF-urile. Cu `renumber`, numerele de pagina tiparite („8/8”, „1/2”) devin
    numere continue in documentul final („8/10”, „9/10”). Intoarce cate numere a inlocuit."""
    from . import pagenumbers

    if len(sources) < 2:
        raise PdfError("Pentru unire ai nevoie de cel puțin două PDF-uri.")
    writer = PdfWriter()
    found: list[list[pagenumbers.PageNumber]] = []
    for src in sources:
        reader = open_pdf(src)
        if renumber:
            try:
                numbers = pagenumbers.find_page_numbers(src)
            except Exception as exc:  # o sursa ciudata nu trebuie sa opreasca unirea
                log.warning("Nu am putut citi numerele de pagina din %s: %s", src, exc)
                numbers = []
            found.extend(numbers or [[] for _ in reader.pages])
        writer.append(reader)

    replaced = 0
    if renumber:
        total = len(writer.pages)
        for page_no, (page, numbers) in enumerate(zip(writer.pages, found, strict=True), 1):
            if numbers:
                pagenumbers.stamp(page, numbers, page_no, total)
                replaced += len(numbers)
    _write(writer, dst)
    return replaced


PageRange = tuple[int, int]  # prima si ultima pagina, numerotate de la 1, inclusiv


def ranges_every(total: int, size: int) -> list[PageRange]:
    """Bucati de cate `size` pagini: 600 de pagini cate 200 -> 1-200, 201-400, 401-600."""
    if size < 1:
        raise PdfError("Numărul de pagini pe fișier trebuie să fie cel puțin 1.")
    return [(start, min(start + size - 1, total)) for start in range(1, total + 1, size)]


def ranges_parts(total: int, parts: int) -> list[PageRange]:
    """`parts` bucati cat mai egale; primele primesc cate o pagina in plus (10 in 3 -> 4,3,3)."""
    if parts < 1:
        raise PdfError("Numărul de părți trebuie să fie cel puțin 1.")
    if parts > total:
        raise PdfError(f"PDF-ul are doar {total} pagini, nu se poate împărți în {parts} părți.")
    base, extra = divmod(total, parts)
    result, start = [], 1
    for i in range(parts):
        end = start + base + (i < extra) - 1
        result.append((start, end))
        start = end + 1
    return result


_RANGE = re.compile(r"^(\d+)\s*(?:(-)\s*(\d*))?$")


def parse_ranges(text: str, total: int) -> list[PageRange]:
    """„1-200, 201-450, 451-” -> [(1, 200), (201, 450), (451, total)].

    Accepta virgula sau punct si virgula intre intervale, „5” pentru o singura pagina si
    „451-” pentru „pana la sfarsit”."""
    parts = [p.strip() for p in re.split(r"[,;]", text) if p.strip()]
    if not parts:
        raise PdfError("Scrie cel puțin un interval, de exemplu 1-200, 201-400.")
    result = []
    for part in parts:
        match = _RANGE.match(part.replace("–", "-"))
        if not match:
            raise PdfError(f"Interval neînțeles: „{part}”. Exemplu corect: 1-200, 201-400.")
        start = int(match.group(1))
        end = start if not match.group(2) else int(match.group(3) or total)
        if start < 1 or end < start:
            raise PdfError(f"Interval invalid: „{part}”.")
        if end > total:
            raise PdfError(f"Intervalul „{part}” depășește PDF-ul, care are {total} pagini.")
        result.append((start, end))
    return result


def describe_ranges(ranges: Sequence[PageRange], limit: int = 6) -> str:
    shown = ", ".join(f"{a}" if a == b else f"{a}–{b}" for a, b in ranges[:limit])
    return shown + (f" … (+{len(ranges) - limit})" if len(ranges) > limit else "")


def split(src: Path, out_dir: Path, ranges: Sequence[PageRange] | None = None) -> list[Path]:
    """Imparte PDF-ul dupa intervalele date (implicit: fiecare pagina separat)."""
    reader = open_pdf(src)
    total = len(reader.pages)
    if ranges is None:
        ranges = ranges_every(total, 1)
    out_dir.mkdir(parents=True, exist_ok=True)
    width = max(3, len(str(total)))
    written = []
    for start, end in ranges:
        if not 1 <= start <= end <= total:
            raise PdfError(f"Intervalul {start}-{end} nu există în PDF ({total} pagini).")
        writer = PdfWriter()
        for index in range(start - 1, end):
            writer.add_page(reader.pages[index])
        name = (f"{src.stem} - pagina {start:0{width}}" if start == end else
                f"{src.stem} - paginile {start:0{width}}-{end:0{width}}")
        dst = out_dir / f"{name}.pdf"
        _write(writer, dst)
        written.append(dst)
    return written


def protect(src: Path, dst: Path, password: str) -> None:
    if not password:
        raise PdfError("Parola nu poate fi goală.")
    writer = PdfWriter(clone_from=open_pdf(src))
    writer.encrypt(user_password=password, owner_password=None, algorithm="AES-256")
    _write(writer, dst)


# --- watermark ----------------------------------------------------------------------------


@cache
def _watermark_font() -> str:
    """Arial suporta diacriticele romanesti; Helvetica din PDF nu are ș și ț."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    for path in (r"C:\Windows\Fonts\arialbd.ttf", r"C:\Windows\Fonts\arial.ttf"):
        if Path(path).exists():
            pdfmetrics.registerFont(TTFont("Watermark", path))
            return "Watermark"
    return "Helvetica-Bold"


def _watermark_page(width: float, height: float, text: str, opacity: float) -> PageObject:
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas

    font = _watermark_font()
    angle = math.degrees(math.atan2(height, width))
    diagonal = math.hypot(width, height)
    size = 0.75 * diagonal / max(stringWidth(text, font, 1), 1)
    size = max(12.0, min(size, 160.0))

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    c.setFillColorRGB(0.55, 0.55, 0.55, alpha=opacity)
    c.translate(width / 2, height / 2)
    c.rotate(angle)
    c.setFont(font, size)
    c.drawCentredString(0, -size / 3, text)
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def watermark(src: Path, dst: Path, text: str, opacity: float = 0.15) -> None:
    text = text.strip()
    if not text:
        raise PdfError("Textul pentru watermark nu poate fi gol.")
    writer = PdfWriter(clone_from=open_pdf(src))
    stamps: dict[tuple[float, float], PageObject] = {}
    for page in writer.pages:
        box = page.mediabox
        size = (float(box.width), float(box.height))
        if size not in stamps:
            stamps[size] = _watermark_page(*size, text, opacity)
        page.merge_transformed_page(
            stamps[size], Transformation().translate(float(box.left), float(box.bottom))
        )
    _write(writer, dst)


# --- OCR ----------------------------------------------------------------------------------


def find_tesseract() -> Path | None:
    import os

    candidates = [
        shutil.which("tesseract"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return Path(candidate)
    return None


@cache
def tesseract_languages(exe: Path) -> tuple[str, ...]:
    try:
        out = subprocess.run(
            [str(exe), "--list-langs"],
            capture_output=True,
            text=True,
            timeout=30,
            creationflags=CREATE_NO_WINDOW,
        ).stdout
    except (OSError, subprocess.TimeoutExpired):
        return ()
    langs = [line.strip() for line in out.splitlines()[1:] if line.strip()]
    return tuple(lang for lang in langs if lang != "osd")


def ocr_languages(exe: Path) -> str:
    available = tesseract_languages(exe)
    chosen = [lang for lang in PREFERRED_OCR_LANGS if lang in available]
    return "+".join(chosen or available[:1] or ["eng"])


def ocr(
    src: Path,
    dst: Path,
    progress: Progress | None = None,
    cancel: threading.Event | None = None,
    page_timeout: int = 180,
) -> None:
    """Face un PDF cautabil: fiecare pagina devine imagine + strat de text invizibil."""
    import pypdfium2 as pdfium

    exe = find_tesseract()
    if exe is None:
        raise PdfError("Tesseract OCR nu este instalat.")
    langs = ocr_languages(exe)
    open_pdf(src)  # mesaj clar pentru PDF-uri cu parola

    document = pdfium.PdfDocument(str(src))
    try:
        total = len(document)
        writer = PdfWriter()
        with tempfile.TemporaryDirectory(prefix="convertor-ocr-") as tmp:
            for index in range(total):
                if cancel is not None and cancel.is_set():
                    raise Cancelled
                image = document[index].render(scale=OCR_DPI / 72).to_pil().convert("RGB")
                img_path = Path(tmp) / f"p{index:05}.jpg"
                image.save(img_path, quality=80, dpi=(OCR_DPI, OCR_DPI))
                base = Path(tmp) / f"p{index:05}"
                try:
                    subprocess.run(
                        [str(exe), str(img_path), str(base), "-l", langs,
                         "--dpi", str(OCR_DPI), "pdf"],
                        check=True,
                        capture_output=True,
                        timeout=page_timeout,
                        creationflags=CREATE_NO_WINDOW,
                    )
                except subprocess.TimeoutExpired as exc:
                    raise PdfError(
                        f"OCR a depășit {page_timeout} s la pagina {index + 1}."
                    ) from exc
                except subprocess.CalledProcessError as exc:
                    detail = exc.stderr.decode(errors="replace").strip()
                    raise PdfError(f"Tesseract a eșuat la pagina {index + 1}: {detail}") from exc
                writer.append(PdfReader(base.with_suffix(".pdf")))
                if progress:
                    progress(index + 1, total)
            _write(writer, dst)
    finally:
        document.close()
    log.info("OCR %s -> %s (%d pagini, %s)", src, dst, total, langs)
