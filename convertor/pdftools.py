"""Unelte PDF: unire, impartire, parola, watermark, comprimare, imagini <-> PDF si OCR.

Bibliotecile PDF (pypdf, reportlab, pypdfium2) se incarca abia la prima folosire, ca
aplicatia sa porneasca repede.
"""

from __future__ import annotations

import logging
import math
import os
import re
import shutil
import subprocess
import tempfile
import threading
from collections.abc import Callable, Sequence
from functools import cache
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING

from .i18n import _
from .office import CREATE_NO_WINDOW
from .paths import state_dir

if TYPE_CHECKING:
    from PIL import Image
    from pypdf import PageObject, PdfReader, PdfWriter

log = logging.getLogger(__name__)

Progress = Callable[[int, int], None]

OCR_DPI = 300
PREFERRED_OCR_LANGS = ("ron", "eng")
TESSDATA_URL = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/{lang}.traineddata"
IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp")
A4 = (595.28, 841.89)


class PdfError(Exception):
    """Eroare cu mesaj prietenos."""


class Cancelled(Exception):
    pass


def open_pdf(path: Path) -> PdfReader:
    from pypdf import PdfReader

    try:
        reader = PdfReader(path)
    except Exception as exc:
        raise PdfError(_("Nu pot citi PDF-ul: {error}").format(error=exc)) from exc
    if reader.is_encrypted and not reader.decrypt(""):
        raise PdfError(_("PDF-ul este protejat cu parolă."))
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


def is_image(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTENSIONS


# --- unire si impartire --------------------------------------------------------------------


def merge(sources: Sequence[Path], dst: Path, renumber: bool = False) -> int:
    """Uneste PDF-urile (si imaginile, ca pagini). Cu `renumber`, numerele de pagina tiparite
    („8/8”, „1/2”) devin numere continue in documentul final („8/10”, „9/10”).
    Intoarce cate numere a inlocuit."""
    from pypdf import PdfReader, PdfWriter

    from . import pagenumbers

    if len(sources) < 2:
        raise PdfError(_("Pentru unire ai nevoie de cel puțin două fișiere."))
    writer = PdfWriter()
    found: list[list[pagenumbers.PageNumber]] = []
    for src in sources:
        if is_image(src):
            buf = BytesIO()
            images_to_pdf([src], buf)
            buf.seek(0)
            reader = PdfReader(buf)
            found.extend([] for _page in reader.pages)
            writer.append(reader)
            continue
        reader = open_pdf(src)
        if renumber:
            try:
                numbers = pagenumbers.find_page_numbers(src)
            except Exception as exc:  # o sursa ciudata nu trebuie sa opreasca unirea
                log.warning("Nu am putut citi numerele de pagina din %s: %s", src, exc)
                numbers = []
            found.extend(numbers or [[] for _page in reader.pages])
        else:
            found.extend([] for _page in reader.pages)
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
        raise PdfError(_("Numărul de pagini pe fișier trebuie să fie cel puțin 1."))
    return [(start, min(start + size - 1, total)) for start in range(1, total + 1, size)]


def ranges_parts(total: int, parts: int) -> list[PageRange]:
    """`parts` bucati cat mai egale; primele primesc cate o pagina in plus (10 in 3 -> 4,3,3)."""
    if parts < 1:
        raise PdfError(_("Numărul de părți trebuie să fie cel puțin 1."))
    if parts > total:
        raise PdfError(_("PDF-ul are doar {total} pagini, nu se poate împărți în {parts} părți.")
                       .format(total=total, parts=parts))
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
        raise PdfError(_("Scrie cel puțin un interval, de exemplu 1-200, 201-400."))
    result = []
    for part in parts:
        match = _RANGE.match(part.replace("–", "-"))
        if not match:
            raise PdfError(_("Interval neînțeles: „{part}”. Exemplu corect: 1-200, 201-400.")
                           .format(part=part))
        start = int(match.group(1))
        end = start if not match.group(2) else int(match.group(3) or total)
        if start < 1 or end < start:
            raise PdfError(_("Interval invalid: „{part}”.").format(part=part))
        if end > total:
            raise PdfError(_("Intervalul „{part}” depășește PDF-ul, care are {total} pagini.")
                           .format(part=part, total=total))
        result.append((start, end))
    return result


def describe_ranges(ranges: Sequence[PageRange], limit: int = 6) -> str:
    shown = ", ".join(f"{a}" if a == b else f"{a}–{b}" for a, b in ranges[:limit])
    return shown + (f" … (+{len(ranges) - limit})" if len(ranges) > limit else "")


def split(src: Path, out_dir: Path, ranges: Sequence[PageRange] | None = None) -> list[Path]:
    """Imparte PDF-ul dupa intervalele date (implicit: fiecare pagina separat)."""
    from pypdf import PdfWriter

    reader = open_pdf(src)
    total = len(reader.pages)
    if ranges is None:
        ranges = ranges_every(total, 1)
    out_dir.mkdir(parents=True, exist_ok=True)
    width = max(3, len(str(total)))
    written = []
    for start, end in ranges:
        if not 1 <= start <= end <= total:
            raise PdfError(_("Intervalul {start}-{end} nu există în PDF ({total} pagini).")
                           .format(start=start, end=end, total=total))
        writer = PdfWriter()
        for index in range(start - 1, end):
            writer.add_page(reader.pages[index])
        name = (f"{src.stem} - {_('pagina')} {start:0{width}}" if start == end else
                f"{src.stem} - {_('paginile')} {start:0{width}}-{end:0{width}}")
        dst = out_dir / f"{name}.pdf"
        _write(writer, dst)
        written.append(dst)
    return written


def protect(src: Path, dst: Path, password: str) -> None:
    from pypdf import PdfWriter

    if not password:
        raise PdfError(_("Parola nu poate fi goală."))
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
    from pypdf import PdfReader
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
    from pypdf import PdfWriter, Transformation

    text = text.strip()
    if not text:
        raise PdfError(_("Textul pentru watermark nu poate fi gol."))
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


# --- comprimare ---------------------------------------------------------------------------

# nivel -> (latura maxima a unei imagini in pixeli, calitate JPEG)
COMPRESSION = {"light": (2400, 85), "balanced": (1600, 72), "strong": (1100, 55)}


def compress(src: Path, dst: Path, level: str = "balanced",
             cancel: threading.Event | None = None) -> tuple[int, int]:
    """Micsoreaza PDF-ul: imaginile mari sunt reduse si recomprimate JPEG, continutul paginilor
    comprimat, obiectele identice unificate. Intoarce (marimea initiala, marimea finala).

    Daca rezultatul n-ar fi mai mic, se pastreaza o copie a originalului."""
    from PIL import Image
    from pypdf import PdfWriter

    max_side, quality = COMPRESSION.get(level, COMPRESSION["balanced"])
    writer = PdfWriter(clone_from=open_pdf(src))
    for page in writer.pages:
        if cancel is not None and cancel.is_set():
            raise Cancelled
        try:
            images = list(page.images)
        except Exception as exc:  # imagini intr-un format pe care pypdf nu il citeste
            log.debug("Imagini ilizibile pe o pagina din %s: %s", src, exc)
            images = []
        for image in images:
            try:
                pil = image.image
                if pil is None or pil.mode in ("1", "P") or "A" in pil.mode:
                    continue  # imagini alb-negru/paleta/transparente: deja mici sau riscante
                if max(pil.size) > max_side:
                    pil.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
                if pil.mode not in ("RGB", "L"):
                    pil = pil.convert("RGB")
                image.replace(pil, quality=quality)
            except Exception as exc:
                log.debug("Imagine lasata neschimbata in %s: %s", src, exc)
        page.compress_content_streams()
    writer.compress_identical_objects()  # implicit: duplicate si orfane
    _write(writer, dst)
    before, after = src.stat().st_size, dst.stat().st_size
    if after >= before:
        shutil.copyfile(src, dst)
        after = before
    return before, after


# --- imagini <-> PDF ----------------------------------------------------------------------


def _frames(path: Path) -> list[Image.Image]:
    """Toate cadrele unei imagini (TIFF-urile scanate pot avea mai multe pagini)."""
    from PIL import Image, ImageOps

    try:
        image = Image.open(path)
    except OSError as exc:
        raise PdfError(_("Nu pot citi imaginea {name}: {error}").format(name=path.name,
                                                                         error=exc)) from exc
    frames = []
    for index in range(getattr(image, "n_frames", 1)):
        image.seek(index)
        frame = ImageOps.exif_transpose(image.copy())  # poze facute cu telefonul rotit
        if frame.mode not in ("RGB", "L"):
            background = Image.new("RGB", frame.size, "white")
            frame = frame.convert("RGBA")
            background.paste(frame, mask=frame.getchannel("A"))
            frame = background
        frames.append(frame)
    return frames


def images_to_pdf(sources: Sequence[Path], dst: Path | BytesIO, page: str = "a4") -> int:
    """Imaginile devin pagini PDF. `page="a4"`: fiecare imagine e incadrata pe o pagina A4
    (portret sau peisaj dupa forma ei); `page="image"`: pagina are marimea imaginii.
    Intoarce numarul de pagini."""
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    target = BytesIO() if isinstance(dst, Path) else dst
    c = canvas.Canvas(target)
    pages = 0
    for src in sources:
        for frame in _frames(src):
            w, h = frame.size
            dpi = frame.info.get("dpi", (150, 150))[0] or 150
            if page == "image":
                width, height = w * 72 / dpi, h * 72 / dpi
                box = (0.0, 0.0, width, height)
            else:
                width, height = A4 if h >= w else (A4[1], A4[0])
                margin = 18.0
                scale = min((width - 2 * margin) / w, (height - 2 * margin) / h, 72 / dpi * 2)
                iw, ih = w * scale, h * scale
                box = ((width - iw) / 2, (height - ih) / 2, iw, ih)
            c.setPageSize((width, height))
            buf = BytesIO()
            if frame.mode == "L" or src.suffix.lower() in (".jpg", ".jpeg"):
                frame.save(buf, "JPEG", quality=90)
            else:
                frame.save(buf, "PNG", optimize=True)
            buf.seek(0)
            c.drawImage(ImageReader(buf), *box)
            c.showPage()
            pages += 1
    c.save()
    if isinstance(dst, Path):
        dst.parent.mkdir(parents=True, exist_ok=True)
        tmp = dst.with_name(dst.name + ".part")
        tmp.write_bytes(target.getvalue())
        tmp.replace(dst)
    return pages


def pdf_to_images(src: Path, out_dir: Path, fmt: str = "png", dpi: int = 150,
                  progress: Progress | None = None,
                  cancel: threading.Event | None = None) -> list[Path]:
    """Fiecare pagina devine o imagine PNG sau JPG."""
    import pypdfium2 as pdfium

    open_pdf(src)  # mesaj clar pentru PDF-uri cu parola
    document = pdfium.PdfDocument(str(src))
    written = []
    try:
        total = len(document)
        width = max(3, len(str(total)))
        out_dir.mkdir(parents=True, exist_ok=True)
        for index in range(total):
            if cancel is not None and cancel.is_set():
                raise Cancelled
            image = document[index].render(scale=dpi / 72).to_pil()
            dst = out_dir / f"{src.stem} - {_('pagina')} {index + 1:0{width}}.{fmt}"
            if fmt == "jpg":
                image.convert("RGB").save(dst, quality=88, dpi=(dpi, dpi))
            else:
                image.save(dst, optimize=True, dpi=(dpi, dpi))
            written.append(dst)
            if progress:
                progress(index + 1, total)
    finally:
        document.close()
    return written


# --- OCR ----------------------------------------------------------------------------------


def find_tesseract() -> Path | None:
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


def user_tessdata() -> Path:
    """Limbi OCR descarcate din aplicatie (fara drepturi de admin)."""
    return state_dir() / "tessdata"


def _tessdata_args() -> list[str]:
    folder = user_tessdata()
    return ["--tessdata-dir", str(folder)] if any(folder.glob("*.traineddata")) else []


@cache
def _list_langs(exe: Path, args: tuple[str, ...]) -> tuple[str, ...]:
    try:
        out = subprocess.run([str(exe), *args, "--list-langs"], capture_output=True, text=True,
                             timeout=30, creationflags=CREATE_NO_WINDOW).stdout
    except (OSError, subprocess.TimeoutExpired):
        return ()
    langs = [line.strip() for line in out.splitlines()[1:] if line.strip()]
    return tuple(lang for lang in langs if lang != "osd")


def tesseract_languages(exe: Path) -> tuple[str, ...]:
    return _list_langs(exe, tuple(_tessdata_args()))


def ocr_languages(exe: Path) -> str:
    available = tesseract_languages(exe)
    chosen = [lang for lang in PREFERRED_OCR_LANGS if lang in available]
    return "+".join(chosen or available[:1] or ["eng"])


def download_ocr_language(lang: str = "ron", progress: Progress | None = None) -> Path:
    """Descarca o limba OCR (tessdata_fast) in folderul utilizatorului. Limbile instalate
    deja cu Tesseract sunt copiate alaturi, pentru ca Tesseract foloseste un singur folder."""
    import urllib.request

    exe = find_tesseract()
    if exe is None:
        raise PdfError(_("Tesseract OCR nu este instalat."))
    folder = user_tessdata()
    folder.mkdir(parents=True, exist_ok=True)
    installed = exe.parent / "tessdata"
    for existing in installed.glob("*.traineddata"):
        target = folder / existing.name
        if not target.exists():
            shutil.copyfile(existing, target)

    dst = folder / f"{lang}.traineddata"
    tmp = dst.with_name(dst.name + ".part")
    request = urllib.request.Request(TESSDATA_URL.format(lang=lang),
                                     headers={"User-Agent": "DocumentConverter"})
    try:
        with urllib.request.urlopen(request, timeout=60) as response, open(tmp, "wb") as fh:
            total = int(response.headers.get("Content-Length") or 0)
            done = 0
            while chunk := response.read(64 * 1024):
                fh.write(chunk)
                done += len(chunk)
                if progress and total:
                    progress(done, total)
    except OSError as exc:
        tmp.unlink(missing_ok=True)
        raise PdfError(_("Descărcarea a eșuat: {error}").format(error=exc)) from exc
    if tmp.stat().st_size < 100_000:  # pagina de eroare, nu un model OCR
        tmp.unlink(missing_ok=True)
        raise PdfError(_("Fișierul descărcat nu este un model OCR valid."))
    tmp.replace(dst)
    _list_langs.cache_clear()
    log.info("Limba OCR %s descarcata in %s", lang, dst)
    return dst


def ocr(
    src: Path,
    dst: Path,
    progress: Progress | None = None,
    cancel: threading.Event | None = None,
    page_timeout: int = 180,
) -> None:
    """Face un PDF cautabil: fiecare pagina devine imagine + strat de text invizibil."""
    import pypdfium2 as pdfium
    from pypdf import PdfReader, PdfWriter

    exe = find_tesseract()
    if exe is None:
        raise PdfError(_("Tesseract OCR nu este instalat."))
    langs = ocr_languages(exe)
    extra = _tessdata_args()
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
                        [str(exe), *extra, str(img_path), str(base), "-l", langs,
                         "--dpi", str(OCR_DPI), "pdf"],
                        check=True,
                        capture_output=True,
                        timeout=page_timeout,
                        creationflags=CREATE_NO_WINDOW,
                    )
                except subprocess.TimeoutExpired as exc:
                    raise PdfError(_("OCR a depășit {seconds} s la pagina {page}.")
                                   .format(seconds=page_timeout, page=index + 1)) from exc
                except subprocess.CalledProcessError as exc:
                    detail = exc.stderr.decode(errors="replace").strip()
                    raise PdfError(_("Tesseract a eșuat la pagina {page}: {detail}")
                                   .format(page=index + 1, detail=detail)) from exc
                writer.append(PdfReader(base.with_suffix(".pdf")))
                if progress:
                    progress(index + 1, total)
            _write(writer, dst)
    finally:
        document.close()
    log.info("OCR %s -> %s (%d pagini, %s)", src, dst, total, langs)
