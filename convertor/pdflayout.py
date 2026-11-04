"""PDF -> DOCX cu aspect fidel, pentru rapoarte pe coloane (balante, jurnale, extrase).

„PDF Reflow” din Word e facut pentru text curgator; pe rapoarte cu multe coloane ghiceste
gresit tabelele, muta sume in dreptul altor randuri si reface paginarea. Aici citim fiecare
caracter cu pozitia lui exacta, refacem randurile si coloanele, iar fiecare rand devine un
paragraf Word cu opriri de tab exact in dreptul coloanelor (sumele aliniate la dreapta).
Rezultatul arata ca originalul pagina cu pagina, ramane editabil, iar randurile copiate in
Excel se impart singure pe coloane (separatorul e tab-ul).

DOCX-ul e scris direct (WordprocessingML), fara Word si fara biblioteci suplimentare.
"""

from __future__ import annotations

import ctypes
import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from functools import cache
from pathlib import Path
from xml.sax.saxutils import escape

TWIPS = 20  # 1 pt = 20 twips
LINE_TOLERANCE = 0.35  # cat se pot abate liniile de baza ale aceluiasi rand (x marimea fontului)
COLUMN_GAP = 0.9  # un spatiu mai mare de atat (x marimea fontului) incepe o coloana noua
WORD_GAP = 0.15  # peste atat, intre doua caractere punem un spatiu
MARGIN = 14.0  # pt; mic, ca toata pagina originala sa incapa
NUMERIC = re.compile(r"^[-+(]?[\d.,]*\d[\d.,]*%?\)?$")
TABULAR_SHARE = 0.3  # cat din randuri trebuie sa aiba 3+ coloane ca PDF-ul sa fie „raport”
INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


@dataclass
class Segment:
    x0: float
    x1: float
    text: str
    size: float
    bold: bool
    font: str = ""  # numele fontului din PDF (de ex. „ABCDEF+Aptos”, „Helvetica-Bold”)

    @property
    def numeric(self) -> bool:
        return bool(NUMERIC.match(self.text))


@dataclass
class Line:
    baseline: float  # distanta de la marginea de sus a paginii
    size: float
    segments: list[Segment] = field(default_factory=list)


@dataclass
class Shape:
    """Dreptunghi sau linie dreapta (chenare de tabel, sublinieri, fundaluri)."""

    x0: float
    top: float  # de la marginea de sus a paginii
    x1: float
    bottom: float
    stroke: str | None  # culoare RRGGBB sau None
    fill: str | None
    width: float  # grosimea liniei, pt


@dataclass
class Page:
    width: float
    height: float
    lines: list[Line]
    shapes: list[Shape] = field(default_factory=list)


@dataclass
class _Char:
    ch: str
    x0: float
    x1: float
    baseline: float
    size: float
    bold: bool
    space_after: bool = False  # pdfium a detectat un spatiu intre cuvinte dupa caracter
    font: str = ""


# --- citire ---------------------------------------------------------------------------------


def _chars(page: object) -> list[_Char]:
    """Caracterele paginii cu pozitie. Latimea e cea de avans a fontului („loose box”), nu
    conturul desenat: altfel cifre inguste ca „1” par urmate de spatii."""
    import pypdfium2.raw as pdfium_c

    textpage = page.get_textpage()  # type: ignore[attr-defined]
    raw = textpage.raw
    top = page.get_size()[1]  # type: ignore[attr-defined]
    ox, oy = ctypes.c_double(), ctypes.c_double()
    chars: list[_Char] = []
    clip_by_object: dict[int, tuple[float, float, float, float] | None] = {}

    def object_key(index: int) -> int:
        return ctypes.cast(pdfium_c.FPDFText_GetTextObject(raw, index), ctypes.c_void_p).value or 0

    def clip_box(index: int) -> tuple[float, float, float, float] | None:
        """Masca de taiere a textului (de ex. celulele din PDF-urile exportate din Excel:
        textul lung continua dupa celula, dar nu se vede). Intersectia cailor mastii."""
        key = object_key(index)
        if key in clip_by_object:
            return clip_by_object[key]
        box = None
        clip = pdfium_c.FPDFPageObj_GetClipPath(pdfium_c.FPDFText_GetTextObject(raw, index))
        if clip:
            fx, fy = ctypes.c_float(), ctypes.c_float()
            for p in range(pdfium_c.FPDFClipPath_CountPaths(clip)):
                xs, ys = [], []
                for s in range(pdfium_c.FPDFClipPath_CountPathSegments(clip, p)):
                    segment = pdfium_c.FPDFClipPath_GetPathSegment(clip, p, s)
                    pdfium_c.FPDFPathSegment_GetPoint(segment, ctypes.byref(fx), ctypes.byref(fy))
                    xs.append(fx.value)
                    ys.append(fy.value)
                if xs:
                    path = (min(xs), min(ys), max(xs), max(ys))
                    box = path if box is None else (max(box[0], path[0]), max(box[1], path[1]),
                                                    min(box[2], path[2]), min(box[3], path[3]))
        clip_by_object[key] = box
        return box

    font_by_object: dict[int, tuple[str, bool]] = {}

    def font_of(index: int) -> tuple[str, bool]:
        """(nume font, aldin). Fonturile standard (Helvetica-Bold) raporteaza greutatea 0,
        deci ne uitam si la nume; o singura data pe obiect de text."""
        key = object_key(index)
        if key not in font_by_object:
            flags = ctypes.c_int()
            size = pdfium_c.FPDFText_GetFontInfo(raw, index, None, 0, ctypes.byref(flags))
            name = ""
            if size:
                buf = ctypes.create_string_buffer(size)
                pdfium_c.FPDFText_GetFontInfo(raw, index, buf, size, ctypes.byref(flags))
                name = buf.value.decode("utf-8", errors="replace")
            bold = (pdfium_c.FPDFText_GetFontWeight(raw, index) >= 600
                    or "bold" in name.lower() or "black" in name.lower())
            font_by_object[key] = (name, bold)
        return font_by_object[key]

    try:
        for i in range(textpage.count_chars()):
            code = pdfium_c.FPDFText_GetUnicode(raw, i)
            if not code or code in (0x0A, 0x0D):
                continue
            # pdfium marcheaza cu 0x02 cratima de la capat de rand; in pagina ea se vede
            ch = "-" if code == 0x02 else chr(code)
            if pdfium_c.FPDFText_IsGenerated(raw, i) == 1:
                if ch.isspace() and chars:
                    chars[-1].space_after = True
                continue
            left, _bottom, right, _top = textpage.get_charbox(i, loose=True)
            pdfium_c.FPDFText_GetCharOrigin(raw, i, ctypes.byref(ox), ctypes.byref(oy))
            clip = clip_box(i)
            if clip is not None:
                center = (left + right) / 2
                if not (clip[0] - 0.5 <= center <= clip[2] + 0.5
                        and clip[1] - 0.5 <= oy.value <= clip[3] + 0.5):
                    continue  # ascuns de masca: nu se vede nici in original
            size = pdfium_c.FPDFText_GetFontSize(raw, i) or 10.0
            if ch.isspace():
                if chars:
                    chars[-1].space_after = True
                continue
            font, bold = font_of(i)
            chars.append(_Char(ch, left, max(right, left), top - oy.value, size, bold,
                               font=font))
    finally:
        textpage.close()
    return chars


def _color(getter: object, obj: object) -> str | None:
    r, g, b, a = (ctypes.c_uint() for _ in range(4))
    if not getter(obj, *(ctypes.byref(v) for v in (r, g, b, a))) or a.value == 0:  # type: ignore[operator]
        return None
    return f"{r.value:02X}{g.value:02X}{b.value:02X}"


def _shapes(page: object) -> list[Shape]:
    """Dreptunghiurile si liniile drepte (orizontale/verticale) ale paginii. Cai complexe
    (grafice, logo-uri vectoriale) nu sunt transpuse."""
    import pypdfium2.raw as pdfium_c

    height = page.get_size()[1]  # type: ignore[attr-defined]
    shapes = []
    for obj in page.get_objects():  # type: ignore[attr-defined]
        if obj.type != pdfium_c.FPDF_PAGEOBJ_PATH:
            continue
        x0, y0, x1, y1 = obj.get_bounds()
        if not _axis_aligned(obj, (x0, y0, x1, y1)):
            continue
        fill_mode, stroke_on = ctypes.c_int(), ctypes.c_int()
        pdfium_c.FPDFPath_GetDrawMode(obj.raw, ctypes.byref(fill_mode), ctypes.byref(stroke_on))
        stroke = _color(pdfium_c.FPDFPageObj_GetStrokeColor, obj.raw) if stroke_on.value else None
        fill = _color(pdfium_c.FPDFPageObj_GetFillColor, obj.raw) if fill_mode.value else None
        if fill == "FFFFFF" and stroke is None:
            continue  # fundal alb: nu se vede
        if stroke is None and fill is None:
            continue
        width = ctypes.c_float()
        pdfium_c.FPDFPageObj_GetStrokeWidth(obj.raw, ctypes.byref(width))
        line = max(width.value, 0.25)
        if stroke:
            # chenarul raportat de pdfium include grosimea liniei; Word o deseneaza centrata
            # pe contur, deci revenim la conturul geometric
            half = line / 2
            if x1 - x0 > line:
                x0, x1 = x0 + half, x1 - half
            if y1 - y0 > line:
                y0, y1 = y0 + half, y1 - half
        shapes.append(Shape(x0, height - y1, x1, height - y0, stroke, fill, line))
    return shapes


def _axis_aligned(obj: object, bounds: tuple[float, float, float, float],
                  tolerance: float = 1.0) -> bool:
    """Toate varfurile caii cad in colturile chenarului: dreptunghi sau linie dreapta
    orizontala/verticala (nu triunghiuri, curbe sau linii oblice)."""
    import pypdfium2.raw as pdfium_c

    raw = obj.raw  # type: ignore[attr-defined]
    count = pdfium_c.FPDFPath_CountSegments(raw)
    if not 2 <= count <= 6:
        return False
    x0, y0, x1, y1 = bounds
    if (x1 - x0) <= 0.5 and (y1 - y0) <= 0.5:
        return False
    a, b, c, d, e, f = obj.get_matrix().get()  # type: ignore[attr-defined]
    px, py = ctypes.c_float(), ctypes.c_float()
    for i in range(count):
        segment = pdfium_c.FPDFPath_GetPathSegment(raw, i)
        if pdfium_c.FPDFPathSegment_GetType(segment) == pdfium_c.FPDF_SEGMENT_BEZIERTO:
            return False
        pdfium_c.FPDFPathSegment_GetPoint(segment, ctypes.byref(px), ctypes.byref(py))
        x = a * px.value + c * py.value + e
        y = b * px.value + d * py.value + f
        near_x = min(abs(x - x0), abs(x - x1)) <= tolerance
        near_y = min(abs(y - y0), abs(y - y1)) <= tolerance
        if not (near_x and near_y):
            return False
    return True


def _lines(chars: list[_Char]) -> list[Line]:
    visible = sorted((c for c in chars), key=lambda c: (c.baseline, c.x0))
    groups: list[list[_Char]] = []
    for c in visible:
        if groups:
            ref = groups[-1]
            base = sum(x.baseline for x in ref) / len(ref)
            if abs(c.baseline - base) <= LINE_TOLERANCE * max(c.size, ref[0].size):
                ref.append(c)
                continue
        groups.append([c])

    lines = []
    for group in groups:
        group.sort(key=lambda c: c.x0)
        line = Line(baseline=sum(c.baseline for c in group) / len(group),
                    size=max(c.size for c in group))
        current: list[_Char] = []
        prev: _Char | None = None
        for c in group:
            if prev is not None:
                gap = c.x0 - prev.x1
                if gap > COLUMN_GAP * c.size:
                    line.segments.append(_segment(current))
                    current = []
                elif prev.space_after or gap > WORD_GAP * c.size:
                    current.append(_Char(" ", prev.x1, c.x0, c.baseline, c.size, c.bold))
            current.append(c)
            prev = c
        if current:
            line.segments.append(_segment(current))
        lines.append(line)
    return lines


def _segment(chars: list[_Char]) -> Segment:
    text = "".join(c.ch for c in chars).strip()
    visible = [c for c in chars if not c.ch.isspace()]
    bold = sum(c.bold for c in visible) * 2 > len(visible)
    font = Counter(c.font for c in visible).most_common(1)[0][0]
    return Segment(chars[0].x0, chars[-1].x1, text, max(c.size for c in visible), bold, font)


# --- fonturi --------------------------------------------------------------------------------


@cache
def word_font(pdf_font: str, bold: bool) -> tuple[str, str | None]:
    """(familia pentru Word, fisierul TTF instalat) pentru un font din PDF.

    Helvetica devine Arial (aceleasi dimensiuni); fonturile Office descarcate (Aptos) sunt
    gasite si ele. Daca fontul nu e instalat: Arial."""
    from PIL import ImageFont

    from .pagenumbers import _font_file

    path = _font_file(pdf_font, bold)
    if path is None:
        return "Arial", None
    try:
        family = ImageFont.truetype(str(path), 10).getname()[0] or "Arial"
    except OSError:
        return "Arial", None
    return family, str(path)


@cache
def _measuring_font(path: str) -> object:
    from PIL import ImageFont

    return ImageFont.truetype(path, 100)


def text_width(text: str, path: str | None, size: float) -> float | None:
    """Latimea textului (pt) in fontul dat; None daca nu se poate masura."""
    if path is None:
        return None
    try:
        return _measuring_font(path).getlength(text) * size / 100  # type: ignore[attr-defined]
    except OSError:
        return None


def read_pages(path: Path, limit: int | None = None) -> list[Page]:
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(path))
    try:
        pages = []
        for index in range(len(document) if limit is None else min(limit, len(document))):
            page = document[index]
            try:
                width, height = page.get_size()
                pages.append(Page(width, height, _lines(_chars(page)), _shapes(page)))
            finally:
                page.close()
        return pages
    finally:
        document.close()


def looks_tabular(path: Path, sample_pages: int = 8) -> bool:
    """Raport pe coloane (multe randuri cu 3+ coloane), nu text curgator."""
    pages = read_pages(path, limit=sample_pages)
    lines = [line for page in pages for line in page.lines]
    if len(lines) < 5:
        return False
    columnar = sum(1 for line in lines if len(line.segments) >= 3)
    return columnar / len(lines) >= TABULAR_SHARE


# --- scriere DOCX ---------------------------------------------------------------------------

_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" '
    'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    "</Types>"
)
_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)
_DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/styles" Target="styles.xml"/>'
    "</Relationships>"
)
_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_DOC_NS = (
    _W + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)
EMU = 12700  # 1 pt
_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f"<w:styles {_W}><w:docDefaults><w:rPrDefault><w:rPr>"
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="ro-RO"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:qFormat/></w:style></w:styles>'
)

# Cu interliniere „exact”, Word pune linia de baza la 80% din inaltimea randului, de sus
# (masurat exportand prin Word randuri cu marimi si distante cunoscute).
BASELINE_RATIO = 0.8
LINE_HEIGHT = 1.2  # inaltimea normala a unui rand, x marimea fontului
MIN_LINE_HEIGHT = 0.5  # randurile aproape suprapuse primesc cel putin atat


def _twips(points: float) -> int:
    return max(0, round(points * TWIPS))


def _snap(points: float) -> float:
    """Valoarea exact cum o va folosi Word (rotunjita la twips)."""
    return _twips(points) / TWIPS


MIN_SCALE = 50  # % - scalarea orizontala minima acceptata de Word


def _run(text: str, size: float, bold: bool, family: str = "Arial", scale: int = 100) -> str:
    text = INVALID_XML.sub("", text)
    half_points = max(2, round(size * 2))
    props = ""
    if family != "Arial":  # Arial e implicit in stiluri
        name = escape(family, {'"': "&quot;"})
        props += f'<w:rFonts w:ascii="{name}" w:hAnsi="{name}" w:cs="{name}"/>'
    if bold:
        props += "<w:b/><w:bCs/>"
    if scale < 100:
        props += f'<w:w w:val="{scale}"/>'
    props += f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/>'
    return f'<w:r><w:rPr>{props}</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def _segment_run(seg: Segment, prefix: str = "") -> str:
    """Segmentul in fontul original; daca ar iesi mai lat decat in PDF (font lipsa sau
    inlocuit), il ingustam cu scalarea orizontala, ca sa nu impinga coloanele urmatoare."""
    family, path = word_font(seg.font, seg.bold)
    text = prefix + seg.text
    scale = 100
    measured = text_width(text, path, seg.size)
    original = seg.x1 - seg.x0
    if measured and original > 0 and measured > original * 1.02:
        scale = max(MIN_SCALE, int(100 * original / measured))
    return _run(text, seg.size, seg.bold, family, scale)


def _paragraph(line: Line, height: float, before: float, page_break: bool,
               margin: float) -> str:
    stops, runs, last_stop = [], [], 0
    for seg in line.segments:
        if seg.numeric:
            kind, pos = "right", _twips(seg.x1 - margin)
        else:
            kind, pos = "left", _twips(seg.x0 - margin)
        if not runs and kind == "left" and pos <= 0:
            # textul incepe chiar la margine: fara tab (Word ar sari la oprirea urmatoare)
            runs.append(_segment_run(seg))
            continue
        if pos <= last_stop:  # coloane suprapuse: lipim de segmentul anterior
            runs.append(_segment_run(seg, prefix=" "))
            continue
        stops.append(f'<w:tab w:val="{kind}" w:pos="{pos}"/>')
        runs.append("<w:r><w:tab/></w:r>")
        runs.append(_segment_run(seg))
        last_stop = pos
    ppr = (
        ("<w:pageBreakBefore/>" if page_break else "")
        + f"<w:tabs>{''.join(stops)}</w:tabs>"
        + f'<w:spacing w:before="{_twips(before)}" w:after="0" '
        f'w:line="{max(_twips(height), 40)}" w:lineRule="exact"/>'
    )
    return f"<w:p><w:pPr>{ppr}</w:pPr>{''.join(runs)}</w:p>"


def _shape_xml(shape: Shape, shape_id: int) -> str:
    """Forma ancorata pe pagina, in spatele textului, exact unde era in PDF."""
    width, height = shape.x1 - shape.x0, shape.bottom - shape.top
    if width < 0.5 or height < 0.5:  # linie dreapta
        geometry = "line"
        width, height = max(width, 0.0), max(height, 0.0)
    else:
        geometry = "rect"
    fill = (f'<a:solidFill><a:srgbClr val="{shape.fill}"/></a:solidFill>' if shape.fill
            else "<a:noFill/>")
    line = (f'<a:ln w="{round(shape.width * EMU)}"><a:solidFill><a:srgbClr '
            f'val="{shape.stroke}"/></a:solidFill></a:ln>' if shape.stroke
            else "<a:ln><a:noFill/></a:ln>")
    cx, cy = round(width * EMU), round(height * EMU)
    return (
        "<w:r><w:drawing>"
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{round(shape.x0 * EMU)}'
        "</wp:posOffset></wp:positionH>"
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{round(shape.top * EMU)}'
        "</wp:posOffset></wp:positionV>"
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:wrapNone/><wp:docPr id="{shape_id}" name="Forma {shape_id}"/>'
        "<wp:cNvGraphicFramePr/>"
        '<a:graphic><a:graphicData '
        'uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        "<wps:wsp><wps:cNvSpPr/><wps:spPr>"
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{geometry}"><a:avLst/></a:prstGeom>{fill}{line}'
        "</wps:spPr><wps:bodyPr/></wps:wsp>"
        "</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


def _section(page: Page, margin: float) -> str:
    orient = ' w:orient="landscape"' if page.width > page.height else ""
    m = _twips(margin)
    # fara margine in dreapta: textul de la marginea paginii originale nu trebuie sa treaca
    # pe randul urmator
    return (
        f'<w:sectPr><w:pgSz w:w="{_twips(page.width)}" w:h="{_twips(page.height)}"{orient}/>'
        f'<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="{m}" '
        'w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'
    )


def _with_shapes(paragraph: str, shapes: list[Shape], first_id: int) -> str:
    """Formele paginii se ancoreaza in primul ei paragraf (pozitia lor e fata de pagina)."""
    if not shapes:
        return paragraph
    drawings = "".join(_shape_xml(s, first_id + i) for i, s in enumerate(shapes))
    return paragraph.replace("</w:pPr>", "</w:pPr>" + drawings, 1)


def _page_body(page: Page, first_page: bool, margin: float) -> list[str]:
    body = []
    if not page.lines:
        return [f"<w:p><w:pPr>{'' if first_page else '<w:pageBreakBefore/>'}</w:pPr></w:p>"]
    # Fiecare rand are inaltimea lui normala; golurile vin din spatierea de dinainte. Tinem
    # minte unde s-a terminat efectiv randul anterior, ca erorile sa nu se adune.
    placed_end = 0.0  # marginea de sus e 0
    for index, line in enumerate(page.lines):
        height = _snap(LINE_HEIGHT * line.size)
        before = line.baseline - BASELINE_RATIO * height - placed_end
        if before < 0:  # rand foarte apropiat de cel anterior: il facem mai scund
            height = _snap(max((line.baseline - placed_end) / BASELINE_RATIO,
                               MIN_LINE_HEIGHT * line.size))
            before = 0.0
        before = _snap(before)
        placed_end += before + height
        body.append(_paragraph(line, height, before,
                               page_break=index == 0 and not first_page, margin=margin))
    return body


def _margin(pages: list[Page]) -> float:
    """Marginea stanga: cel mult MARGIN, dar niciodata la dreapta textului cel mai din stanga
    (altfel primele coloane s-ar suprapune)."""
    lefts = [seg.x0 for page in pages for line in page.lines for seg in line.segments]
    return max(0.0, min(MARGIN, min(lefts) - 0.5)) if lefts else MARGIN


def write_docx(pages: list[Page], dst: Path) -> None:
    margin = _margin(pages)
    body: list[str] = []
    shape_id = 1
    for index, page in enumerate(pages):
        page_body = _page_body(page, first_page=index == 0, margin=margin)
        page_body[0] = _with_shapes(page_body[0], page.shapes, shape_id)
        shape_id += len(page.shapes)
        body.extend(page_body)
        next_page = pages[index + 1] if index + 1 < len(pages) else None
        if next_page is not None and (next_page.width, next_page.height) != (page.width,
                                                                              page.height):
            # alta dimensiune de pagina: sectiune noua, care se incheie in ultimul paragraf
            body[-1] = body[-1].replace("<w:pPr>", f"<w:pPr>{_section(page, margin)}", 1)
    final = _section(pages[-1], margin) if pages else ""
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_DOC_NS}><w:body>{''.join(body)}{final}</w:body></w:document>"
    )
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_name(dst.name + ".part")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CONTENT_TYPES)
        z.writestr("_rels/.rels", _ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        z.writestr("word/styles.xml", _STYLES)
        z.writestr("word/document.xml", document)
    tmp.replace(dst)


def convert(src: Path, dst: Path) -> int:
    """Converteste PDF-ul; intoarce numarul de pagini."""
    pages = read_pages(src)
    write_docx(pages, dst)
    return len(pages)
