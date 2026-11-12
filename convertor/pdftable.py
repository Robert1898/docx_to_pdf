"""PDF -> Excel: rapoartele pe coloane devin tabele cu valori reale (numere, date).

Pleaca de la citirea din „aspect fidel” (randuri si segmente cu pozitii exacte):
  * coloanele sunt grupate pe tot documentul: numerele dupa marginea dreapta, textul dupa
    cea stanga; un segment intra in coloana cu care se suprapune cel mai mult;
  * randurile care se repeta identic pe multe pagini (antetul raportului) apar o singura data;
  * sumele devin numere („4,134,700.00” sau „1.234,56”), datele devin date Excel, iar
    codurile cu zerouri in fata („0012”) raman text.

Fisierul .xlsx e scris direct (SpreadsheetML), fara biblioteci suplimentare.
"""

from __future__ import annotations

import datetime as dt
import re
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape

from .pdflayout import Line, Page, Segment, read_pages

COLUMN_TOLERANCE = 4.0  # pt: pozitii mai apropiate de atat sunt aceeasi coloana
REPEATED_SHARE = 0.5  # randurile prezente identic pe cel putin jumatate din pagini = antet

MONTHS = {m: i + 1 for i, m in enumerate(
    ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"))}
MONTHS.update({"ian": 1, "mai": 5, "iun": 6, "iul": 7, "noi": 11})
_DATE_TEXT = re.compile(r"^(\d{1,2})[-./ ]([A-Za-z]{3})[a-z]*[-./ ](\d{4})$")
_DATE_NUM = re.compile(r"^(\d{1,2})[./](\d{1,2})[./](\d{4})$")
_DATE_ISO = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")
_US_NUMBER = re.compile(r"^-?\d{1,3}(,\d{3})+(\.\d+)?$|^-?\d+\.\d+$")
_RO_NUMBER = re.compile(r"^-?\d{1,3}(\.\d{3})+(,\d+)?$|^-?\d+,\d+$")
_INTEGER = re.compile(r"^-?\d+$")
EXCEL_EPOCH = dt.date(1899, 12, 30)


@dataclass
class Cell:
    value: str | float | int | dt.date
    bold: bool = False
    decimals: int = 0


@dataclass
class _Column:
    left: float
    right: float


# --- valori --------------------------------------------------------------------------------


def parse_value(text: str) -> tuple[str | float | int | dt.date, int]:
    """(valoare, zecimale). Textul care nu e clar numar sau data ramane text."""
    raw = text.strip()
    negative = raw.startswith("(") and raw.endswith(")")  # (1,234.00) = negativ in contabilitate
    core = raw[1:-1] if negative else raw
    if core.endswith("-") and core[:-1].replace(",", "").replace(".", "").isdigit():
        core, negative = core[:-1], True  # 1,234.00- (unele rapoarte ERP)

    if (m := _DATE_TEXT.match(core)) and m.group(2).lower() in MONTHS:
        try:
            return dt.date(int(m.group(3)), MONTHS[m.group(2).lower()], int(m.group(1))), 0
        except ValueError:
            return raw, 0
    for pattern, order in ((_DATE_NUM, (3, 2, 1)), (_DATE_ISO, (1, 2, 3))):
        if m := pattern.match(core):
            try:
                return dt.date(*(int(m.group(i)) for i in order)), 0
            except ValueError:
                return raw, 0

    if _INTEGER.match(core):
        if len(core.lstrip("-")) > 1 and core.lstrip("-").startswith("0"):
            return raw, 0  # cod cu zerouri in fata
        if len(core.lstrip("-")) > 15:
            return raw, 0  # CNP, IBAN-uri numerice: Excel ar pierde cifre
        number = int(core)
        return (-number if negative else number), 0
    if _US_NUMBER.match(core):
        whole, _sep, frac = core.replace(",", "").partition(".")
    elif _RO_NUMBER.match(core):
        whole, _sep, frac = core.replace(".", "").partition(",")
    else:
        return raw, 0
    number = float(f"{whole}.{frac}" if frac else whole)
    return (-number if negative else number), len(frac)


# --- coloane si randuri -------------------------------------------------------------------


RIVER_SHARE = 0.05  # o coloana verticala goala pe 95% din randuri desparte doua coloane
MIN_RIVER = 2  # pt
TYPICAL_SHARE = 0.3  # randurile „tipice” trebuie sa fie macar atat din randurile cu 3+ coloane


def _columns(lines: list[Line]) -> list[_Column]:
    """Coloanele tabelului.

    Rapoartele au randuri de date cu o structura fixa (aici: nr., inventar, denumire, data,
    durata, patru sume), plus randuri de sectiune si de total cu alta asezare. Randurile
    „tipice” (cel mai des intalnit numar de coloane) dau intervalul exact al fiecarei
    coloane. Fara o structura dominanta, coloanele sunt zonele dintre „culoarele”
    verticale goale pe aproape toate randurile."""
    shaped = [line for line in lines if len(line.segments) >= 3]
    if shaped:
        count, freq = Counter(len(line.segments) for line in shaped).most_common(1)[0]
        if freq >= TYPICAL_SHARE * len(shaped):
            typical = [line.segments for line in shaped if len(line.segments) == count]
            return [_Column(min(row[i].x0 for row in typical), max(row[i].x1 for row in typical))
                    for i in range(count)]
    return _rivers([line.words for line in lines])


def _rivers(lines: list[list[Segment]]) -> list[_Column]:
    if not lines:
        return []
    width = int(max(s.x1 for line in lines for s in line) + 2)
    cover = [0] * (width + 1)
    for line in lines:
        hit = bytearray(width + 1)
        for seg in line:
            for x in range(max(0, int(seg.x0)), min(width, int(seg.x1) + 1)):
                hit[x] = 1
        for x, flag in enumerate(hit):
            cover[x] += flag
    threshold = max(1, RIVER_SHARE * len(lines))
    columns, start, gap = [], None, 0
    for x, count in enumerate(cover):
        if count > threshold:
            if start is None:
                start = x
            gap = 0
        elif start is not None:
            gap += 1
            if gap >= MIN_RIVER:
                columns.append(_Column(start, x - gap + 1))
                start, gap = None, 0
    if start is not None:
        columns.append(_Column(start, width))
    return columns


def _column_of(seg: Segment, columns: list[_Column]) -> int:
    def overlap(col: _Column) -> float:
        return min(seg.x1, col.right) - max(seg.x0, col.left)

    best = max(range(len(columns)), key=lambda i: overlap(columns[i]))
    if overlap(columns[best]) > 0:
        return best
    middle = (seg.x0 + seg.x1) / 2
    return min(range(len(columns)),
               key=lambda i: abs(middle - (columns[i].left + columns[i].right) / 2))


def _line_key(segments: list[Segment]) -> tuple:
    """Randul fara cifre (numarul paginii difera de la o pagina la alta)."""
    return tuple((round(s.x0), re.sub(r"\d", "#", s.text)) for s in segments)


def _is_continuation(row: list[Cell | None], previous: list[Cell | None] | None) -> int | None:
    """Randul e continuarea unui text rupt pe doua randuri (de ex. o descriere lunga)?
    Intoarce coloana in care se lipeste, sau None."""
    if previous is None:
        return None
    filled = [i for i, cell in enumerate(row) if cell is not None]
    if len(filled) != 1:
        return None
    index = filled[0]
    cell, above = row[index], previous[index]
    if not isinstance(cell.value, str) or above is None or not isinstance(above.value, str):
        return None
    # randul de deasupra trebuie sa fie un rand de date (cu cel putin inca o valoare)
    return index if sum(c is not None for c in previous) >= 2 else None


def rows_from_pages(pages: list[Page], drop_repeated: bool = True,
                    join_wrapped: bool = True) -> list[list[Cell | None]]:
    if not pages:
        return []
    repeated: set[tuple] = set()
    if drop_repeated and len(pages) >= 3:
        counts: Counter[tuple] = Counter()
        for page in pages:
            counts.update({_line_key(line.segments) for line in page.lines})
        repeated = {k for k, n in counts.items() if n >= REPEATED_SHARE * len(pages)}

    data_lines = [line for page in pages for line in page.lines
                  if _line_key(line.segments) not in repeated]
    columns = _columns(data_lines or [line for p in pages for line in p.lines])
    if not columns:
        return []

    rows: list[list[Cell | None]] = []
    seen_repeated: set[tuple] = set()
    for page in pages:
        for line in page.lines:
            key = _line_key(line.segments)
            if key in repeated:
                if key in seen_repeated:
                    continue
                seen_repeated.add(key)
            row: list[Cell | None] = [None] * len(columns)
            by_column: dict[int, list[Segment]] = {}
            for word in line.words:
                by_column.setdefault(_column_of(word, columns), []).append(word)
            for index, words in by_column.items():
                value, decimals = parse_value(" ".join(w.text for w in words))
                bold = sum(w.bold for w in words) * 2 > len(words)
                row[index] = Cell(value, bold, decimals)
            target = _is_continuation(row, rows[-1] if rows else None) if join_wrapped else None
            if target is not None:
                above = rows[-1][target]
                assert above is not None
                rows[-1][target] = Cell(f"{above.value} {row[target].value}", above.bold)
                continue
            rows.append(row)
    return rows


# --- scriere XLSX --------------------------------------------------------------------------

_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" '
    'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    "</Types>"
)
_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
)
_BOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/styles" Target="styles.xml"/></Relationships>'
)
_MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
_REL = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

# stiluri: 0 normal, 1 aldin, 2 numar cu 2 zecimale, 3 intreg cu separator, 4 data,
# 5 numar cu 2 zecimale aldin, 6 numar cu 1 zecimala, 7..9 = 3..4 zecimale / text aldin
_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f"<styleSheet {_MAIN}>"
    '<numFmts count="4"><numFmt numFmtId="164" formatCode="#,##0.00"/>'
    '<numFmt numFmtId="165" formatCode="dd.mm.yyyy"/>'
    '<numFmt numFmtId="166" formatCode="#,##0.0"/>'
    '<numFmt numFmtId="167" formatCode="#,##0.000"/></numFmts>'
    '<fonts count="2"><font><sz val="10"/><name val="Arial"/></font>'
    '<font><b/><sz val="10"/><name val="Arial"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border/></borders>'
    '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
    '<cellXfs count="8">'
    '<xf fontId="0"/>'
    '<xf fontId="1" applyFont="1"/>'
    '<xf numFmtId="164" fontId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="3" fontId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="165" fontId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="164" fontId="1" applyNumberFormat="1" applyFont="1"/>'
    '<xf numFmtId="166" fontId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="167" fontId="0" applyNumberFormat="1"/>'
    "</cellXfs></styleSheet>"
)
INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def _column_letter(index: int) -> str:
    letters = ""
    index += 1
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def _cell_xml(ref: str, cell: Cell) -> str:
    value = cell.value
    if isinstance(value, dt.date):
        return f'<c r="{ref}" s="4"><v>{(value - EXCEL_EPOCH).days}</v></c>'
    if isinstance(value, int | float) and not isinstance(value, bool):
        if isinstance(value, int):
            style = 1 if cell.bold else 3
        else:
            style = {0: 3, 1: 6, 2: 5 if cell.bold else 2}.get(cell.decimals, 7)
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    text = escape(INVALID_XML.sub("", str(value)))
    style = ' s="1"' if cell.bold else ""
    return f'<c r="{ref}" t="inlineStr"{style}><is><t xml:space="preserve">{text}</t></is></c>'


def _width(cells: list[Cell | None]) -> float:
    longest = 0
    for cell in cells:
        if cell is None:
            continue
        value = cell.value
        if isinstance(value, dt.date):
            longest = max(longest, 10)
        elif isinstance(value, float):
            longest = max(longest, len(f"{value:,.2f}"))
        else:
            longest = max(longest, len(str(value)))
    return min(max(longest + 2, 6), 60)


def write_xlsx(rows: list[list[Cell | None]], dst: Path, sheet: str = "Date") -> None:
    columns = max((len(r) for r in rows), default=0)
    widths = [_width([r[i] for r in rows if i < len(r)]) for i in range(columns)]
    cols = "".join(f'<col min="{i + 1}" max="{i + 1}" width="{w:.1f}" customWidth="1"/>'
                   for i, w in enumerate(widths))
    body = []
    for r, row in enumerate(rows, 1):
        cells = "".join(_cell_xml(f"{_column_letter(c)}{r}", cell)
                        for c, cell in enumerate(row) if cell is not None)
        body.append(f'<row r="{r}">{cells}</row>')
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<worksheet {_MAIN} {_REL}>"
        + (f"<cols>{cols}</cols>" if cols else "")
        + f"<sheetData>{''.join(body)}</sheetData></worksheet>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook {_MAIN} {_REL}><sheets><sheet name="{escape(sheet)}" sheetId="1" '
        'r:id="rId1"/></sheets></workbook>'
    )
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_name(dst.name + ".part")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CONTENT_TYPES)
        z.writestr("_rels/.rels", _ROOT_RELS)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", _BOOK_RELS)
        z.writestr("xl/styles.xml", _STYLES)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    tmp.replace(dst)


def convert(src: Path, dst: Path, drop_repeated: bool = True) -> int:
    """Converteste PDF-ul; intoarce numarul de randuri scrise."""
    rows = rows_from_pages(read_pages(src, shapes=False), drop_repeated=drop_repeated)
    write_xlsx(rows, dst)
    return len(rows)
