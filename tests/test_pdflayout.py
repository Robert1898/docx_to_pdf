from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest
from reportlab.pdfgen import canvas

from convertor import pdflayout

ROWS = [("1", "10_M", "Licenta program", "31-jan-2004", "36", "19,259.69", "0.00"),
        ("2", "119_M", "LICENTA TRAMVAI", "30-apr-2010", "60", "4,134,700.00", "0.00"),
        ("3", "11_M", "Design interior", "31-mar-2004", "36", "1,200.00", "0.00")]
# coloane: stanga pentru text, dreapta pentru numere (ca in rapoarte)
LEFT = (40, 70, 120, 330)
RIGHT = (420, 520, 580)


def report_pdf(path: Path, rows=ROWS, pages: int = 1) -> Path:
    c = canvas.Canvas(str(path), pagesize=(612, 792))
    for _ in range(pages):
        c.setFont("Helvetica-Bold", 10)
        c.drawString(40, 750, "Jurnal Balanta")
        c.rect(36, 700, 550, 14)
        c.setFont("Helvetica", 7)
        y = 690
        for row in rows * 4:
            for x, text in zip(LEFT, row[:4], strict=True):
                c.drawString(x, y, text)
            for x, text in zip(RIGHT, row[4:], strict=True):
                c.drawRightString(x, y, text)
            y -= 10
        c.showPage()
    c.save()
    return path


def prose_pdf(path: Path) -> Path:
    c = canvas.Canvas(str(path), pagesize=(612, 792))
    c.setFont("Helvetica", 11)
    y = 740
    for i in range(30):
        c.drawString(72, y, f"Acesta este un paragraf obisnuit de text, randul {i}, fara coloane.")
        y -= 14
    c.save()
    return path


def test_reads_rows_and_columns(tmp_path: Path) -> None:
    [page] = pdflayout.read_pages(report_pdf(tmp_path / "r.pdf"))
    title, *rows = page.lines
    assert [s.text for s in title.segments] == ["Jurnal Balanta"]
    assert title.segments[0].bold
    assert [s.text for s in rows[0].segments] == list(ROWS[0])
    numbers = rows[1].segments[4:]
    assert [s.numeric for s in numbers] == [True, True, True]
    assert abs(numbers[1].x1 - RIGHT[1]) < 1  # numerele aliniate la dreapta


def test_reads_rectangles(tmp_path: Path) -> None:
    [page] = pdflayout.read_pages(report_pdf(tmp_path / "r.pdf"))
    [box] = page.shapes
    assert (round(box.x0), round(box.x1)) == (36, 586)
    assert box.stroke == "000000" and box.fill is None


def test_detects_reports_vs_prose(tmp_path: Path) -> None:
    assert pdflayout.looks_tabular(report_pdf(tmp_path / "r.pdf"))
    assert not pdflayout.looks_tabular(prose_pdf(tmp_path / "p.pdf"))


def test_hidden_overflow_text_is_dropped(tmp_path: Path) -> None:
    """Ca in PDF-urile din Excel: textul lung continua, dar e taiat de masca celulei."""
    path = tmp_path / "clip.pdf"
    c = canvas.Canvas(str(path), pagesize=(612, 792))
    c.saveState()
    clip = c.beginPath()
    clip.rect(40, 690, 60, 20)
    c.clipPath(clip, stroke=0, fill=0)
    c.setFont("Helvetica", 9)
    c.drawString(42, 700, "Vizibil ascunsdupamargine")
    c.restoreState()
    c.drawString(200, 700, "Coloana")
    c.save()
    [page] = pdflayout.read_pages(path)
    texts = [s.text for line in page.lines for s in line.segments]
    assert texts[-1] == "Coloana"
    assert texts[0].startswith("Vizibil") and "margine" not in texts[0]


def test_writes_valid_docx_with_tab_stops(tmp_path: Path) -> None:
    dst = tmp_path / "r.docx"
    assert pdflayout.convert(report_pdf(tmp_path / "r.pdf", pages=2), dst) == 2
    with zipfile.ZipFile(dst) as z:
        assert z.testzip() is None
        xml = z.read("word/document.xml").decode("utf-8")
    assert xml.count("<w:pageBreakBefore/>") == 1  # a doua pagina incepe pe pagina noua
    assert 'w:val="right"' in xml and 'w:val="left"' in xml  # numere la dreapta, text la stanga
    assert "4,134,700.00" in xml
    assert xml.count("<wps:wsp>") == 2  # chenarul de pe fiecare pagina
    assert re.search(r'<w:pgSz w:w="12240" w:h="15840"/>', xml)  # Letter, ca in PDF


def test_empty_input_is_rejected_cleanly(tmp_path: Path) -> None:
    with pytest.raises(Exception):  # noqa: B017 - pdfium ridica erori proprii
        pdflayout.read_pages(tmp_path / "lipsa.pdf")
