from __future__ import annotations

import datetime as dt
import zipfile
from pathlib import Path

import pytest
from reportlab.pdfgen import canvas

from convertor import pdftable
from convertor.pdflayout import read_pages


@pytest.mark.parametrize(("text", "value", "decimals"), [
    ("1,234.56", 1234.56, 2),
    ("1.234,56", 1234.56, 2),
    ("(1,000.00)", -1000.0, 2),
    ("250.00-", -250.0, 2),
    ("42", 42, 0),
    ("-7", -7, 0),
    ("0123", "0123", 0),  # cod cu zero in fata ramane text
    ("1234567890123456", "1234567890123456", 0),  # prea multe cifre pentru Excel
    ("15-SEP-2026", dt.date(2026, 9, 15), 0),
    ("15.09.2026", dt.date(2026, 9, 15), 0),
    ("2026-09-15", dt.date(2026, 9, 15), 0),
    ("31.02.2026", "31.02.2026", 0),  # data imposibila
    ("Cheltuieli", "Cheltuieli", 0),
])
def test_parse_value(text: str, value: object, decimals: int) -> None:
    assert pdftable.parse_value(text) == (value, decimals)


def table_pdf(path: Path, rows: int = 12) -> Path:
    """Raport tipic: data, descriere, suma aliniata la dreapta."""
    c = canvas.Canvas(str(path))
    c.setFont("Helvetica-Bold", 9)
    y = 760
    for x, title in ((60, "Data"), (150, "Descriere")):
        c.drawString(x, y, title)
    c.drawRightString(500, y, "Suma")
    c.setFont("Helvetica", 9)
    for i in range(rows):
        y -= 16
        c.drawString(60, y, f"{i + 1:02d}.09.2026")
        c.drawString(150, y, f"Factura client {i + 1}")
        c.drawRightString(500, y, f"{(i + 1) * 1234.5:,.2f}")
    c.save()
    return path


def test_rows_from_pages_finds_columns_and_numbers(tmp_path: Path) -> None:
    rows = pdftable.rows_from_pages(read_pages(table_pdf(tmp_path / "t.pdf"), shapes=False))
    assert len(rows) == 13
    assert [c.value for c in rows[0] if c] == ["Data", "Descriere", "Suma"]
    assert rows[0][0].bold
    first = rows[1]
    assert first[0].value == dt.date(2026, 9, 1)
    assert first[1].value == "Factura client 1"
    assert first[2].value == pytest.approx(1234.5)
    assert first[2].decimals == 2
    # toate sumele in aceeasi coloana, ca numere
    assert all(isinstance(r[2].value, float) for r in rows[1:])


def test_convert_writes_valid_xlsx(tmp_path: Path) -> None:
    dst = tmp_path / "t.xlsx"
    assert pdftable.convert(table_pdf(tmp_path / "t.pdf", rows=3), dst) == 4
    with zipfile.ZipFile(dst) as z:
        names = set(z.namelist())
        assert {"[Content_Types].xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml",
                "xl/styles.xml"} <= names
        sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert "Factura client 2" in sheet
    assert "<v>2469</v>" in sheet or "<v>2469.0</v>" in sheet


def test_column_letters() -> None:
    assert [pdftable._column_letter(i) for i in (0, 25, 26, 27, 701, 702)] == [
        "A", "Z", "AA", "AB", "ZZ", "AAA"]
