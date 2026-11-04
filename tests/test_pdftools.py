from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader

from convertor import pdftools


def text_of(path: Path) -> str:
    return " ".join(page.extract_text() or "" for page in PdfReader(path).pages)


def test_merge_keeps_order(tmp_path: Path, pdf_factory) -> None:
    a = pdf_factory("a.pdf", ["A1", "A2"])
    b = pdf_factory("b.pdf", ["B1"])
    dst = tmp_path / "out" / "merged.pdf"
    pdftools.merge([b, a], dst)
    reader = PdfReader(dst)
    assert [p.extract_text().strip() for p in reader.pages] == ["B1", "A1", "A2"]


def test_merge_needs_two_files(tmp_path: Path, pdf_factory) -> None:
    with pytest.raises(pdftools.PdfError):
        pdftools.merge([pdf_factory("a.pdf", ["A"])], tmp_path / "m.pdf")


def test_split_writes_one_file_per_page(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("doc.pdf", ["P1", "P2", "P3"])
    written = pdftools.split(src, tmp_path / "pages")
    assert [p.name for p in written] == [
        "doc - pagina 001.pdf", "doc - pagina 002.pdf", "doc - pagina 003.pdf"]
    assert text_of(written[1]).strip() == "P2"


def test_ranges_every() -> None:
    assert pdftools.ranges_every(600, 200) == [(1, 200), (201, 400), (401, 600)]
    assert pdftools.ranges_every(5, 2) == [(1, 2), (3, 4), (5, 5)]
    assert pdftools.ranges_every(3, 10) == [(1, 3)]
    with pytest.raises(pdftools.PdfError):
        pdftools.ranges_every(10, 0)


def test_ranges_parts() -> None:
    assert pdftools.ranges_parts(600, 3) == [(1, 200), (201, 400), (401, 600)]
    assert pdftools.ranges_parts(10, 3) == [(1, 4), (5, 7), (8, 10)]
    with pytest.raises(pdftools.PdfError, match="doar 2 pagini"):
        pdftools.ranges_parts(2, 3)


def test_parse_ranges() -> None:
    assert pdftools.parse_ranges("1-200, 201-450, 451-", 600) == [(1, 200), (201, 450),
                                                                   (451, 600)]
    assert pdftools.parse_ranges("5; 7 – 9", 10) == [(5, 5), (7, 9)]
    for bad in ("", "1-2OO", "5-3", "0-4"):
        with pytest.raises(pdftools.PdfError):
            pdftools.parse_ranges(bad, 600)
    with pytest.raises(pdftools.PdfError, match="600 pagini"):
        pdftools.parse_ranges("1-700", 600)


def test_split_by_ranges(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("doc.pdf", [f"P{i}" for i in range(1, 8)])
    written = pdftools.split(src, tmp_path / "parts", pdftools.ranges_every(7, 3))
    assert [p.name for p in written] == [
        "doc - paginile 001-003.pdf", "doc - paginile 004-006.pdf", "doc - pagina 007.pdf"]
    assert [len(PdfReader(p).pages) for p in written] == [3, 3, 1]
    assert text_of(written[1]).split() == ["P4", "P5", "P6"]


def test_protect_requires_the_password(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("doc.pdf", ["secret"])
    dst = tmp_path / "locked.pdf"
    pdftools.protect(src, dst, "parola-ș")
    reader = PdfReader(dst)
    assert reader.is_encrypted
    assert reader.decrypt("gresit") == 0
    assert reader.decrypt("parola-ș") != 0
    assert "secret" in reader.pages[0].extract_text()
    with pytest.raises(pdftools.PdfError, match="parolă"):
        pdftools.page_count(dst)


def test_protect_rejects_empty_password(tmp_path: Path, pdf_factory) -> None:
    with pytest.raises(pdftools.PdfError):
        pdftools.protect(pdf_factory("doc.pdf", ["x"]), tmp_path / "o.pdf", "")


def test_watermark_adds_text_on_every_page(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("doc.pdf", ["Pagina unu", "Pagina doi"])
    dst = tmp_path / "wm.pdf"
    pdftools.watermark(src, dst, "CONFIDENȚIAL", 0.2)
    reader = PdfReader(dst)
    assert len(reader.pages) == 2
    for page in reader.pages:
        text = page.extract_text()
        assert "CONFIDEN" in text
        assert "Pagina" in text


def scanned_pdf(path: Path, text: str) -> Path:
    img = Image.new("RGB", (1700, 2200), "white")
    font = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 72)
    ImageDraw.Draw(img).text((150, 300), text, fill="black", font=font)
    img.save(path, resolution=200)
    return path


def test_is_scanned(tmp_path: Path, pdf_factory) -> None:
    assert not pdftools.is_scanned(pdf_factory("t.pdf", ["Un text suficient de lung pentru test"]))
    assert pdftools.is_scanned(scanned_pdf(tmp_path / "s.pdf", "Hello"))


@pytest.mark.ocr
@pytest.mark.skipif(pdftools.find_tesseract() is None, reason="Tesseract nu e instalat")
def test_ocr_makes_scanned_pdf_searchable(tmp_path: Path) -> None:
    src = scanned_pdf(tmp_path / "scan.pdf", "Invoice number 4512\nTotal amount due 990")
    dst = tmp_path / "scan (OCR).pdf"
    pages = []
    pdftools.ocr(src, dst, progress=lambda i, n: pages.append((i, n)))
    assert pages == [(1, 1)]
    assert "4512" in text_of(dst)
    assert not pdftools.is_scanned(dst)
