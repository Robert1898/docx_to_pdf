from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image
from pypdf import PdfReader

from convertor import pageedit
from convertor.pdftools import PdfError


def test_reorder_rotate_delete(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("doc.pdf", ["unu", "doi", "trei"])
    edits = pageedit.initial(src)
    assert [e.source for e in edits] == [0, 1, 2]
    edits = [edits[2], edits[0]]  # pagina 2 stearsa, 3 mutata in fata
    edits[1].rotation = 90
    dst = tmp_path / "out.pdf"
    pageedit.save(src, edits, dst)
    pages = PdfReader(dst).pages
    assert ["trei" in pages[0].extract_text(), "unu" in pages[1].extract_text()] == [True, True]
    assert pages[1].rotation == 90
    assert len(pages) == 2


def test_stamp_adds_an_image(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("doc.pdf", ["semneaza aici"])
    stamp = tmp_path / "semnatura.png"
    Image.new("RGBA", (200, 80), (0, 0, 255, 200)).save(stamp)
    edits = pageedit.initial(src)
    edits[0].stamps.append(pageedit.Stamp(stamp, 300, 100, 500, 180))
    dst = tmp_path / "out.pdf"
    pageedit.save(src, edits, dst)
    page = PdfReader(dst).pages[0]
    assert len(page.images) == 1
    assert "semneaza aici" in page.extract_text()  # textul original ramane


def test_empty_document_is_refused(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("doc.pdf", ["x"])
    with pytest.raises(PdfError):
        pageedit.save(src, [], tmp_path / "out.pdf")


def test_unreadable_stamp_image(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("doc.pdf", ["x"])
    bad = tmp_path / "bad.png"
    bad.write_bytes(b"not an image")
    edits = pageedit.initial(src)
    edits[0].stamps.append(pageedit.Stamp(bad, 0, 0, 10, 10))
    with pytest.raises(PdfError, match=r"bad\.png"):
        pageedit.save(src, edits, tmp_path / "out.pdf")


def test_device_to_page_maps_corners(pdf_factory) -> None:
    import pypdfium2 as pdfium

    src = pdf_factory("doc.pdf", ["x"])
    document = pdfium.PdfDocument(str(src))
    page = document[0]
    try:
        w, h = (round(v) for v in page.get_size())
        x, y = pageedit.device_to_page(page, 0, (w, h), (0, 0))
        assert (round(x), round(y)) == (0, h)  # sus-stanga in imagine = sus-stanga PDF
        x, y = pageedit.device_to_page(page, 0, (w, h), (w, h))
        assert (round(x), round(y)) == (w, 0)
        x, y = pageedit.device_to_page(page, 90, (h, w), (0, 0))
        assert (round(x), round(y)) == (0, 0)  # rotita 90: sus-stanga = jos-stanga PDF
        assert pageedit.page_rotation(page) == 0
    finally:
        page.close()
        document.close()
