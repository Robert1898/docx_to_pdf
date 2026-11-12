from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest
from PIL import Image
from pypdf import PdfReader

from convertor import pdftools
from convertor.core import Kind, kind_of


def image(path: Path, size: tuple[int, int], color: str = "red") -> Path:
    Image.new("RGB", size, color).save(path)
    return path


def test_images_are_a_kind() -> None:
    assert kind_of(Path("scan.JPG")) is Kind.IMAGE
    assert pdftools.is_image(Path("a.webp"))
    assert not pdftools.is_image(Path("a.pdf"))


def test_images_to_pdf_a4_orientation(tmp_path: Path) -> None:
    portrait = image(tmp_path / "p.png", (600, 900))
    landscape = image(tmp_path / "l.jpg", (900, 600), "blue")
    dst = tmp_path / "out.pdf"
    assert pdftools.images_to_pdf([portrait, landscape], dst) == 2
    pages = PdfReader(dst).pages
    sizes = [(round(float(p.mediabox.width)), round(float(p.mediabox.height))) for p in pages]
    assert sizes == [(595, 842), (842, 595)]


def test_images_to_pdf_image_size_and_multiframe_tiff(tmp_path: Path) -> None:
    frames = [Image.new("RGB", (200, 100), c) for c in ("red", "green", "blue")]
    tiff = tmp_path / "scan.tif"
    frames[0].save(tiff, save_all=True, append_images=frames[1:], dpi=(100, 100))
    buffer = BytesIO()
    assert pdftools.images_to_pdf([tiff], buffer, page="image") == 3
    page = PdfReader(BytesIO(buffer.getvalue())).pages[0]
    assert float(page.mediabox.width) > float(page.mediabox.height)


def test_pdf_to_images(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("doc.pdf", ["unu", "doi"])
    out = tmp_path / "imgs"
    seen = []
    paths = pdftools.pdf_to_images(src, out, fmt="jpg", dpi=50,
                                   progress=lambda page, total: seen.append((page, total)))
    assert [p.suffix for p in paths] == [".jpg", ".jpg"]
    assert all(p.exists() for p in paths)
    assert seen[-1] == (2, 2)
    with Image.open(paths[0]) as img:
        assert img.width == pytest.approx(595 * 50 / 72, abs=2)  # A4 din reportlab


def test_merge_accepts_images(tmp_path: Path, pdf_factory) -> None:
    pdf = pdf_factory("a.pdf", ["text"])
    dst = tmp_path / "merged.pdf"
    pdftools.merge([pdf, image(tmp_path / "b.png", (300, 400))], dst)
    assert len(PdfReader(dst).pages) == 2


def test_compress_shrinks_big_images(tmp_path: Path) -> None:
    noisy = Image.effect_noise((2400, 3200), 60).convert("RGB")
    big = tmp_path / "big.png"
    noisy.save(big)
    src = tmp_path / "src.pdf"
    pdftools.images_to_pdf([big], src, page="image")
    dst = tmp_path / "small.pdf"
    before, after = pdftools.compress(src, dst, "strong")
    assert before == src.stat().st_size
    assert after == dst.stat().st_size
    assert after < before
    assert len(PdfReader(dst).pages) == 1


def test_compress_keeps_original_when_not_smaller(tmp_path: Path, pdf_factory) -> None:
    src = pdf_factory("text.pdf", ["doar text"])
    dst = tmp_path / "out.pdf"
    before, after = pdftools.compress(src, dst, "light")
    assert after <= before
    assert dst.read_bytes() == src.read_bytes() or after < before
