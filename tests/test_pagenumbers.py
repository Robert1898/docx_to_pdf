from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader
from reportlab.pdfgen import canvas

from convertor import pagenumbers, pdftools


def numbered_pdf(path: Path, pages: int, footer: str, align: str = "center",
                 body: str = "Text obisnuit") -> Path:
    """PDF cu numar de pagina in subsol; `footer` foloseste {n} si {total}."""
    c = canvas.Canvas(str(path), pagesize=(595, 842))
    for n in range(1, pages + 1):
        c.setFont("Helvetica", 11)
        c.drawString(72, 770, f"{body} {n}/{pages} in mijlocul frazei, anul 2026.")
        text = footer.format(n=n, total=pages)
        c.setFont("Helvetica", 9)
        if align == "center":
            c.drawCentredString(297.5, 30, text)
        elif align == "right":
            c.drawRightString(523, 30, text)
        else:
            c.drawString(72, 30, text)
        c.showPage()
    c.save()
    return path


def page_texts(path: Path) -> list[str]:
    return [page.extract_text() for page in PdfReader(path).pages]


def test_detects_footer_numbers_but_not_body_numbers(tmp_path: Path) -> None:
    src = numbered_pdf(tmp_path / "a.pdf", 3, "{n}/{total}", align="right")
    found = pagenumbers.find_page_numbers(src)
    assert [len(p) for p in found] == [1, 1, 1]
    number = found[1][0]
    assert number.template == "{page}/{total}"
    assert number.align == "right"
    assert round(number.font_size) == 9
    assert number.background == (1.0, 1.0, 1.0)


def test_detects_word_style_text(tmp_path: Path) -> None:
    src = numbered_pdf(tmp_path / "b.pdf", 2, "Pagina {n} din {total}")
    [[first], [second]] = pagenumbers.find_page_numbers(src)
    assert first.template == second.template == "Pagina {page} din {total}"
    assert first.align == "center"


def test_ignores_numbers_that_do_not_match_the_page(tmp_path: Path) -> None:
    # „1/5” pe fiecare pagina nu e numar de pagina (pagina 2 ar trebui sa fie 2/5)
    src = numbered_pdf(tmp_path / "c.pdf", 2, "Anexa 1/5")
    found = pagenumbers.find_page_numbers(src)
    assert found == [[], []]


def test_merge_renumbers_continuously(tmp_path: Path) -> None:
    a = numbered_pdf(tmp_path / "a.pdf", 8, "{n} / {total}", align="right")
    b = numbered_pdf(tmp_path / "b.pdf", 2, "Pagina {n} din {total}")
    dst = tmp_path / "unit.pdf"
    assert pdftools.merge([a, b], dst, renumber=True) == 10
    texts = page_texts(dst)
    assert len(texts) == 10
    assert "8 / 10" in texts[7]
    assert "Pagina 9 din 10" in texts[8]
    assert "Pagina 10 din 10" in texts[9]
    # textul din corp ramane neatins
    assert "1/2 in mijlocul frazei" in texts[8]


def test_merge_without_renumber_keeps_pages_untouched(tmp_path: Path) -> None:
    a = numbered_pdf(tmp_path / "a.pdf", 2, "{n}/{total}")
    b = numbered_pdf(tmp_path / "b.pdf", 1, "{n}/{total}")
    dst = tmp_path / "unit.pdf"
    assert pdftools.merge([a, b], dst) == 0
    assert "3/3" not in "".join(page_texts(dst))


def test_new_number_covers_old_one(tmp_path: Path) -> None:
    import pypdfium2 as pdfium

    a = numbered_pdf(tmp_path / "a.pdf", 1, "{n}/{total}")
    b = numbered_pdf(tmp_path / "b.pdf", 1, "{n}/{total}")
    dst = tmp_path / "unit.pdf"
    pdftools.merge([a, b], dst, renumber=True)
    # ce se vede in subsolul paginii 2: doar „2/2”, nu si „1/1” dedesubt
    def dark_pixels_in_footer(path: Path, index: int) -> int:
        image = pdfium.PdfDocument(str(path))[index].render(scale=2).to_pil().convert("L")
        footer = image.crop((500, image.height - 80, 700, image.height - 40))
        return sum(footer.histogram()[:128])

    dark, reference = dark_pixels_in_footer(dst, 1), dark_pixels_in_footer(a, 0)
    assert 0.6 * reference < dark < 1.6 * reference  # un singur numar, nu doua suprapuse
