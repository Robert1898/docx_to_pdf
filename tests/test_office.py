"""Teste cu Microsoft Office real. Ruleaza doar unde e instalat: pytest -m office"""

from __future__ import annotations

import os
import time
import winreg
from pathlib import Path

import pytest
from pypdf import PdfReader

from convertor.core import ExportOptions, Kind, plan_jobs
from convertor.engine import Engine, Status
from convertor.office import PdfWarningGuard, ResiliencyCleaner, WordApp

pytestmark = [
    pytest.mark.office,
    pytest.mark.skipif(not WordApp.available(), reason="Microsoft Word nu e instalat"),
]

WORD_OPTIONS = r"Software\Microsoft\Office\16.0\Word\Options"


@pytest.fixture(scope="module")
def samples(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    import pythoncom
    import win32com.client

    out = tmp_path_factory.mktemp("samples")
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Add()
        doc.Content.Text = "Capitol unu\nText cu diacritice: ăîșțâ.\nCapitol doi\nFinal."
        doc.Paragraphs(1).Style = -2  # Titlu 1
        doc.Paragraphs(3).Style = -2
        doc.SaveAs2(str(out / "raport.docx"), FileFormat=16)
        doc.Password = "1234"
        doc.SaveAs2(str(out / "secret.docx"), FileFormat=16)
        doc.Close(False)
    finally:
        word.Quit()

    files = {"docx": out / "raport.docx", "secret": out / "secret.docx"}
    from convertor.office import ExcelApp, PowerPointApp

    if ExcelApp.available():
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Add()
            wb.Worksheets(1).Cells(1, 1).Value = "Buget 2026"
            wb.SaveAs(str(out / "buget.xlsx"), FileFormat=51)
            wb.Close(False)
            files["xlsx"] = out / "buget.xlsx"
        finally:
            excel.Quit()
    if PowerPointApp.available():
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            pres = ppt.Presentations.Add(WithWindow=False)
            slide = pres.Slides.Add(1, 1)
            slide.Shapes(1).TextFrame.TextRange.Text = "Prezentare test"
            pres.SaveAs(str(out / "prezentare.pptx"), 24)
            pres.Close()
            files["pptx"] = out / "prezentare.pptx"
        finally:
            ppt.Quit()
    return files


def registry_value() -> int | None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, WORD_OPTIONS) as key:
            return winreg.QueryValueEx(key, PdfWarningGuard.VALUE)[0]
    except FileNotFoundError:
        return None


def word_marks_mentioning(folder: Path) -> list[str]:
    """Toate marcajele Word (StartupItems, DisabledItems, DocumentRecovery) din folder."""
    wanted = {os.path.normcase(str(folder))}
    found: list[str] = []

    def scan(path: str) -> None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                for name, data in ResiliencyCleaner._values(key):
                    if ResiliencyCleaner.mentions(data, wanted):
                        found.append(rf"{path}\{name}")
                for sub in ResiliencyCleaner._subkeys(key):
                    scan(rf"{path}\{sub}")
        except FileNotFoundError:
            pass

    scan(ResiliencyCleaner.root())
    return found


def run(files: list[Path], out: Path, **options) -> list:
    return Engine(ExportOptions(**options)).run(plan_jobs(files, out))


def test_word_to_pdf_with_bookmarks_and_pdfa(samples, tmp_path: Path) -> None:
    before = registry_value()
    [result] = run([samples["docx"]], tmp_path, bookmarks=True, pdfa=True)
    assert result.status is Status.DONE, result.message
    reader = PdfReader(result.job.dst)
    assert [item.title for item in reader.outline] == ["Capitol unu", "Capitol doi"]
    assert "ăîșțâ" in reader.pages[0].extract_text()
    assert reader.xmp_metadata is not None and reader.xmp_metadata.pdfaid_part
    assert registry_value() == before


def test_pdf_to_docx_roundtrip(samples, tmp_path: Path) -> None:
    [pdf] = run([samples["docx"]], tmp_path / "pdf")
    [docx] = run([pdf.job.dst], tmp_path / "docx")
    assert docx.status is Status.DONE, docx.message
    assert docx.job.kind is Kind.PDF and docx.job.dst.suffix == ".docx"
    assert docx.job.dst.stat().st_size > 0


def test_excel_and_powerpoint(samples, tmp_path: Path) -> None:
    files = [samples[k] for k in ("xlsx", "pptx") if k in samples]
    if not files:
        pytest.skip("Excel/PowerPoint nu sunt instalate")
    for result in run(files, tmp_path):
        assert result.status is Status.DONE, result.message
        assert PdfReader(result.job.dst).pages


def test_excel_prints_only_real_content(tmp_path: Path) -> None:
    """Zona „folosita” umflata (borduri pe mii de randuri) nu mai da pagini goale."""
    import pythoncom
    import win32com.client

    from convertor.office import ExcelApp

    if not ExcelApp.available():
        pytest.skip("Excel nu e instalat")
    src = tmp_path / "umflat.xlsx"
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Add()
        while wb.Worksheets.Count < 3:
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        data, empty, custom = (wb.Worksheets(i) for i in (1, 2, 3))
        data.Range("A1:D1").Value = ("Produs", "Cantitate", "Pret", "Total")
        data.Range("A2").Value = "Produs 1"
        data.Range("A1:Z3000").Borders.LineStyle = 1
        data.Range("E2:E3000").Formula = '=IF(A2="nimic","x","")'
        empty.Range("A1:J500").Borders.LineStyle = 1
        custom.Range("A1").Value = "Zona aleasa"
        custom.Range("A200").Value = "In afara zonei"
        custom.PageSetup.PrintArea = "$A$1:$B$3"
        wb.SaveAs(str(src), FileFormat=51)
        wb.Close(False)
    finally:
        excel.Quit()

    [result] = run([src], tmp_path / "out")
    assert result.status is Status.DONE, result.message
    pages = [page.extract_text() for page in PdfReader(result.job.dst).pages]
    assert len(pages) == 2  # foaia cu date + zona aleasa; foaia goala e sarita
    assert "Produs" in pages[0]
    assert "Zona aleasa" in pages[1] and "In afara zonei" not in pages[1]


def test_excel_far_away_cell_and_merged_title(tmp_path: Path) -> None:
    """O celula uitata la randul 20000 nu mai aduce mii de pagini goale si nu se pierde;
    un titlu unit peste mai multe coloane nu e taiat."""
    import pythoncom
    import win32com.client

    from convertor.office import ExcelApp

    if not ExcelApp.available():
        pytest.skip("Excel nu e instalat")
    src = tmp_path / "ratacit.xlsx"
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Range("A1:H1").Merge()
        ws.Range("A1").Value = "Raport lunar foarte important"
        ws.Range("A1").HorizontalAlignment = -4108  # centrat
        ws.Range("A2:C2").Value = ("Cont", "Nume", "Cod")
        for r in range(3, 13):
            ws.Range(f"A{r}:C{r}").Value = (r, f"Persoana {r}", f"COD_{r}")
        ws.Range("C20000").Value = "COD_RATACIT"
        wb.SaveAs(str(src), FileFormat=51)
        wb.Close(False)
    finally:
        excel.Quit()

    [result] = run([src], tmp_path / "out")
    assert result.status is Status.DONE, result.message
    pages = [page.extract_text() for page in PdfReader(result.job.dst).pages]
    assert len(pages) == 2  # tabelul + celula ratacita pe pagina ei
    assert "Raport lunar foarte important" in pages[0]
    assert "COD_RATACIT" in pages[1]


def test_excel_multiple_sheets_and_wide_tables(tmp_path: Path) -> None:
    """Toate foile intr-un PDF, in ordine; tabelul foarte lat iese in peisaj la 60%, cu
    antetul si coloana de identificare repetate pe fiecare pagina."""
    import pythoncom
    import win32com.client

    from convertor.office import ExcelApp

    if not ExcelApp.available():
        pytest.skip("Excel nu e instalat")
    src = tmp_path / "foi.xlsx"
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Add()
        while wb.Worksheets.Count < 3:
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        narrow, wide, slightly = (wb.Worksheets(i) for i in (1, 2, 3))
        narrow.Name, wide.Name, slightly.Name = "Ingust", "Lat", "PutinLat"
        narrow.Range("A1:B2").Value = (("Foaia", "intai"), ("ok", 1))
        wide.Range(wide.Cells(1, 1), wide.Cells(1, 40)).Value = tuple(f"H{c}" for c in range(1, 41))
        ids = tuple((f"ID{r}",) for r in range(2, 122))
        wide.Range(wide.Cells(2, 1), wide.Cells(121, 1)).Value = ids
        wide.Range(wide.Cells(2, 2), wide.Cells(121, 40)).Value = 7
        slightly.Range(slightly.Cells(1, 1), slightly.Cells(5, 12)).Value = "x"
        wb.SaveAs(str(src), FileFormat=51)
        wb.Close(False)
    finally:
        excel.Quit()

    [result] = run([src], tmp_path / "out")
    assert result.status is Status.DONE, result.message
    pages = PdfReader(result.job.dst).pages
    texts = [p.extract_text() for p in pages]
    landscape = [float(p.mediabox.width) > float(p.mediabox.height) for p in pages]

    assert "Foaia" in texts[0] and not landscape[0]  # foaia ingusta: portret, prima
    assert landscape[-1] and "x" in texts[-1]  # foaia putin lata: o pagina in peisaj
    wide_pages = texts[1:-1]
    assert len(wide_pages) >= 2 and all(landscape[1:-1])
    # pagina din dreapta a primei benzi: are antetul de la capat si ID-urile din coloana A
    right = next(t for t in wide_pages if "H40" in t)
    assert "ID2" in right
    # o pagina din banda de jos: antetul se repeta
    lower = next(t for t in wide_pages if "ID121" in t)
    assert "H1" in lower


def test_pdf_report_round_trip_keeps_rows_and_columns(tmp_path: Path) -> None:
    """Raport pe coloane -> DOCX (automat: aspect fidel) -> PDF prin Word: aceleasi pagini,
    aceleasi randuri, fiecare valoare in coloana si pe randul ei."""
    from convertor import pdflayout
    from tests.test_pdflayout import report_pdf

    src = report_pdf(tmp_path / "raport.pdf", pages=3)
    [docx] = run([src], tmp_path / "docx")
    assert docx.status is Status.DONE, docx.message
    assert "aspect fidel" in docx.message
    [back] = run([docx.job.dst], tmp_path / "pdf")
    original, again = pdflayout.read_pages(src), pdflayout.read_pages(back.job.dst)
    assert len(again) == len(original) == 3
    for page_a, page_b in zip(original, again, strict=True):
        assert len(page_a.lines) == len(page_b.lines)
        for line_a, line_b in zip(page_a.lines, page_b.lines, strict=True):
            assert [s.text for s in line_a.segments] == [s.text for s in line_b.segments]
            assert abs(line_a.baseline - line_b.baseline) < 1
            for seg_a, seg_b in zip(line_a.segments, line_b.segments, strict=True):
                edge = (seg_a.x1, seg_b.x1) if seg_a.numeric else (seg_a.x0, seg_b.x0)
                assert abs(edge[0] - edge[1]) < 1.5


def test_pdf_prose_still_goes_through_word(samples, tmp_path: Path) -> None:
    [pdf] = run([samples["docx"]], tmp_path / "pdf")
    [docx] = run([pdf.job.dst], tmp_path / "docx")
    assert docx.status is Status.DONE, docx.message
    assert "prin Word" in docx.message


def test_password_protected_document_fails_fast(samples, tmp_path: Path) -> None:
    start = time.monotonic()
    [result] = run([samples["secret"]], tmp_path, timeout=60)
    assert result.status is Status.FAILED
    assert "parolă" in result.message
    assert time.monotonic() - start < 45


def test_timeout_kills_word_then_recovers_and_cleans_up(
    samples, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from convertor import office

    before = registry_value()
    jobs = plan_jobs([samples["docx"]], tmp_path / "a") + plan_jobs([samples["docx"]],
                                                                    tmp_path / "b")
    real_convert = office.OfficeSession.convert
    calls: list[object] = []

    def convert(self, job, options, src=None):
        # primul fisier primeste o limita imposibila: Word e oprit in timp ce il deschide
        calls.append(job)
        if len(calls) == 1:
            options = ExportOptions(timeout=0.05)  # type: ignore[arg-type]
        return real_convert(self, job, options, src)

    monkeypatch.setattr(office.OfficeSession, "convert", convert)
    results = Engine(ExportOptions(timeout=60)).run(jobs)

    assert results[0].status is Status.FAILED
    assert "limita" in results[0].message
    assert results[1].status is Status.DONE, results[1].message
    assert registry_value() == before
    assert word_marks_mentioning(samples["docx"].parent) == []
