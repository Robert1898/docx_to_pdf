"""Word: documente -> PDF si PDF -> DOCX."""

from __future__ import annotations

from pathlib import Path

from ..core import ExportOptions, Job, Kind
from .base import MSO_AUTOMATION_SECURITY_FORCE_DISABLE, NO_PASSWORD, OfficeApp
from .registry import PdfWarningGuard, ResiliencyCleaner

WD_FORMAT_DOCX = 16
WD_EXPORT_PDF = 17


class WordApp(OfficeApp):
    progid = "Word.Application"
    image = "WINWORD.EXE"
    name = "Microsoft Word"
    # optiunea din registry e „imprumutata” o singura data per instanta, chiar daca Word e
    # repornit dupa un timeout; se elibereaza la quit()
    _guarded = False

    def _guard(self) -> None:
        if not self._guarded:
            PdfWarningGuard.suppress()
            self._guarded = True

    def _unguard(self) -> None:
        if self._guarded:
            self._guarded = False
            PdfWarningGuard.restore()

    def start(self) -> None:
        self._guard()
        ResiliencyCleaner.clean()  # inainte ca Word sa-si incarce lista la pornire
        try:
            super().start()
        except Exception:
            self._unguard()
            raise
        # dupa un timeout, Word tocmai a marcat fisierul la care l-am oprit
        ResiliencyCleaner.clean()

    def configure(self) -> None:
        self.app.Visible = False
        self.app.DisplayAlerts = 0
        self.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

    def quit(self) -> None:
        try:
            super().quit()
        finally:
            self._unguard()
            ResiliencyCleaner.clean()

    def convert(self, job: Job, options: ExportOptions, src: Path) -> None:
        doc = self.app.Documents.Open(
            FileName=str(src),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            PasswordDocument=NO_PASSWORD,
            PasswordTemplate=NO_PASSWORD,
            Revert=False,
            Visible=False,
            NoEncodingDialog=True,
        )
        try:
            if job.kind is Kind.PDF:
                doc.SaveAs2(
                    FileName=str(job.dst), FileFormat=WD_FORMAT_DOCX, AddToRecentFiles=False
                )
            else:
                doc.ExportAsFixedFormat(
                    OutputFileName=str(job.dst),
                    ExportFormat=WD_EXPORT_PDF,
                    OpenAfterExport=False,
                    OptimizeFor=1 if options.for_screen else 0,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=1 if options.bookmarks else 0,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=options.pdfa,
                )
        finally:
            doc.Close(SaveChanges=0)
