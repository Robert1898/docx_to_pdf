"""PowerPoint -> PDF. PowerPoint are o singura instanta pe calculator."""

from __future__ import annotations

import logging
import threading
from pathlib import Path

from ..core import ExportOptions, Job
from .base import MSO_AUTOMATION_SECURITY_FORCE_DISABLE, OfficeApp, describe_error

log = logging.getLogger(__package__)  # un singur jurnal pentru tot pachetul

PP_FIXED_FORMAT_PDF = 2
PP_SAVE_AS_PDF = 32
PP_ALERTS_NONE = 1
POWERPOINT_LOCK = threading.Lock()


class PowerPointApp(OfficeApp):
    progid = "PowerPoint.Application"
    image = "POWERPNT.EXE"
    name = "Microsoft PowerPoint"
    # Application.HWND e marcat „restricted” si da „Member not found” cat timp PowerPoint nu
    # are fereastra vizibila, deci pid-ul vine din tasklist (_window_handle ramane None).

    def configure(self) -> None:
        # PowerPoint nu poate fi ascuns; prezentarile se deschid insa fara fereastra
        self.app.DisplayAlerts = PP_ALERTS_NONE
        self.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

    def convert(self, job: Job, options: ExportOptions, src: Path) -> None:
        pres = self.app.Presentations.Open(
            FileName=str(src), ReadOnly=True, Untitled=False, WithWindow=False
        )
        try:
            try:
                # Pozitional: cu argumente numite, pywin32 trimite PrintRange ca „lipsa”,
                # iar PowerPoint il refuza. None ajunge ca VT_EMPTY si e acceptat.
                pres.ExportAsFixedFormat(
                    str(job.dst),
                    PP_FIXED_FORMAT_PDF,
                    2 if options.for_screen else 1,  # Intent: ecran / tipar
                    0,  # FrameSlides
                    1,  # HandoutOrder
                    1,  # OutputType: slide-uri
                    0,  # PrintHiddenSlides
                    None,  # PrintRange
                    1,  # RangeType: toate
                    "",  # SlideShowName
                    True,  # IncludeDocProperties
                    True,  # KeepIRMSettings
                    True,  # DocStructureTags
                    True,  # BitmapMissingFonts
                    options.pdfa,  # UseISO19005_1
                )
            except Exception as exc:
                # ultima solutie: SaveAs merge mereu, dar fara optiuni
                log.info("ExportAsFixedFormat a esuat (%s), incerc SaveAs", describe_error(exc))
                pres.SaveAs(str(job.dst), PP_SAVE_AS_PDF)
        finally:
            pres.Close()
