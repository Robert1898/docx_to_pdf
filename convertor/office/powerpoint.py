"""PowerPoint -> PDF. PowerPoint are o singura instanta pe calculator."""

from __future__ import annotations

import logging
import threading
from pathlib import Path

from ..core import ExportOptions, Job
from .base import MSO_AUTOMATION_SECURITY_FORCE_DISABLE, OfficeApp, describe_error
from .process import _visible_windows

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

    # PowerPoint are un singur proces, comun tuturor firelor: prezentarile deschise acum de
    # aplicatie se numara la nivel de clasa, sub POWERPOINT_LOCK.
    open_presentations = 0

    def __init__(self) -> None:
        super().__init__()
        self.foreign_seen = False  # ultima numaratoare (COM) a gasit prezentari straine

    def configure(self) -> None:
        # PowerPoint nu poate fi ascuns; prezentarile se deschid insa fara fereastra
        self.app.DisplayAlerts = PP_ALERTS_NONE
        self.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        if self.owned:  # proces nou: nimic deschis inca, nici de noi
            PowerPointApp.open_presentations = 0
        self.foreign_seen = False

    def _foreign_count(self) -> int:
        """Prezentarile din instanta care nu sunt ale noastre (COM: doar din firul nostru)."""
        try:
            count = int(self.app.Presentations.Count)
        except Exception as exc:  # intentionat larg: aplicatia poate fi deja moarta
            log.debug("Presentations.Count %s: %s", self.name, describe_error(exc))
            return 0
        return count - PowerPointApp.open_presentations

    def _foreign_window(self) -> bool:
        # Prezentarile noastre se deschid fara fereastra; una deschisa de utilizator (de ex.
        # din Explorer, care o trimite in procesul nostru) are fereastra vizibila PPTFrameClass.
        return bool(self.pid) and bool(_visible_windows(self.pid, {"PPTFrameClass"}))

    def has_foreign_documents(self) -> bool:
        return self.foreign_seen or self._foreign_window()

    def prepare_quit(self) -> bool:
        """Refuza inchiderea daca instanta are prezentari straine: Quit sau oprirea fortata
        le-ar inchide. PowerPoint ramane deschis pentru utilizator, care il inchide singur."""
        foreign = self._foreign_count()  # numaratoare proaspata, nu foreign_seen
        if foreign <= 0 and not self._foreign_window():
            return True
        log.warning("%s (pid %s) are prezentari deschise de altcineva (%d dupa numaratoare); nu "
                    "il inchid si nu il opresc", self.name, self.pid, max(foreign, 0))
        return False

    def quit(self) -> None:
        with POWERPOINT_LOCK:  # nu in timp ce alt fir converteste in acelasi proces
            super().quit()

    def convert(self, job: Job, options: ExportOptions, src: Path) -> None:
        self.foreign_seen = self._foreign_count() > 0
        pres = self.app.Presentations.Open(
            FileName=str(src), ReadOnly=True, Untitled=False, WithWindow=False
        )
        PowerPointApp.open_presentations += 1
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
            except Exception as exc:  # intentionat larg: orice esec al exportului incearca SaveAs
                if self.abandoned:  # apel anulat la timeout: PowerPoint e inca blocat
                    raise
                # ultima solutie: SaveAs merge mereu, dar fara optiuni
                log.info("ExportAsFixedFormat a esuat (%s), incerc SaveAs", describe_error(exc))
                pres.SaveAs(str(job.dst), PP_SAVE_AS_PDF)
        finally:
            if not self.abandoned:  # Close ar astepta si el PowerPoint-ul blocat
                pres.Close()
                PowerPointApp.open_presentations -= 1
