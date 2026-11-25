"""Excel -> PDF, cu zona de tiparire limitata la continutul real al fiecarei foi."""

from __future__ import annotations

import contextlib
import logging
from pathlib import Path
from typing import Any

from .. import excelarea
from ..core import ExportOptions, Job
from .base import (
    MSO_AUTOMATION_SECURITY_FORCE_DISABLE,
    NO_PASSWORD,
    OfficeApp,
    describe_error,
)
from .process import _visible_windows

log = logging.getLogger(__package__)  # un singur jurnal pentru tot pachetul

XL_TYPE_PDF = 0
XL_VALUES = -4163
XL_PART = 2
XL_BY_ROWS, XL_BY_COLUMNS = 1, 2
XL_NEXT, XL_PREVIOUS = 1, 2
XL_SHEET_VISIBLE, XL_SHEET_HIDDEN = -1, 0
MAX_WHITESPACE_SKIPS = 2000


def _find_content(ws: Any, order: int, direction: int) -> Any:
    """Prima/ultima celula cu o valoare vizibila (nu doar spatii) in ordinea data.

    Cautarea dupa valori (nu formule) ignora formulele care intorc "" si celulele doar
    formatate, adica exact ce umfla zona „folosita”."""
    cells = ws.Cells
    start = cells(1, 1) if direction == XL_PREVIOUS else cells(ws.Rows.Count, ws.Columns.Count)
    found = cells.Find(What="*", After=start, LookIn=XL_VALUES, LookAt=XL_PART,
                       SearchOrder=order, SearchDirection=direction, MatchCase=False)
    if found is None:
        return None
    first = found.Address
    for _attempt in range(MAX_WHITESPACE_SKIPS):
        if str(found.Text).strip():
            return found
        found = cells.FindPrevious(found) if direction == XL_PREVIOUS else cells.FindNext(found)
        if found is None or found.Address == first:
            return None
    return found


MAX_SCAN_CELLS = 1_000_000  # peste atat, citirea valorilor dintr-o data ar fi prea grea
MAX_MERGE_CHECKS = 5_000
MAX_PRINT_AREA_CHARS = 255


def _find_bounds(ws: Any) -> excelarea.Rect | None:
    """Varianta pentru foi uriase: marginile continutului prin cautarea lui Excel."""
    rows, cols = [], []
    for order, collect in ((XL_BY_ROWS, rows), (XL_BY_COLUMNS, cols)):
        for direction in (XL_NEXT, XL_PREVIOUS):
            cell = _find_content(ws, order, direction)
            if cell is not None:
                collect.append(cell.Row if order == XL_BY_ROWS else cell.Column)
    if not rows or not cols:
        return None
    return min(rows), min(cols), max(rows), max(cols)


def _rect(rng: Any) -> excelarea.Rect:
    return (rng.Row, rng.Column, rng.Row + rng.Rows.Count - 1, rng.Column + rng.Columns.Count - 1)


def _content_rects(ws: Any) -> list[excelarea.Rect]:
    """Dreptunghiurile cu continut real: celule cu valori, celule unite, grafice si imagini.

    Valorile se citesc dintr-o data din zona folosita si se analizeaza in Python, ceea ce e
    determinist (cautarea lui Excel depinde de starea aplicatiei si a dat o data rezultate
    gresite). Nimic nu poate iesi din zona folosita, in afara de forme."""
    used = ws.UsedRange
    used_rect = _rect(used)
    rows, cols = used.Rows.Count, used.Columns.Count
    if rows * cols <= MAX_SCAN_CELLS:
        rects = excelarea.cells_with_content(used.Value2, used_rect[0], used_rect[1])
        if rects and used.MergeCells is not False:  # None = unele celule sunt unite
            for r, c, _r, _c in rects[:MAX_MERGE_CHECKS]:
                cell = ws.Cells(r, c)
                if cell.MergeCells:
                    rects.append(_rect(cell.MergeArea))
    else:
        bounds = _find_bounds(ws)
        rects = [excelarea.clamp(bounds, used_rect)] if bounds else []
    for shape in ws.Shapes:
        # intentionat larg: o forma ciudata (COM) e doar sarita
        with contextlib.suppress(Exception):
            if shape.Visible:
                top_left, bottom_right = shape.TopLeftCell, shape.BottomRightCell
                rects.append((top_left.Row, top_left.Column, bottom_right.Row,
                              bottom_right.Column))
    return rects


XL_PORTRAIT, XL_LANDSCAPE = 1, 2
XL_DOWN_THEN_OVER, XL_OVER_THEN_DOWN = 1, 2
# dimensiuni in puncte (latime, inaltime) pentru formatele uzuale; restul ~ Letter
PAPER_SIZES = {1: (612, 792), 5: (612, 1008), 8: (842, 1191), 9: (595, 842), 11: (420, 595)}
MIN_SCALE = 60  # % sub care textul devine greu de citit pe hartie
TITLE_COLUMN_MAX_SHARE = 0.3  # coloana repetata nu are voie sa manance mai mult din pagina


def _layout_wide(ws: Any, width: float, header: tuple[int, int] | None,
                 label: int | None) -> str:
    """Asezarea tabelelor mai late decat pagina (varianta „lizibil”):

    * peisaj de indata ce continutul nu incape pe portret;
    * daca incape pe latime la cel putin 60%, o singura pagina in latime;
    * altfel 60%, paginile merg intai spre dreapta, apoi in jos, iar antetul si coloana de
      identificare (de ex. CNP, nume) se repeta pe fiecare pagina, ca fiecare rand sa ramana
      de recunoscut.

    Scalarea, orientarea peisaj si randurile/coloanele repetate alese de utilizator in Excel
    nu sunt atinse."""
    setup = ws.PageSetup
    if setup.Zoom is False or setup.Zoom != 100:
        return ""
    paper_w, paper_h = PAPER_SIZES.get(setup.PaperSize, PAPER_SIZES[1])
    margins = setup.LeftMargin + setup.RightMargin
    portrait, landscape = paper_w - margins, paper_h - margins
    if setup.Orientation == XL_PORTRAIT and width <= portrait:
        return ""
    setup.Orientation = XL_LANDSCAPE
    scale = landscape / width
    if scale >= 1:
        return "peisaj"
    if scale * 100 >= MIN_SCALE:
        setup.Zoom = False
        setup.FitToPagesWide = 1
        setup.FitToPagesTall = False
        return f"peisaj, incadrat pe latime ({scale:.0%})"

    setup.Zoom = MIN_SCALE
    setup.Order = XL_OVER_THEN_DOWN
    notes = [f"peisaj {MIN_SCALE}%"]
    if header and not setup.PrintTitleRows:
        setup.PrintTitleRows = ws.Rows(f"{header[0]}:{header[1]}").Address
        notes.append(f"antet repetat {setup.PrintTitleRows}")
    if label is not None and not setup.PrintTitleColumns:
        column = ws.Columns(label)
        if float(column.Width) * MIN_SCALE / 100 <= TITLE_COLUMN_MAX_SHARE * landscape:
            setup.PrintTitleColumns = column.Address
            notes.append(f"coloana repetata {setup.PrintTitleColumns}")
    return ", ".join(notes)


class ExcelApp(OfficeApp):
    progid = "Excel.Application"
    image = "EXCEL.EXE"
    name = "Microsoft Excel"
    # Dupa un export in PDF, Excel ramane uneori ~60 s in memorie dupa Quit (de obicei iese in
    # 2-4 s). Registrul e deja inchis fara salvare, deci oprirea nu pierde nimic. Sub limita de
    # 10 s cu care interfata asteapta inchiderea pool-ului.
    quit_grace = 8.0

    def __init__(self) -> None:
        super().__init__()
        self.open_workbooks = 0  # registrele deschise acum de noi in aceasta instanta
        self.foreign_seen = False  # ultima numaratoare (COM) a gasit registre straine

    def _window_handle(self) -> int | None:
        return int(self.app.Hwnd)  # fereastra principala, exista si cand Excel e ascuns

    def configure(self) -> None:
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.ScreenUpdating = False
        self.app.AskToUpdateLinks = False
        self.app.EnableEvents = False
        self.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        self.open_workbooks = 0
        self.foreign_seen = False

    def has_foreign_documents(self) -> bool:
        # Fara COM (apelat din alt fir, cat timp Excel poate fi blocat): instanta noastra e
        # ascunsa, deci o fereastra XLMAIN vizibila inseamna ca altcineva lucreaza in ea; plus
        # numaratoarea facuta de firul nostru inainte de ultima deschidere.
        if self.foreign_seen:
            return True
        return bool(self.pid) and bool(_visible_windows(self.pid, {"XLMAIN"}))

    def release_abandoned(self) -> None:
        # registrele straine raman la vedere, in grija utilizatorului
        self.prepare_quit()

    def prepare_quit(self) -> bool:
        """Refuza inchiderea daca instanta are registre straine.

        Un registru deschis de altcineva in instanta noastra nu are voie sa fie pierdut: cu
        DisplayAlerts = False, Quit l-ar inchide fara sa intrebe si fara sa-l salveze, iar
        oprirea fortata la fel. Instanta e facuta vizibila si lasata utilizatorului, care o
        inchide singur."""
        try:
            count = int(self.app.Workbooks.Count)
        except Exception as exc:  # intentionat larg: aplicatia poate fi deja moarta
            log.debug("Workbooks.Count %s: %s", self.name, describe_error(exc))
            count = 0
        foreign = count - self.open_workbooks
        if foreign <= 0:
            return True
        log.warning("%s (pid %s) are %d registru(e) deschis(e) de altcineva; nu il inchid si "
                    "nu il opresc, il las vizibil utilizatorului", self.name, self.pid, foreign)
        # fiecare separat: o setare refuzata nu are voie sa lase instanta ascunsa
        with contextlib.suppress(Exception):  # intentionat larg: COM
            self.app.DisplayAlerts = True
        with contextlib.suppress(Exception):  # intentionat larg: COM
            self.app.ScreenUpdating = True
        with contextlib.suppress(Exception):  # intentionat larg: COM
            self.app.EnableEvents = True
        with contextlib.suppress(Exception):  # intentionat larg: COM
            self.app.Visible = True
        with contextlib.suppress(Exception):  # intentionat larg: COM
            self.app.UserControl = True  # ramane deschis dupa ce eliberam referinta
        return False

    def convert(self, job: Job, options: ExportOptions, src: Path) -> None:
        try:
            self.foreign_seen = int(self.app.Workbooks.Count) > self.open_workbooks
        except Exception as exc:  # intentionat larg: o eroare aici apare oricum la Open
            log.debug("Workbooks.Count %s: %s", self.name, describe_error(exc))
        wb = self.app.Workbooks.Open(
            Filename=str(src),
            UpdateLinks=0,
            ReadOnly=True,
            Password=NO_PASSWORD,
            WriteResPassword=NO_PASSWORD,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
        self.open_workbooks += 1
        try:
            self._limit_to_content(wb)
            wb.ExportAsFixedFormat(
                Type=XL_TYPE_PDF,
                Filename=str(job.dst),
                Quality=1 if options.for_screen else 0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        finally:
            if not self.abandoned:  # Close ar astepta si el Excel-ul blocat
                wb.Close(SaveChanges=False)
                self.open_workbooks -= 1  # doar daca s-a inchis: altfel e tot al nostru, deschis

    # Excel tipareste toata zona „folosita” a foii, care poate fi mult mai mare decat datele:
    # formatare pe coloane intregi, formule trase in jos care intorc "", un spatiu uitat.
    # Rezultatul: sute sau mii de pagini goale. Setam zona de tiparire pe continutul real
    # (registrul e deschis doar pentru citire si nu se salveaza, deci fisierul nu se schimba).

    def _limit_to_content(self, wb: Any) -> None:
        with contextlib.suppress(Exception):  # lipseste doar in Excel foarte vechi
            self.app.PrintCommunication = False  # PageSetup e mult mai rapid asa
        try:
            visible = [ws for ws in wb.Worksheets if ws.Visible == XL_SHEET_VISIBLE]
            empty = []
            for ws in visible:
                try:
                    if ws.PageSetup.PrintArea:
                        continue  # zona de tiparire definita de utilizator ramane
                    rects = _content_rects(ws)
                    areas = excelarea.content_areas(rects)
                    if not areas:
                        empty.append(ws)
                        continue
                    ranges = [ws.Range(ws.Cells(a[0], a[1]), ws.Cells(a[2], a[3]))
                              for a in areas]
                    print_area = ",".join(r.Address for r in ranges)
                    if len(print_area) > MAX_PRINT_AREA_CHARS:  # limita Excel: o singura zona
                        areas = [excelarea.bounding_box(areas)]
                        a = areas[0]
                        ranges = [ws.Range(ws.Cells(a[0], a[1]), ws.Cells(a[2], a[3]))]
                        print_area = ranges[0].Address
                    ws.PageSetup.PrintArea = print_area

                    # antetul si coloana de identificare se cauta in zona cea mai lata
                    widest = max(range(len(ranges)), key=lambda i: float(ranges[i].Width))
                    main = areas[widest]
                    cells = [r for r in rects if r[0] == r[2] and r[1] == r[3]]
                    header = excelarea.header_rows(cells, main)
                    label = excelarea.label_column(cells, main, header[1] if header else main[0])
                    layout = _layout_wide(ws, float(ranges[widest].Width), header, label)
                    log.info("Foaia %r: zona de tiparire %s%s", ws.Name, print_area,
                             f", {layout}" if layout else "")
                except Exception as exc:  # o foaie ciudata nu opreste conversia
                    log.debug("Zona de tiparire pentru %r: %s", ws.Name, describe_error(exc))
            # foile fara continut ar da doar pagini goale; macar una ramane vizibila
            for ws in empty[: len(empty) - (len(empty) == len(visible))]:
                with contextlib.suppress(Exception):  # intentionat larg: COM, foaia ramane vizibila
                    ws.Visible = XL_SHEET_HIDDEN
                    log.info("Foaia %r e goala, o sar", ws.Name)
        finally:
            with contextlib.suppress(Exception):  # intentionat larg: COM, doar viteza
                self.app.PrintCommunication = True
