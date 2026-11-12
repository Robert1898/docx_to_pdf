"""Conversie prin Microsoft Office (COM), cu timeout per fisier.

Fiecare aplicatie Office porneste o singura data si e refolosita pentru toate fisierele.
Daca un fisier depaseste timpul alocat (de obicei un dialog invizibil care asteapta
raspuns), procesul Office pornit de noi e oprit fortat si repornit pentru urmatorul fisier.
Nu oprim niciodata o instanta Office pe care utilizatorul o avea deja deschisa.
"""

from __future__ import annotations

import contextlib
import json
import logging
import os
import signal
import subprocess
import threading
import winreg
from pathlib import Path
from typing import Any

from . import excelarea
from .core import ExportOptions, Job, Kind
from .i18n import _
from .paths import state_dir

log = logging.getLogger(__name__)

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# Parola gresita intentionat: un document neprotejat o ignora, unul protejat da eroare
# imediat in loc sa deschida un dialog invizibil de parola.
NO_PASSWORD = "\u0001convertor-no-password"

MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3  # macro-urile nu ruleaza la deschidere
WD_FORMAT_DOCX = 16
WD_EXPORT_PDF = 17
XL_TYPE_PDF = 0
PP_FIXED_FORMAT_PDF = 2
PP_SAVE_AS_PDF = 32
PP_ALERTS_NONE = 1

# coduri HRESULT care inseamna ca procesul Office a murit sau nu mai raspunde
OFFICE_DIED = "Aplicația Office s-a închis neașteptat."
POWERPOINT_LOCK = threading.Lock()

DEAD_SERVER_CODES = {
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2147023170,  # RPC_S_CALL_FAILED
    -2147417848,  # RPC_E_DISCONNECTED
    -2147417851,  # RPC_E_SERVERFAULT
    -2147418111,  # RPC_E_CALL_REJECTED
}


class ConversionError(Exception):
    """Eroare cu mesaj prietenos, afisabil direct utilizatorului."""


class ConversionTimeout(ConversionError):
    pass


def progid_available(progid: str) -> bool:
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid))
        return True
    except OSError:
        return False


def describe_error(exc: BaseException) -> str:
    try:
        import pywintypes
    except ImportError:  # pragma: no cover - doar pe sisteme fara pywin32
        return str(exc)
    if isinstance(exc, pywintypes.com_error):
        hresult, message, excepinfo, _arg = exc.args
        detail = (excepinfo[2] if excepinfo and excepinfo[2] else message) or str(exc)
        low = detail.lower()
        if "password" in low or "parol" in low:
            return _("Documentul este protejat cu parolă.")
        if hresult in DEAD_SERVER_CODES or (excepinfo and excepinfo[5] in DEAD_SERVER_CODES):
            return _(OFFICE_DIED)
        return " ".join(detail.split())
    return str(exc) or exc.__class__.__name__


# --- optiunea lui Word care afiseaza un mesaj la deschiderea PDF-urilor -----------------


class PdfWarningGuard:
    """Dezactiveaza temporar mesajul „Word va converti PDF-ul...”.

    Valoarea initiala e salvata intr-un fisier, ca sa poata fi refacuta si dupa un crash
    (la urmatoarea pornire a aplicatiei, prin `restore()`).
    """

    VALUE = "DisableConvertPdfWarning"
    # mai multe instante Word pot lucra in paralel: optiunea se reface abia dupa ultima
    _users = 0
    _lock = threading.RLock()

    @staticmethod
    def _marker() -> Path:
        return state_dir() / "word_pdf_warning.json"

    @staticmethod
    def _options_key() -> str:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer") as key:
            major = winreg.QueryValue(key, None).rsplit(".", 1)[1]  # "Word.Application.16"
        return rf"Software\Microsoft\Office\{major}.0\Word\Options"

    @classmethod
    def suppress(cls) -> None:
        with cls._lock:
            cls._users += 1
            if cls._users == 1:
                cls._suppress()

    @classmethod
    def restore(cls) -> None:
        """Apelat la inchiderea unei instante Word si la pornirea aplicatiei (dupa un crash)."""
        with cls._lock:
            if cls._users > 0:
                cls._users -= 1
                if cls._users > 0:
                    return
            cls._restore()

    @classmethod
    def _suppress(cls) -> None:
        try:
            key_path = cls._options_key()
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                marker = cls._marker()
                if not marker.exists():  # dupa un crash pastram valoarea originala
                    try:
                        old: int | None = winreg.QueryValueEx(key, cls.VALUE)[0]
                    except FileNotFoundError:
                        old = None
                    marker.write_text(json.dumps({"key": key_path, "value": old}))
                winreg.SetValueEx(key, cls.VALUE, 0, winreg.REG_DWORD, 1)
        except OSError as exc:
            log.warning("Nu am putut dezactiva mesajul PDF al lui Word: %s", exc)

    @classmethod
    def _restore(cls) -> None:
        marker = cls._marker()
        try:
            data = json.loads(marker.read_text())
        except FileNotFoundError:
            return
        except (OSError, ValueError) as exc:
            log.warning("Marcaj registry ilizibil, il sterg: %s", exc)
            marker.unlink(missing_ok=True)
            return
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, data["key"]) as key:
                if data.get("value") is None:
                    with contextlib.suppress(FileNotFoundError):
                        winreg.DeleteValue(key, cls.VALUE)
                else:
                    winreg.SetValueEx(key, cls.VALUE, 0, winreg.REG_DWORD, int(data["value"]))
            marker.unlink(missing_ok=True)
            log.info("Am refacut optiunea %s", cls.VALUE)
        except (OSError, KeyError, TypeError, ValueError) as exc:
            log.warning("Nu am putut reface optiunea %s: %s", cls.VALUE, exc)


# --- marcajele „document periculos” ale lui Word -----------------------------------------


class ResiliencyCleaner:
    """Sterge marcajele pe care Word le pune fisierelor deschise cand l-am oprit fortat.

    In timp ce deschide un fisier, Word il noteaza in Resiliency\\StartupItems. Daca e oprit
    in acel moment, la urmatoarea pornire citeste nota, muta fisierul in DisabledItems si
    intreaba (invizibil, deci blocant) „Ultima dată când ați deschis fișierul a cauzat o
    eroare gravă”. Marcajele au aparut din cauza noastra, asa ca le stergem imediat dupa
    oprire, inainte de repornire. Lista fisierelor e salvata pe disc, ca sa putem curata si
    dupa un crash al aplicatiei.
    """

    ROOT = r"Software\Microsoft\Office\{version}\Word\Resiliency"
    VALUE_KEYS = ("StartupItems", "DisabledItems")
    RECOVERY_KEY = "DocumentRecovery"
    CRASH_VALUE = "PotentialDataLossInfo2"
    _lock = threading.RLock()  # mai multe fire Word pot scrie lista in acelasi timp

    @staticmethod
    def _list_file() -> Path:
        return state_dir() / "word_killed_files.json"

    @classmethod
    def _load(cls) -> list[str]:
        try:
            data = json.loads(cls._list_file().read_text(encoding="utf-8"))
            return [str(p) for p in data] if isinstance(data, list) else []
        except (OSError, ValueError):
            return []

    @classmethod
    def _store(cls, paths: list[str]) -> None:
        path = cls._list_file()
        if paths:
            path.write_text(json.dumps(paths, ensure_ascii=False), encoding="utf-8")
        else:
            path.unlink(missing_ok=True)

    @classmethod
    def remember(cls, path: Path) -> None:
        with cls._lock:
            paths = cls._load()
            if str(path) not in paths:
                cls._store([*paths, str(path)])

    @staticmethod
    def mentions(data: object, wanted: set[str]) -> bool:
        """Calea apare in valoare? Word o scrie in UTF-16, nu neaparat la offset par."""
        if isinstance(data, str):
            texts = [data]
        elif isinstance(data, bytes):
            texts = [data.decode("utf-16-le", "ignore"), data[1:].decode("utf-16-le", "ignore")]
        else:
            return False
        return any(path in text.lower() for text in texts for path in wanted)

    @classmethod
    def root(cls) -> str:
        version = PdfWarningGuard._options_key().split("\\")[3]  # "16.0"
        return cls.ROOT.format(version=version)

    @staticmethod
    def _values(key: winreg.HKEYType) -> list[tuple[str, object]]:
        values, index = [], 0
        while True:
            try:
                name, data, _type = winreg.EnumValue(key, index)
            except OSError:
                return values
            values.append((name, data))
            index += 1

    @staticmethod
    def _subkeys(key: winreg.HKEYType) -> list[str]:
        names, index = [], 0
        while True:
            try:
                names.append(winreg.EnumKey(key, index))
            except OSError:
                return names
            index += 1

    @classmethod
    def clean(cls) -> None:
        with cls._lock:
            cls._clean()

    @classmethod
    def _clean(cls) -> None:
        paths = cls._load()
        if not paths:
            return
        wanted = {os.path.normcase(p) for p in paths}
        access = winreg.KEY_READ | winreg.KEY_SET_VALUE
        try:
            root = cls.root()
        except OSError:
            return
        startup_cleared = False
        try:
            for sub in cls.VALUE_KEYS:
                with contextlib.suppress(FileNotFoundError), winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER, rf"{root}\{sub}", 0, access
                ) as key:
                    for name, data in cls._values(key):
                        if cls.mentions(data, wanted):
                            winreg.DeleteValue(key, name)
                            startup_cleared |= sub == "StartupItems"
                            log.info("Am sters marcajul Word %s\\%s", sub, name)

            recovery = rf"{root}\{cls.RECOVERY_KEY}"
            with contextlib.suppress(FileNotFoundError), winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, recovery
            ) as key:
                for sub in cls._subkeys(key):
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rf"{recovery}\{sub}") as item:
                        hit = any(cls.mentions(d, wanted) for _n, d in cls._values(item))
                    if hit:
                        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, rf"{recovery}\{sub}")
                        log.info("Am sters inregistrarea de recuperare Word %s", sub)

            if startup_cleared:
                # nota „Word s-a inchis neasteptat” a aparut tot din cauza opririi noastre
                with contextlib.suppress(FileNotFoundError), winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER, root, 0, access
                ) as key:
                    winreg.DeleteValue(key, cls.CRASH_VALUE)
        except OSError as exc:
            log.warning("Nu am putut curata marcajele Word: %s", exc)
        # pastram lista (scurta): Word poate muta o nota in DisabledItems si mai tarziu
        cls._store(paths[-50:])


# --- procese ------------------------------------------------------------------------------


def _pids(image: str) -> set[int]:
    out = subprocess.run(
        ["tasklist", "/FI", f"IMAGENAME eq {image}", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        creationflags=CREATE_NO_WINDOW,
    ).stdout
    pids = set()
    for line in out.splitlines():
        parts = line.strip().strip('"').split('","')
        if len(parts) > 1 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def _kill(pid: int) -> None:
    try:
        os.kill(pid, signal.SIGTERM)  # pe Windows: TerminateProcess
    except OSError as exc:
        log.debug("Procesul %s nu mai exista: %s", pid, exc)


# --- dialoguri ascunse --------------------------------------------------------------------

# Citeste dialogul prin UI Automation (.NET, prezent pe orice Windows 10/11) si apasa primul
# buton: „Da” / „OK” / „Doar citire”. Butonul de inchidere din bara de titlu e ignorat.
_ANSWER_DIALOG_PS = r"""
param([long]$Handle)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$A = [Windows.Automation.AutomationElement]
$root = $A::FromHandle([IntPtr]$Handle)
$all = $root.FindAll([Windows.Automation.TreeScope]::Descendants,
                     [Windows.Automation.Condition]::TrueCondition)
$texts = @(); $buttons = @()
foreach ($e in $all) {
    $c = $e.Current
    if ($c.ControlType -eq [Windows.Automation.ControlType]::Text -and $c.Name) {
        $texts += $c.Name
    }
    if ($c.ControlType -eq [Windows.Automation.ControlType]::Button -and $c.Name -and
        $c.AutomationId -notmatch 'Close' -and -not $c.IsOffscreen) { $buttons += $e }
}
$pressed = ''
if ($buttons.Count -gt 0) {
    $pressed = $buttons[0].Current.Name
    $buttons[0].GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
}
@{ text = ($texts -join ' '); button = $pressed } | ConvertTo-Json -Compress
"""

DIALOG_CLASSES = {"#32770", "NUIDialog", "bosa_sdm_msword", "bosa_sdm_XL9", "bosa_sdm_Mso96"}


def _visible_dialogs(pid: int) -> list[int]:
    import win32gui
    import win32process

    found: list[int] = []

    def visit(hwnd: int, _: object) -> bool:
        if (
            win32gui.IsWindowVisible(hwnd)
            and win32gui.GetClassName(hwnd) in DIALOG_CLASSES
            and win32process.GetWindowThreadProcessId(hwnd)[1] == pid
        ):
            found.append(hwnd)
        return True

    win32gui.EnumWindows(visit, None)
    return found


def answer_dialog(hwnd: int) -> tuple[str, str] | None:
    """Raspunde la un dialog: intoarce (textul dialogului, butonul apasat)."""
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-Command", f"& {{ {_ANSWER_DIALOG_PS} }} {hwnd}"],
            capture_output=True,
            timeout=30,
            creationflags=CREATE_NO_WINDOW,
        )
        data = json.loads(out.stdout.decode("utf-8", errors="replace") or "{}")
    except (OSError, ValueError, subprocess.TimeoutExpired) as exc:
        log.warning("Nu am putut raspunde la dialogul %s: %s", hwnd, exc)
        return None
    return str(data.get("text", "")), str(data.get("button", ""))


class DialogWatcher(threading.Thread):
    """Cat timp un fisier se converteste, raspunde la dialogurile instantei Office.

    Aplicatia Office e invizibila, deci un dialog (de ex. „Ultima dată când ați deschis
    fișierul a cauzat o eroare gravă. Îl deschideți oricum?”) ar bloca totul pana la timeout.
    """

    def __init__(self, pid: int, interval: float = 1.0) -> None:
        super().__init__(name=f"dialogs-{pid}", daemon=True)
        self.pid = pid
        self.interval = interval
        self.answers: list[str] = []
        self._stop_event = threading.Event()
        self._seen: set[int] = set()

    def run(self) -> None:
        while not self._stop_event.wait(self.interval):
            try:
                dialogs = _visible_dialogs(self.pid)
            except Exception as exc:  # procesul poate disparea oricand
                log.debug("Enumerare dialoguri: %s", exc)
                continue
            for hwnd in dialogs:
                if hwnd in self._seen or self._stop_event.is_set():
                    continue
                self._seen.add(hwnd)
                answered = answer_dialog(hwnd)
                if answered:
                    text, button = answered
                    log.warning("Dialog Office: %r -> am apasat %r", text, button)
                    self.answers.append(text)

    def stop(self) -> None:
        self._stop_event.set()


# --- aplicatii Office ---------------------------------------------------------------------


class OfficeApp:
    progid = ""
    image = ""
    name = ""

    def __init__(self) -> None:
        self.app: Any = None
        self.pid: int | None = None
        self.owned = False
        self.timed_out = False

    @classmethod
    def available(cls) -> bool:
        return progid_available(cls.progid)

    def ensure(self) -> Any:
        if self.app is None:
            self.start()
        return self.app

    def start(self) -> None:
        import win32com.client

        before = _pids(self.image)
        self.app = win32com.client.DispatchEx(self.progid)
        new = _pids(self.image) - before
        # PowerPoint permite o singura instanta: daca era deja deschis, il impartim
        self.owned = bool(new)
        self.pid = next(iter(new)) if len(new) == 1 else None
        self.timed_out = False
        log.info("%s pornit (pid=%s, propriu=%s)", self.name, self.pid, self.owned)
        self.configure()

    def configure(self) -> None:
        pass

    def quit(self) -> None:
        if self.app is None:
            return
        try:
            if self.owned and not self.timed_out:
                self.app.Quit()
        except Exception as exc:  # aplicatia poate fi deja moarta
            log.debug("Quit %s: %s", self.name, exc)
        finally:
            self.app = None

    def kill(self) -> bool:
        """Opreste fortat procesul (apelat din alt fir). False daca nu e al nostru."""
        if not (self.owned and self.pid):
            return False
        self.timed_out = True
        _kill(self.pid)
        return True

    def reset(self) -> None:
        """Dupa un timeout sau un crash: uitam instanta, se reporneste la nevoie."""
        self.app = None
        self.pid = None

    def convert(self, job: Job, options: ExportOptions, src: Path) -> None:
        raise NotImplementedError


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

    def configure(self) -> None:
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.ScreenUpdating = False
        self.app.AskToUpdateLinks = False
        self.app.EnableEvents = False
        self.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

    def convert(self, job: Job, options: ExportOptions, src: Path) -> None:
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
            wb.Close(SaveChanges=False)

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
                with contextlib.suppress(Exception):
                    ws.Visible = XL_SHEET_HIDDEN
                    log.info("Foaia %r e goala, o sar", ws.Name)
        finally:
            with contextlib.suppress(Exception):
                self.app.PrintCommunication = True


class PowerPointApp(OfficeApp):
    progid = "PowerPoint.Application"
    image = "POWERPNT.EXE"
    name = "Microsoft PowerPoint"

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


OFFICE_FOR_KIND: dict[Kind, type[OfficeApp]] = {
    Kind.WORD: WordApp,
    Kind.PDF: WordApp,
    Kind.EXCEL: ExcelApp,
    Kind.POWERPOINT: PowerPointApp,
}


class OfficeSession:
    """Tine aplicatiile Office pornite pe durata unui lot. Se foloseste dintr-un singur fir."""

    def __init__(self) -> None:
        self._apps: dict[type[OfficeApp], OfficeApp] = {}
        self._lock = threading.Lock()
        self._com_ready = False
        self._current: tuple[OfficeApp, Path] | None = None

    def __enter__(self) -> OfficeSession:
        import pythoncom

        pythoncom.CoInitialize()
        self._com_ready = True
        return self

    def __exit__(self, *exc: object) -> None:
        import pythoncom

        for app in list(self._apps.values()):
            app.quit()
        self._apps.clear()
        if self._com_ready:
            pythoncom.CoUninitialize()
            self._com_ready = False

    def _app(self, cls: type[OfficeApp]) -> OfficeApp:
        with self._lock:
            if cls not in self._apps:
                self._apps[cls] = cls()
            return self._apps[cls]

    def kill_all(self) -> None:
        """Oprire de urgenta (de exemplu la inchiderea ferestrei). Sigur din orice fir."""
        with self._lock:
            apps = list(self._apps.values())
            current = self._current
        if current is not None and isinstance(current[0], WordApp):
            ResiliencyCleaner.remember(current[1])
        for app in apps:
            app.kill()
        ResiliencyCleaner.clean()

    def convert(self, job: Job, options: ExportOptions, src: Path | None = None) -> str:
        """Converteste un fisier. Intoarce o nota (de ex. dialogurile la care am raspuns)."""
        cls = OFFICE_FOR_KIND[job.kind]
        # PowerPoint are o singura instanta pe calculator: firele paralele il folosesc pe rand
        with POWERPOINT_LOCK if cls is PowerPointApp else contextlib.nullcontext():
            return self._convert(cls, job, options, src)

    def _convert(self, cls: type[OfficeApp], job: Job, options: ExportOptions,
                 src: Path | None) -> str:
        app = self._app(cls)
        opened = src or job.src
        timer: threading.Timer | None = None
        watcher: DialogWatcher | None = None
        try:
            app.ensure()
            timer = threading.Timer(options.timeout, self._on_timeout, args=(app, opened))
            timer.daemon = True
            if app.owned and app.pid:
                watcher = DialogWatcher(app.pid)
                watcher.start()
            with self._lock:
                self._current = (app, opened)
            timer.start()
            app.convert(job, options, opened)
        except Exception as exc:
            dialog = (" " + _("Dialog: „{text}”").format(text=watcher.answers[-1])
                      if watcher and watcher.answers else "")
            if app.timed_out:
                app.reset()
                raise ConversionTimeout(
                    _("A depășit limita de {seconds} s; {app} a fost oprit și va fi repornit.")
                    .format(seconds=options.timeout, app=app.name) + dialog
                ) from exc
            message = describe_error(exc)
            if message == _(OFFICE_DIED):
                app.reset()
            raise ConversionError(message + dialog) from exc
        finally:
            if timer is not None:
                timer.cancel()
            if watcher is not None:
                watcher.stop()
            with self._lock:
                self._current = None
        if watcher and watcher.answers:
            return _("am răspuns automat la {count} dialog(uri) {app}").format(
                count=len(watcher.answers), app=app.name)
        return ""

    @staticmethod
    def _on_timeout(app: OfficeApp, opened: Path) -> None:
        if isinstance(app, WordApp):
            ResiliencyCleaner.remember(opened)
        if app.kill():
            log.warning("Timeout la %s: am oprit %s (pid %s)", opened, app.name, app.pid)
            if isinstance(app, WordApp):
                ResiliencyCleaner.clean()  # inainte ca Word sa fie repornit
        else:
            log.warning(
                "Timeout la %s, dar %s era deschis de utilizator si nu il opresc", opened, app.name
            )
