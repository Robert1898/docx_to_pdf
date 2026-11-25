"""Dialogurile ascunse ale instantelor Office: citite prin UI Automation si raspunse doar
cu butoane din lista alba."""

from __future__ import annotations

import json
import logging
import subprocess
import threading

from .process import CREATE_NO_WINDOW

log = logging.getLogger(__package__)  # un singur jurnal pentru tot pachetul

# Butoanele pe care avem voie sa le apasam, in ordinea preferintei. „Doar citire” e primul:
# e raspunsul cel mai sigur, si un dialog care il ofera are de obicei si „OK” (de ex. parola
# de scriere, unde „OK” cu parola goala ar da doar o alta eroare). Fisierele sunt deschise
# doar pentru citire si nu sunt salvate niciodata, asa ca un raspuns afirmativ e sigur.
# Orice alt dialog (de ex. doar „Nu” / „Anulare”) ramane neatins: e notat in jurnal, iar
# limita de timp opreste instanta.
ALLOWED_DIALOG_BUTTONS = ("Doar citire", "Read-Only", "Da", "Yes", "Deschide", "Open", "OK")


def _button_key(name: str) -> str:
    """„&Doar citire”, „Read-only ”, „DA” -> aceeasi cheie: fara accelerator, spatii, majuscule."""
    return " ".join(name.replace("&", "").split()).casefold()


def choose_button(names: list[str], allowed: tuple[str, ...] = ALLOWED_DIALOG_BUTTONS
                  ) -> str | None:
    """Butonul de apasat dintre `names` (asa cum apar in dialog), sau None daca niciunul nu e
    in lista alba. La mai multe potriviri castiga cel mai devreme in `allowed`."""
    keys = {_button_key(name): name for name in reversed(names)}  # primul cu acel nume
    for wanted in allowed:
        name = keys.get(_button_key(wanted))
        if name is not None:
            return name
    return None


# Citeste dialogul prin UI Automation (.NET, prezent pe orice Windows 10/11): textul si numele
# butoanelor. Butonul de inchidere din bara de titlu e ignorat.
_FIND_BUTTONS_PS = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$root = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Handle)
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
"""

_READ_DIALOG_PS = r"""
param([long]$Handle)
""" + _FIND_BUTTONS_PS + r"""
@{ text = ($texts -join ' '); buttons = @($buttons | ForEach-Object { $_.Current.Name }) } |
    ConvertTo-Json -Compress
"""

# Apasa butonul cu numele exact ales in Python (numele vine din lista alba, nu din dialog)
_PRESS_BUTTON_PS = r"""
param([long]$Handle, [string]$Name)
""" + _FIND_BUTTONS_PS + r"""
$pressed = ''
$button = $buttons | Where-Object { $_.Current.Name -ceq $Name } | Select-Object -First 1
if ($button) {
    $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
    $pressed = $Name
}
@{ button = $pressed } | ConvertTo-Json -Compress
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


def _run_dialog_script(script: str, *args: str) -> dict:
    out = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
         "-Command", f"& {{ {script} }} {' '.join(args)}"],
        capture_output=True,
        timeout=30,
        creationflags=CREATE_NO_WINDOW,
    )
    data = json.loads(out.stdout.decode("utf-8", errors="replace") or "{}")
    return data if isinstance(data, dict) else {}


def _names(value: object) -> list[str]:
    # ConvertTo-Json face dintr-o lista cu un singur element un text simplu
    if isinstance(value, str):
        return [value]
    return [str(v) for v in value] if isinstance(value, list) else []


def answer_dialog(hwnd: int) -> tuple[str, list[str], str | None] | None:
    """Citeste un dialog si apasa un buton din lista alba, daca are unul.

    Intoarce (textul dialogului, butoanele lui, butonul apasat sau None), ori None daca
    dialogul nu a putut fi citit."""
    try:
        data = _run_dialog_script(_READ_DIALOG_PS, str(hwnd))
    except (OSError, ValueError, subprocess.TimeoutExpired) as exc:
        log.warning("Nu am putut citi dialogul %s: %s", hwnd, exc)
        return None
    text, buttons = str(data.get("text", "")), _names(data.get("buttons"))
    choice = choose_button(buttons)
    if choice is None:
        return text, buttons, None
    quoted = "'" + choice.replace("'", "''") + "'"
    try:
        pressed = str(_run_dialog_script(_PRESS_BUTTON_PS, str(hwnd), quoted).get("button", ""))
    except (OSError, ValueError, subprocess.TimeoutExpired) as exc:
        log.warning("Nu am putut apasa %r in dialogul %s: %s", choice, hwnd, exc)
        return text, buttons, None
    return text, buttons, pressed or None  # gol: dialogul s-a inchis intre timp


class DialogWatcher(threading.Thread):
    """Cat timp un fisier se converteste, raspunde la dialogurile instantei Office.

    Aplicatia Office e invizibila, deci un dialog (de ex. „Ultima dată când ați deschis
    fișierul a cauzat o eroare gravă. Îl deschideți oricum?”) ar bloca totul pana la timeout.
    Se apasa doar un buton din ALLOWED_DIALOG_BUTTONS; un dialog fara niciunul e doar notat,
    iar limita de timp opreste instanta.
    """

    def __init__(self, pid: int, interval: float = 1.0) -> None:
        super().__init__(name=f"dialogs-{pid}", daemon=True)
        self.pid = pid
        self.interval = interval
        self.answers: list[str] = []  # textele dialogurilor la care am raspuns
        self.unanswered: list[str] = []  # textele dialogurilor lasate neatinse
        self._stop_event = threading.Event()
        self._seen: set[int] = set()

    @property
    def last_text(self) -> str | None:
        """Ultimul dialog vazut (pentru mesajul de eroare), raspuns sau nu."""
        texts = self.answers + self.unanswered
        return texts[-1] if texts else None

    def run(self) -> None:
        while not self._stop_event.wait(self.interval):
            try:
                dialogs = _visible_dialogs(self.pid)
            except Exception as exc:  # intentionat larg: procesul poate disparea oricand
                log.debug("Enumerare dialoguri: %s", exc)
                continue
            for hwnd in dialogs:
                if hwnd in self._seen or self._stop_event.is_set():
                    continue
                self._seen.add(hwnd)
                answered = answer_dialog(hwnd)
                if answered is None:
                    continue
                text, buttons, pressed = answered
                if pressed:
                    log.warning("Dialog Office: %r, butoane %r -> am apasat %r",
                                text, buttons, pressed)
                    self.answers.append(text)
                else:
                    log.warning("Dialog Office: %r, butoane %r -> niciun buton permis, nu apas "
                                "nimic (limita de timp va opri instanta)", text, buttons)
                    self.unanswered.append(text)

    def stop(self) -> None:
        self._stop_event.set()
