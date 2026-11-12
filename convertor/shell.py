"""Integrare in Windows Explorer: meniul click-dreapta si „Trimite către”.

Totul se scrie doar pentru utilizatorul curent (HKCU / %APPDATA%), fara drepturi de admin.
Installer-ul apeleaza aceleasi functii (prin docconv.exe), ca sa existe o singura sursa.
"""

from __future__ import annotations

import contextlib
import logging
import os
import subprocess
import sys
import winreg
from pathlib import Path

from . import APP_ID, APP_NAME, GUI_EXE, LEGACY_APP_NAME, LEGACY_MENU_KEY
from .core import EXTENSIONS, SUPPORTED, Kind
from .paths import is_frozen, resource_path

log = logging.getLogger(__name__)

MENU_KEY = APP_ID
CLASSES = r"Software\Classes\SystemFileAssociations"


def gui_command() -> list[str]:
    """Comanda care porneste interfata grafica, din exe sau din sursa."""
    if is_frozen():
        return [str(Path(sys.executable).with_name(GUI_EXE))]
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    interpreter = pythonw if pythonw.exists() else Path(sys.executable)
    return [str(interpreter), str(resource_path("run_gui.py"))]


def _icon() -> str:
    return gui_command()[0] if is_frozen() else str(resource_path("icon.ico"))


def _menu_label(ext: str) -> str:
    return "Convertește în DOCX" if ext in EXTENSIONS[Kind.PDF] else "Convertește în PDF"


def _notify_shell() -> None:
    try:
        from win32com.shell import shell, shellcon

        shell.SHChangeNotify(shellcon.SHCNE_ASSOCCHANGED, shellcon.SHCNF_IDLIST, None, None)
    except Exception as exc:  # doar reimprospatare, nu e critic
        log.debug("SHChangeNotify: %s", exc)


def register_context_menu() -> None:
    command = subprocess.list2cmdline([*gui_command(), "--auto"]) + ' "%1"'
    for ext in SUPPORTED:
        base = rf"{CLASSES}\{ext}\shell\{MENU_KEY}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base) as key:
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, _menu_label(ext))
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, _icon())
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base + r"\command") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)
    _notify_shell()
    log.info("Meniu click-dreapta inregistrat: %s", command)


def unregister_context_menu() -> None:
    """Elimina meniul, inclusiv cel inregistrat sub numele vechi."""
    for key_name in (MENU_KEY, LEGACY_MENU_KEY):
        for ext in SUPPORTED:
            base = rf"{CLASSES}\{ext}\shell\{key_name}"
            for sub in (base + r"\command", base):
                with contextlib.suppress(FileNotFoundError):
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, sub)
    _notify_shell()
    log.info("Meniu click-dreapta eliminat")


def context_menu_registered(key_name: str = MENU_KEY) -> bool:
    try:
        winreg.CloseKey(
            winreg.OpenKey(winreg.HKEY_CURRENT_USER, rf"{CLASSES}\.docx\shell\{key_name}")
        )
        return True
    except OSError:
        return False


def sendto_path(name: str = APP_NAME) -> Path:
    appdata = os.environ.get("APPDATA", str(Path.home() / "AppData" / "Roaming"))
    return Path(appdata) / "Microsoft" / "Windows" / "SendTo" / f"{name}.lnk"


def create_sendto() -> None:
    import win32com.client

    target, *args = gui_command()
    link = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(sendto_path()))
    link.TargetPath = target
    link.Arguments = subprocess.list2cmdline([*args, "--auto"])
    link.WorkingDirectory = str(Path(target).parent)
    link.IconLocation = _icon()
    link.Description = "Convertește documentele selectate"
    link.Save()
    log.info("Scurtatura Trimite catre creata: %s", sendto_path())


def remove_sendto() -> None:
    sendto_path().unlink(missing_ok=True)
    sendto_path(LEGACY_APP_NAME).unlink(missing_ok=True)


def sendto_registered() -> bool:
    return sendto_path().exists()


RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"


def set_autostart(enabled: bool) -> None:
    """Pornire minimizata la logarea in Windows (util pentru folderul urmarit)."""
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
        if enabled:
            command = subprocess.list2cmdline([*gui_command(), "--minimized"])
            winreg.SetValueEx(key, MENU_KEY, 0, winreg.REG_SZ, command)
        else:
            with contextlib.suppress(FileNotFoundError):
                winreg.DeleteValue(key, MENU_KEY)


def autostart_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            winreg.QueryValueEx(key, MENU_KEY)
        return True
    except OSError:
        return False


def migrate_legacy() -> None:
    """Integrarile activate sub numele vechi („Convertor DOCX-PDF”) sunt mutate pe cel nou."""
    try:
        if context_menu_registered(LEGACY_MENU_KEY):
            unregister_context_menu()
            register_context_menu()
            log.info("Meniul click-dreapta a fost mutat pe numele nou")
        if sendto_path(LEGACY_APP_NAME).exists():
            remove_sendto()
            create_sendto()
            log.info("Scurtatura Trimite catre a fost mutata pe numele nou")
    except Exception as exc:  # migrarea nu are voie sa impiedice pornirea
        log.warning("Migrarea integrarii cu Explorer a esuat: %s", exc)
