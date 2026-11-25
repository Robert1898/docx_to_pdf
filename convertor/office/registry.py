"""Registry: ce aplicatii Office exista si setarile Word pe care le atingem temporar
(mesajul la deschiderea PDF-urilor, marcajele „document periculos”)."""

from __future__ import annotations

import contextlib
import json
import logging
import os
import threading
import winreg
from pathlib import Path

from ..paths import state_dir

log = logging.getLogger(__package__)  # un singur jurnal pentru tot pachetul


def progid_available(progid: str) -> bool:
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid))
        return True
    except OSError:
        return False


def word_version() -> str:
    """Versiunea Office a lui Word instalat, de ex. „16.0” (din cheia CurVer a lui
    Word.Application). OSError daca Word nu e instalat sau versiunea nu poate fi citita."""
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer") as key:
        progid = winreg.QueryValue(key, None)  # "Word.Application.16"
    major = progid.rpartition(".")[2]
    if not major.isdigit():
        raise OSError(f"Versiune Word necunoscuta: {progid!r}")
    return f"{major}.0"


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
        return rf"Software\Microsoft\Office\{word_version()}\Word\Options"

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
        return cls.ROOT.format(version=word_version())

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
