"""Traducerea interfetei. Textele sunt scrise in romana in cod si trec prin `_()`.

Textele cu valori variabile folosesc acolade: `_("Am creat {name}").format(name=...)`, ca
traducerea sa poata muta valorile unde e firesc in limba respectiva. Jurnalul ramane in romana.
"""

from __future__ import annotations

import ctypes

LANGUAGES = ("system", "ro", "en")
_current = "ro"


def system_language() -> str:
    """Limba interfetei Windows: romana sau, pentru orice altceva, engleza."""
    try:
        lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
    except (AttributeError, OSError):
        return "ro"
    return "ro" if lang_id & 0x3FF == 0x18 else "en"  # LANG_ROMANIAN


def set_language(choice: str) -> str:
    global _current
    _current = system_language() if choice == "system" else choice if choice in LANGUAGES else "ro"
    return _current


def current() -> str:
    return _current


def _(text: str) -> str:
    if _current == "en":
        from .i18n_en import EN

        return EN.get(text, text)
    return text
