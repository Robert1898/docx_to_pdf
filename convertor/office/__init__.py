"""Conversie prin Microsoft Office (COM), cu timeout per fisier.

Fiecare aplicatie Office porneste o singura data si e refolosita pentru toate fisierele.
Daca un fisier depaseste timpul alocat (de obicei un dialog invizibil care asteapta
raspuns), procesul Office pornit de noi e oprit fortat si repornit pentru urmatorul fisier.
Nu oprim niciodata o instanta Office pe care utilizatorul o avea deja deschisa.

Module:
  base.py        erorile pentru utilizator si aplicatia Office generica (OfficeApp)
  word.py        Word: documente -> PDF, PDF -> DOCX
  excel.py       Excel -> PDF, zona de tiparire pe continutul real
  powerpoint.py  PowerPoint -> PDF
  session.py     OfficeSession: aplicatiile unui lot, limita de timp per fisier
  dialogs.py     raspunsul la dialogurile ascunse (lista alba de butoane)
  registry.py    setarile Word atinse temporar si marcajele lasate dupa oprire
  process.py     gasirea, asteptarea si oprirea proceselor Office
"""

from __future__ import annotations

from .base import OFFICE_DIED, ConversionError, ConversionTimeout, OfficeApp, describe_error
from .dialogs import ALLOWED_DIALOG_BUTTONS, DialogWatcher, answer_dialog, choose_button
from .excel import ExcelApp
from .powerpoint import PowerPointApp
from .process import CREATE_NO_WINDOW
from .registry import PdfWarningGuard, ResiliencyCleaner, progid_available
from .session import OFFICE_FOR_KIND, OfficeSession
from .word import WordApp

__all__ = [
    "ALLOWED_DIALOG_BUTTONS",
    "CREATE_NO_WINDOW",
    "OFFICE_DIED",
    "OFFICE_FOR_KIND",
    "ConversionError",
    "ConversionTimeout",
    "DialogWatcher",
    "ExcelApp",
    "OfficeApp",
    "OfficeSession",
    "PdfWarningGuard",
    "PowerPointApp",
    "ResiliencyCleaner",
    "WordApp",
    "answer_dialog",
    "choose_button",
    "describe_error",
    "progid_available",
]
