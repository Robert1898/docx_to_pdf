"""Document Converter: conversie in masa intre documente Office si PDF, plus unelte PDF.

Numele de aici sunt singura sursa: spec-ul PyInstaller, build.ps1 si installer-ul le citesc.
"""

__version__ = "2.2.1"

APP_NAME = "Document Converter"
APP_TITLE = APP_NAME
APP_ID = "DocumentConverter"  # foldere de setari/jurnale, chei de registry
PUBLISHER = "Robert-Raul Mîrza"
GUI_EXE = f"{APP_NAME}.exe"
CLI_NAME = "docconv"
CLI_EXE = f"{CLI_NAME}.exe"

# numele vechi (pana la 2.0), pentru migrarea setarilor si a integrarii cu Explorer
LEGACY_APP_ID = "ConvertorDOCX-PDF"
LEGACY_APP_NAME = "Convertor DOCX-PDF"
LEGACY_MENU_KEY = "ConvertorDOCXPDF"
