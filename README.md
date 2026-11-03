# Convertor DOCX ⇄ PDF

Aplicație Windows cu interfață grafică (tkinter) care convertește în masă documente Word în PDF și PDF-uri în DOCX.

## Funcții
- Adaugi fișiere sau un întreg folder (se caută recursiv)
- Direcția se alege automat după tipul fișierului:
  - `.docx`, `.doc`, `.rtf`, `.odt` → `.pdf`
  - `.pdf` → `.docx`
- Rezultatul se salvează lângă original sau într-un folder ales
- Avertizează înainte să suprascrie fișiere existente
- Bară de progres și raport de erori la final

## Cerințe
Pentru Word → PDF ai nevoie de **unul dintre**:
- Microsoft Word instalat (folosit prioritar, prin `pywin32`)
- [LibreOffice](https://www.libreoffice.org/) (folosit automat dacă Word lipsește)

Pentru PDF → DOCX e nevoie de **Microsoft Word** (2013 sau mai nou). Calitatea depinde de PDF:
documentele cu text se convertesc bine, PDF-urile scanate (imagini) nu devin text editabil.
În timpul conversiei, aplicația dezactivează temporar mesajul lui Word „Word va converti PDF-ul…”
(opțiunea `DisableConvertPdfWarning`) și o readuce la valoarea inițială la final.

## Utilizare

### Varianta 1: executabilul
Descarcă `Convertor DOCX-PDF.exe` din secțiunea [Releases](../../releases) și rulează-l.
Executabilul nu este semnat digital, așa că Windows SmartScreen poate afișa un avertisment
(„More info" → „Run anyway").

### Varianta 2: din sursă
```bash
pip install -r requirements.txt
python docx_to_pdf_gui.py
```

### Construire .exe
```bash
pyinstaller "Convertor DOCX-PDF.spec"
```
Rezultatul apare în `dist/`.

## Licență
MIT, vezi fișierul [LICENSE](LICENSE).
