# Convertor DOCX → PDF

Aplicație Windows cu interfață grafică (tkinter) care convertește în masă documente Word în PDF.

## Funcții
- Adaugi fișiere sau un întreg folder (se caută recursiv)
- Formate acceptate: `.docx`, `.doc`, `.rtf`, `.odt`
- PDF-ul se salvează lângă original sau într-un folder ales
- Avertizează înainte să suprascrie PDF-uri existente
- Bară de progres și raport de erori la final

## Cerințe
Pentru conversie ai nevoie de **unul dintre**:
- Microsoft Word instalat (folosit prioritar, prin `pywin32`)
- [LibreOffice](https://www.libreoffice.org/) (folosit automat dacă Word lipsește)

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
