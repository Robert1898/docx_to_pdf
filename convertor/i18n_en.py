"""Traducerea in engleza. Cheia este textul romanesc din cod, exact cum apare in `_()`.

Un text lipsa de aici apare in romana; testul `test_i18n` verifica ca nu lipseste niciunul.
"""

from __future__ import annotations

EN: dict[str, str] = {
    # --- linia de comanda -------------------------------------------------------------------
    "EROARE": "ERROR",
    "OMIS": "SKIPPED",
    "ANULAT": "CANCELLED",
    "integrare Windows": "Windows integration",
    "fișiere sau foldere (recursiv)": "files or folders (recursive)",
    "folder destinație (implicit: lângă original)":
        "output folder (default: next to the original)",
    "suprascrie rezultatele existente (implicit sunt omise)":
        "overwrite existing results (skipped by default)",
    "PDF/A pentru arhivare": "PDF/A for archiving",
    "fără semne de carte din titluri": "no bookmarks from headings",
    "optimizat pentru ecran (fișier mai mic)": "optimized for screen (smaller file)",
    "OCR pe PDF-urile scanate înainte de DOCX": "OCR scanned PDFs before converting to DOCX",
    "în ce se convertesc PDF-urile: docx (implicit) sau xlsx (tabele)":
        "what PDFs are converted to: docx (default) or xlsx (tables)",
    "PDF → DOCX: auto (implicit), layout = aspect fidel (rapoarte, tabele), word = text "
    "curgător prin Word":
        "PDF → DOCX: auto (default), layout = exact layout (reports, tables), word = flowing "
        "text through Word",
    "imagine → PDF: a4 = încadrată pe A4 (implicit), image = pagina are mărimea imaginii":
        "image → PDF: a4 = fitted on A4 (default), image = page has the image's size",
    "limită de timp per fișier (implicit 180)": "time limit per file (default 180)",
    "urmărește folderul și convertește fișierele noi (Ctrl+C oprește); rezultatele merg în "
    "--out sau în FOLDER\\Convertite":
        "watch the folder and convert new files (Ctrl+C stops); results go to --out or to "
        "FOLDER\\Convertite",
    "adaugă „Convertește” în meniul click-dreapta": "add “Convert” to the right-click menu",
    "adaugă aplicația în „Trimite către”": "add the app to “Send to”",
    "elimină integrările (meniu, Trimite către, pornire cu Windows)":
        "remove the integrations (menu, Send to, start with Windows)",
    "--timeout trebuie să fie cel puțin 10 secunde.": "--timeout must be at least 10 seconds.",
    "Nu am găsit fișiere suportate.": "No supported files found.",
    "Urmăresc {folder} → {out}. Ctrl+C oprește.": "Watching {folder} → {out}. Ctrl+C stops.",
    "Oprit.": "Stopped.",
    "Întrerupt.": "Interrupted.",
    "{app}: documente Word, Excel, PowerPoint și imagini în PDF; PDF în DOCX sau Excel.":
        "{app}: Word, Excel, PowerPoint documents and images to PDF; PDF to DOCX or Excel.",
    "conversii Office simultane, 1-{max} (implicit 2)":
        "simultaneous Office conversions, 1-{max} (default 2)",
    "Folderul urmărit nu există: {folder}": "The watched folder does not exist: {folder}",
    "--jobs trebuie să fie între 1 și {max}.": "--jobs must be between 1 and {max}.",
    "{done} convertite, {skipped} omise, {failed} cu erori.":
        "{done} converted, {skipped} skipped, {failed} failed.",
    # --- motor ------------------------------------------------------------------------------
    "Imagine": "Image",
    "În așteptare": "Waiting",
    "Se convertește…": "Converting…",
    "Gata": "Done",
    "Eroare": "Error",
    "Omis (există deja)": "Skipped (already exists)",
    "Anulat": "Cancelled",
    "prin LibreOffice": "via LibreOffice",
    "Nu am găsit Microsoft Office sau LibreOffice instalat.":
        "Neither Microsoft Office nor LibreOffice is installed.",
    "cu OCR": "with OCR",
    "PDF scanat, dar Tesseract nu e instalat: fără OCR":
        "scanned PDF, but Tesseract is not installed: no OCR",
    "PDF-ul e scanat (conține doar imagini). Activează OCR (necesită Tesseract) ca să obții "
    "tabele.":
        "The PDF is scanned (images only). Turn on OCR (requires Tesseract) to get tables.",
    "{rows} rânduri": "{rows} rows",
    "PDF scanat: prin Word, care păstrează imaginile":
        "scanned PDF: through Word, which keeps the images",
    "Word lipsește": "Word is missing",
    "aspect fidel": "exact layout",
    "prin Word": "through Word",
    "{count} pagini": "{count} pages",
    "PDF-ul e scanat (conține doar imagini). Activează OCR (necesită Tesseract) ca să obții "
    "text, sau instalează Microsoft Word.":
        "The PDF is scanned (images only). Turn on OCR (requires Tesseract) to get text, or "
        "install Microsoft Word.",
    "Aplicația Office s-a închis neașteptat.": "The Office application closed unexpectedly.",
    "Documentul este protejat cu parolă.": "The document is password protected.",
    "am răspuns automat la {count} dialog(uri) {app}":
        "automatically answered {count} {app} dialog(s)",
    "Dialog: „{text}”": "Dialog: “{text}”",
    "A depășit limita de {seconds} s; {app} a fost oprit și va fi repornit.":
        "Exceeded the {seconds} s limit; {app} was stopped and will be restarted.",
    "Documentul nu mai are nicio pagină.": "The document has no pages left.",
    "Nu pot citi imaginea {name}: {error}": "Cannot read the image {name}: {error}",
    "PDF-ul este protejat cu parolă.": "The PDF is password protected.",
    "Pentru unire ai nevoie de cel puțin două fișiere.": "Merging needs at least two files.",
    "Numărul de pagini pe fișier trebuie să fie cel puțin 1.":
        "The number of pages per file must be at least 1.",
    "Numărul de părți trebuie să fie cel puțin 1.": "The number of parts must be at least 1.",
    "Scrie cel puțin un interval, de exemplu 1-200, 201-400.":
        "Enter at least one range, for example 1-200, 201-400.",
    "Parola nu poate fi goală.": "The password cannot be empty.",
    "Textul pentru watermark nu poate fi gol.": "The watermark text cannot be empty.",
    "Fișierul descărcat nu este un model OCR valid.":
        "The downloaded file is not a valid OCR model.",
    "PDF-ul are doar {total} pagini, nu se poate împărți în {parts} părți.":
        "The PDF has only {total} pages and cannot be split into {parts} parts.",
    "pagina": "page",
    "paginile": "pages",
    "Nu pot citi PDF-ul: {error}": "Cannot read the PDF: {error}",
    "Interval neînțeles: „{part}”. Exemplu corect: 1-200, 201-400.":
        "Unrecognized range: “{part}”. Correct example: 1-200, 201-400.",
    "Interval invalid: „{part}”.": "Invalid range: “{part}”.",
    "Intervalul „{part}” depășește PDF-ul, care are {total} pagini.":
        "The range “{part}” goes past the end of the PDF, which has {total} pages.",
    "Intervalul {start}-{end} nu există în PDF ({total} pagini).":
        "The range {start}-{end} does not exist in the PDF ({total} pages).",
    "OCR a depășit {seconds} s la pagina {page}.": "OCR exceeded {seconds} s on page {page}.",
    "Tesseract a eșuat la pagina {page}: {detail}": "Tesseract failed on page {page}: {detail}",
    # --- fereastra principala ---------------------------------------------------------------
    "Eroare neașteptată": "Unexpected error",
    "Conversie": "Conversion",
    "Unelte PDF": "PDF tools",
    "Pagini": "Pages",
    "Setări": "Settings",
    "{done} convertite, {failed} eșuate din folderul urmărit.":
        "{done} converted, {failed} failed from the watched folder.",
    "Urmăresc „{folder}”.": "Watching “{folder}”.",
    "{error}\n\nDetaliile sunt în jurnal (Setări → Jurnal).":
        "{error}\n\nThe details are in the log (Settings → Log).",
    "Folder urmărit": "Watched folder",
    "Lucru în desfășurare": "Work in progress",
    "O operație e încă în curs. Închizi oricum? Fișierul curent nu va fi terminat.":
        "An operation is still running. Close anyway? The current file will not be finished.",
    # --- tab Conversie ----------------------------------------------------------------------
    "Trage aici documente Word, Excel, PowerPoint, PDF sau imagini\n(sau foldere întregi), "
    "ori folosește butoanele de mai sus.":
        "Drag Word, Excel, PowerPoint, PDF documents or images here\n(or whole folders), or "
        "use the buttons above.",
    "Toate documentele": "All documents",
    "Imagini": "Images",
    "Automat (recomandat)": "Automatic (recommended)",
    "Aspect fidel (rapoarte, tabele)": "Exact layout (reports, tables)",
    "Text editabil (Word)": "Editable text (Word)",
    "Word (DOCX)": "Word (DOCX)",
    "Excel (XLSX)": "Excel (XLSX)",
    "încadrată pe A4": "fitted on A4",
    "la mărimea imaginii": "at the image's size",
    "{m} min {s:02d} s": "{m} min {s:02d} s",
    "PDF-ul primește un cuprins navigabil din titlurile documentului Word.":
        "The PDF gets a navigable outline from the Word document's headings.",
    "Format standard pentru arhivare pe termen lung, cerut uneori de instituții.":
        "Standard format for long-term archiving, sometimes required by institutions.",
    "Imagini la rezoluție mai mică: fișier mai mic, potrivit pentru e-mail.":
        "Lower-resolution images: a smaller file, suitable for e-mail.",
    "Dacă un fișier durează mai mult (de obicei un dialog ascuns), aplicația Office pornită "
    "de program e oprită și se trece la următorul fișier.":
        "If a file takes longer (usually because of a hidden dialog), the Office application "
        "started by the program is stopped and the next file follows.",
    "Câte fișiere se convertesc în același timp, fiecare cu propria instanță Office. Ajută "
    "la loturi mari; folosește mai multă memorie.":
        "How many files are converted at the same time, each with its own Office instance. "
        "Helps with large batches; uses more memory.",
    "Aspect fidel: fiecare pagină arată ca în PDF, cu coloanele aliniate exact; potrivit "
    "pentru rapoarte, balanțe, extrase, tabele.\nText editabil (Word): textul curge liber, "
    "potrivit pentru scrisori și contracte; Word ghicește tabelele și poate muta valori.\n"
    "Automat: rapoartele pe coloane merg pe aspect fidel, restul prin Word.":
        "Exact layout: every page looks like in the PDF, with columns aligned exactly; good "
        "for reports, balances, statements, tables.\nEditable text (Word): text flows freely, "
        "good for letters and contracts; Word guesses the tables and may move values.\n"
        "Automatic: column reports use exact layout, everything else goes through Word.",
    "Alege…": "Choose…",
    "Anulează": "Cancel",
    "Convertește": "Convert",
    "Pornesc aplicațiile necesare…": "Starting the required applications…",
    "{s} s": "{s} s",
    "Adaugă fișiere…": "Add files…",
    "Adaugă folder…": "Add folder…",
    "Elimină": "Remove",
    "Golește lista": "Clear list",
    "Fișier": "File",
    "Folder": "Folder",
    "Status": "Status",
    "Opțiuni": "Options",
    "Semne de carte din titluri": "Bookmarks from headings",
    "PDF/A (arhivare)": "PDF/A (archiving)",
    "Optimizat pentru ecran (fișier mai mic)": "Optimized for screen (smaller file)",
    "OCR pentru PDF-uri scanate": "OCR for scanned PDFs",
    "Necesită Tesseract OCR instalat (vezi README).": "Requires Tesseract OCR (see README).",
    "PDF-urile scanate (imagini) sunt trecute prin OCR înainte de conversia în DOCX sau "
    "Excel, ca să rezulte text editabil.":
        "Scanned PDFs (images) go through OCR before being converted to DOCX or Excel, so the "
        "result has editable text.",
    "Destinație": "Destination",
    "Nu am găsit fișiere noi suportate.": "No new supported files found.",
    "Alege documente": "Choose documents",
    "Alege un folder": "Choose a folder",
    "Alege folderul de destinație": "Choose the destination folder",
    "Deschide originalul": "Open the original",
    "Deschide folderul originalului": "Open the original's folder",
    "Valoare invalidă": "Invalid value",
    "Lipsă fișiere": "No files",
    "Adaugă cel puțin un document.": "Add at least one document.",
    "Limita per fișier trebuie să fie între 10 și 3600 de secunde.":
        "The per-file limit must be between 10 and 3600 seconds.",
    "Terminat cu erori": "Finished with errors",
    "Se oprește după fișierele care se convertesc acum…":
        "Stopping after the files being converted now…",
    "Adaugă documente pentru a începe.": "Add documents to begin.",
    "Gata de conversie. Dublu-click pe un rând deschide rezultatul.":
        "Ready to convert. Double-click a row to open the result.",
    "Niciun fișier": "No files",
    "Nu am găsit documente noi în folderul ales.":
        "No new documents found in the chosen folder.",
    "Nu pot deschide": "Cannot open",
    "Deschide rezultatul": "Open the result",
    "Deschide folderul rezultatului": "Open the result's folder",
    "Vezi eroarea": "Show the error",
    "Elimină din listă": "Remove from list",
    "Lipsă destinație": "No destination",
    "Alege folderul de destinație.": "Choose the destination folder.",
    "Fișiere existente": "Existing files",
    "încă ~{time}": "~{time} left",
    "Eroare: {message}": "Error: {message}",
    "{n} convertite": "{n} converted",
    " în {time}.": " in {time}.",
    "Gata: {summary} Dublu-click pe un rând deschide rezultatul.":
        "Done: {summary} Double-click a row to open the result.",
    "Conversie terminată": "Conversion finished",
    "Limită per fișier:": "Limit per file:",
    "secunde": "seconds",
    "Conversii simultane:": "Simultaneous conversions:",
    "PDF în:": "PDF to:",
    "Mod DOCX:": "DOCX mode:",
    "Imagine → PDF:": "Image → PDF:",
    "Lângă fișierul original": "Next to the original file",
    "Deschide folderul la final": "Open the folder when finished",
    "Am primit {n} fișier(e); le adaug după conversia curentă.":
        "Received {n} file(s); they will be added after the current conversion.",
    "Conversiile simultane trebuie să fie între 1 și {n}.":
        "Simultaneous conversions must be between 1 and {n}.",
    "Destinație invalidă": "Invalid destination",
    "{n} omise": "{n} skipped",
    "{n} anulate": "{n} cancelled",
    "{n} cu erori": "{n} failed",
    "{n} fișier(e)": "{n} file(s)",
    "{n} fișier(e) rezultat există deja.\n\nDa – le suprascriu\nNu – le omit\nAnulează – nu "
    "pornesc conversia":
        "{n} result file(s) already exist.\n\nYes – overwrite them\nNo – skip them\nCancel – "
        "don't start the conversion",
    "{done}/{total}: {name}": "{done}/{total}: {name}",
    "{done}/{total} terminate": "{done}/{total} finished",
    "… și încă {n}": "… and {n} more",
    # --- tab Pagini -------------------------------------------------------------------------
    "Deschide PDF…": "Open PDF…",
    "Salvează ca…": "Save as…",
    " (modificat, nesalvat)": " (modified, unsaved)",
    "Se salvează…": "Saving…",
    "Semnătură sau ștampilă": "Signature or stamp",
    "Deschide un PDF ca să-i organizezi paginile.": "Open a PDF to organize its pages.",
    "Rotește spre stânga": "Rotate left",
    "Rotește spre dreapta": "Rotate right",
    "Șterge paginile selectate (Delete)": "Delete the selected pages (Delete)",
    "Mută mai în față": "Move earlier",
    "Mută mai în spate": "Move later",
    "Pune o imagine (semnătură, ștampilă) pe pagina selectată":
        "Place an image (signature, stamp) on the selected page",
    "Salvează doar paginile selectate într-un PDF nou":
        "Save only the selected pages into a new PDF",
    "Revino la documentul original": "Go back to the original document",
    "Alege un PDF": "Choose a PDF",
    "Modificări nesalvate": "Unsaved changes",
    "Renunți la modificările făcute paginilor?": "Discard the changes made to the pages?",
    "Ștergere": "Delete",
    "Documentul trebuie să păstreze cel puțin o pagină.":
        "The document must keep at least one page.",
    "Semnătură": "Signature",
    "Selectează o singură pagină.": "Select a single page.",
    "Alege imaginea semnăturii sau a ștampilei": "Choose the signature or stamp image",
    "Salvează PDF-ul": "Save the PDF",
    "Salvare": "Save",
    "Salvează cu alt nume; originalul rămâne neatins.":
        "Save under another name; the original stays untouched.",
    "Salvează paginile selectate": "Save the selected pages",
    "Aplică": "Apply",
    "Semnătură…": "Signature…",
    "Extrage…": "Extract…",
    "Renunță": "Discard",
    "„{name}”: {n} pagini. Se generează miniaturile…":
        "“{name}”: {n} pages. Generating thumbnails…",
    "Deschide un PDF sau trage-l aici.\nPaginile se pot apoi roti, șterge, muta (trăgându-le),"
    " extrage sau semna.":
        "Open a PDF or drag it here.\nPages can then be rotated, deleted, moved (by dragging), "
        "extracted or signed.",
    "„{name}”: {n} pagini{changed}. Selectate: {s}.":
        "“{name}”: {n} pages{changed}. Selected: {s}.",
    "Am salvat {path}.": "Saved {path}.",
    "Am salvat {name}.": "Saved {name}.",
    "Trage cu mouse-ul dreptunghiul în care să apară imaginea.":
        "Drag with the mouse the rectangle where the image should appear.",
    "{name} (editat).pdf": "{name} (edited).pdf",
    "{name} (extras).pdf": "{name} (extract).pdf",
    # --- tab Unelte PDF ---------------------------------------------------------------------
    "Toate fișierele din listă devin un singur PDF, în ordinea din tabel (trage rândurile "
    "sau folosește ↑ ↓). Imaginile devin pagini.":
        "All files in the list become a single PDF, in table order (drag the rows or use "
        "↑ ↓). Images become pages.",
    "Fiecare parte devine un PDF separat, într-un folder nou lângă original.":
        "Each part becomes a separate PDF, in a new folder next to the original.",
    "Imaginile mari sunt micșorate și recomprimate; textul rămâne intact. Dacă rezultatul "
    "n-ar fi mai mic, se păstrează originalul.":
        "Large images are downscaled and recompressed; text stays intact. If the result "
        "would not be smaller, the original is kept.",
    "Fiecare pagină devine o imagine, într-un folder nou lângă original.":
        "Each page becomes an image, in a new folder next to the original.",
    "PDF-ul nu se poate deschide fără parolă (criptare AES-256). Parola nu se salvează.":
        "The PDF cannot be opened without the password (AES-256 encryption). The password "
        "is not saved.",
    "Textul apare în diagonală, semitransparent, pe fiecare pagină.":
        "The text appears diagonally, semi-transparent, on every page.",
    "PDF-urile scanate primesc un strat de text: devin căutabile și se pot copia.":
        "Scanned PDFs get a text layer: they become searchable and copyable.",
    "Unește": "Merge",
    "Împarte": "Split",
    "Comprimă": "Compress",
    "În imagini": "To images",
    "Parolă": "Password",
    "Watermark": "Watermark",
    "OCR": "OCR",
    "Ușor (calitate maximă)": "Light (best quality)",
    "Echilibrat": "Balanced",
    "Puternic (fișier minim)": "Strong (smallest file)",
    "Mută în sus": "Move up",
    "Mută în jos": "Move down",
    "Trage aici fișiere PDF (sau imagini, pentru unire).":
        "Drag PDF files here (or images, for merging).",
    "Execută": "Run",
    "Numerele de pagină tipărite în antet sau subsol („8/8”, „Pagina 1 din 2”) sunt "
    "înlocuite cu numerotarea documentului unit, în același font și loc.":
        "Page numbers printed in the header or footer (“8/8”, “Page 1 of 2”) are replaced "
        "with the merged document's numbering, in the same font and place.",
    "Fiecare pagină separat": "Every page separately",
    "Câte": "Every",
    "În": "Into",
    "Intervale:": "Ranges:",
    "Se oprește după fișierul curent…": "Stopping after the current file…",
    "Adaugă PDF-uri pentru a începe.": "Add PDFs to begin.",
    "Operație": "Operation",
    "Renumerotează paginile (ex. 8/8 și 1/2 devin 8/10 și 9/10)":
        "Renumber the pages (e.g. 8/8 and 1/2 become 8/10 and 9/10)",
    "Adaugă un PDF ca să vezi cum va fi împărțit.": "Add a PDF to see how it will be split.",
    "Alege PDF-uri sau imagini": "Choose PDFs or images",
    "Adaugă cel puțin un PDF.": "Add at least one PDF.",
    "Imagini în listă": "Images in the list",
    "Imaginile pot fi doar unite cu PDF-uri. Pentru a le transforma în PDF-uri separate, "
    "folosește tab-ul Conversie.":
        "Images can only be merged with PDFs. To turn them into separate PDFs, use the "
        "Conversion tab.",
    "Se lucrează…": "Working…",
    "fără previzualizare": "no preview",
    "„{name}” ({total} pagini) → {count} fișier(e): {ranges}":
        "“{name}” ({total} pages) → {count} file(s): {ranges}",
    "imagine": "image",
    "Unire": "Merge",
    "Ai nevoie de cel puțin două fișiere.": "You need at least two files.",
    "Salvează PDF-ul unit": "Save the merged PDF",
    "Alege un nume diferit de fișierele din listă.":
        "Choose a name different from the files in the list.",
    "Introdu o parolă.": "Enter a password.",
    "Parolele nu coincid.": "The passwords do not match.",
    "Scrie textul pentru watermark.": "Enter the watermark text.",
    "Nu am găsit numere de pagină tipărite de renumerotat.":
        "No printed page numbers found to renumber.",
    "Previzualizare": "Preview",
    "pagini într-un fișier": "pages per file",
    "părți egale": "equal parts",
    "ex. 1-200, 201-450, 451-  (gol după cratimă = până la sfârșit)":
        "e.g. 1-200, 201-450, 451-  (empty after the dash = to the end)",
    "Nivel:": "Level:",
    "Format:": "Format:",
    "Rezoluție:": "Resolution:",
    "Parolă:": "Password:",
    "Confirmă:": "Confirm:",
    "Text:": "Text:",
    "Opacitate:": "Opacity:",
    "paginile pe fișier": "the pages per file",
    "numărul de părți": "the number of parts",
    "„{name}” nu poate fi împărțit (nu e PDF, e protejat sau deteriorat).":
        "“{name}” cannot be split (not a PDF, protected or damaged).",
    "{n} fișier(e) în listă.": "{n} file(s) in the list.",
    "PDF și imagini": "PDF and images",
    "{before} → {after} (−{saved}%)": "{before} → {after} (−{saved}%)",
    "{n} imagini în „{folder}”": "{n} images in “{folder}”",
    "{n} fișiere în „{folder}”": "{n} files in “{folder}”",
    "Se unește…": "Merging…",
    "Se procesează…": "Processing…",
    "Limbi folosite: {langs}.": "Languages used: {langs}.",
    "{name} (unit).pdf": "{name} (merged).pdf",
    "Împărțire": "Split",
    "{name} (OCR).pdf": "{name} (OCR).pdf",
    "Am renumerotat {n} numere de pagină.": "Renumbered {n} page numbers.",
    "Scrie un număr întreg pentru {what}.": "Enter a whole number for {what}.",
    "{name} (protejat).pdf": "{name} (protected).pdf",
    "{name} (watermark).pdf": "{name} (watermark).pdf",
    "{name} (comprimat).pdf": "{name} (compressed).pdf",
    "{name} - imagini": "{name} - images",
    "{name} - pagini": "{name} - pages",
    "Gata → {name}": "Done → {name}",
    "Am creat {path}.": "Created {path}.",
    "Gata: {ok} reușite, {failed} cu erori.": "Done: {ok} succeeded, {failed} failed.",
    "Gata: {ok} reușite.": "Done: {ok} succeeded.",
    "parolă": "password",
    "OCR pagina {page}/{total}…": "OCR page {page}/{total}…",
    "Pagina {page}/{total}…": "Page {page}/{total}…",
    # --- tab Setari -------------------------------------------------------------------------
    "Mai multe fișiere selectate deodată ajung în aceeași fereastră. Pe Windows 11, opțiunile "
    "apar la „Afișați mai multe opțiuni” (Shift + F10).":
        "Several files selected at once arrive in the same window. On Windows 11 the options "
        "are under “Show more options” (Shift + F10).",
    "Aspect": "Appearance",
    "Actualizări": "Updates",
    "Verifică acum": "Check now",
    "Descarcă și instalează": "Download and install",
    "OCR (text din PDF-uri scanate)": "OCR (text from scanned PDFs)",
    "Descarcă limba română": "Download Romanian",
    "Integrare cu Windows Explorer": "Windows Explorer integration",
    "Componente detectate": "Detected components",
    "Jurnal": "Log",
    "Despre": "About",
    "Limbă": "Language",
    "Limba se schimbă după repornire. Repornesc aplicația acum?":
        "The language changes after a restart. Restart the application now?",
    "Verific…": "Checking…",
    "Ca sistemul": "Same as system",
    "Luminoasă": "Light",
    "Întunecată": "Dark",
    "Folder urmărit:": "Watched folder:",
    "Rezultate în:": "Results in:",
    "Rulezi din sursă: comanda va porni scriptul Python curent.":
        "Running from source: the command will start the current Python script.",
    "Word → PDF și PDF → DOCX prin Word": "Word → PDF and PDF → DOCX through Word",
    "Excel → PDF": "Excel → PDF",
    "PowerPoint → PDF": "PowerPoint → PDF",
    "alternativă pentru documente → PDF când lipsește Office":
        "alternative for documents → PDF when Office is missing",
    "necesar pentru OCR": "needed for OCR",
    "Alege folderul urmărit": "Choose the watched folder",
    "Alege folderul pentru rezultate": "Choose the folder for results",
    "Alege întâi folderul urmărit.": "Choose the watched folder first.",
    "Nu am putut verifica (fără internet sau depozit privat).":
        "Could not check (no internet or private repository).",
    "Actualizare disponibilă": "Update available",
    "Actualizare": "Update",
    "Rulezi din sursă; actualizează cu git.": "Running from source; update with git.",
    "Tesseract OCR nu este instalat.": "Tesseract OCR is not installed.",
    "Jurnalul este încă gol.": "The log is still empty.",
    "Deschide jurnalul": "Open the log",
    "Deschide folderul": "Open the folder",
    "Pornire cu Windows": "Start with Windows",
    "Versiune nouă: {v} (ai {current}).": "New version: {v} (you have {current}).",
    "Limbi disponibile: {langs}.": "Available languages: {langs}.",
    "Integrare Windows": "Windows integration",
    "Versiunea instalată: {v}.": "Installed version: {v}.",
    "Temă:": "Theme:",
    "Limbă:": "Language:",
    "Convertește automat fișierele puse în folderul de mai jos":
        "Automatically convert files placed in the folder below",
    "Pornește aplicația (minimizată) odată cu Windows":
        "Start the application (minimized) with Windows",
    "Folosește opțiunile din tab-ul Conversie. Funcționează cât timp aplicația e deschisă "
    "(și minimizată).":
        "Uses the options from the Conversion tab. Works while the application is open (even "
        "minimized).",
    "Verifică automat (cel mult o dată pe zi)": "Check automatically (at most once a day)",
    "„Convertește în PDF / DOCX” în meniul click-dreapta":
        "“Convert to PDF / DOCX” in the right-click menu",
    "Aplicația în „Trimite către”": "The application in “Send to”",
    "„Aspect fidel”, PDF → Excel, imaginile și uneltele PDF nu au nevoie de Office.":
        "“Exact layout”, PDF → Excel, images and the PDF tools don't need Office.",
    "Util când ceva nu merge: conține detaliile fiecărei erori.":
        "Useful when something goes wrong: it holds the details of every error.",
    "© 2026 Robert-Raul Mîrza · Licență MIT": "© 2026 Robert-Raul Mîrza · MIT License",
    "Scurtături: Ctrl+O adaugă fișiere, Ctrl+Shift+O adaugă un folder, Ctrl+Enter pornește, "
    "Esc anulează, Delete elimină, Ctrl+1…4 schimbă tab-ul, Ctrl+S salvează paginile.":
        "Shortcuts: Ctrl+O adds files, Ctrl+Shift+O adds a folder, Ctrl+Enter starts, Esc "
        "cancels, Delete removes, Ctrl+1…4 switches tabs, Ctrl+S saves the pages.",
    "Ai cea mai nouă versiune ({v}).": "You have the latest version ({v}).",
    "{app} {v} poate fi instalat din Setări.": "{app} {v} can be installed from Settings.",
    "Descarc versiunea {v} și pornesc installer-ul? Aplicația se va închide.":
        "Download version {v} and start the installer? The application will close.",
    "Descarc… {p}%": "Downloading… {p}%",
    "Descărcarea a eșuat: {error}": "The download failed: {error}",
}
