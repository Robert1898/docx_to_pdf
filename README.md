# Document Converter

Aplicație Windows care convertește în masă documente Office și imagini în PDF, iar PDF-uri în
DOCX sau Excel, plus unelte PDF (unire, împărțire, compresie, organizare pagini, semnătură, OCR).
Totul rulează **local, offline**: documentele nu pleacă de pe calculator (singura conexiune la
internet este verificarea opțională a actualizărilor și descărcarea limbii române pentru OCR).
Interfața este în română sau engleză.

## Funcții

**Conversie** (direcția se alege automat după tipul fișierului)

| Din | În | Cu |
|---|---|---|
| Word: `.docx` `.doc` `.docm` `.rtf` `.odt` | PDF | Word sau LibreOffice |
| Excel: `.xlsx` `.xls` `.xlsm` `.ods` | PDF | Excel sau LibreOffice |
| PowerPoint: `.pptx` `.ppt` `.pptm` `.odp` | PDF | PowerPoint sau LibreOffice |
| PDF | DOCX | „Aspect fidel” sau Word (opțional cu OCR pentru PDF-uri scanate) |
| PDF | Excel (`.xlsx`) | tabelele din rapoarte, cu numere și date ca valori reale (fără Office) |
| Imagini: `.jpg` `.png` `.bmp` `.gif` `.tif` `.webp` | PDF | încadrate pe A4 sau la mărimea imaginii (fără Office) |

**PDF → DOCX** are trei moduri (opțiunea „PDF → DOCX” din tab-ul Conversie, `--pdf-mode`):
- **Aspect fidel** (rapoarte, balanțe, jurnale, extrase, tabele): fiecare pagină arată ca în PDF.
  Rândurile și coloanele sunt refăcute din poziția exactă a fiecărui caracter, sumele sunt
  aliniate la dreapta, iar chenarele și liniile drepte sunt păstrate. Textul rămâne editabil, iar
  rândurile copiate în Excel se împart singure pe coloane. Nu are nevoie de Word. Textul ascuns
  în PDF (de ex. depășirea unei celule din Excel) rămâne ascuns. Imaginile (logo-uri, ștampile,
  poze) sunt puse exact în locul lor; grafica vectorială complexă (grafice) nu este transpusă.
- **Text editabil (Word)**: „PDF Reflow” din Word, potrivit pentru text curgător (scrisori,
  contracte). Pe rapoarte cu multe coloane Word ghicește tabelele și poate muta valori pe alte
  rânduri.
- **Automat** (implicit): rapoartele pe coloane merg pe aspect fidel, textul curgător și PDF-urile
  scanate prin Word.

**PDF → Excel** (opțiunea „PDF în: Excel”, `--pdf-to xlsx`): pentru rapoarte, jurnale,
balanțe și extrase. Coloanele sunt găsite după poziția textului pe pagină, antetul repetat pe
fiecare pagină apare o singură dată, iar textul rupt pe două rânduri e lipit la loc. Sumele
(`1,234.56`, `1.234,56`, `(1,000.00)` negativ) devin numere, datele (`15.09.2026`, `15-SEP-2026`)
devin date Excel; codurile cu zero în față și numerele foarte lungi (CNP, IBAN) rămân text.

**Imagini → PDF**: fiecare imagine devine un PDF (TIFF-urile cu mai multe pagini devin un PDF cu
mai multe pagini). Pentru un singur PDF din mai multe imagini, folosește **Unelte PDF → Unește**.

**Conversii în paralel**: „Conversii simultane” (1–4, `--jobs`) pornește mai multe instanțe Word
sau Excel pentru loturile mari. Pentru câteva fișiere se folosește una singură, iar Word rămâne
pornit câteva minute după conversie, ca lotul următor să înceapă imediat. PowerPoint lucrează
mereu pe rând (Windows permite o singură instanță).

- Opțiuni PDF: semne de carte din titluri, **PDF/A** pentru arhivare, calitate pentru ecran (fișier mai mic)
- **Excel fără pagini goale**: se tipărește doar conținutul real al fiecărei foi (valori, grafice,
  imagini, celule unite), nu toată zona „folosită” umflată de formatări sau formule care întorc
  gol. O celulă uitată departe de date ajunge pe pagina ei, fără miile de rânduri goale dintre ele.
- **Mai multe foi**: toate foile vizibile ajung în același PDF, în ordinea tab-urilor, fiecare
  începând pe o pagină nouă; foile ascunse și cele goale sunt sărite.
- **Tabele late**: ce nu încape pe portret trece pe peisaj. Dacă încape pe lățime la cel puțin 60%,
  iese pe o singură pagină în lățime. Altfel rămâne la 60%, paginile merg întâi spre dreapta, apoi
  în jos, iar antetul și coloana de identificare (de ex. CNP, nume, cod) se repetă pe fiecare
  pagină, ca fiecare rând să rămână de recunoscut.
- Zonele de tipărire, scalarea și rândurile/coloanele repetate setate de tine în Excel sunt
  respectate.
- Fișiere individuale, foldere întregi (recursiv) sau **drag & drop** din Explorer; ordinea din
  listă se schimbă trăgând rândurile
- Tabel cu status pe fiecare fișier, progres și timp estimat (și pe bara de activități);
  dublu-click deschide rezultatul, click-dreapta oferă mai multe opțiuni
- Rezultatul se salvează lângă original sau într-un folder ales; nu suprascrie nimic fără să întrebe
- **Dialogurile ascunse ale lui Office** (de ex. „Ultima dată când ați deschis fișierul a cauzat o
  eroare gravă. Îl deschideți oricum?”) primesc răspuns automat, în loc să blocheze conversia;
  textul lor apare în jurnal
- Buton **Anulează** și **limită de timp per fișier**: dacă Office tot se blochează, instanța
  pornită de aplicație e oprită și se trece la următorul fișier (o instanță deschisă de tine nu
  este oprită niciodată)
- Documentele protejate cu parolă dau eroare imediat, în loc să blocheze conversia
- Macro-urile din documente nu rulează în timpul conversiei

**Unelte PDF**: unire (în ordinea aleasă, inclusiv cu imagini), împărțire, **compresie**,
**PDF în imagini** (PNG/JPG, 72–300 DPI), protejare cu parolă (AES-256), watermark text și **OCR**
(transformă PDF-urile scanate în PDF-uri căutabile). În dreapta listei apare o previzualizare a
primei pagini din fișierul selectat.

**Compresia** micșorează imaginile mari și le recomprimă (trei niveluri: ușor, echilibrat,
puternic); textul și fonturile rămân intacte. Dacă rezultatul n-ar fi mai mic, se păstrează
originalul. Ajută mult la PDF-urile scanate sau cu poze, foarte puțin la cele cu text.

Împărțirea are patru moduri: fiecare pagină separat, **câte N pagini** într-un fișier (600 de
pagini câte 200 → 3 fișiere), **în N părți egale** sau **intervale alese** (`1-200, 201-450, 451-`;
un interval fără sfârșit merge până la ultima pagină). Sub opțiuni apare o previzualizare cu
fișierele care vor rezulta, iar o greșeală de scriere e semnalată înainte de a porni.

La unire, numerele de pagină tipărite în antet sau subsol pot fi **renumerotate**: un PDF de 8
pagini („8/8”) urmat de unul de 2 („1/2”) devine un document cu „8/10”, „9/10”, „10/10”. Numărul
vechi e acoperit cu un petic de culoarea fundalului, iar cel nou e scris în același loc, cu același
font, mărime și culoare. Ca să nu fie atinse alte cifre, un număr e înlocuit doar dacă se potrivește
cu pagina (pagina 8 din 8), stă în antet/subsol și e la capătul rândului lui. Formate recunoscute:
`8/8`, `8 / 8`, `Pagina 8 din 8`, `Page 8 of 8`, `Pagina 8` și un număr singur în subsol. Numărul
vechi rămâne în stratul de text de sub petic, deci o căutare în PDF îl mai poate găsi.

**Pagini** (organizator): deschizi un PDF și vezi miniaturile tuturor paginilor. Paginile se pot
**roti**, **șterge**, **muta** (trăgându-le sau cu săgețile), **extrage** într-un PDF nou și se
poate pune o **semnătură sau ștampilă** (o imagine PNG, ideal cu fundal transparent) pe o pagină:
dublu-click pe pagină, apoi tragi dreptunghiul unde să apară. Rezultatul se salvează cu alt nume
(Ctrl+S); originalul rămâne neatins.

**Folder urmărit** (Setări): fișierele puse într-un folder ales sunt convertite automat, cu
opțiunile din tab-ul Conversie, în subfolderul `Convertite` (sau altul ales). Un fișier e preluat
abia după ce s-a terminat de copiat, iar unul deja convertit nu e reluat. Merge cât timp aplicația
e deschisă, și minimizată; opțiunea „Pornește odată cu Windows” o deschide minimizată la logare.

**Actualizări**: aplicația verifică cel mult o dată pe zi dacă există o versiune nouă în
[GitHub Releases](https://github.com/Robert1898/docx_to_pdf/releases) și o poate descărca și
instala din Setări. Verificarea se poate opri. Funcționează doar dacă depozitul și release-urile
sunt publice.

**Scurtături**: Ctrl+O adaugă fișiere (sau deschide un PDF în Pagini), Ctrl+Shift+O adaugă un
folder, Ctrl+Enter pornește, Esc anulează, Delete elimină, Ctrl+1…4 schimbă tab-ul, Ctrl+S
salvează paginile.

**Altele**: interfață în română (implicit) sau engleză, ori după limba Windows, temă luminoasă /
întunecată / ca sistemul, afișare clară pe ecrane cu scalare (125–200%), setări memorate între
porniri, notificări la final, jurnal pentru depanare, **meniu click-dreapta** și **„Trimite
către”** în Explorer (mai multe fișiere selectate deodată ajung în aceeași fereastră, nu câte una
pentru fiecare), plus o **linie de comandă**.

## Instalare

### Installer (recomandat)
Rulează `Document-Converter-Setup-<versiune>.exe`. Implicit se instalează doar pentru
utilizatorul curent, fără drepturi de administrator. Installer-ul poate adăuga și integrarea cu
Explorer (meniul click-dreapta și „Trimite către”) și o scurtătură pe desktop.

Executabilul nu este semnat digital, așa că Windows SmartScreen poate afișa un avertisment
(„Mai multe informații” → „Rulează oricum”).

### Actualizare de la „Convertor DOCX-PDF” (până la v2.0)
Aplicația s-a numit înainte „Convertor DOCX-PDF”. Installer-ul nou dezinstalează automat
versiunea veche, iar la prima pornire setările și integrarea cu Explorer (meniul click-dreapta,
„Trimite către”) sunt mutate pe numele nou.

### Portabil
Copiază folderul `Document Converter` (din `dist/`) oriunde și pornește `Document Converter.exe`.
Integrarea cu Explorer se activează din tab-ul **Setări**.

### Ce trebuie să fie instalat pe calculator
- **Microsoft Office** (Word, Excel, PowerPoint): folosit prioritar. Pentru PDF → DOCX e nevoie de
  Word 2013 sau mai nou.
- sau **[LibreOffice](https://www.libreoffice.org/)**: alternativă pentru documente → PDF când Office lipsește.
- opțional **[Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)** pentru OCR.
  Pentru diacritice românești, bifează limba *Romanian* la instalarea Tesseract (sau copiază
  `ron.traineddata` în `C:\Program Files\Tesseract-OCR\tessdata`), ori apasă **Setări → OCR →
  Descarcă limba română** (fără drepturi de administrator). Aplicația folosește automat `ron+eng`
  când sunt disponibile.

Tab-ul **Setări → Componente detectate** arată ce a găsit aplicația.

## Utilizare

### Interfața grafică
Trage fișiere sau foldere în fereastră (sau folosește butoanele), alege opțiunile și apasă
**Convertește**.

### Din Explorer
- **Click-dreapta** pe un document → „Convertește în PDF” / „Convertește în DOCX”.
  Pe Windows 11 opțiunea apare la „Afișați mai multe opțiuni” (Shift + F10).
- Pentru mai multe fișiere deodată: selectează-le → click-dreapta → **Trimite către** →
  Document Converter.

### Linia de comandă
```bash
docconv "D:\Documente" --out "D:\PDF"
```
```bash
docconv raport.docx buget.xlsx scan.pdf --ocr --pdfa
```
```bash
docconv jurnal.pdf --pdf-to xlsx
```
```bash
docconv --watch "D:\De convertit" --jobs 3
```

| Opțiune | Efect |
|---|---|
| `-o, --out DIR` | folder destinație (implicit lângă original) |
| `--overwrite` | suprascrie rezultatele existente (implicit sunt omise) |
| `--pdfa` | PDF/A pentru arhivare |
| `--screen` | optimizat pentru ecran (fișier mai mic) |
| `--no-bookmarks` | fără semne de carte din titluri |
| `--ocr` | OCR pe PDF-urile scanate înainte de conversia în DOCX |
| `--pdf-to FORMAT` | în ce se convertesc PDF-urile: `docx` (implicit) sau `xlsx` |
| `--pdf-mode MOD` | PDF → DOCX: `auto` (implicit), `layout` (aspect fidel), `word` |
| `--image-page MOD` | imagine → PDF: `a4` (implicit) sau `image` (pagina de mărimea imaginii) |
| `-j, --jobs N` | conversii Office simultane, 1–4 (implicit 2) |
| `--timeout SEC` | limită de timp per fișier (implicit 180) |
| `--watch FOLDER` | urmărește folderul și convertește fișierele noi până la Ctrl+C (rezultate în `--out` sau `FOLDER\Convertite`) |
| `--register-shell` / `--register-sendto` / `--unregister-shell` | integrarea cu Explorer (`--unregister-shell` oprește și pornirea cu Windows) |

Cod de ieșire: `0` totul a mers, `1` unele fișiere au eșuat, `2` argumente greșite sau niciun fișier găsit.

## Depanare
- **Setări → Jurnal** deschide jurnalul (`%LOCALAPPDATA%\DocumentConverter\logs`), cu detaliile fiecărei erori.
- Setările sunt în `%APPDATA%\DocumentConverter\settings.json`.
- În timpul conversiei PDF → DOCX, aplicația dezactivează temporar mesajul lui Word „Word va converti
  PDF-ul…” (`DisableConvertPdfWarning`) și îl readuce la valoarea inițială la final, chiar și după
  o închidere bruscă (la următoarea pornire).
- Cât timp un fișier se convertește, aplicația urmărește ferestrele instanței Office pornite de
  ea. Dacă apare un dialog, îl citește (UI Automation), îl notează în jurnal și apasă primul buton
  („Da” / „OK” / „Doar citire”). Fișierele sunt deschise doar pentru citire și nu sunt salvate
  niciodată, așa că răspunsul afirmativ e sigur.
- Dacă Word trebuie totuși oprit forțat, el marchează fișierul ca „periculos” (în
  `Resiliency\DisabledItems`). Aplicația șterge marcajele vizibile în registry, iar dialogul pe care
  Word l-ar afișa la redeschiderea fișierului primește răspuns automat.

## Dezvoltare

```bash
python -m venv .venv
```
```bash
.venv\Scripts\pip install -r requirements-dev.txt
```
```bash
.venv\Scripts\python run_gui.py
```

### Teste și verificări
```bash
.venv\Scripts\python -m ruff check .
```
```bash
.venv\Scripts\python -m pytest -m "not office and not ocr"
```
```bash
.venv\Scripts\python -m pytest -m office
```
Testele marcate `office` pornesc Microsoft Office real și durează mai mult; testul de OCR rulează
doar dacă Tesseract e instalat.

### Construire
```bash
powershell -ExecutionPolicy Bypass -File build.ps1
```
Rulează ruff și testele rapide, apoi PyInstaller (`document-converter.spec`, varianta *onedir* cu două
executabile: interfața și `docconv.exe`) și Inno Setup (`installer\document-converter.iss`).
Rezultate în `dist\`. Opțiuni: `-SkipTests`, `-NoInstaller`. Inno Setup 6 trebuie să fie instalat.

**Semnare digitală** (opțional): setează `SIGN_CERT_THUMBPRINT` (certificat din Windows
Certificate Store) sau `SIGN_PFX` + `SIGN_PFX_PASSWORD` înainte de `build.ps1`; ambele executabile
și installer-ul se semnează cu `signtool` (Windows SDK) și marcaj de timp. Fără ele, pasul e sărit.

Versiunea se schimbă într-un singur loc: `convertor/__init__.py`.

### Structura
```
convertor/
  core.py         tipuri de fișiere, opțiuni, nume pentru rezultate
  engine.py       rulează un lot de conversii (folosit de GUI și CLI)
  office.py       Word / Excel / PowerPoint prin COM, limită de timp, curățenie în registry
  libreoffice.py  alternativa fără Office
  pdflayout.py    PDF → DOCX cu aspect fidel (rânduri, coloane, chenare, imagini)
  pdftable.py     PDF → Excel (coloane, numere, date)
  excelarea.py    ce se tipărește din fiecare foaie Excel
  pagenumbers.py  renumerotarea paginilor la unire
  pdftools.py     unire, împărțire, compresie, imagini, parolă, watermark, OCR
  pageedit.py     organizatorul de pagini și semnătura
  shell.py        meniul click-dreapta, „Trimite către”, pornirea cu Windows
  watch.py        folderul urmărit
  updater.py      actualizări din GitHub Releases
  i18n.py         traducerea (textele englezești sunt în i18n_en.py)
  settings.py     setări memorate
  cli.py          linia de comandă
  gui/            interfața (ttkbootstrap + tkinterdnd2), instanță unică, progres în taskbar
tests/            pytest
installer/        script Inno Setup + traducerea română a installer-ului
```

## Licență
MIT, vezi fișierul [LICENSE](LICENSE). Traducerea română a installer-ului (`installer/Romanian.isl`)
provine din depozitul oficial Inno Setup.
