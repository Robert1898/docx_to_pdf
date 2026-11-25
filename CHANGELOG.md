# Istoricul versiunilor

Toate schimbările importante ale aplicației sunt notate aici. Formatul urmează
[Keep a Changelog](https://keepachangelog.com/ro/1.1.0/), iar versiunile respectă
[Semantic Versioning](https://semver.org/lang/ro/).

## [2.4.0] - 2026-11-25

### Securitate
- Actualizările sunt verificate înainte de instalare: installer-ul descărcat trebuie să aibă
  suma SHA-256 din `SHA256SUMS.txt`, publicat în același release. Altfel este șters și nu
  se instalează.
- Descărcările se fac numai prin https, numai de pe domeniile GitHub, inclusiv după
  redirecționări.
- Fiecare release are atașat un SBOM CycloneDX (`*.cdx.json`), adică lista exactă a
  bibliotecilor incluse în aplicație.

### Adăugat
- Folderul urmărit reîncearcă fișierele eșuate de cel mult 3 ori, cu pauze tot mai lungi.
  Ca făcute sunt marcate doar fișierele convertite cu succes. Starea veche
  (`watch_done.json`) este preluată automat.
- La dialogurile ascunse Office se apasă doar butoane dintr-o listă albă: Da, OK, Deschide,
  Doar citire. Un dialog fără astfel de butoane e notat în jurnal și nu se apasă nimic.

### Reparat
- Excel pornit de aplicație nu mai rămâne în memorie după conversie.
- Instanța Office pornită de aplicație este recunoscută după fereastra ei (Word, Excel), nu
  după lista proceselor. Aplicația nu mai poate opri, din greșeală, un Office pornit de
  utilizator în același moment.
- Prezentările PowerPoint și registrele Excel deschise de utilizator în instanța folosită de
  aplicație nu mai sunt închise:
  - fără `Quit` și fără oprire forțată, nici la timeout, nici la Anulare;
  - fișierul aflat în conversie este marcat eșuat, iar Office rămâne pornit.
- La timeout sau la Anulare, un apel Office blocat nu mai ține pe loc conversiile următoare.
- PowerPoint se închide la sfârșitul fiecărui lot, nu după 5 minute de inactivitate.

### Eliminat
- Setarea globală Excel `IgnoreRemoteRequests` nu mai este modificată. Pe Office 16, Explorer
  deschide oricum registrele într-un proces separat, iar setarea putea rămâne salvată.

### Schimbat
- Necesită Python 3.11 sau mai nou.
- Codul pentru Office este împărțit în pachetul `convertor.office`.
- CI:
  - teste pe Python 3.11–3.14, cu acoperire;
  - verificare de tipuri cu pyright;
  - acțiuni GitHub fixate pe commit;
  - Dependabot pentru acțiuni și dependențe.
- Release-ul se construiește cu dependențe fixate și verificate prin hash
  (`requirements.lock`) și cu Inno Setup la o versiune fixă.

## [2.3.0] - 2026-11-12

### Adăugat
- PDF → Excel, compresie PDF, imagini ↔ PDF.
- Tab nou „Pagini”: miniaturi, rotire, ștergere, mutare, extragere, semnătură sau ștampilă.
- Folder urmărit, pornire cu Windows, descărcarea limbii române pentru OCR.
- Previzualizare, progres cu timp rămas în bara de activități, scurtături, iconițe.
- Interfață în engleză (implicit: română).
- Conversie în paralel, cu Office păstrat pornit între fișiere.
- O singură fereastră pentru mai multe fișiere selectate în Explorer.
- Actualizări automate din GitHub Releases.
- CI și release cu GitHub Actions.

### Schimbat
- pycryptodome în loc de cryptography.
- Bibliotecile sunt încărcate abia când e nevoie de ele, deci installer-ul e mai mic.

### Reparat
- Setarea Word din registry nu era refăcută după un timeout.

[Nelansat]: https://github.com/Robert1898/docx_to_pdf/compare/v2.3.0...HEAD
[2.3.0]: https://github.com/Robert1898/docx_to_pdf/compare/v2.2.1...v2.3.0
