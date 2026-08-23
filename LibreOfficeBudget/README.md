# Budget fra CSV — LibreOffice-udvidelse

Laver et færdigt månedsbudget ud fra en CSV-fil fra banken. Posteringerne
læses, kategoriseres automatisk og skrives ind i et regneark, der følger det
klassiske "Monthly Budget"-opsæt (Oversigt + Transaktioner) — i danske kroner.

Alt sker i to klik: **Budget ▸ Importér CSV-fil til budget…**, bekræft de
kolonner udvidelsen har fundet, og budgettet er der.

---

## Indhold

| Mappe/fil | Hvad det er |
|---|---|
| `oxt/` | Selve udvidelsen (menu, dialoger, registrering) |
| `budget_core/` | Motoren: CSV-genkendelse, tal/dato-parsing, kategorier, budget. Ren Python, ingen UNO |
| `office/budget_office.py` | Skriver regnearket via UNO (ark, formler, formater) |
| `budget_cli.py` | Samme motor som kommandolinjeværktøj (afløser for det oprindelige script) |
| `build_oxt.py` | Bygger `dist/budget-fra-csv-1.0.0.oxt` |
| `tests/` | Enhedstest (uden LibreOffice) + en ende-til-ende-test (med LibreOffice) |

## Installation

```bash
python3 build_oxt.py            # -> dist/budget-fra-csv-1.0.0.oxt
python3 build_oxt.py --install  # bygger og installerer med unopkg
```

Alternativt: dobbeltklik på `.oxt`-filen, eller **Funktioner ▸ Extension
Manager ▸ Tilføj**. Genstart LibreOffice bagefter — så ligger menuen
**Budget** i menulinjen.

* **Windows/macOS:** virker uden videre (LibreOffice har Python indbygget).
* **Linux:** kræver Python-understøttelsen: `sudo apt install python3-uno`.

## Sådan bruges den

1. **Budget ▸ Importér CSV-fil til budget…** — vælg filen fra netbanken.
2. Dialogen viser, hvad der er genkendt (dato-, tekst- og beløbskolonne,
   decimaltegn, datoformat). Ret det, hvis noget er gættet forkert, og tryk
   **Opret budget**.
3. Der åbnes et nyt regneark med:

   | Ark | Indhold |
   |---|---|
   | **Oversigt** (ét pr. måned) | Budget / Faktisk / Diff. pr. kategori, startsaldo, slutsaldo og nøgletal — samme opsæt som skabelonen, bare i kroner |
   | **Transaktioner** | Udgifter i kolonne B:E, indtægter i G:J (Dato, Beløb, Beskrivelse, Kategori) |
   | **Alle måneder** | Kategorier × måneder med totaler og gennemsnit (kun når filen dækker flere måneder) |
   | **Kategorier** | Nøgleordene, der styrer den automatiske kategorisering |

4. Ret en kategori i **Transaktioner** (der er en rulleliste i
   kategori-kolonnen) — oversigten opdaterer sig selv, fordi "Faktisk" er
   `SUMIF`/`SUMIFS`-formler.
5. Skriv dine egne budgettal i den lyserøde **Budget**-kolonne. Har filen
   flere måneder, er kolonnen udfyldt med gennemsnittet pr. måned som forslag.

### Når reglerne skal rettes

Tilføj linjer i arket **Kategorier** (nøgleord + kategori) og vælg
**Budget ▸ Opdatér kategorier og budget**. Så køres alle posteringer igennem
igen, og oversigterne bygges op på ny — dine budgettal og din startsaldo
bevares. Reglerne gemmes samtidig i din brugerprofil
(`~/.config/budget-fra-csv/kategoriregler.csv`, på Windows
`%APPDATA%\BudgetFraCSV\`), så de bruges næste gang du importerer.

* Et nøgleord matcher, når det indgår i posteringsteksten (uden hensyn til
  store/små bogstaver og accenter). Korte nøgleord (≤ 4 tegn) matcher kun
  hele ord — så "Irma" ikke fanger "firma".
* Længste nøgleord vinder, så `Circle K` slår `K`.
* Skriv `re:` foran for et regulært udtryk, fx `re:faktura\s*\d+`.
* En kategori, du selv har valgt i rullelisten, bliver ikke overskrevet af en
  regel, der alligevel ikke rammer noget.

### Andre kommandoer i menuen

* **Lav budget ud fra det aktive ark** — samme funktion, men på et ark der
  allerede er åbent. Brug den til `.xlsx`/`.ods`-filer: åbn dem i Calc først.
* **Vis kategoriregler** / **Nulstil kategoriregler**.

## Hvilke CSV-filer virker?

Genkendelsen prøver at klare det, netbanker rent faktisk eksporterer:

* **Tegnsæt:** UTF-8 (med og uden BOM), UTF-16/UTF-32, Windows-1252, Latin-1.
* **Skilletegn:** `;`, `,`, TAB, `|` — inkl. Excels `sep=;`-linje.
* **Overskrifter:** genkendes automatisk, også når der står et par
  informationslinjer over tabellen (kontonummer, periode m.m.). Filer helt
  uden overskriftsrække virker også.
* **Tal:** `1.234,56`, `1,234.56`, `1 234,56`, `1'234.56`, `kr. 99,50`,
  `$1,234.56`, `(120,45)`, `198,50-`, `1.234,56 DR`, minustegn som `−`.
* **Datoer:** `31-05-2025`, `31/05/25`, `31.05.2025`, `2025-05-31`,
  `20250531`, `31052025`, `31. maj 2025`, `May 31, 2025`, med eller uden
  klokkeslæt. Dag/måned-rækkefølgen afgøres ud fra hele kolonnen.
* **Beløbskolonner:** én kolonne med fortegn, eller to kolonner
  (`Hævet`/`Indsat`, `Debit`/`Credit`). En `Saldo`-kolonne genkendes og
  bruges som startsaldo i stedet for at blive forvekslet med beløbet.
* **Tekstkolonner:** flere tekstkolonner (fx `Navn` + `Titel`) slås sammen,
  så kategorireglerne har mere at arbejde med.
* Tomme linjer, ujævne rækker og "I alt"-linjer i bunden springes over — og
  tælles med i den kvittering, du får til sidst.

Er noget gættet forkert, kan alt rettes i importdialogen.

## Kommandolinje

Samme motor uden LibreOffice:

```bash
python3 budget_cli.py kontoudtog.csv                 # rapport + to CSV-filer
python3 budget_cli.py kontoudtog.csv --kun-rapport   # kun rapporten
python3 budget_cli.py kontoudtog.csv --interaktiv --gem-regler
python3 budget_cli.py kontoudtog.csv --decimal , --datoformat dmy --fortegn positiv
```

Der skrives `<navn>_kategoriseret.csv` (med kategori, type og måned) og
`<navn>_budget.csv` (kategorier × måneder).

## Test

```bash
python3 -m unittest discover -s tests      # 38 test, ingen LibreOffice nødvendig

# ende-til-ende mod en kørende LibreOffice:
soffice --headless --norestore --accept="socket,host=localhost,port=2002;urp;" &
python3 tests/office_smoke.py
```

`tests/data/` indeholder eksempelfiler i otte forskellige CSV-varianter
(Danske Bank-stil, Nordea med indledning, engelsk/amerikansk, hævet/indsat,
UTF-16 med tabulator, uden overskrifter, `sep=`-linje og en rodet fil).

## Kendte begrænsninger

* Oversigtsarkene bygges helt om, når du vælger **Opdatér kategorier og
  budget**. Budgettal, startsaldo og kategorinavne bevares — andre egne
  ændringer på oversigtsarket gør ikke.
* Formlerne kigger på række 5–5000 i `Transaktioner`; ved flere end ca. 5000
  posteringer skal området udvides (`TX_MAX_ROW` i `office/budget_office.py`).
* Beløb behandles i kroner. Er der en valutakolonne med andre valutaer, læses
  beløbene som de står — der omregnes ikke.
