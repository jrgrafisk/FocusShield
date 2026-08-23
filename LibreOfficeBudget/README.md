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
| `budget_core/` | Motoren: CSV-genkendelse, tal/dato-parsing, kategorier, budget og budgetforslag. Ren Python, ingen UNO |
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
   | **Budgetforslag** | Et udkast til et fast månedsbudget ud fra hele perioden — faste, variable og periodiske udgifter hver for sig, med en Mål-kolonne du selv kan skrive i |
   | **Prognose** | Saldoen 24 måneder frem, hvis alt fortsætter som nu — og hvis du rammer dine mål. Med graf |
   | **Oversigt** (ét pr. måned) | Budget / Faktisk / Diff. pr. kategori, startsaldo, slutsaldo og nøgletal — samme opsæt som skabelonen, bare i kroner |
   | **Transaktioner** | Udgifter i kolonne B:E, indtægter i G:J (Dato, Beløb, Beskrivelse, Kategori) |
   | **Alle måneder** | Kategorier × måneder med totaler og gennemsnit (kun når filen dækker flere måneder) |
   | **Kategorier** | Nøgleordene, der styrer den automatiske kategorisering |

   *Budgetforslag og Prognose laves kun, når filen dækker mere end én måned.*

4. Ret en kategori i **Transaktioner** (der er en rulleliste i
   kategori-kolonnen) — oversigten opdaterer sig selv, fordi "Faktisk" er
   `SUMIF`/`SUMIFS`-formler.
5. Skriv dine egne budgettal i den lyserøde **Budget**-kolonne. Har filen
   flere måneder, er kolonnen udfyldt med gennemsnittet pr. måned som forslag.

## Budgetforslag ud fra et helt år

Giver du den en CSV med hele året, analyserer den hver kategori måned for
måned og deler udgifterne i tre:

| Gruppe | Hvornår | Forslaget bliver |
|---|---|---|
| **Faste udgifter** | Samme beløb næsten hver måned (husleje, forsikring, lån, telefon, abonnementer) | Gennemsnittet, rundet op |
| **Variable udgifter** | Der hver måned, men beløbet svinger (dagligvarer, restaurant, transport, shopping) | Den **typiske** måned (median) — så én dyr december ikke sætter dagligvarebudgettet |
| **Periodiske udgifter** | Få gange om året (el-afregning hvert kvartal, ferie, jul) | Årets samlede beløb delt ud på 12 — altså det du skal henlægge hver måned |

Indtægten budgetteres forsigtigt: den typiske måned, ikke måneden med
bonussen. Kvartalsvise indtægter (fx børneydelse) fordeles på månederne.

Tabellen har seks talkolonner:

`Gns./md` · `Typisk md` · `Laveste` · `Højeste` · **`Mål`** (din egen, lyserød) · `Forskel`

**Forskel = Mål − Gns./md.** Sætter du dagligvarer til 4.000 kr. under dit
nuværende forbrug, står der `-4.000 kr.` — så du kan se præcis hvad du skal
finde. Nederst summeres det hele til **Til opsparing**, og du kan skrive et
**opsparingsmål pr. måned**; linjen under viser, om målet hænger sammen.

## Prognose

Arket **Prognose** fremskriver saldoen 24 måneder ud fra din nuværende
startsaldo (som du kan rette i den lyserøde celle), med to linjer i grafen:

* **Saldo nu** — hvis alt fortsætter præcis som de sidste måneder.
* **Saldo m. mål** — hvis du rammer tallene i Mål-kolonnen.

Begge kolonner er formler, der peger på Budgetforslag, så grafen flytter sig,
i samme øjeblik du ændrer et mål.

## Kategorisering

Tre måder, alt efter hvor meget du vil gøre:

1. **Rullelisten i Transaktioner** — vælg kategori direkte i kolonnen
   Kategori. Oversigterne opdaterer sig selv.
2. **`Budget ▸ Kategorisér poster uden kategori…`** — en liste over alle
   ukendte posteringer (samlet pr. forretning, med antal og beløb). Marker en
   eller flere, vælg eller skriv en kategori, tryk **Tildel**. Sæt flueben i
   "Gem som regel", så kender den forretningen næste gang.
3. **`Budget ▸ Kategorier - opret, omdøb, slet…`** — opret en ny kategori,
   omdøb en (posteringer, regler og lister følger med) eller slet en
   (posteringerne bliver ukategoriserede igen, så du kan sætte dem et andet
   sted hen). Alle ark bygges om bagefter.

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
* **Kategorisér poster uden kategori…** og **Kategorier - opret, omdøb, slet…**
  (se ovenfor).
* **Opdatér kategorier og budget** — kører reglerne igennem igen og bygger
  alle oversigter, budgetforslaget og prognosen om.
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
python3 -m unittest discover -s tests      # 51 test, ingen LibreOffice nødvendig

# ende-til-ende mod en kørende LibreOffice:
soffice --headless --norestore --accept="socket,host=localhost,port=2002;urp;" &
python3 tests/office_smoke.py
```

`tests/data/` indeholder eksempelfiler i ni forskellige CSV-varianter
(Danske Bank-stil, Nordea med indledning, engelsk/amerikansk, hævet/indsat,
UTF-16 med tabulator, uden overskrifter, `sep=`-linje, en rodet fil — og
`aar_2024.csv` med et helt års posteringer, som budgetforslaget testes på).

## Kendte begrænsninger

* Oversigtsarkene, budgetforslaget og prognosen bygges helt om, når du vælger
  **Opdatér kategorier og budget** (og efter kategoriændringer). Budgettal,
  mål, opsparingsmål og startsaldoer bevares — andre egne ændringer på de ark
  gør ikke.
* Tallene i Budgetforslag er et øjebliksbillede af den importerede periode
  (ikke formler). De genberegnes, når du opdaterer.
* Formlerne kigger på række 5–5000 i `Transaktioner`; ved flere end ca. 5000
  posteringer skal området udvides (`TX_MAX_ROW` i `office/budget_office.py`).
* Beløb behandles i kroner. Er der en valutakolonne med andre valutaer, læses
  beløbene som de står — der omregnes ikke.
