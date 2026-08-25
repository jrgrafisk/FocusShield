# Budget fra CSV - til Google Sheets

Et lille script til Google Sheets, der laver et færdigt budget ud fra et
bank-CSV-udtræk direkte i dit regneark - **ingen installation på
computeren, ingen kommandoprompt, ingen Python, intet program at
downloade.** Du indsætter scriptet i et Google Sheets-ark én gang, og
derefter foregår alt inde i browseren.

Det er den samme motor (kategorisering, budgetklassificering osv.), som
ligger bag LibreOffice- og Excel-udgaverne af "Budget fra CSV" - bare
skrevet om til Google Apps Script (JavaScript), fordi Sheets ikke kan køre
Python eller VBA.

## Sådan installeres det (én gang, i browseren)

1. Opret et nyt, tomt Google Sheets-ark (sheets.new).
2. Gå til **Udvidelser (Extensions) → Apps Script**. Der åbner en
   kodeeditor i en ny fane.
3. Der ligger en fil `Code.gs` med lidt forudfyldt eksempelkode - marker alt
   indholdet (Ctrl+A / Cmd+A) og slet det.
4. Åbn filen **`BudgetFraCSV.gs`** fra denne pakke, kopiér hele indholdet,
   og indsæt det i den tomme `Code.gs` i Apps Script-editoren.
5. Opret én fil mere: klik på **+** ud for "Filer" til venstre → **HTML** →
   navngiv den præcis **`ImportDialog`** (Apps Script tilføjer selv
   `.html`). Slet den forudfyldte kode, åbn filen **`ImportDialog.html`**
   fra denne pakke, kopiér hele indholdet, og indsæt det.
6. Gem projektet (Ctrl+S / Cmd+S). Giv det evt. et navn, fx "Budget fra
   CSV".
7. Gå tilbage til regnearket og genindlæs siden (F5). Der kommer nu en ny
   menu i menulinjen: **"Budget fra CSV"**.
8. Første gang du bruger et menupunkt, beder Google om godkendelse ("Denne
   app er ikke verificeret" - det er normalt for et script, du selv har
   indsat; klik "Avanceret" → "Gå til [projektnavn] (usikker)" → "Tillad").
   Det sker kun én gang.

Det er det - ingen `.exe`, ingen zip-fil at pakke ud, ingen "Trust
Center", ingen VBA-editor. Bare to filer indsat i browseren, og fra da af
findes budgettet i det Google Sheets-ark, du lige har lavet - helt
almindeligt at dele, åbne på mobilen, eller kopiere til et nyt ark.

Vil du bruge det i endnu et ark senere, gentager du bare trin 1-8 i det nye
ark - scriptet følger ikke automatisk med til andre ark (det er sådan
Apps Script fungerer: ét script pr. ark, medmindre du selv kopierer det
over).

## Sådan bruges det

Menuen **"Budget fra CSV"** har tre punkter:

- **Nyt budget fra CSV-fil…** - vælg CSV-filen fra netbanken i dialogen,
  der åbner. Scriptet gætter dato/tekst/beløb-kolonnerne og viser sit gæt
  til godkendelse (ret det, hvis det gættede forkert), og bygger derefter
  budgettet direkte ind i dette ark: posteringer, kategorier, et udkast til
  et fast månedsbudget og en 24-måneders prognose med graf.
- **Tilføj flere posteringer (CSV) til budget…** - vælg en ny CSV-fil (fx
  den seneste måneds udtræk). Nye posteringer kategoriseres med de regler,
  du allerede har i arket. Poster, der er identiske med noget, der allerede
  findes (samme dato, tekst og beløb), springes automatisk over.
- **Opdatér kategorier i budget…** - kør dette efter at have tilføjet
  eller rettet en regel i arket "Kategorier". Alle posteringer får deres
  kategori genberegnet efter de opdaterede regler.

Der er ingen "gem"-dialog - Google Sheets gemmer automatisk, ligesom det
altid gør.

### Kategorisering

Arket **Kategorier** har to tabeller side om side:

- **Nøgleord / Kategori** (venstre) - dine regler. Tilføj en ny række for
  en ny regel, fx "netto" ▸ "Dagligvarer". Kør derefter **Opdatér
  kategorier i budget…**, så den slår igennem på alle posteringer med det
  samme, ikke kun nye.
- **Kategori / Konto** (højre) - en liste over alle kategorier og hvilken
  konto, de trækkes fra.

Du kan også rette kategorien direkte i kolonnen "Kategori" i arket
"Transaktioner" for én enkelt postering (der er en rulleliste med alle
kendte kategorier).

### Interne overførsler mellem egne konti

Bankudtræk viser ofte en overførsel mellem to af dine egne konti som en
almindelig indtægt eller udgift. Sæt dine egne kontonumre ind under
"Kontonummer(e)" på arket **Konti** (flere numre adskilles med komma), og
kør **Opdatér kategorier i budget…** - så bliver enhver postering, hvis
tekst indeholder et af numrene, automatisk sat til "Ignoreret" og tæller
ikke længere med i budgettet.

## Hvad er anderledes end de andre udgaver

- Budgettet laves direkte i det ark, du har åbent - der er intet begreb om
  en separat fil at gemme eller åbne, sådan som LibreOffice/Excel/Python-
  udgaverne har.
- Ingen "kategorisér ukategoriserede poster"-guide med grupperede forslag -
  ukategoriserede poster får kategorien "Ukategoriseret" og kan rettes
  direkte i Transaktioner-arket eller ved at tilføje en regel.
- Ingen separate "Oversigt"-ark pr. måned - "Alle måneder" viser i stedet
  en samlet oversigt, måned for måned, i ét ark.
- Formlerne (Budgetforslag, Prognose) er almindelige Sheets-formler og
  genberegnes altid live - retter du et Mål-tal i Budgetforslag, opdaterer
  Prognosen sig selv med det samme, uden at du skal køre noget script.
- Rettet en kategori direkte i Transaktioner-arket overlever ikke
  nødvendigvis en kørsel af "Opdatér kategorier": hvis der findes en regel,
  der matcher posteringens tekst, vinder reglen. Vil en postering have en
  fast kategori uanset regler, skal reglen selv rettes eller fjernes.

## Fejlfinding

- **"Denne app er ikke verificeret"**: normalt for et script, du selv har
  indsat i dit eget ark - det er ikke publiceret i Google Workspace
  Marketplace, så Google kan ikke automatisk bekræfte, hvem der har lavet
  det. Klik "Avanceret" → "Gå til [projektnavn] (usikker)" → "Tillad".
- **Menuen "Budget fra CSV" kommer ikke frem**: genindlæs regnearket
  (F5) efter at have gemt scriptet - menuen oprettes, når arket åbnes.
- **"Fandt ingen datarækker i filen"**: CSV-filen er tom, eller kunne ikke
  læses som CSV - tjek at det rent faktisk er den fil, du eksporterede fra
  netbanken.
- **Dialogen fryser på "Arbejder…"**: luk dialogen og prøv igen - store
  CSV-filer (mange tusinde rækker) kan tage nogle sekunder. Sker det igen,
  tjek Apps Scripts køre-log under "Udførelser" i Apps Script-editoren for
  en fejlbesked.

## Til udviklere

`src/` indeholder de samme filer opdelt enkeltvis, hver med en lille
`module.exports`-blok i bunden, så de kan køres og testes med almindelig
Node.js uden nogen forbindelse til Google - selve parsing/kategoriserings-
logikken (alt undtagen `SheetWriter.gs` og `Code.gs`) er derfor testet med
et rigtigt testsuite, ikke kun læst igennem:

```
node --test tests/
```

`SheetWriter.gs` (som skriver til selve regnearket) kan ikke køre uden en
rigtig Google-forbindelse, men `tests/mock_spreadsheet_app.js` er en let
efterligning af `SpreadsheetApp`-API'et, så `BudgetBook.build()` /
`append_transactions()` / `refresh()` alligevel kan køres rigtigt igennem
og tjekkes - se `tests/sheetwriter.test.js`.

`BudgetFraCSV.gs` (den fil, brugeren rent faktisk indsætter) er `src/`-
filerne samlet i én, i den rækkefølge `build.py` definerer. Redigér aldrig
`BudgetFraCSV.gs` direkte - redigér den tilsvarende fil under `src/` og kør:

```
python3 build.py
```

`LibreOfficeBudget/budget_core/` og `BudgetScript/budget_core/` er den
samme motor skrevet i Python - denne udgave er en selvstændig JavaScript-
port af samme logik (Apps Script kan ikke køre Python), holdt i sync
manuelt. Ændrer du kategoriseringsreglerne eller klassificeringsreglerne
ét sted, så ret dem tilsvarende de andre steder.
