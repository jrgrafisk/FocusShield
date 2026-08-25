# Budget fra CSV - script-udgave

Et lille program, der laver et færdigt Excel-budget ud fra et bank-CSV-udtræk
- **ingen makroer, ingen tilføjelsesprogrammer, ingen VBA-editor.** Du kører
scriptet, vælger din CSV-fil, og får en helt almindelig `.xlsx`-fil med
alt regnet ud: kategoriserede posteringer, et udkast til et fast
månedsbudget, og en 24-måneders prognose med graf.

Det er den samme motor (kategorisering, budgetklassificering osv.), som
ligger bag LibreOffice-udvidelsen "Budget fra CSV" - bare uden nogen
tilføjelse installeret i selve regnearksprogrammet.

## Sådan installeres det (én gang)

1. **Installér Python**, hvis du ikke allerede har det: hent den fra
   [python.org/downloads](https://www.python.org/downloads/) og kør
   installationsprogrammet. **Vigtigt:** sæt flueben i "Add python.exe to
   PATH" på den første side af installationsguiden.
2. Åbn en kommandoprompt i den mappe, du har udpakket disse filer til
   (Shift + højreklik i mappen ▸ "Åbn PowerShell-vindue her" eller
   "Åbn kommandoprompt her"), og kør:

   ```
   py -3 -m pip install -r requirements.txt
   ```

   (Det installerer `openpyxl`, det eneste, scriptet har brug for udover
   Python selv - `tkinter`, som viser vinduerne, følger med Python.)

Det er det. Intet Trust Center, ingen makro-sikkerhedsadvarsler, ingen
VBA-editor.

## Sådan bruges det

Dobbeltklik på **`start_budget.bat`** (eller kør `py -3 budget_fra_csv.py`
fra en kommandoprompt). Der åbner et lille vindue med tre knapper:

- **Nyt budget fra CSV-fil…** - vælg CSV-filen fra netbanken, bekræft
  kolonnevalget (dato/tekst/beløb - det er som regel gættet rigtigt, men
  du kan rette det i vinduet, der åbner), vælg hvor budgettet skal gemmes,
  og du har et færdigt Excel-ark.
- **Tilføj flere posteringer (CSV) til budget…** - vælg et budget, du
  allerede har lavet, og en ny CSV-fil (fx den seneste måneds udtræk). Nye
  posteringer kategoriseres med de regler, du allerede har i budgettet.
  Poster, der er identiske med noget, der allerede findes (samme dato,
  tekst og beløb), springes automatisk over.
- **Opdatér kategorier i budget…** - kør dette efter at have tilføjet
  eller rettet en regel i arket "Kategorier" i et eksisterende budget. Alle
  posteringer får deres kategori genberegnet efter de opdaterede regler.

Alle tre gemmer en (ny eller opdateret) `.xlsx`-fil, du derefter bare åbner
i Excel som et helt almindeligt regneark.

### Kategorisering

Arket **Kategorier** i den færdige fil har to tabeller side om side:

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

## Hvad er anderledes end LibreOffice-udgaven

- Ingen "kategorisér ukategoriserede poster"-guide med grupperede forslag -
  ukategoriserede poster får kategorien "Ukategoriseret" og kan rettes
  direkte i Transaktioner-arket eller ved at tilføje en regel.
- Ingen separate "Oversigt"-ark pr. måned (Budget/Faktisk/Diff. pr.
  kategori) - "Alle måneder" viser i stedet en samlet oversigt, måned for
  måned, i ét ark.
- Ingen "Opret bankbudget"-øjebliksbillede endnu.
- Ingen opdeling af prognosen pr. konto.
- Formlerne er almindelige Excel-formler (SUM, +, -) - ingen live
  genberegning, når du retter noget direkte i Excel og gemmer; kør
  **Opdatér kategorier i budget…** for at få ændringer i regler til at
  slå igennem, eller **Tilføj flere posteringer…** for nye data. Retter du
  bare et Mål-tal i Budgetforslag, opdaterer Excel selv de formler, der
  bruger det (fx Prognosen), med det samme.

## Fejlfinding

- **"ModuleNotFoundError: No module named 'openpyxl'"**: kør
  `py -3 -m pip install -r requirements.txt` i mappen med scriptet.
- **"py" er ikke genkendt som kommando**: Python blev sandsynligvis
  installeret uden "Add python.exe to PATH" - kør installationsprogrammet
  igen og sæt fluebenet, eller brug `python` i stedet for `py -3`.
- **Vinduet åbner slet ikke**: kør scriptet fra en kommandoprompt (ikke ved
  dobbeltklik) for at se en eventuel fejlbesked: `py -3 budget_fra_csv.py`.

## Til udviklere

`budget_core/` er en kopi af den motor, der også bruges af LibreOffice-
udvidelsen (`LibreOfficeBudget/budget_core/` i samme repo) - hold dem i
sync, hvis den ene rettes. `xlsx_writer.py` er "office"-laget (skriver
arkene med openpyxl); `budget_fra_csv.py` er GUI'en (tkinter). Kør
testene med:

```
python3 -m unittest discover -s tests -v
```
