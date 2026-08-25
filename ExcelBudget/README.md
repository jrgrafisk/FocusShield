# Budget fra CSV - Excel-udgave

En VBA-udgave af LibreOffice-udvidelsen "Budget fra CSV": importerer et
bank-CSV-udtræk, kategoriserer posteringerne automatisk, og bygger et
udkast til et fast månedsbudget med en 24-måneders prognose.

## Vigtigt at vide, før du starter

**Kun Windows.** Koden bruger tre Windows-komponenter (`ADODB.Stream` til
robust indlæsning af CSV-filer uanset tegnsæt, `VBScript.RegExp` til
mønstergenkendelse, og `Scripting.Dictionary` som opslagstabel). De findes
ikke i Excel til Mac, som ikke har adgang til Windows' COM-komponenter.
Denne udgave virker altså kun i Excel til Windows (skrivebordsudgaven -
ikke Excel til nettet).

**Ingen udleveret `.xlam`-fil.** En Excel-tilføjelsesprograms kompilerede
makro-lager (`vbaProject.bin`) kan kun genereres af Excel selv - det kan
ikke bygges uden for et rigtigt Excel. Derfor leveres koden som almindelige
VBA-kildefiler (`.bas`/`.cls`), som du importerer én gang i VBA-editoren -
der er en installations-makro, der gør det meste af arbejdet for dig (se
"Sådan installeres den" nedenfor). Tager et par minutter.

**Ikke testet i et rigtigt Excel.** Koden er skrevet omhyggeligt og er en
tro oversættelse af den LibreOffice-udvidelse, den er baseret på - men er
skrevet uden adgang til et rigtigt Excel til at afprøve den undervejs (i
modsætning til LibreOffice-udgaven, som er afprøvet grundigt mod en kørende
LibreOffice for hver eneste funktion). Der kan derfor være fejl, der først
viser sig, når du rent faktisk bruger den. Stød du på en fejl - en makro,
der ikke starter, eller opfører sig forkert - så sig til, så retter vi det.

**Ingen pop op-vinduer under import.** LibreOffice-udgaven viser en dialog,
hvor du kan rette kolonnevalget (dato/tekst/beløb), før budgettet bygges.
At genskabe den slags rigtige pop op-vinduer i VBA kræver en "UserForm",
som skal designes visuelt i VBA-editoren - det kan ikke leveres pålideligt
som en tekstfil uden adgang til Excel til at lave layoutet i. Denne udgave
bruger i stedet automatisk kolonnegenkendelse direkte, uden bekræftelse.
Gættede den forkert (sjældent, men kan ske med usædvanlige CSV-filer), kan
du rette det direkte i arket "Transaktioner" og "Kategorier" bagefter, og
trykke "Opdatér kategorier og budget".

## Sådan installeres den (én gang)

Der er 26 filer i `vba`-mappen. Du behøver **ikke** importere dem én ad
gangen - der er en lille installations-makro (`_Installer.bas`), der gør
resten for dig.

### Trin 1: slå adgang til VBA-projektet til (én gang, i alle Excel-projekter)

I Excel: **Filer ▸ Indstillinger ▸ Sikkerhedscenter ▸ Indstillinger for
Sikkerhedscenter… ▸ Makroindstillinger** ▸ sæt flueben i **"Vis
tillid til adgang til VBA-projektobjektmodellen"** ▸ **OK** ▸ **OK**.
(På engelsk: Trust Center ▸ Macro Settings ▸ "Trust access to the VBA
project object model".) Det er det, der giver en makro lov til at
importere andre moduler for dig - uden det kan installations-makroen
ikke køre.

### Trin 2: importér installations-filen og kør den

1. Åbn Excel, opret en ny, tom projektmappe.
2. Tryk **Alt+F11** for at åbne VBA-editoren.
3. Højreklik på **VBAProject (Mappe1)** i Project Explorer (til venstre -
   ikke synlig? tryk **Ctrl+R**) ▸ **Filer** ▸ **Importér fil…** ▸ vælg
   **kun** `_Installer.bas` fra `vba`-mappen ▸ **Åbn**.
4. Dobbeltklik på det nye modul **modInstaller** i Project Explorer, klik
   et sted inde i `Public Sub ImportAllModules()`, og tryk **F5** (Kør).
5. Der åbner en mappevælger - vælg selve `vba`-mappen (den samme, du hentede
   `_Installer.bas` fra). Den importerer nu de resterende 23 filer og
   indsætter ThisWorkbook-koden automatisk.
6. En besked bekræfter, hvor mange filer der blev importeret. Er der noget,
   der fejlede, står det i beskeden - de kan importeres manuelt bagefter
   på samme måde som i trin 3.

### Trin 3: gem som tilføjelsesprogram og aktivér

7. Tryk **Ctrl+S** for at gemme. Excel spørger, om du vil gemme som
   makro-aktiveret. Vælg **Filtype: Excel-tilføjelsesprogram (*.xlam)**,
   giv den et navn (fx `BudgetFraCSV.xlam`), og gem den i den mappe, Excel
   selv foreslår (så den automatisk findes igen senere).
8. Luk projektmappen. I Excel: **Filer ▸ Indstillinger ▸
   Tilføjelsesprogrammer ▸ Nederst: "Administrer: Excel-tilføjelsesprogrammer"
   ▸ Kør…** ▸ sæt flueben ud for **BudgetFraCSV** ▸ **OK**.
   (Er den ikke på listen: **Gennemse…** og find den `.xlam`-fil, du gemte
   i trin 7.)
9. Der kommer måske en sikkerhedsadvarsel om makroer - det er forventet,
   siden det er en makro-baseret tilføjelse; vælg at aktivere/tillade den.
10. Åbn (eller opret) en ny projektmappe. Der skulle nu være et menupunkt
    **Budget** i topmenuen (under fanen **Tilføjelsesprogrammer**, hvis din
    Excel viser den fane - ellers direkte i topmenulinjen).

Hvis menuen ikke dukker op: tryk Alt+F11, find `modMenu` i Project
Explorer, dobbeltklik, og kør `BuildMenu` manuelt (F5, med markøren inde i
`Public Sub BuildMenu()`).

### Alternativ: importér alle filer manuelt

Vil du hellere undgå at slå Trust Center-indstillingen til, kan du
importere filerne én for én i stedet: **Filer ▸ Importér fil…** for hver
`.bas`- og `.cls`-fil (Ctrl+klik virker som regel til at markere flere ad
gangen inden for samme mappe, men ikke alle Excel-udgaver understøtter
det lige godt - importér dem enkeltvis, hvis det driller). **Undtagen**
`ThisWorkbook.cls` og `_Installer.bas`: dobbeltklik på **ThisWorkbook**
(findes allerede under Microsoft Excel Objects i Project Explorer), åbn
`ThisWorkbook.cls` i en teksteditor, kopiér alt indholdet, og indsæt det i
det tomme kodevindue for ThisWorkbook. `_Installer.bas` skal slet ikke med
i denne fremgangsmåde. Fortsæt derefter fra trin 7 ovenfor.

## Sådan bruges den

| Menupunkt | Hvad den gør |
|---|---|
| **Importér CSV-fil til budget…** | Vælg en CSV-fil fra netbanken. Opretter en ny projektmappe med hele budgettet. |
| **Tilføj flere posteringer (CSV)…** | Importér endnu en CSV-fil ind i det budget, der allerede er åbent - fx den seneste måneds udtræk. Poster, der allerede findes (samme dato, tekst og beløb), springes automatisk over. |
| **Lav budget ud fra det aktive ark…** | Samme som CSV-import, men læser data fra det ark, der er åbent lige nu (fx efter at have limet data ind manuelt). |
| **Opdatér kategorier og budget** | Kører kategoriseringsreglerne igennem alle posteringer igen (fx efter at have tilføjet en ny regel i arket "Kategorier"), og bygger Budgetforslag/Prognose/"Alle måneder" om. |
| **Konti - ret saldi…** | Hopper til (og opretter om nødvendigt) arket "Konti" - dine rigtige konti og saldi, samt kontonumre til at fange interne overførsler (se nedenfor). |
| **Opret bankbudget (øjebliksbillede)** | Laver et nyt ark med de aktuelle mål fra Budgetforslag som faste tal - til at sende til banken. Opdateres ikke automatisk; kør kommandoen igen for en ny version. |
| **Vis kategoriregler** | Hopper til arket "Kategorier". |
| **Nulstil kategoriregler** | Nulstiller alle regler til standardsættet (dine egne regler går tabt). |
| **Om Budget fra CSV** | Kort om-boks. |

### Kategorisering

Arket **Kategorier** har to tabeller side om side:

- **Nøgleord / Kategori** (venstre) - dine regler. Skriv fx "netto" i
  Nøgleord-kolonnen og "Dagligvarer" i Kategori-kolonnen for en ny regel.
  Kør **Opdatér kategorier og budget** bagefter, så den slår igennem på
  alle posteringer med det samme, ikke kun nye.
- **Kategori / Konto** (højre) - en liste over alle kategorier, der findes,
  og hvilken konto de trækkes fra (bruges ikke til andet end at vise det på
  arket lige nu).

Du kan også bare rette kategorien direkte i kolonnen "Kategori" i arket
"Transaktioner" for én enkelt postering (der er en rulleliste med alle
kendte kategorier). Det ændrer kun den ene postering; en regel i
Kategorier-arket slår igennem på alle posteringer med samme nøgleord.

### Interne overførsler mellem egne konti

Ligesom LibreOffice-udgaven: bankudtræk viser ofte en overførsel mellem to
af dine egne konti som en almindelig indtægt eller udgift. Sæt dine egne
kontonumre ind under "Kontonummer(e)" på Konti-arket (flere numre
adskilles med komma), og tryk **Opdatér kategorier og budget** - så bliver
enhver postering, hvis tekst indeholder et af numrene, automatisk sat til
"Ignoreret" og tæller ikke længere med i budgettet.

## Hvad er anderledes end LibreOffice-udgaven

- Ingen pop op-dialoger (se ovenfor).
- Ingen "kategorisér ukategoriserede poster"-guide med grupperede forslag -
  ukategoriserede poster får kategorien "Ukategoriseret" og kan rettes
  direkte i Transaktioner-arket eller ved at tilføje en regel.
- Ingen separate "Oversigt"-ark pr. måned (Budget/Faktisk/Diff. pr.
  kategori) - "Alle måneder" viser i stedet en samlet oversigt, måned for
  måned, i ét ark.
- Ingen kommentarer med de største bidragende posteringer ved hover på
  Budgetforslag-linjer.
- Ingen opdeling af prognosen pr. konto.
- Kun Windows (se ovenfor).

## Fejlfinding

- **"Compile error" ved import/kørsel**: du har sandsynligvis glemt at
  importere en af filerne, eller ThisWorkbook-koden mangler. Kør
  `ImportAllModules` igen (den springer filer, der allerede findes, over
  med en fejlbesked du kan se) - eller tjek manuelt i Project Explorer at
  alle 24 moduler/klasser fra `vba`-mappen er der, plus koden i
  ThisWorkbook.
- **To ens `Option Explicit`-fejl i ThisWorkbook**: sker kun hvis
  ThisWorkbook allerede havde en linje `Option Explicit` fra før - slet
  den ene af de to.
- **Danske bogstaver (æøå) ser forkerte ud i menuer/beskeder**: det tyder
  på, at en fil er blevet gemt om med et forkert tegnsæt undervejs. Prøv at
  importere filen igen fra den oprindelige udleverede kildefil.
- **Makroen kører ikke, "makroer er deaktiveret"**: Filer ▸ Indstillinger ▸
  Sikkerhedscenter ▸ Indstillinger for Sikkerhedscenter ▸ Makroindstillinger
  ▸ vælg "Deaktiver alle makroer med underretning" (så du kan vælge at
  aktivere dem fra sag til sag) - eller stol på tilføjelsesprogrammer i din
  organisations betroede placering.
