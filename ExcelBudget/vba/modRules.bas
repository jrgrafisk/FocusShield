Attribute VB_Name = "modRules"
Option Explicit
' Category rules: keyword -> category. Port of budget_core/rules.py.
' Keywords/categories with Danish letters use the ~ae~/~oe~/~aa~/~ee~
' markers decoded by modI18n.DK - see that module for why.

Public Const DEFAULT_CATEGORY As String = "Ukategoriseret"

' A posting set to this category counts towards nothing: not an expense,
' not an income, not "uncategorised" either. For internal transfers,
' refunds that would otherwise double up a purchase, test postings and
' the like - things a budget should simply not see.
Public Const IGNORED_CATEGORY As String = "Ignoreret"

' Keywords this short (or shorter) only match whole words.
Public Const SHORT_KEYWORD As Long = 4

Public Function IncomeCategories() As Variant
    IncomeCategories = Array(DK("L~oe~n"), "Offentlige ydelser", _
                             "Renter og udbytte", "Refusion", DK("Indt~ae~gt"))
End Function

' Every (keyword, category) pair as one big table, one row per line,
' fields separated by "|". Danish letters use the ~xx~ markers.
Private Function DefaultRulesRaw() As String
    Dim t As String
    t = ""
    ' --- Dagligvarer ---
    t = t & "netto|Dagligvarer" & vbLf
    t = t & "f~oe~tex|Dagligvarer" & vbLf
    t = t & "bilka|Dagligvarer" & vbLf
    t = t & "rema|Dagligvarer" & vbLf
    t = t & "lidl|Dagligvarer" & vbLf
    t = t & "aldi|Dagligvarer" & vbLf
    t = t & "kvickly|Dagligvarer" & vbLf
    t = t & "superbrugsen|Dagligvarer" & vbLf
    t = t & "dagli'brugsen|Dagligvarer" & vbLf
    t = t & "brugsen|Dagligvarer" & vbLf
    t = t & "coop|Dagligvarer" & vbLf
    t = t & "irma|Dagligvarer" & vbLf
    t = t & "meny|Dagligvarer" & vbLf
    t = t & "spar |Dagligvarer" & vbLf
    t = t & "fakta|Dagligvarer" & vbLf
    t = t & "365discount|Dagligvarer" & vbLf
    t = t & "l~oe~vbjerg|Dagligvarer" & vbLf
    t = t & "min k~oe~bmand|Dagligvarer" & vbLf
    t = t & "nemlig.com|Dagligvarer" & vbLf
    t = t & "bager|Dagligvarer" & vbLf
    t = t & "slagter|Dagligvarer" & vbLf
    t = t & "supermarked|Dagligvarer" & vbLf
    t = t & "grocery|Dagligvarer" & vbLf
    ' --- Restaurant og takeaway ---
    t = t & "restaurant|Restaurant" & vbLf
    t = t & "bistro|Restaurant" & vbLf
    t = t & "cafe|Restaurant" & vbLf
    t = t & "caf~ee~|Restaurant" & vbLf
    t = t & "kro |Restaurant" & vbLf
    t = t & "pizza|Restaurant" & vbLf
    t = t & "sushi|Restaurant" & vbLf
    t = t & "burger|Restaurant" & vbLf
    t = t & "mcdonald|Restaurant" & vbLf
    t = t & "kebab|Restaurant" & vbLf
    t = t & "grill|Restaurant" & vbLf
    t = t & "wolt|Restaurant" & vbLf
    t = t & "just eat|Restaurant" & vbLf
    t = t & "hungry|Restaurant" & vbLf
    t = t & "takeaway|Restaurant" & vbLf
    t = t & "starbucks|Restaurant" & vbLf
    t = t & "espresso house|Restaurant" & vbLf
    t = t & "joe & the juice|Restaurant" & vbLf
    t = t & "baresso|Restaurant" & vbLf
    t = t & "bodega|Restaurant" & vbLf
    t = t & "bar |Restaurant" & vbLf
    ' --- Transport ---
    t = t & "dsb|Transport" & vbLf
    t = t & "rejsekort|Transport" & vbLf
    t = t & "metro|Transport" & vbLf
    t = t & "movia|Transport" & vbLf
    t = t & "arriva|Transport" & vbLf
    t = t & "taxa|Transport" & vbLf
    t = t & "taxi|Transport" & vbLf
    t = t & "dantaxi|Transport" & vbLf
    t = t & "4x35|Transport" & vbLf
    t = t & "uber|Transport" & vbLf
    t = t & "bolt.eu|Transport" & vbLf
    t = t & "donkey republic|Transport" & vbLf
    t = t & "brobizz|Transport" & vbLf
    t = t & "storeb~ae~lt|Transport" & vbLf
    t = t & "~oe~resund|Transport" & vbLf
    t = t & "f~ae~rge|Transport" & vbLf
    t = t & "molslinjen|Transport" & vbLf
    t = t & "scandlines|Transport" & vbLf
    t = t & "parkering|Transport" & vbLf
    t = t & "easypark|Transport" & vbLf
    t = t & "parkpark|Transport" & vbLf
    t = t & "apcoa|Transport" & vbLf
    t = t & "q-park|Transport" & vbLf
    t = t & "circle k|Transport" & vbLf
    t = t & "shell|Transport" & vbLf
    t = t & "q8|Transport" & vbLf
    t = t & "ok benzin|Transport" & vbLf
    t = t & "uno-x|Transport" & vbLf
    t = t & "ingo|Transport" & vbLf
    t = t & "benzin|Transport" & vbLf
    t = t & "diesel|Transport" & vbLf
    t = t & "ladestander|Transport" & vbLf
    t = t & "clever|Transport" & vbLf
    t = t & "spirii|Transport" & vbLf
    t = t & "fdm|Transport" & vbLf
    t = t & "bilsyn|Transport" & vbLf
    t = t & "autov~ae~rksted|Transport" & vbLf
    t = t & "d~ae~k|Transport" & vbLf
    ' --- Rejser og hotel ---
    t = t & "hotel|Rejser" & vbLf
    t = t & "resort|Rejser" & vbLf
    t = t & "booking.com|Rejser" & vbLf
    t = t & "airbnb|Rejser" & vbLf
    t = t & "hotels.com|Rejser" & vbLf
    t = t & "expedia|Rejser" & vbLf
    t = t & "momondo|Rejser" & vbLf
    t = t & "travel|Rejser" & vbLf
    t = t & "flybillet|Rejser" & vbLf
    t = t & "ryanair|Rejser" & vbLf
    t = t & "norwegian|Rejser" & vbLf
    t = t & "lufthansa|Rejser" & vbLf
    t = t & "sas |Rejser" & vbLf
    t = t & "airport|Rejser" & vbLf
    t = t & "lufthavn|Rejser" & vbLf
    t = t & "hostel|Rejser" & vbLf
    t = t & "camping|Rejser" & vbLf
    ' --- Bolig ---
    t = t & "husleje|Bolig" & vbLf
    t = t & "boligselskab|Bolig" & vbLf
    t = t & "andelsbolig|Bolig" & vbLf
    t = t & "ejerforening|Bolig" & vbLf
    t = t & "grundejerforening|Bolig" & vbLf
    t = t & "realkredit|Bolig" & vbLf
    t = t & "totalkredit|Bolig" & vbLf
    t = t & "nykredit|Bolig" & vbLf
    t = t & "ejendomsskat|Bolig" & vbLf
    t = t & "grundskyld|Bolig" & vbLf
    t = t & "bolig l~aa~n|Bolig" & vbLf
    t = t & "vicev~ae~rt|Bolig" & vbLf
    t = t & "renovation|Bolig" & vbLf
    t = t & "skorstensfejer|Bolig" & vbLf
    ' --- El, vand og varme ---
    t = t & "~oe~rsted|El, vand og varme" & vbLf
    t = t & "norlys|El, vand og varme" & vbLf
    t = t & "andel energi|El, vand og varme" & vbLf
    t = t & "ewii|El, vand og varme" & vbLf
    t = t & "nrgi|El, vand og varme" & vbLf
    t = t & "energi fyn|El, vand og varme" & vbLf
    t = t & "hofor|El, vand og varme" & vbLf
    t = t & "fjernvarme|El, vand og varme" & vbLf
    t = t & "vandforsyning|El, vand og varme" & vbLf
    t = t & "elforsyning|El, vand og varme" & vbLf
    t = t & "naturgas|El, vand og varme" & vbLf
    t = t & "radius|El, vand og varme" & vbLf
    t = t & "vand og spildevand|El, vand og varme" & vbLf
    ' --- Telefon og internet ---
    t = t & "yousee|Telefon og internet" & vbLf
    t = t & "tdc|Telefon og internet" & vbLf
    t = t & "telenor|Telefon og internet" & vbLf
    t = t & "telia|Telefon og internet" & vbLf
    t = t & "cbb|Telefon og internet" & vbLf
    t = t & "oister|Telefon og internet" & vbLf
    t = t & "hiper|Telefon og internet" & vbLf
    t = t & "stofa|Telefon og internet" & vbLf
    t = t & "fastspeed|Telefon og internet" & vbLf
    t = t & "callme|Telefon og internet" & vbLf
    t = t & "lebara|Telefon og internet" & vbLf
    t = t & "bredb~aa~nd|Telefon og internet" & vbLf
    t = t & "mobilabonnement|Telefon og internet" & vbLf
    ' --- Abonnementer ---
    t = t & "netflix|Abonnementer" & vbLf
    t = t & "spotify|Abonnementer" & vbLf
    t = t & "hbo|Abonnementer" & vbLf
    t = t & "max.com|Abonnementer" & vbLf
    t = t & "disney|Abonnementer" & vbLf
    t = t & "viaplay|Abonnementer" & vbLf
    t = t & "tv 2 play|Abonnementer" & vbLf
    t = t & "tv2 play|Abonnementer" & vbLf
    t = t & "youtube|Abonnementer" & vbLf
    t = t & "icloud|Abonnementer" & vbLf
    t = t & "apple.com/bill|Abonnementer" & vbLf
    t = t & "google one|Abonnementer" & vbLf
    t = t & "google storage|Abonnementer" & vbLf
    t = t & "dropbox|Abonnementer" & vbLf
    t = t & "adobe|Abonnementer" & vbLf
    t = t & "microsoft|Abonnementer" & vbLf
    t = t & "openai|Abonnementer" & vbLf
    t = t & "anthropic|Abonnementer" & vbLf
    t = t & "claude.ai|Abonnementer" & vbLf
    t = t & "patreon|Abonnementer" & vbLf
    t = t & "audible|Abonnementer" & vbLf
    t = t & "storytel|Abonnementer" & vbLf
    t = t & "mofibo|Abonnementer" & vbLf
    t = t & "zwift|Abonnementer" & vbLf
    t = t & "strava|Abonnementer" & vbLf
    t = t & "licens|Abonnementer" & vbLf
    t = t & "abonnement|Abonnementer" & vbLf
    t = t & "dr licens|Abonnementer" & vbLf
    ' --- Forsikring og pension ---
    t = t & "forsikring|Forsikring og pension" & vbLf
    t = t & "tryg|Forsikring og pension" & vbLf
    t = t & "topdanmark|Forsikring og pension" & vbLf
    t = t & "alka|Forsikring og pension" & vbLf
    t = t & "codan|Forsikring og pension" & vbLf
    t = t & "alm brand|Forsikring og pension" & vbLf
    t = t & "if skade|Forsikring og pension" & vbLf
    t = t & "gf |Forsikring og pension" & vbLf
    t = t & "pfa|Forsikring og pension" & vbLf
    t = t & "velliv|Forsikring og pension" & vbLf
    t = t & "danica|Forsikring og pension" & vbLf
    t = t & "sampension|Forsikring og pension" & vbLf
    t = t & "pensiondanmark|Forsikring og pension" & vbLf
    t = t & "pension|Forsikring og pension" & vbLf
    ' --- Sundhed ---
    t = t & "apotek|Sundhed" & vbLf
    t = t & "l~ae~ge|Sundhed" & vbLf
    t = t & "tandl~ae~ge|Sundhed" & vbLf
    t = t & "fysioterapi|Sundhed" & vbLf
    t = t & "kiropraktor|Sundhed" & vbLf
    t = t & "psykolog|Sundhed" & vbLf
    t = t & "sygesikring|Sundhed" & vbLf
    t = t & "danmark sygefor|Sundhed" & vbLf
    t = t & "optiker|Sundhed" & vbLf
    t = t & "briller|Sundhed" & vbLf
    t = t & "hospital|Sundhed" & vbLf
    t = t & "klinik|Sundhed" & vbLf
    ' --- Fritid og sport ---
    t = t & "fitness|Fritid og sport" & vbLf
    t = t & "sats|Fritid og sport" & vbLf
    t = t & "puregym|Fritid og sport" & vbLf
    t = t & "loop|Fritid og sport" & vbLf
    t = t & "sv~oe~mmehal|Fritid og sport" & vbLf
    t = t & "idr~ae~tsforening|Fritid og sport" & vbLf
    t = t & "kontingent|Fritid og sport" & vbLf
    t = t & "biograf|Fritid og sport" & vbLf
    t = t & "nordisk film|Fritid og sport" & vbLf
    t = t & "cinemaxx|Fritid og sport" & vbLf
    t = t & "tivoli|Fritid og sport" & vbLf
    t = t & "legoland|Fritid og sport" & vbLf
    t = t & "zoo|Fritid og sport" & vbLf
    t = t & "museum|Fritid og sport" & vbLf
    t = t & "billetlugen|Fritid og sport" & vbLf
    t = t & "ticketmaster|Fritid og sport" & vbLf
    t = t & "koncert|Fritid og sport" & vbLf
    t = t & "bibliotek|Fritid og sport" & vbLf
    ' --- Shopping ---
    t = t & "h&m|Shopping" & vbLf
    t = t & "zara|Shopping" & vbLf
    t = t & "zalando|Shopping" & vbLf
    t = t & "bestseller|Shopping" & vbLf
    t = t & "magasin|Shopping" & vbLf
    t = t & "matas|Shopping" & vbLf
    t = t & "normal|Shopping" & vbLf
    t = t & "elgiganten|Shopping" & vbLf
    t = t & "power.dk|Shopping" & vbLf
    t = t & "proshop|Shopping" & vbLf
    t = t & "komplett|Shopping" & vbLf
    t = t & "ikea|Shopping" & vbLf
    t = t & "jysk|Shopping" & vbLf
    t = t & "ilva|Shopping" & vbLf
    t = t & "bauhaus|Shopping" & vbLf
    t = t & "silvan|Shopping" & vbLf
    t = t & "xl-byg|Shopping" & vbLf
    t = t & "stark|Shopping" & vbLf
    t = t & "harald nyborg|Shopping" & vbLf
    t = t & "biltema|Shopping" & vbLf
    t = t & "jem & fix|Shopping" & vbLf
    t = t & "amazon|Shopping" & vbLf
    t = t & "ebay|Shopping" & vbLf
    t = t & "temu|Shopping" & vbLf
    t = t & "shein|Shopping" & vbLf
    t = t & "wish|Shopping" & vbLf
    t = t & "aliexpress|Shopping" & vbLf
    t = t & "boghandel|Shopping" & vbLf
    t = t & "saxo.com|Shopping" & vbLf
    t = t & "blomster|Shopping" & vbLf
    t = t & "t~oe~j|Shopping" & vbLf
    t = t & "sportmaster|Shopping" & vbLf
    t = t & "intersport|Shopping" & vbLf
    ' --- Boern og uddannelse ---
    t = t & "daginstitution|B~oe~rn og uddannelse" & vbLf
    t = t & "vuggestue|B~oe~rn og uddannelse" & vbLf
    t = t & "b~oe~rnehave|B~oe~rn og uddannelse" & vbLf
    t = t & "dagpleje|B~oe~rn og uddannelse" & vbLf
    t = t & "sfo|B~oe~rn og uddannelse" & vbLf
    t = t & "skolefritids|B~oe~rn og uddannelse" & vbLf
    t = t & "efterskole|B~oe~rn og uddannelse" & vbLf
    t = t & "studie|B~oe~rn og uddannelse" & vbLf
    t = t & "undervisning|B~oe~rn og uddannelse" & vbLf
    t = t & "kursus|B~oe~rn og uddannelse" & vbLf
    ' --- Gebyrer og renter ---
    t = t & "gebyr|Gebyrer og renter" & vbLf
    t = t & "rentetilskrivning|Gebyrer og renter" & vbLf
    t = t & "renter|Gebyrer og renter" & vbLf
    t = t & "overtr~ae~k|Gebyrer og renter" & vbLf
    t = t & "valutatill~ae~g|Gebyrer og renter" & vbLf
    t = t & "kortgebyr|Gebyrer og renter" & vbLf
    t = t & "betalingsservice|Gebyrer og renter" & vbLf
    t = t & "rykker|Gebyrer og renter" & vbLf
    ' --- Laan og afdrag ---
    t = t & "afdrag|L~aa~n og afdrag" & vbLf
    t = t & "bill~aa~n|L~aa~n og afdrag" & vbLf
    t = t & "forbrugsl~aa~n|L~aa~n og afdrag" & vbLf
    t = t & "studiel~aa~n|L~aa~n og afdrag" & vbLf
    t = t & "su-l~aa~n|L~aa~n og afdrag" & vbLf
    t = t & "l~aa~nebetaling|L~aa~n og afdrag" & vbLf
    t = t & "santander|L~aa~n og afdrag" & vbLf
    t = t & "ekspres bank|L~aa~n og afdrag" & vbLf
    t = t & "resurs bank|L~aa~n og afdrag" & vbLf
    t = t & "basisbank|L~aa~n og afdrag" & vbLf
    t = t & "lendo|L~aa~n og afdrag" & vbLf
    t = t & "sparxpres|L~aa~n og afdrag" & vbLf
    t = t & "kreditforening|L~aa~n og afdrag" & vbLf
    ' --- Opsparing og investering ---
    t = t & "opsparing|Opsparing" & vbLf
    t = t & "nordnet|Opsparing" & vbLf
    t = t & "saxo bank|Opsparing" & vbLf
    t = t & "aktier|Opsparing" & vbLf
    t = t & "investering|Opsparing" & vbLf
    t = t & "coinbase|Opsparing" & vbLf
    t = t & "puljeinvest|Opsparing" & vbLf
    ' --- Overfoersler ---
    t = t & "mobilepay|MobilePay" & vbLf
    t = t & "overf~oe~rsel|Overf~oe~rsel" & vbLf
    t = t & "overf~oe~relse|Overf~oe~rsel" & vbLf
    t = t & "egen konto|Overf~oe~rsel" & vbLf
    t = t & "straksoverf~oe~rsel|Overf~oe~rsel" & vbLf
    t = t & "transfer|Overf~oe~rsel" & vbLf
    ' --- Indtaegter ---
    t = t & "l~oe~n|L~oe~n" & vbLf
    t = t & "l~oe~noverf~oe~rsel|L~oe~n" & vbLf
    t = t & "l~oe~n overf~oe~rsel|L~oe~n" & vbLf
    t = t & "l~oe~nudbetaling|L~oe~n" & vbLf
    t = t & "l~oe~n udbetaling|L~oe~n" & vbLf
    t = t & "salary|L~oe~n" & vbLf
    t = t & "m~aa~nedsl~oe~n|L~oe~n" & vbLf
    t = t & "a-conto l~oe~n|L~oe~n" & vbLf
    t = t & "udbetaling danmark|Offentlige ydelser" & vbLf
    t = t & "su |Offentlige ydelser" & vbLf
    t = t & "su-|Offentlige ydelser" & vbLf
    t = t & "b~oe~rne- og ungeydelse|Offentlige ydelser" & vbLf
    t = t & "b~oe~rnepenge|Offentlige ydelser" & vbLf
    t = t & "boligst~oe~tte|Offentlige ydelser" & vbLf
    t = t & "dagpenge|Offentlige ydelser" & vbLf
    t = t & "pensionsudbetaling|Offentlige ydelser" & vbLf
    t = t & "skattestyrelsen|Refusion" & vbLf
    t = t & "overskydende skat|Refusion" & vbLf
    t = t & "refusion|Refusion" & vbLf
    t = t & "tilbagebetaling|Refusion" & vbLf
    t = t & "udbytte|Renter og udbytte" & vbLf
    t = t & "renteindt~ae~gt|Renter og udbytte" & vbLf

    DefaultRulesRaw = t
End Function

' [(keyword, category), ...] as a Collection of 2-element String arrays,
' Danish letters already decoded.
Public Function DefaultRulesRows() As Collection
    Dim out As New Collection
    Dim lines() As String, i As Long, parts() As String
    lines = Split(DefaultRulesRaw(), vbLf)
    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            parts = Split(lines(i), "|")
            out.Add Array(DK(parts(0)), DK(parts(1)))
        End If
    Next i
    Set DefaultRulesRows = out
End Function

' All categories used by the default rules, sorted, income last, with
' DEFAULT_CATEGORY/IGNORED_CATEGORY always appended.
Public Function DefaultCategories() As Collection
    Dim rs As New clsRuleSet
    rs.InitDefault
    Set DefaultCategories = rs.Categories()
End Function
