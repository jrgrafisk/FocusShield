"""Category rules: keyword -> category.

A rule is a keyword that is matched case/accent insensitively against the
transaction text.  A keyword starting with ``re:`` is treated as a regular
expression instead.  Longer keywords are tried first so that
``"Circle K"`` wins over ``"K"``.
"""

from __future__ import annotations

import json
import os
import re
from typing import Iterable, List, Optional, Sequence, Tuple

from .textutils import fold

__all__ = ["RuleSet", "DEFAULT_RULES", "DEFAULT_CATEGORY", "INCOME_CATEGORIES",
           "default_categories", "user_rules_path"]

DEFAULT_CATEGORY = "Ukategoriseret"

# Keywords this short (or shorter) only match whole words.
SHORT_KEYWORD = 4

# Categories that are income even when the amount happens to be positive by
# accident (used for ordering and for the summary sheet).
INCOME_CATEGORIES = ("Løn", "Offentlige ydelser", "Renter og udbytte",
                     "Refusion", "Indtægt")

# (keyword, category).  Danish first, then the international ones.
DEFAULT_RULES: Tuple[Tuple[str, str], ...] = (
    # --- Dagligvarer ---
    ("netto", "Dagligvarer"), ("føtex", "Dagligvarer"), ("bilka", "Dagligvarer"),
    ("rema", "Dagligvarer"), ("lidl", "Dagligvarer"), ("aldi", "Dagligvarer"),
    ("kvickly", "Dagligvarer"), ("superbrugsen", "Dagligvarer"),
    ("dagli'brugsen", "Dagligvarer"), ("brugsen", "Dagligvarer"),
    ("coop", "Dagligvarer"), ("irma", "Dagligvarer"), ("meny", "Dagligvarer"),
    ("spar ", "Dagligvarer"), ("fakta", "Dagligvarer"), ("365discount", "Dagligvarer"),
    ("løvbjerg", "Dagligvarer"), ("min købmand", "Dagligvarer"),
    ("nemlig.com", "Dagligvarer"), ("bager", "Dagligvarer"), ("slagter", "Dagligvarer"),
    ("supermarked", "Dagligvarer"), ("grocery", "Dagligvarer"),

    # --- Restaurant og takeaway ---
    ("restaurant", "Restaurant"), ("bistro", "Restaurant"), ("cafe", "Restaurant"),
    ("café", "Restaurant"), ("kro ", "Restaurant"), ("pizza", "Restaurant"),
    ("sushi", "Restaurant"), ("burger", "Restaurant"), ("mcdonald", "Restaurant"),
    ("kebab", "Restaurant"), ("grill", "Restaurant"), ("wolt", "Restaurant"),
    ("just eat", "Restaurant"), ("hungry", "Restaurant"), ("takeaway", "Restaurant"),
    ("starbucks", "Restaurant"), ("espresso house", "Restaurant"),
    ("joe & the juice", "Restaurant"), ("baresso", "Restaurant"),
    ("bodega", "Restaurant"), ("bar ", "Restaurant"),

    # --- Transport ---
    ("dsb", "Transport"), ("rejsekort", "Transport"), ("metro", "Transport"),
    ("movia", "Transport"), ("arriva", "Transport"), ("taxa", "Transport"),
    ("taxi", "Transport"), ("dantaxi", "Transport"), ("4x35", "Transport"),
    ("uber", "Transport"), ("bolt.eu", "Transport"), ("donkey republic", "Transport"),
    ("brobizz", "Transport"), ("storebælt", "Transport"), ("øresund", "Transport"),
    ("færge", "Transport"), ("molslinjen", "Transport"), ("scandlines", "Transport"),
    ("parkering", "Transport"), ("easypark", "Transport"), ("parkpark", "Transport"),
    ("apcoa", "Transport"), ("q-park", "Transport"),
    ("circle k", "Transport"), ("shell", "Transport"), ("q8", "Transport"),
    ("ok benzin", "Transport"), ("uno-x", "Transport"), ("ingo", "Transport"),
    ("benzin", "Transport"), ("diesel", "Transport"), ("ladestander", "Transport"),
    ("clever", "Transport"), ("spirii", "Transport"), ("fdm", "Transport"),
    ("bilsyn", "Transport"), ("autoværksted", "Transport"), ("dæk", "Transport"),

    # --- Rejser og hotel ---
    ("hotel", "Rejser"), ("resort", "Rejser"), ("booking.com", "Rejser"),
    ("airbnb", "Rejser"), ("hotels.com", "Rejser"), ("expedia", "Rejser"),
    ("momondo", "Rejser"), ("travel", "Rejser"), ("flybillet", "Rejser"),
    ("ryanair", "Rejser"), ("norwegian", "Rejser"), ("lufthansa", "Rejser"),
    ("sas ", "Rejser"), ("airport", "Rejser"), ("lufthavn", "Rejser"),
    ("hostel", "Rejser"), ("camping", "Rejser"),

    # --- Bolig ---
    ("husleje", "Bolig"), ("boligselskab", "Bolig"), ("andelsbolig", "Bolig"),
    ("ejerforening", "Bolig"), ("grundejerforening", "Bolig"),
    ("realkredit", "Bolig"), ("totalkredit", "Bolig"), ("nykredit", "Bolig"),
    ("ejendomsskat", "Bolig"), ("grundskyld", "Bolig"), ("boliglån", "Bolig"),
    ("vicevært", "Bolig"), ("renovation", "Bolig"), ("skorstensfejer", "Bolig"),

    # --- El, vand og varme ---
    ("ørsted", "El, vand og varme"), ("norlys", "El, vand og varme"),
    ("andel energi", "El, vand og varme"), ("ewii", "El, vand og varme"),
    ("nrgi", "El, vand og varme"), ("energi fyn", "El, vand og varme"),
    ("hofor", "El, vand og varme"), ("fjernvarme", "El, vand og varme"),
    ("vandforsyning", "El, vand og varme"), ("elforsyning", "El, vand og varme"),
    ("naturgas", "El, vand og varme"), ("radius", "El, vand og varme"),
    ("vand og spildevand", "El, vand og varme"),

    # --- Telefon og internet ---
    ("yousee", "Telefon og internet"), ("tdc", "Telefon og internet"),
    ("telenor", "Telefon og internet"), ("telia", "Telefon og internet"),
    ("cbb", "Telefon og internet"), ("oister", "Telefon og internet"),
    ("hiper", "Telefon og internet"), ("stofa", "Telefon og internet"),
    ("fastspeed", "Telefon og internet"), ("callme", "Telefon og internet"),
    ("lebara", "Telefon og internet"), ("bredbånd", "Telefon og internet"),
    ("mobilabonnement", "Telefon og internet"),

    # --- Abonnementer ---
    ("netflix", "Abonnementer"), ("spotify", "Abonnementer"), ("hbo", "Abonnementer"),
    ("max.com", "Abonnementer"), ("disney", "Abonnementer"), ("viaplay", "Abonnementer"),
    ("tv 2 play", "Abonnementer"), ("tv2 play", "Abonnementer"),
    ("youtube", "Abonnementer"), ("icloud", "Abonnementer"),
    ("apple.com/bill", "Abonnementer"), ("google one", "Abonnementer"),
    ("google storage", "Abonnementer"), ("dropbox", "Abonnementer"),
    ("adobe", "Abonnementer"), ("microsoft", "Abonnementer"),
    ("openai", "Abonnementer"), ("anthropic", "Abonnementer"),
    ("claude.ai", "Abonnementer"), ("patreon", "Abonnementer"),
    ("audible", "Abonnementer"), ("storytel", "Abonnementer"),
    ("mofibo", "Abonnementer"), ("zwift", "Abonnementer"), ("strava", "Abonnementer"),
    ("licens", "Abonnementer"), ("abonnement", "Abonnementer"),
    ("dr licens", "Abonnementer"),

    # --- Forsikring og pension ---
    ("forsikring", "Forsikring og pension"), ("tryg", "Forsikring og pension"),
    ("topdanmark", "Forsikring og pension"), ("alka", "Forsikring og pension"),
    ("codan", "Forsikring og pension"), ("alm brand", "Forsikring og pension"),
    ("if skade", "Forsikring og pension"), ("gf ", "Forsikring og pension"),
    ("pfa", "Forsikring og pension"), ("velliv", "Forsikring og pension"),
    ("danica", "Forsikring og pension"), ("sampension", "Forsikring og pension"),
    ("pensiondanmark", "Forsikring og pension"), ("pension", "Forsikring og pension"),

    # --- Sundhed ---
    ("apotek", "Sundhed"), ("læge", "Sundhed"), ("tandlæge", "Sundhed"),
    ("fysioterapi", "Sundhed"), ("kiropraktor", "Sundhed"), ("psykolog", "Sundhed"),
    ("sygesikring", "Sundhed"), ("danmark sygefor", "Sundhed"), ("optiker", "Sundhed"),
    ("briller", "Sundhed"), ("hospital", "Sundhed"), ("klinik", "Sundhed"),

    # --- Fritid og sport ---
    ("fitness", "Fritid og sport"), ("sats", "Fritid og sport"),
    ("puregym", "Fritid og sport"), ("loop", "Fritid og sport"),
    ("svømmehal", "Fritid og sport"), ("idrætsforening", "Fritid og sport"),
    ("kontingent", "Fritid og sport"), ("biograf", "Fritid og sport"),
    ("nordisk film", "Fritid og sport"), ("cinemaxx", "Fritid og sport"),
    ("tivoli", "Fritid og sport"), ("legoland", "Fritid og sport"),
    ("zoo", "Fritid og sport"), ("museum", "Fritid og sport"),
    ("billetlugen", "Fritid og sport"), ("ticketmaster", "Fritid og sport"),
    ("koncert", "Fritid og sport"), ("bibliotek", "Fritid og sport"),

    # --- Shopping ---
    ("h&m", "Shopping"), ("zara", "Shopping"), ("zalando", "Shopping"),
    ("bestseller", "Shopping"), ("magasin", "Shopping"), ("matas", "Shopping"),
    ("normal", "Shopping"), ("elgiganten", "Shopping"), ("power.dk", "Shopping"),
    ("proshop", "Shopping"), ("komplett", "Shopping"), ("ikea", "Shopping"),
    ("jysk", "Shopping"), ("ilva", "Shopping"), ("bauhaus", "Shopping"),
    ("silvan", "Shopping"), ("xl-byg", "Shopping"), ("stark", "Shopping"),
    ("harald nyborg", "Shopping"), ("biltema", "Shopping"), ("jem & fix", "Shopping"),
    ("amazon", "Shopping"), ("ebay", "Shopping"), ("temu", "Shopping"),
    ("shein", "Shopping"), ("wish", "Shopping"), ("aliexpress", "Shopping"),
    ("boghandel", "Shopping"), ("saxo.com", "Shopping"), ("blomster", "Shopping"),
    ("tøj", "Shopping"), ("sportmaster", "Shopping"), ("intersport", "Shopping"),

    # --- Børn og uddannelse ---
    ("daginstitution", "Børn og uddannelse"), ("vuggestue", "Børn og uddannelse"),
    ("børnehave", "Børn og uddannelse"), ("dagpleje", "Børn og uddannelse"),
    ("sfo", "Børn og uddannelse"), ("skolefritids", "Børn og uddannelse"),
    ("efterskole", "Børn og uddannelse"), ("studie", "Børn og uddannelse"),
    ("undervisning", "Børn og uddannelse"), ("kursus", "Børn og uddannelse"),

    # --- Gebyrer og renter ---
    ("gebyr", "Gebyrer og renter"), ("rentetilskrivning", "Gebyrer og renter"),
    ("renter", "Gebyrer og renter"), ("overtræk", "Gebyrer og renter"),
    ("valutatillæg", "Gebyrer og renter"), ("kortgebyr", "Gebyrer og renter"),
    ("betalingsservice", "Gebyrer og renter"), ("rykker", "Gebyrer og renter"),

    # --- Lån og afdrag ---
    ("afdrag", "Lån og afdrag"), ("billån", "Lån og afdrag"),
    ("forbrugslån", "Lån og afdrag"), ("studielån", "Lån og afdrag"),
    ("su-lån", "Lån og afdrag"), ("lånebetaling", "Lån og afdrag"),
    ("santander", "Lån og afdrag"), ("ekspres bank", "Lån og afdrag"),
    ("resurs bank", "Lån og afdrag"), ("basisbank", "Lån og afdrag"),
    ("lendo", "Lån og afdrag"), ("sparxpres", "Lån og afdrag"),
    ("kreditforening", "Lån og afdrag"),

    # --- Opsparing og investering ---
    ("opsparing", "Opsparing"), ("nordnet", "Opsparing"), ("saxo bank", "Opsparing"),
    ("aktier", "Opsparing"), ("investering", "Opsparing"), ("coinbase", "Opsparing"),
    ("puljeinvest", "Opsparing"),

    # --- Overførsler ---
    ("mobilepay", "MobilePay"), ("overførsel", "Overførsel"),
    ("overførelse", "Overførsel"), ("egen konto", "Overførsel"),
    ("straksoverførsel", "Overførsel"), ("transfer", "Overførsel"),

    # --- Indtægter ---
    ("løn", "Løn"), ("lønoverførsel", "Løn"), ("løn overførsel", "Løn"),
    ("lønudbetaling", "Løn"), ("løn udbetaling", "Løn"), ("salary", "Løn"),
    ("månedsløn", "Løn"), ("a-conto løn", "Løn"),
    ("udbetaling danmark", "Offentlige ydelser"), ("su ", "Offentlige ydelser"),
    ("su-", "Offentlige ydelser"), ("børne- og ungeydelse", "Offentlige ydelser"),
    ("børnepenge", "Offentlige ydelser"), ("boligstøtte", "Offentlige ydelser"),
    ("dagpenge", "Offentlige ydelser"), ("pensionsudbetaling", "Offentlige ydelser"),
    ("skattestyrelsen", "Refusion"), ("overskydende skat", "Refusion"),
    ("refusion", "Refusion"), ("tilbagebetaling", "Refusion"),
    ("udbytte", "Renter og udbytte"), ("renteindtægt", "Renter og udbytte"),
)


def default_categories() -> List[str]:
    """All categories used by the default rules, sorted, income last."""
    seen: List[str] = []
    for _keyword, category in DEFAULT_RULES:
        if category not in seen:
            seen.append(category)
    expenses = sorted(c for c in seen if c not in INCOME_CATEGORIES)
    income = [c for c in INCOME_CATEGORIES if c in seen]
    return expenses + income + [DEFAULT_CATEGORY]


class RuleSet(object):
    """An ordered collection of keyword rules."""

    def __init__(self, rules: Optional[Iterable[Sequence[str]]] = None,
                 default: str = DEFAULT_CATEGORY):
        self.default = default
        self._rules: List[Tuple[str, str]] = []
        self._compiled: List[Tuple[str, object, str]] = []
        for rule in (rules if rules is not None else DEFAULT_RULES):
            if not rule:
                continue
            keyword = str(rule[0]).strip()
            category = str(rule[1]).strip() if len(rule) > 1 else ""
            if keyword and category:
                self._rules.append((keyword, category))
        self._compile()

    # -- construction -----------------------------------------------------
    @classmethod
    def defaults(cls) -> "RuleSet":
        return cls(DEFAULT_RULES)

    @classmethod
    def from_rows(cls, rows: Iterable[Sequence[object]]) -> "RuleSet":
        """Read rules from ``[[keyword, category], ...]`` (e.g. a sheet)."""
        parsed: List[Tuple[str, str]] = []
        for row in rows:
            if not row:
                continue
            keyword = str(row[0]).strip() if row[0] is not None else ""
            category = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
            if not keyword or not category:
                continue
            if fold(keyword) in ("nogleord", "noegleord", "keyword", "søgeord"):
                continue  # header row
            parsed.append((keyword, category))
        return cls(parsed)

    def to_rows(self) -> List[List[str]]:
        return [[keyword, category] for keyword, category in self._rules]

    # -- persistence ------------------------------------------------------
    @classmethod
    def load(cls, path: str) -> "RuleSet":
        with open(path, "r", encoding="utf-8") as fp:
            if path.lower().endswith(".json"):
                data = json.load(fp)
                if isinstance(data, dict):
                    return cls(list(data.items()))
                return cls(data)
            rows = [line.rstrip("\n").split(";") for line in fp if line.strip()]
        return cls.from_rows(rows)

    def save(self, path: str) -> None:
        directory = os.path.dirname(os.path.abspath(path))
        if directory and not os.path.isdir(directory):
            os.makedirs(directory)
        with open(path, "w", encoding="utf-8") as fp:
            if path.lower().endswith(".json"):
                json.dump(self.to_rows(), fp, ensure_ascii=False, indent=2)
            else:
                fp.write("Nøgleord;Kategori\n")
                for keyword, category in self._rules:
                    fp.write("%s;%s\n" % (keyword, category))

    # -- matching ---------------------------------------------------------
    def _compile(self) -> None:
        compiled = []
        for keyword, category in self._rules:
            if keyword.lower().startswith("re:"):
                try:
                    pattern = re.compile(keyword[3:].strip(), re.IGNORECASE)
                except re.error:
                    continue
                compiled.append((len(keyword), "re", pattern, category))
                continue
            needle = fold(keyword)
            if not needle:
                continue
            if len(needle) <= SHORT_KEYWORD and " " not in needle:
                # A short keyword must not match inside a longer word:
                # "Irma" may not turn "firma" into groceries.  Digits after
                # the keyword are fine ("REMA1000", "NETTO8021").
                pattern = re.compile(r"(?<![a-z0-9])%s(?![a-z])"
                                     % re.escape(needle))
                compiled.append((len(needle), "word", pattern, category))
            else:
                compiled.append((len(needle), "kw", needle, category))
        # Longest keyword first so specific rules beat generic ones.
        compiled.sort(key=lambda item: item[0], reverse=True)
        self._compiled = [item[1:] for item in compiled]

    def categorise(self, text: object) -> str:
        """Return the category for ``text`` (or the default category)."""
        folded = fold(text)
        if not folded:
            return self.default
        for kind, needle, category in self._compiled:
            if kind == "kw":
                if needle in folded:
                    return category
            elif kind == "word":
                if needle.search(folded):
                    return category
            elif needle.search(str(text)):
                return category
        return self.default

    # -- editing ----------------------------------------------------------
    def add(self, keyword: str, category: str) -> None:
        keyword = keyword.strip()
        category = category.strip()
        if not keyword or not category:
            return
        folded = fold(keyword)
        self._rules = [(k, c) for k, c in self._rules if fold(k) != folded]
        self._rules.append((keyword, category))
        self._compile()

    def categories(self) -> List[str]:
        out: List[str] = []
        for _keyword, category in self._rules:
            if category not in out:
                out.append(category)
        expenses = sorted(c for c in out if c not in INCOME_CATEGORIES)
        income = [c for c in INCOME_CATEGORIES if c in out]
        return expenses + income + [self.default]

    def __len__(self) -> int:
        return len(self._rules)

    def __iter__(self):
        return iter(self._rules)


def user_rules_path() -> str:
    """Where the user's own rules are stored between sessions."""
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "BudgetFraCSV", "kategoriregler.csv")
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(
        os.path.expanduser("~"), ".config")
    return os.path.join(base, "budget-fra-csv", "kategoriregler.csv")
