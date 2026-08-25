// Category rules: keyword -> category.
//
// Port of budget_core/rules.py. A rule is a keyword that is matched
// case/accent insensitively against the transaction text. A keyword
// starting with "re:" is treated as a regular expression instead. Longer
// keywords are tried first so that "Circle K" wins over "K".

var DEFAULT_CATEGORY = "Ukategoriseret";

// A posting set to this category counts towards nothing: not an expense,
// not an income, not "uncategorised" either. For internal transfers,
// refunds that would otherwise double up a purchase, test postings and
// the like - things a budget should simply not see.
var IGNORED_CATEGORY = "Ignoreret";

// Keywords this short (or shorter) only match whole words.
var SHORT_KEYWORD = 4;

// Categories that are income even when the amount happens to be positive
// by accident (used for ordering and for the summary sheet).
var INCOME_CATEGORIES = ["Løn", "Offentlige ydelser", "Renter og udbytte", "Refusion", "Indtægt"];

// [keyword, category]. Danish first, then the international ones.
var DEFAULT_RULES = [

    // --- Dagligvarer ---
    ["netto", "Dagligvarer"], ["føtex", "Dagligvarer"], ["bilka", "Dagligvarer"],
    ["rema", "Dagligvarer"], ["lidl", "Dagligvarer"], ["aldi", "Dagligvarer"],
    ["kvickly", "Dagligvarer"], ["superbrugsen", "Dagligvarer"],
    ["dagli'brugsen", "Dagligvarer"], ["brugsen", "Dagligvarer"],
    ["coop", "Dagligvarer"], ["irma", "Dagligvarer"], ["meny", "Dagligvarer"],
    ["spar ", "Dagligvarer"], ["fakta", "Dagligvarer"], ["365discount", "Dagligvarer"],
    ["løvbjerg", "Dagligvarer"], ["min købmand", "Dagligvarer"],
    ["nemlig.com", "Dagligvarer"], ["bager", "Dagligvarer"], ["slagter", "Dagligvarer"],
    ["supermarked", "Dagligvarer"], ["grocery", "Dagligvarer"],

    // --- Restaurant og takeaway ---
    ["restaurant", "Restaurant"], ["bistro", "Restaurant"], ["cafe", "Restaurant"],
    ["café", "Restaurant"], ["kro ", "Restaurant"], ["pizza", "Restaurant"],
    ["sushi", "Restaurant"], ["burger", "Restaurant"], ["mcdonald", "Restaurant"],
    ["kebab", "Restaurant"], ["grill", "Restaurant"], ["wolt", "Restaurant"],
    ["just eat", "Restaurant"], ["hungry", "Restaurant"], ["takeaway", "Restaurant"],
    ["starbucks", "Restaurant"], ["espresso house", "Restaurant"],
    ["joe & the juice", "Restaurant"], ["baresso", "Restaurant"],
    ["bodega", "Restaurant"], ["bar ", "Restaurant"],

    // --- Transport ---
    ["dsb", "Transport"], ["rejsekort", "Transport"], ["metro", "Transport"],
    ["movia", "Transport"], ["arriva", "Transport"], ["taxa", "Transport"],
    ["taxi", "Transport"], ["dantaxi", "Transport"], ["4x35", "Transport"],
    ["uber", "Transport"], ["bolt.eu", "Transport"], ["donkey republic", "Transport"],
    ["brobizz", "Transport"], ["storebælt", "Transport"], ["øresund", "Transport"],
    ["færge", "Transport"], ["molslinjen", "Transport"], ["scandlines", "Transport"],
    ["parkering", "Transport"], ["easypark", "Transport"], ["parkpark", "Transport"],
    ["apcoa", "Transport"], ["q-park", "Transport"],
    ["circle k", "Transport"], ["shell", "Transport"], ["q8", "Transport"],
    ["ok benzin", "Transport"], ["uno-x", "Transport"], ["ingo", "Transport"],
    ["benzin", "Transport"], ["diesel", "Transport"], ["ladestander", "Transport"],
    ["clever", "Transport"], ["spirii", "Transport"], ["fdm", "Transport"],
    ["bilsyn", "Transport"], ["autoværksted", "Transport"], ["dæk", "Transport"],

    // --- Rejser og hotel ---
    ["hotel", "Rejser"], ["resort", "Rejser"], ["booking.com", "Rejser"],
    ["airbnb", "Rejser"], ["hotels.com", "Rejser"], ["expedia", "Rejser"],
    ["momondo", "Rejser"], ["travel", "Rejser"], ["flybillet", "Rejser"],
    ["ryanair", "Rejser"], ["norwegian", "Rejser"], ["lufthansa", "Rejser"],
    ["sas ", "Rejser"], ["airport", "Rejser"], ["lufthavn", "Rejser"],
    ["hostel", "Rejser"], ["camping", "Rejser"],

    // --- Bolig ---
    ["husleje", "Bolig"], ["boligselskab", "Bolig"], ["andelsbolig", "Bolig"],
    ["ejerforening", "Bolig"], ["grundejerforening", "Bolig"],
    ["realkredit", "Bolig"], ["totalkredit", "Bolig"], ["nykredit", "Bolig"],
    ["ejendomsskat", "Bolig"], ["grundskyld", "Bolig"], ["boliglån", "Bolig"],
    ["vicevært", "Bolig"], ["renovation", "Bolig"], ["skorstensfejer", "Bolig"],

    // --- El, vand og varme ---
    ["ørsted", "El, vand og varme"], ["norlys", "El, vand og varme"],
    ["andel energi", "El, vand og varme"], ["ewii", "El, vand og varme"],
    ["nrgi", "El, vand og varme"], ["energi fyn", "El, vand og varme"],
    ["hofor", "El, vand og varme"], ["fjernvarme", "El, vand og varme"],
    ["vandforsyning", "El, vand og varme"], ["elforsyning", "El, vand og varme"],
    ["naturgas", "El, vand og varme"], ["radius", "El, vand og varme"],
    ["vand og spildevand", "El, vand og varme"],

    // --- Telefon og internet ---
    ["yousee", "Telefon og internet"], ["tdc", "Telefon og internet"],
    ["telenor", "Telefon og internet"], ["telia", "Telefon og internet"],
    ["cbb", "Telefon og internet"], ["oister", "Telefon og internet"],
    ["hiper", "Telefon og internet"], ["stofa", "Telefon og internet"],
    ["fastspeed", "Telefon og internet"], ["callme", "Telefon og internet"],
    ["lebara", "Telefon og internet"], ["bredbånd", "Telefon og internet"],
    ["mobilabonnement", "Telefon og internet"],

    // --- Abonnementer ---
    ["netflix", "Abonnementer"], ["spotify", "Abonnementer"], ["hbo", "Abonnementer"],
    ["max.com", "Abonnementer"], ["disney", "Abonnementer"], ["viaplay", "Abonnementer"],
    ["tv 2 play", "Abonnementer"], ["tv2 play", "Abonnementer"],
    ["youtube", "Abonnementer"], ["icloud", "Abonnementer"],
    ["apple.com/bill", "Abonnementer"], ["google one", "Abonnementer"],
    ["google storage", "Abonnementer"], ["dropbox", "Abonnementer"],
    ["adobe", "Abonnementer"], ["microsoft", "Abonnementer"],
    ["openai", "Abonnementer"], ["anthropic", "Abonnementer"],
    ["claude.ai", "Abonnementer"], ["patreon", "Abonnementer"],
    ["audible", "Abonnementer"], ["storytel", "Abonnementer"],
    ["mofibo", "Abonnementer"], ["zwift", "Abonnementer"], ["strava", "Abonnementer"],
    ["licens", "Abonnementer"], ["abonnement", "Abonnementer"],
    ["dr licens", "Abonnementer"],

    // --- Forsikring og pension ---
    ["forsikring", "Forsikring og pension"], ["tryg", "Forsikring og pension"],
    ["topdanmark", "Forsikring og pension"], ["alka", "Forsikring og pension"],
    ["codan", "Forsikring og pension"], ["alm brand", "Forsikring og pension"],
    ["if skade", "Forsikring og pension"], ["gf ", "Forsikring og pension"],
    ["pfa", "Forsikring og pension"], ["velliv", "Forsikring og pension"],
    ["danica", "Forsikring og pension"], ["sampension", "Forsikring og pension"],
    ["pensiondanmark", "Forsikring og pension"], ["pension", "Forsikring og pension"],

    // --- Sundhed ---
    ["apotek", "Sundhed"], ["læge", "Sundhed"], ["tandlæge", "Sundhed"],
    ["fysioterapi", "Sundhed"], ["kiropraktor", "Sundhed"], ["psykolog", "Sundhed"],
    ["sygesikring", "Sundhed"], ["danmark sygefor", "Sundhed"], ["optiker", "Sundhed"],
    ["briller", "Sundhed"], ["hospital", "Sundhed"], ["klinik", "Sundhed"],

    // --- Fritid og sport ---
    ["fitness", "Fritid og sport"], ["sats", "Fritid og sport"],
    ["puregym", "Fritid og sport"], ["loop", "Fritid og sport"],
    ["svømmehal", "Fritid og sport"], ["idrætsforening", "Fritid og sport"],
    ["kontingent", "Fritid og sport"], ["biograf", "Fritid og sport"],
    ["nordisk film", "Fritid og sport"], ["cinemaxx", "Fritid og sport"],
    ["tivoli", "Fritid og sport"], ["legoland", "Fritid og sport"],
    ["zoo", "Fritid og sport"], ["museum", "Fritid og sport"],
    ["billetlugen", "Fritid og sport"], ["ticketmaster", "Fritid og sport"],
    ["koncert", "Fritid og sport"], ["bibliotek", "Fritid og sport"],

    // --- Shopping ---
    ["h&m", "Shopping"], ["zara", "Shopping"], ["zalando", "Shopping"],
    ["bestseller", "Shopping"], ["magasin", "Shopping"], ["matas", "Shopping"],
    ["normal", "Shopping"], ["elgiganten", "Shopping"], ["power.dk", "Shopping"],
    ["proshop", "Shopping"], ["komplett", "Shopping"], ["ikea", "Shopping"],
    ["jysk", "Shopping"], ["ilva", "Shopping"], ["bauhaus", "Shopping"],
    ["silvan", "Shopping"], ["xl-byg", "Shopping"], ["stark", "Shopping"],
    ["harald nyborg", "Shopping"], ["biltema", "Shopping"], ["jem & fix", "Shopping"],
    ["amazon", "Shopping"], ["ebay", "Shopping"], ["temu", "Shopping"],
    ["shein", "Shopping"], ["wish", "Shopping"], ["aliexpress", "Shopping"],
    ["boghandel", "Shopping"], ["saxo.com", "Shopping"], ["blomster", "Shopping"],
    ["tøj", "Shopping"], ["sportmaster", "Shopping"], ["intersport", "Shopping"],

    // --- Børn og uddannelse ---
    ["daginstitution", "Børn og uddannelse"], ["vuggestue", "Børn og uddannelse"],
    ["børnehave", "Børn og uddannelse"], ["dagpleje", "Børn og uddannelse"],
    ["sfo", "Børn og uddannelse"], ["skolefritids", "Børn og uddannelse"],
    ["efterskole", "Børn og uddannelse"], ["studie", "Børn og uddannelse"],
    ["undervisning", "Børn og uddannelse"], ["kursus", "Børn og uddannelse"],

    // --- Gebyrer og renter ---
    ["gebyr", "Gebyrer og renter"], ["rentetilskrivning", "Gebyrer og renter"],
    ["renter", "Gebyrer og renter"], ["overtræk", "Gebyrer og renter"],
    ["valutatillæg", "Gebyrer og renter"], ["kortgebyr", "Gebyrer og renter"],
    ["betalingsservice", "Gebyrer og renter"], ["rykker", "Gebyrer og renter"],

    // --- Lån og afdrag ---
    ["afdrag", "Lån og afdrag"], ["billån", "Lån og afdrag"],
    ["forbrugslån", "Lån og afdrag"], ["studielån", "Lån og afdrag"],
    ["su-lån", "Lån og afdrag"], ["lånebetaling", "Lån og afdrag"],
    ["santander", "Lån og afdrag"], ["ekspres bank", "Lån og afdrag"],
    ["resurs bank", "Lån og afdrag"], ["basisbank", "Lån og afdrag"],
    ["lendo", "Lån og afdrag"], ["sparxpres", "Lån og afdrag"],
    ["kreditforening", "Lån og afdrag"],

    // --- Opsparing og investering ---
    ["opsparing", "Opsparing"], ["nordnet", "Opsparing"], ["saxo bank", "Opsparing"],
    ["aktier", "Opsparing"], ["investering", "Opsparing"], ["coinbase", "Opsparing"],
    ["puljeinvest", "Opsparing"],

    // --- Overførsler ---
    ["mobilepay", "MobilePay"], ["overførsel", "Overførsel"],
    ["overførelse", "Overførsel"], ["egen konto", "Overførsel"],
    ["straksoverførsel", "Overførsel"], ["transfer", "Overførsel"],

    // --- Indtægter ---
    ["løn", "Løn"], ["lønoverførsel", "Løn"], ["løn overførsel", "Løn"],
    ["lønudbetaling", "Løn"], ["løn udbetaling", "Løn"], ["salary", "Løn"],
    ["månedsløn", "Løn"], ["a-conto løn", "Løn"],
    ["udbetaling danmark", "Offentlige ydelser"], ["su ", "Offentlige ydelser"],
    ["su-", "Offentlige ydelser"], ["børne- og ungeydelse", "Offentlige ydelser"],
    ["børnepenge", "Offentlige ydelser"], ["boligstøtte", "Offentlige ydelser"],
    ["dagpenge", "Offentlige ydelser"], ["pensionsudbetaling", "Offentlige ydelser"],
    ["skattestyrelsen", "Refusion"], ["overskydende skat", "Refusion"],
    ["refusion", "Refusion"], ["tilbagebetaling", "Refusion"],
    ["udbytte", "Renter og udbytte"], ["renteindtægt", "Renter og udbytte"],];

/** All categories used by the default rules, sorted, income last. */
function defaultCategories() {
  var seen = [];
  DEFAULT_RULES.forEach(function (rule) {
    if (seen.indexOf(rule[1]) === -1) seen.push(rule[1]);
  });
  var expenses = seen.filter(function (c) { return INCOME_CATEGORIES.indexOf(c) === -1; }).sort();
  var income = INCOME_CATEGORIES.filter(function (c) { return seen.indexOf(c) !== -1; });
  var result = expenses.concat(income);
  [DEFAULT_CATEGORY, IGNORED_CATEGORY].forEach(function (special) {
    if (result.indexOf(special) === -1) result.push(special);
  });
  return result;
}

/** An ordered collection of keyword rules. */
class RuleSet {
  constructor(rules, defaultCategory) {
    this.default = defaultCategory || DEFAULT_CATEGORY;
    this._rules = [];
    this._compiled = [];
    var source = rules !== undefined && rules !== null ? rules : DEFAULT_RULES;
    source.forEach((rule) => {
      if (!rule || !rule.length) return;
      var keyword = String(rule[0]).trim();
      var category = rule.length > 1 ? String(rule[1]).trim() : "";
      if (keyword && category) this._rules.push([keyword, category]);
    });
    this._compile();
  }

  static defaults() {
    return new RuleSet(DEFAULT_RULES);
  }

  /** Read rules from [[keyword, category], ...] (e.g. a sheet). */
  static fromRows(rows) {
    var parsed = [];
    rows.forEach(function (row) {
      if (!row) return;
      var keyword = row[0] !== null && row[0] !== undefined ? String(row[0]).trim() : "";
      var category = row.length > 1 && row[1] !== null && row[1] !== undefined ? String(row[1]).trim() : "";
      if (!keyword || !category) return;
      var f = fold(keyword);
      if (f === "nogleord" || f === "noegleord" || f === "keyword" || f === "sogeord") return; // header row
      parsed.push([keyword, category]);
    });
    return new RuleSet(parsed);
  }

  toRows() {
    return this._rules.map(function (r) { return [r[0], r[1]]; });
  }

  _compile() {
    var compiled = [];
    this._rules.forEach(function (rule) {
      var keyword = rule[0], category = rule[1];
      if (keyword.toLowerCase().indexOf("re:") === 0) {
        try {
          var pattern = new RegExp(keyword.slice(3).trim(), "i");
          compiled.push([keyword.length, "re", pattern, category]);
        } catch (e) {
          // invalid regex rule - skip it, same as the Python port.
        }
        return;
      }
      var needle = fold(keyword);
      if (!needle) return;
      if (needle.length <= SHORT_KEYWORD && needle.indexOf(" ") === -1) {
        // A short keyword must not match inside a longer word: "Irma" may
        // not turn "firma" into groceries. Digits after the keyword are
        // fine ("REMA1000", "NETTO8021").
        var escaped = needle.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        var wordPattern = new RegExp("(?<![a-z0-9])" + escaped + "(?![a-z])");
        compiled.push([needle.length, "word", wordPattern, category]);
      } else {
        compiled.push([needle.length, "kw", needle, category]);
      }
    });
    // Longest keyword first so specific rules beat generic ones.
    compiled.sort(function (a, b) { return b[0] - a[0]; });
    this._compiled = compiled.map(function (item) { return item.slice(1); });
  }

  /** Return the category for text (or the default category). */
  categorise(text) {
    var folded = fold(text);
    if (!folded) return this.default;
    for (var i = 0; i < this._compiled.length; i++) {
      var kind = this._compiled[i][0], needle = this._compiled[i][1], category = this._compiled[i][2];
      if (kind === "kw") {
        if (folded.indexOf(needle) !== -1) return category;
      } else if (kind === "word") {
        if (needle.test(folded)) return category;
      } else if (needle.test(String(text))) {
        return category;
      }
    }
    return this.default;
  }

  add(keyword, category) {
    keyword = keyword.trim();
    category = category.trim();
    if (!keyword || !category) return;
    var folded = fold(keyword);
    this._rules = this._rules.filter(function (r) { return fold(r[0]) !== folded; });
    this._rules.push([keyword, category]);
    this._compile();
  }

  categories() {
    var out = [];
    this._rules.forEach(function (r) {
      if (out.indexOf(r[1]) === -1) out.push(r[1]);
    });
    var expenses = out.filter(function (c) { return INCOME_CATEGORIES.indexOf(c) === -1; }).sort();
    var income = INCOME_CATEGORIES.filter(function (c) { return out.indexOf(c) !== -1; });
    var result = expenses.concat(income);
    [this.default, IGNORED_CATEGORY].forEach(function (special) {
      if (result.indexOf(special) === -1) result.push(special);
    });
    return result;
  }

  get length() {
    return this._rules.length;
  }

  [Symbol.iterator]() {
    return this._rules[Symbol.iterator]();
  }
}

if (typeof module !== "undefined") {
  module.exports = {
    RuleSet: RuleSet, DEFAULT_RULES: DEFAULT_RULES, DEFAULT_CATEGORY: DEFAULT_CATEGORY,
    IGNORED_CATEGORY: IGNORED_CATEGORY, INCOME_CATEGORIES: INCOME_CATEGORIES,
    defaultCategories: defaultCategories,
  };
  var fold = require("./TextUtils.gs").fold;
}
