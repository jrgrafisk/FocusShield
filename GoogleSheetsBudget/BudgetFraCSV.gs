// ============================================================================
// Budget fra CSV - til Google Sheets
//
// Dette er alle scriptets filer samlet i én, til nem installation: opret ét
// script-fil i Apps Script-editoren (Extensions > Apps Script), slet den
// forududfyldte kode, og indsæt hele denne fils indhold. Se README.md for
// den fulde installationsvejledning (kun denne fil + ImportDialog.html skal
// oprettes manuelt).
//
// Kildekoden ligger opdelt i enkeltfiler under src/, hvis du vil læse eller
// ændre den - denne fil genereres derfra (build.py) og bør ikke redigeres
// direkte.
// ============================================================================

// ---- TextUtils.gs -------------------------------------------------------------

// Small text helpers shared by the parsing/detection modules.
//
// Port of budget_core/textutils.py - kept behaviourally identical.

var _LETTER_MAP = {
  "æ": "ae", "ø": "o", "ß": "ss", "đ": "d", "ł": "l", "þ": "th",
};

var _WS_RE = /\s+/g;

/**
 * Normalise text for keyword comparisons: case, accents and Nordic
 * letters are removed so "Café Øst", "CAFE OST" and "cafe øst" compare
 * equal.
 */
function fold(text) {
  if (text === null || text === undefined) return "";
  var s = String(text);
  s = s.normalize("NFKD").replace(/[\u0300-\u036f]/g, "");
  s = s.toLowerCase();
  for (var src in _LETTER_MAP) {
    if (s.indexOf(src) !== -1) {
      s = s.split(src).join(_LETTER_MAP[src]);
    }
  }
  return s.replace(_WS_RE, " ").trim();
}

/** Collapse whitespace but keep the original characters. */
function squeeze(text) {
  if (text === null || text === undefined) return "";
  return String(text).replace(_WS_RE, " ").trim();
}

/** True when every cell in row is empty or whitespace. */
function isBlankRow(row) {
  for (var i = 0; i < row.length; i++) {
    if (String(row[i]).trim()) return false;
  }
  return true;
}

// ---- Parsing.gs ---------------------------------------------------------------

// Tolerant parsing of the amounts and dates found in bank CSV exports.
//
// Port of budget_core/parsing.py - deliberately forgiving: real exports
// contain currency symbols, non breaking spaces, trailing minus signs,
// parentheses for negative numbers, several thousand separators and about
// a dozen date layouts. Nothing throws - unparsable input returns null.

var _CURRENCY_RE = /(?:^|(?<=[\s\d]))(?:dkk|sek|nok|isk|eur|usd|gbp|chf|pln|czk|huf|kr\.?|kroner)(?=$|[\s\d.,+-])|[€$£¥₤]/gi;
var _DC_SUFFIX_RE = /\s*(?:dr|db|cr|kr)\.?$/i;
var _NUMBER_BODY_RE = /^[0-9]+(?:[.,'’][0-9]+)*$/;
var _GROUPED_RE_CACHE = {};

function _cleanNumberText(raw) {
  if (raw === null || raw === undefined) return [null, false];
  if (typeof raw === "boolean") return [null, false];
  if (typeof raw === "number") {
    return [String(Math.abs(raw)), raw < 0];
  }

  var s = String(raw).trim();
  if (!s) return [null, false];

  // Normalise exotic spaces and dashes.
  [" ", " ", " ", " "].forEach(function (ch) { s = s.split(ch).join(" "); });
  ["−", "–", "—"].forEach(function (ch) { s = s.split(ch).join("-"); });

  s = s.replace(_CURRENCY_RE, " ");
  var negative = false;

  // "1.234,56 DR" / "1.234,56 CR"
  var m = _DC_SUFFIX_RE.exec(s);
  if (m) {
    var suffix = m[0].trim().toLowerCase();
    if (suffix.indexOf("dr") === 0 || suffix.indexOf("db") === 0) negative = true;
    s = s.slice(0, m.index);
  }

  s = s.trim();
  if (s.length >= 2 && s[0] === "(" && s[s.length - 1] === ")") {
    negative = !negative;
    s = s.slice(1, -1).trim();
  }

  if (s[0] === "-" || s[0] === "+") {
    negative = negative !== (s[0] === "-");
    s = s.slice(1).trim();
  } else if (s[s.length - 1] === "-" || s[s.length - 1] === "+") {
    var last = s[s.length - 1];
    negative = negative !== (last === "-");
    s = s.slice(0, -1).trim();
  }

  s = s.split(" ").join("").split("'").join("").split("’").join("");
  if (!s || !_NUMBER_BODY_RE.test(s)) return [null, false];
  return [s, negative];
}

function _isGrouped(body, sep) {
  var pattern = _GROUPED_RE_CACHE[sep];
  if (!pattern) {
    var escaped = sep.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    pattern = new RegExp("^\\d{1,3}(?:" + escaped + "\\d{3})+$");
    _GROUPED_RE_CACHE[sep] = pattern;
  }
  return pattern.test(body);
}

function _valueDecimalSeparator(body) {
  var hasDot = body.indexOf(".") !== -1;
  var hasComma = body.indexOf(",") !== -1;
  if (hasDot && hasComma) {
    return body.lastIndexOf(".") > body.lastIndexOf(",") ? "." : ",";
  }
  var pairs = [[",", "."], [".", ","]];
  for (var i = 0; i < pairs.length; i++) {
    var sep = pairs[i][0], other = pairs[i][1];
    if (body.indexOf(sep) === -1) continue;
    var count = body.split(sep).length - 1;
    if (count > 1) return other; // repeated separator can only be grouping
    var tail = body.slice(body.lastIndexOf(sep) + 1);
    if (tail.length === 3 && _isGrouped(body, sep)) return other; // "1.234"/"1,234"
    return sep;
  }
  return null;
}

function _assemble(body, decimal) {
  var grouping = decimal === "." ? "," : ".";
  body = body.split(grouping).join("");
  var parts = body.split(decimal);
  if (parts.length === 1) return parts[0];
  if (parts.length > 2) {
    var lastPart = parts[parts.length - 1];
    if (lastPart.length === 1 || lastPart.length === 2) {
      return parts.slice(0, -1).join("") + "." + lastPart;
    }
    return parts.join("");
  }
  return parts[0] + "." + parts[1];
}

/**
 * Parse raw into a float, or return null.
 *
 * decimal forces the decimal separator ("," or "."). When omitted the
 * separator is guessed from the value itself.
 */
function parseAmount(raw, decimal) {
  var cleaned = _cleanNumberText(raw);
  var body = cleaned[0], negative = cleaned[1];
  if (body === null) return null;
  var sep = decimal || _valueDecimalSeparator(body);
  var text = sep === null ? body : _assemble(body, sep);
  var value = parseFloat(text);
  if (isNaN(value)) return null;
  return negative ? -value : value;
}

/** True when raw can be read as a number. */
function looksLikeAmount(raw) {
  return _cleanNumberText(raw)[0] !== null;
}

/** Decide whether a column uses "," or "." as decimal separator. */
function detectDecimalSeparator(samples) {
  var strong = { ",": 0, ".": 0 };
  var weak = { ",": 0, ".": 0 };
  for (var i = 0; i < samples.length; i++) {
    var body = _cleanNumberText(samples[i])[0];
    if (!body) continue;
    var hasDot = body.indexOf(".") !== -1;
    var hasComma = body.indexOf(",") !== -1;
    if (hasDot && hasComma) {
      strong[body.lastIndexOf(".") > body.lastIndexOf(",") ? "." : ","] += 2;
      continue;
    }
    var pairs = [[",", "."], [".", ","]];
    for (var p = 0; p < pairs.length; p++) {
      var sep = pairs[p][0], other = pairs[p][1];
      if (body.indexOf(sep) === -1) continue;
      var count = body.split(sep).length - 1;
      if (count > 1) {
        strong[other] += 1;
      } else {
        var tail = body.slice(body.lastIndexOf(sep) + 1);
        if (tail.length === 1 || tail.length === 2 || tail.length > 3) {
          strong[sep] += 1;
        } else if (_isGrouped(body, sep)) {
          weak[other] += 1;
        } else {
          strong[sep] += 1;
        }
      }
    }
  }
  if (strong[","] !== strong["."]) return strong[","] > strong["."] ? "," : ".";
  if (weak[","] !== weak["."]) return weak[","] > weak["."] ? "," : ".";
  if (strong[","] || weak[","]) return ",";
  return null;
}

// --------------------------------------------------------------------------
// Dates
// --------------------------------------------------------------------------

var MONTH_NAMES_DA = [
  "januar", "februar", "marts", "april", "maj", "juni",
  "juli", "august", "september", "oktober", "november", "december",
];

var _MONTHS = {};
(function () {
  MONTH_NAMES_DA.forEach(function (name, i) {
    _MONTHS[name] = i + 1;
    _MONTHS[name.slice(0, 3)] = i + 1;
  });
  var en = ["january", "february", "march", "april", "may", "june", "july",
    "august", "september", "october", "november", "december"];
  en.forEach(function (name, i) {
    _MONTHS[name] = i + 1;
    _MONTHS[name.slice(0, 3)] = i + 1;
  });
  _MONTHS["mar"] = 3; _MONTHS["okt"] = 10; _MONTHS["oct"] = 10;
  _MONTHS["sept"] = 9; _MONTHS["des"] = 12;
})();

var _TOKEN_RE = /[0-9]+|[a-zæøå]+/gi;
var _COMPACT_RE = /^(\d{4})(\d{2})(\d{2})$/;
var _COMPACT_DMY_RE = /^(\d{2})(\d{2})(\d{4})$/;

function _makeDate(year, month, day) {
  if (year < 100) year += year < 70 ? 2000 : 1900;
  if (!(year >= 1900 && year <= 2200 && month >= 1 && month <= 12 && day >= 1 && day <= 31)) {
    return null;
  }
  var d = new Date(Date.UTC(year, month - 1, day));
  if (d.getUTCFullYear() !== year || d.getUTCMonth() !== month - 1 || d.getUTCDate() !== day) {
    return null; // e.g. 31. februar
  }
  return d;
}

/**
 * Parse a date written in (almost) any common layout. Returns a UTC-anchored
 * Date (midnight) or null. Understood: "2024-05-31", "31-05-2024",
 * "31/05/24", "31.05.2024", "20240531", "31052024", "31. maj 2024",
 * "May 31, 2024" and any of those followed by a clock time.
 */
function parseDate(raw, dayfirst) {
  if (dayfirst === undefined) dayfirst = true;
  if (raw === null || raw === undefined) return null;
  if (raw instanceof Date) {
    return new Date(Date.UTC(raw.getFullYear(), raw.getMonth(), raw.getDate()));
  }

  var s = String(raw).trim();
  if (!s) return null;

  var compact = s.split(" ").join("");
  var m = _COMPACT_RE.exec(compact);
  if (m) {
    var result = _makeDate(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));
    if (result) return result;
  }
  m = _COMPACT_DMY_RE.exec(compact);
  if (m) {
    var r = _makeDate(parseInt(m[3], 10), parseInt(m[2], 10), parseInt(m[1], 10)) ||
      _makeDate(parseInt(m[3], 10), parseInt(m[1], 10), parseInt(m[2], 10));
    if (r) return r;
  }

  var tokens = s.match(_TOKEN_RE) || [];
  if (tokens.length < 3) return null;

  var monthFromName = null;
  var numbers = [];
  for (var i = 0; i < tokens.length; i++) {
    var tok = tokens[i];
    if (/^\d+$/.test(tok)) {
      if (tok.length > 4) return null;
      numbers.push(parseInt(tok, 10));
    } else {
      var name = fold(tok).replace(/\.$/, "");
      if (_MONTHS.hasOwnProperty(name) && monthFromName === null) {
        monthFromName = _MONTHS[name];
      } else if (numbers.length) {
        break; // a word after the numbers - stop, e.g. "31-05-2024 kl"
      }
    }
    if (monthFromName !== null && numbers.length >= 2) break;
    if (monthFromName === null && numbers.length >= 3) break;
  }

  if (monthFromName !== null) {
    if (numbers.length < 2) return null;
    var a0 = numbers[0], b0 = numbers[1];
    var year0 = a0 > 31 ? a0 : b0;
    var day0 = a0 > 31 ? b0 : a0;
    return _makeDate(year0, monthFromName, day0);
  }

  if (numbers.length < 3) return null;
  var a = numbers[0], b = numbers[1], c = numbers[2];
  if (a > 31) return _makeDate(a, b, c); // year first
  if (a > 12) return _makeDate(c, b, a);
  if (b > 12) return _makeDate(c, a, b);
  if (dayfirst) return _makeDate(c, b, a) || _makeDate(c, a, b);
  return _makeDate(c, a, b) || _makeDate(c, b, a);
}

/** True when raw can be read as a date in some layout. */
function looksLikeDate(raw) {
  return parseDate(raw) !== null;
}

/** Decide between "31/05/2024" (day first) and "05/31/2024". */
function detectDayfirst(samples) {
  var dayVotes = 0, monthVotes = 0;
  for (var i = 0; i < samples.length; i++) {
    var raw = samples[i];
    if (raw === null || raw === undefined) continue;
    var s = String(raw).trim();
    if (!s || _COMPACT_RE.test(s.split(" ").join(""))) continue;
    var numbers = (s.match(_TOKEN_RE) || [])
      .filter(function (t) { return /^\d+$/.test(t) && t.length <= 4; })
      .map(function (t) { return parseInt(t, 10); });
    if (numbers.length < 3) continue;
    var a = numbers[0], b = numbers[1];
    if (a > 31) continue; // ISO style, tells us nothing
    if (a > 12) dayVotes += 1;
    else if (b > 12) monthVotes += 1;
  }
  return !(monthVotes > dayVotes);
}

/** Date -> "2024-05" (sorts chronologically). */
function monthKey(date) {
  var y = date.getUTCFullYear();
  var m = date.getUTCMonth() + 1;
  return y + "-" + (m < 10 ? "0" + m : String(m));
}

/** "2024-05" -> "maj 2024" for human readable headings. */
function monthLabel(key) {
  var parts = key.split("-");
  if (parts.length !== 2) return key;
  var idx = parseInt(parts[1], 10) - 1;
  if (isNaN(idx) || idx < 0 || idx >= MONTH_NAMES_DA.length) return key;
  return MONTH_NAMES_DA[idx] + " " + parts[0];
}

// ---- Transfers.gs -------------------------------------------------------------

// Recognise transactions that move money between the user's own accounts.
//
// Port of budget_core/transfers.py. Bank exports have no reliable flag for
// "this is not real income" - a transfer from a savings account into the
// checking account looks exactly like a salary payment. The only thing
// that gives it away is the account number written in the text
// ("Overført fra 1234567890", "Til konto 5301-1234567890"). If the user
// tells us which account numbers are their own (on the Konti sheet), we
// can catch those and keep them out of the budget entirely.

// Danish account numbers run to 10 digits (plus a 4 digit reg. number).
// Anything shorter is too likely to collide with a date or a receipt
// number, so we simply never match on it.
var MIN_ACCOUNT_DIGITS = 6;

var _SPLIT_RE = /[,;/]+/;

/** Strip everything but digits: "1234 5678901" -> "12345678901". */
function normalizeAccountNumber(raw) {
  return String(raw || "").replace(/\D/g, "");
}

/** The account numbers in one Konti cell (comma/semicolon separated). */
function parseAccountNumbers(raw) {
  if (!raw) return [];
  var numbers = [];
  String(raw).split(_SPLIT_RE).forEach(function (part) {
    var digits = normalizeAccountNumber(part);
    if (digits.length >= MIN_ACCOUNT_DIGITS) numbers.push(digits);
  });
  return numbers;
}

/**
 * True when text contains one of accountNumbers.
 *
 * Matches a single token whose digits contain the account number (so
 * "5301-1234567890" and plain "1234567890" both match), and also two
 * adjacent pure-digit tokens concatenated (so "5301 1234567890", reg.
 * number and account number split by a space, matches too).
 */
function textMentionsAccount(text, accountNumbers) {
  var numbers = (accountNumbers || []).filter(function (n) { return n; });
  if (!numbers.length) return false;
  var tokens = squeeze(text).split(" ");
  var tokenDigits = tokens.map(normalizeAccountNumber);
  for (var i = 0; i < tokenDigits.length; i++) {
    var digits = tokenDigits[i];
    if (digits.length < MIN_ACCOUNT_DIGITS) continue;
    if (numbers.some(function (n) { return digits.indexOf(n) !== -1; })) return true;
  }
  for (var j = 0; j < tokens.length - 1; j++) {
    if (/^\d+$/.test(tokens[j]) && /^\d+$/.test(tokens[j + 1])) {
      var combined = tokens[j] + tokens[j + 1];
      if (numbers.some(function (n) { return combined.indexOf(n) !== -1; })) return true;
    }
  }
  return false;
}

// ---- Merchant.gs --------------------------------------------------------------

// Find the recognisable shop or company inside a transaction text.
//
// Port of budget_core/merchant.py. Bank exports rarely write just "Netto".
// They write things like "Dankort-nota 4711 NETTO 8021 KØBENHAVN 21.05 kl.
// 17.42". Amounts, dates, times, card and receipt numbers make every
// single line unique, so grouping on the raw text gives one group per
// transaction. This module strips that noise, so lines like that group as
// "NETTO KØBENHAVN" - and so a saved rule matches next month's purchase in
// the same shop.

// Words that say something about the payment, not about who was paid.
// Banks write both the Danish "kort" and the English "card" spelling for
// most of these ("Debitkort" vs "Debitcard"), so every compound is listed
// in both forms.
var NOISE_WORDS = {};
[
  "dankort", "nota", "dankortnota", "kortkob", "kortkoeb", "kort", "kortnr",
  "visa", "mastercard", "maestro", "eurocard", "kreditkort", "betalingskort",
  "debitkort", "chipkort", "haevekort", "kontokort", "korttype",
  "kortbetaling", "korttransaktion", "kortnota", "kortholder",
  "kortoplysninger", "kortkobsnota",
  "kreditcard", "debitcard", "betalingscard", "chipcard", "bankcard",
  "kontocard", "cardtype", "cardbetaling", "cardtransaktion", "cardnota",
  "cardholder", "cardoplysninger", "cardkobsnota", "cardkob", "cardkoeb",
  "cardnr", "card",
  "americanexpress", "amex", "dinersclub", "diners", "jcb", "unionpay",
  "bs", "pbs", "betaling", "betalingsservice", "indbetalingskort",
  "ref", "refnr", "reference", "referencenr", "id", "idnr", "nr", "no",
  "kl", "den", "d", "dato", "tid", "kvittering", "bilag", "faktura", "fakt",
  "kob", "koeb", "purchase", "payment", "pos", "atm", "automat", "netbank",
  "mobilbank", "onlinebank", "udland", "udl", "valutakurs", "kurs",
  "dk", "dkk", "sek", "nok", "eur", "usd", "gbp", "kr", "kroner",
  "fra", "til", "med", "m", "v", "og", "af", "pa", "konto", "kontonr",
  "konto-nr",
  "transaktion", "postering", "posteringstekst", "tekst", "note",
  "aut", "auto", "automatisk", "straks", "straksoverforsel",
].forEach(function (w) { NOISE_WORDS[w] = true; });

// Tokens that are a date ("21.05", "2025-05-21") or a time ("17.42", "17:42").
var _DATE_RE = /^\d{1,4}[./-]\d{1,2}([./-]\d{2,4})?\.?$/;
var _TIME_RE = /^\d{1,2}[:.]\d{2}$/;
// JS \w/\W are ASCII-only even with the "u" flag, unlike Python's
// Unicode-aware \w - use \p{L}\p{N} (Unicode letters/digits) instead so
// Danish letters (æøå) count as "word" characters, matching the Python
// behaviour exactly.
var _LETTER_RE = /\p{L}/u;
var _TRIM_RE = /^[^\p{L}\p{N}_&]+|[^\p{L}\p{N}_&%]+$/gu;

// How much of a merchant name a saved rule should use.
var RULE_WORDS = 3;
var RULE_CHARS = 32;
var MIN_NAME_LENGTH = 3;

/** True when a token says something about the payment, not the payee. */
function _isNoise(token) {
  if (!_LETTER_RE.test(token)) return true; // pure numbers, "//", "-", "***"
  if (_DATE_RE.test(token) || _TIME_RE.test(token)) return true;
  if (looksLikeAmount(token)) return true; // "245,00", "1.234,56", "99kr"
  var folded = fold(token);
  if (NOISE_WORDS.hasOwnProperty(folded)) return true;
  // "Dankort-nota", "kortkøb/visa": noise when every part is noise.
  var parts = folded.split(/[-/.]/).filter(function (p) { return p; });
  if (parts.length > 1 && parts.every(function (p) { return NOISE_WORDS.hasOwnProperty(p); })) {
    return true;
  }
  var digits = (token.match(/\d/g) || []).length;
  if (digits && digits * 2 >= token.length) return true; // "8021", "4711a", "12-34"
  return false;
}

/** "&", "-", "/" and friends: part of a name, not a word of their own. */
function _isConnector(token) {
  return token.length <= 2 && !_LETTER_RE.test(token) && !/\d/.test(token);
}

/** [tokens, keepFlags, connectorFlags] for the words of text. */
function _analyse(text) {
  var tokens = squeeze(text).split(" ").filter(function (t) { return t; });
  var keep = [];
  var connectors = [];
  tokens.forEach(function (token) {
    var trimmed = token.replace(_TRIM_RE, "") || token;
    var connector = _isConnector(token);
    connectors.push(connector);
    keep.push(!connector && !_isNoise(trimmed));
  });
  return [tokens, keep, connectors];
}

/** The tokens of text that actually name someone. */
function _cleanTokens(text) {
  var r = _analyse(text);
  var tokens = r[0], keep = r[1];
  var out = [];
  for (var i = 0; i < tokens.length; i++) {
    if (keep[i]) out.push(tokens[i].replace(_TRIM_RE, "") || tokens[i]);
  }
  return out;
}

/** The shop/company part of a transaction text (original spelling). */
function merchantName(text) {
  var tokens = _cleanTokens(text);
  var name = tokens.join(" ").trim();
  if (name.length < MIN_NAME_LENGTH) return squeeze(text);
  return name;
}

/** A grouping key: two texts from the same shop give the same key. */
function merchantKey(text) {
  return fold(merchantName(text)) || fold(text);
}

/**
 * A keyword to save as a rule - short enough to match next month too.
 *
 * Rules match as plain substrings, so the keyword has to be an unbroken
 * piece of the original text: "JOE & THE JUICE" must keep its "&".
 */
function ruleKeyword(text) {
  var r = _analyse(text);
  var tokens = r[0], keep = r[1], connectors = r[2];
  if (!keep.some(function (k) { return k; })) return squeeze(text).slice(0, RULE_CHARS);
  var first = keep.indexOf(true);
  var last = first;
  var words = 0;
  for (var index = first; index < tokens.length; index++) {
    if (connectors[index]) continue; // "&" does not end the name
    if (!keep[index]) break; // a shop number ends it: "NETTO 8021"
    words += 1;
    if (words > RULE_WORDS) break;
    last = index;
  }
  var keyword = tokens.slice(first, last + 1).join(" ").replace(_TRIM_RE, "");
  if (keyword.length > RULE_CHARS) keyword = keyword.slice(0, RULE_CHARS).replace(/\s+$/, "");
  return keyword || squeeze(text).slice(0, RULE_CHARS);
}

// ---- Rules.gs -----------------------------------------------------------------

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

// ---- CsvSniff.gs --------------------------------------------------------------

// Turn "some CSV text" into a clean table.
//
// Port of budget_core/csvsniff.py. Handles the parts of real world exports
// that break naive CSV parsing: sep=; hint lines, semicolon/comma/tab/pipe
// separators, bank preambles above the real header, missing headers and
// ragged rows. Character-encoding detection (BOM/UTF-16/cp1252) happens in
// the browser before the text ever reaches here - see Encoding.gs - so
// this module only works with an already-decoded JS string, unlike the
// Python version which also reads raw bytes.

var CANDIDATE_DELIMITERS = [";", ",", "\t", "|"];

// Words that give away a header row. Compared token by token on folded text.
var HEADER_WORDS = {};
[
  "dato", "date", "datum", "bogfort", "bogforingsdato", "posteringsdato",
  "transaktionsdato", "valor", "valordato", "rentedato", "tid", "time",
  "tekst", "text", "beskrivelse", "description", "posteringstekst",
  "narrative", "details", "detaljer", "modtager", "afsender", "payee",
  "merchant", "reference", "memo", "note", "navn", "name", "titel",
  "belob", "amount", "sum", "value", "betrag", "montant", "beloeb",
  "saldo", "balance", "beholdning", "konto", "account", "kontonummer",
  "valuta", "currency", "type", "kategori", "category", "status",
  "debet", "kredit", "debit", "credit", "ind", "ud", "indsat", "haevet",
  "indbetaling", "udbetaling", "withdrawal", "deposit", "posting",
  "transaktion", "transaction", "art", "nummer", "number", "id",
  "afstemt", "gebyr", "fee",
].forEach(function (w) { HEADER_WORDS[w] = true; });

// Sniffing never looks at more than this many rows.
var SAMPLE_ROWS = 200;

/** A rectangular table of strings plus what we learned while reading it. */
class Table {
  constructor(rows, header, opts) {
    opts = opts || {};
    this.rows = rows;
    this.header = header || null;
    this.encoding = opts.encoding || "utf-8";
    this.delimiter = opts.delimiter || ";";
    this.preamble = opts.preamble || [];
    this.notes = opts.notes || [];
    this.source = opts.source || "";
    var widths = rows.map(function (r) { return r.length; });
    widths.push(header ? header.length : 0);
    this.n_columns = Math.max.apply(null, widths.concat([0]));
  }

  /** All values of one column (as strings). */
  column(index, limit) {
    var rows = limit === undefined || limit === null ? this.rows : this.rows.slice(0, limit);
    return rows.map(function (r) { return index < r.length ? r[index] : ""; });
  }

  header_name(index) {
    if (this.header && index < this.header.length) return this.header[index].trim();
    return "Kolonne " + (index + 1);
  }

  header_names() {
    var out = [];
    for (var i = 0; i < this.n_columns; i++) out.push(this.header_name(i));
    return out;
  }

  sample(count) {
    return this.rows.slice(0, count === undefined ? 5 : count);
  }

  get length() {
    return this.rows.length;
  }

  describe() {
    return this.rows.length + " rækker, " + this.n_columns + " kolonner, tegnsæt " +
      this.encoding + ", skilletegn " +
      (this.delimiter === "\t" ? "TAB" : "\"" + this.delimiter + "\"");
  }

  /** Build a table from already split rows (CSV rows or sheet cells). */
  static fromRows(rawRows, hasHeader, opts) {
    opts = opts || {};
    var rows = rawRows.map(function (row) {
      return row.map(function (c) { return c === null || c === undefined ? "" : String(c); });
    });
    var notes = (opts.notes || []).slice();
    var found = _findStart(rows);
    var start = found[0], nColumns = found[1];
    var preamble = rows.slice(0, start).filter(function (r) { return !isBlankRow(r); });
    if (preamble.length) notes.push("Sprang " + preamble.length + " indledende linje(r) over.");

    var body = rows.slice(start).filter(function (r) { return !isBlankRow(r); });
    if (!body.length) {
      return new Table([], null, { preamble: preamble, notes: notes, encoding: opts.encoding,
        delimiter: opts.delimiter, source: opts.source });
    }

    var header = null;
    var first = body[0];
    var useHeader = hasHeader === undefined || hasHeader === null ? looksLikeHeader(first, body.slice(1, 21)) : hasHeader;
    if (useHeader) {
      header = first.map(function (c) { return String(c).trim(); });
      body = body.slice(1);
      header = _uniquify(header);
    } else {
      notes.push("Ingen overskriftsrække fundet - kolonnerne navngives automatisk.");
    }

    var ragged = 0;
    var fixed = [];
    body.forEach(function (row) {
      if (row.length < nColumns) {
        if (row.length < 2) return;
        ragged += 1;
        row = row.concat(new Array(nColumns - row.length).fill(""));
      }
      fixed.push(row.slice());
    });
    if (ragged) notes.push(ragged + " række(r) havde færre kolonner end resten og blev fyldt ud.");

    return new Table(fixed, header, { preamble: preamble, notes: notes, encoding: opts.encoding,
      delimiter: opts.delimiter, source: opts.source });
  }
}

/** Make header names unique and non empty. */
function _uniquify(names) {
  var seen = {};
  var out = [];
  names.forEach(function (name, i) {
    name = (name || "").trim() || "Kolonne " + (i + 1);
    if (seen.hasOwnProperty(name)) {
      seen[name] += 1;
      name = name + " (" + seen[name] + ")";
    } else {
      seen[name] = 1;
    }
    out.push(name);
  });
  return out;
}

// --------------------------------------------------------------------------
// Delimiter
// --------------------------------------------------------------------------

/**
 * Split one CSV line respecting double-quoted fields (RFC 4180 style, "" is
 * an escaped quote). Mirrors Python's csv.reader for a single delimiter.
 */
function _splitCsvLine(line, delimiter) {
  var fields = [];
  var field = "";
  var inQuotes = false;
  var atFieldStart = true;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; }
      } else {
        field += ch;
      }
    } else if (atFieldStart && ch === " ") {
      continue; // skipinitialspace: spaces right after the delimiter (or line start)
    } else if (atFieldStart && ch === '"') {
      inQuotes = true;
      atFieldStart = false;
    } else if (ch === delimiter) {
      fields.push(field);
      field = "";
      atFieldStart = true;
    } else {
      field += ch;
      atFieldStart = false;
    }
  }
  fields.push(field);
  return fields;
}

function _parseCsv(text, delimiter) {
  var lines = text.split("\n");
  if (lines.length && lines[lines.length - 1] === "") lines.pop();
  return lines.map(function (line) { return _splitCsvLine(line.replace(/\r$/, ""), delimiter); });
}

/** Pick the separator that yields the most consistent, widest table. */
function detectDelimiter(text, candidates) {
  candidates = candidates || CANDIDATE_DELIMITERS;
  var lines = text.split(/\r\n|\r|\n/).filter(function (l) { return l.trim(); }).slice(0, 60);
  if (!lines.length) return ";";

  var best = null;
  var bestScore = 0.0;
  candidates.forEach(function (delimiter) {
    var rows = lines.map(function (line) { return _splitCsvLine(line, delimiter); });
    var widths = rows.filter(function (r) { return r.some(function (c) { return String(c).trim(); }); })
      .map(function (r) { return r.length; });
    if (!widths.length) return;
    var counts = {};
    widths.forEach(function (w) { counts[w] = (counts[w] || 0) + 1; });
    var width = 0, freq = -1;
    Object.keys(counts).forEach(function (w) {
      var wi = parseInt(w, 10);
      if (counts[w] > freq || (counts[w] === freq && wi > width)) { freq = counts[w]; width = wi; }
    });
    if (width < 2) return;
    var score = (freq / widths.length) * Math.min(width, 12);
    if (score > bestScore + 1e-9) { bestScore = score; best = delimiter; }
  });
  return best || ";";
}

var _SEP_HINT_PREFIXES = ["sep=", "SEP="];

/** Consume Excel's sep=; first line if present. */
function _stripSepHint(text) {
  var idx = text.indexOf("\n");
  if (idx === -1) return [text, null];
  var first = text.slice(0, idx);
  var rest = text.slice(idx + 1);
  var stripped = first.trim().replace(/^﻿/, "");
  for (var i = 0; i < _SEP_HINT_PREFIXES.length; i++) {
    var prefix = _SEP_HINT_PREFIXES[i];
    if (stripped.indexOf(prefix) === 0 && stripped.length === prefix.length + 1) {
      return [rest, stripped[prefix.length]];
    }
  }
  return [text, null];
}

// --------------------------------------------------------------------------
// Header detection
// --------------------------------------------------------------------------

/** Decide whether row is a header rather than the first data row. */
function looksLikeHeader(row, following) {
  var cells = row.map(function (c) { return String(c).trim(); });
  var nonEmpty = cells.filter(function (c) { return c; });
  if (!nonEmpty.length) return false;

  if (cells.some(function (c) { return looksLikeDate(c); })) return false;

  var hits = 0;
  nonEmpty.forEach(function (cell) {
    var folded = fold(cell);
    if (HEADER_WORDS.hasOwnProperty(folded)) { hits += 1; return; }
    var tokens = folded.split("/").join(" ").split("-").join(" ").split(" ").filter(function (t) { return t; });
    if (tokens.some(function (t) { return HEADER_WORDS.hasOwnProperty(t); })) hits += 1;
  });
  if (hits >= 2 || (hits === 1 && nonEmpty.length <= 3)) return true;

  following = (following || []).filter(function (r) { return !isBlankRow(r); });
  if (following.length) {
    var withDate = following.filter(function (r) { return r.some(function (c) { return looksLikeDate(c); }); }).length;
    if (withDate >= Math.max(1, Math.floor(following.length / 2))) return true; // rows below have dates, this one does not
    var numericBelow = following.filter(function (r) { return r.some(function (c) { return looksLikeAmount(c); }); }).length;
    if (numericBelow >= Math.max(1, Math.floor(following.length / 2)) &&
      !cells.some(function (c) { return looksLikeAmount(c); })) return true;
  }
  return hits > 0;
}

/** Find where the real table starts and how wide it is. */
function _findStart(rows) {
  var counts = {};
  rows.forEach(function (r) { if (!isBlankRow(r)) counts[r.length] = (counts[r.length] || 0) + 1; });
  var keys = Object.keys(counts);
  if (!keys.length) return [0, 0];
  var width = 0, freq = -1;
  keys.forEach(function (w) {
    var wi = parseInt(w, 10);
    if (counts[w] > freq || (counts[w] === freq && wi > width)) { freq = counts[w]; width = wi; }
  });

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (isBlankRow(row) || row.length !== width) continue;
    var following = rows.slice(i + 1, i + 5).filter(function (r) { return !isBlankRow(r); });
    if (!following.length || following.filter(function (r) { return r.length === width; }).length >=
      Math.max(1, following.length - 1)) {
      return [i, width];
    }
  }
  return [0, width];
}

// --------------------------------------------------------------------------
// Public entry point
// --------------------------------------------------------------------------

/**
 * Turn already-decoded CSV text into a Table. Encoding is decided by the
 * caller (the browser dialog, via TextDecoder - see Encoding.gs); pass the
 * resulting JS string plus the encoding name it used (for Table.describe()).
 */
function readTableFromText(text, opts) {
  opts = opts || {};
  text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  var stripped = _stripSepHint(text);
  text = stripped[0];
  var hinted = stripped[1];
  var delimiter = opts.delimiter;
  if (hinted && !delimiter) delimiter = hinted;
  if (!delimiter) delimiter = detectDelimiter(text);

  var rawRows = _parseCsv(text, delimiter);
  return Table.fromRows(rawRows, opts.hasHeader, {
    encoding: opts.encoding || "utf-8", delimiter: delimiter, source: opts.source || "",
  });
}

// ---- Columns.gs ---------------------------------------------------------------

// Figure out which column holds the date, the text and the amount.
//
// Port of budget_core/columns.py. Bank exports differ wildly: one amount
// column with signs, two columns (Hævet/Indsat), an extra Saldo column
// that must not be mistaken for the amount, two date columns
// (Bogført/Valør), text spread over a type column and a description
// column... This module scores every column and picks the most likely
// role, and it explains its choices so the user can correct them in the
// import dialog.

var DATE_WORDS = ["dato", "date", "datum", "bogfort", "bogforingsdato", "posteringsdato",
  "transaktionsdato", "valor", "valordato", "posting", "booking", "fecha"];
var DATE_WORDS_PRIMARY = ["dato", "date", "bogfort", "posteringsdato", "transaktionsdato",
  "posting", "booking"];
var TEXT_WORDS = ["tekst", "text", "beskrivelse", "description", "posteringstekst",
  "narrative", "details", "detaljer", "modtager", "afsender", "payee",
  "merchant", "reference", "memo", "note", "navn", "name", "titel",
  "forklaring", "kommentar", "bemaerkning"];
var AMOUNT_WORDS = ["belob", "beloeb", "amount", "sum", "betrag", "montant", "value",
  "transaktionsbelob", "bevaegelse", "posteringsbelob"];
var IN_WORDS = ["indsat", "indbetaling", "indbetalt", "ind", "kredit", "credit",
  "deposit", "modtaget", "tilgang", "indgaende", "income", "haevet ind"];
var OUT_WORDS = ["haevet", "udbetaling", "udbetalt", "ud", "debet", "debit",
  "withdrawal", "betalt", "afgang", "udgaende", "expense", "traek"];
var BALANCE_WORDS = ["saldo", "balance", "beholdning", "kontosaldo", "running balance", "ny saldo"];
var CURRENCY_WORDS = ["valuta", "currency", "mont", "vaeluta"];
var IGNORE_WORDS = ["kontonummer", "account number", "kortnummer", "id", "reg nr",
  "registreringsnummer", "afstemt", "status"];

function _hasWord(header, words) {
  var folded = fold(header);
  if (!folded) return false;
  if (words.indexOf(folded) !== -1) return true;
  var tokens = folded.split("/").join(" ").split("-").join(" ").split(".").join(" ")
    .split(" ").filter(function (t) { return t; });
  if (tokens.some(function (t) { return words.indexOf(t) !== -1; })) return true;
  return words.some(function (w) { return w.length > 4 && folded.indexOf(w) !== -1; });
}

/** Simple per column statistics used by the detectors. */
class ColumnStats {
  constructor(index, header, values) {
    this.index = index;
    this.header = header || "";
    this.values = values.slice();
    var filled = this.values.filter(function (v) { return String(v).trim(); });
    this.filled = filled;
    var total = filled.length || 1;
    this.fill_ratio = filled.length / (this.values.length || 1);
    this.date_ratio = filled.filter(function (v) { return looksLikeDate(v); }).length / total;
    var numeric = filled.filter(function (v) { return looksLikeAmount(v); });
    this.numeric_ratio = numeric.length / total;
    this.decimal_ratio = (numeric.filter(function (v) {
      return String(v).indexOf(",") !== -1 || String(v).indexOf(".") !== -1;
    }).length) / (numeric.length || 1);
    this.avg_length = filled.length ? filled.reduce(function (s, v) { return s + String(v).length; }, 0) / total : 0.0;
    var distinctSet = {};
    filled.forEach(function (v) { distinctSet[String(v).trim().toLowerCase()] = true; });
    this.distinct_ratio = filled.length ? Object.keys(distinctSet).length / total : 0.0;
    this.long_int_ratio = (numeric.filter(function (v) {
      var s = String(v).trim().split(" ").join("");
      return /^\d+$/.test(s) && s.length >= 8;
    }).length) / (numeric.length || 1);
  }

  is_dateish() { return this.date_ratio >= 0.6; }
  is_numeric() { return this.numeric_ratio >= 0.6 && this.date_ratio < 0.6; }
}

/** Which column plays which role, plus the parsing settings. */
class ColumnMapping {
  constructor() {
    this.date = null;
    this.text = [];
    this.amount = null;
    this.amount_in = null;
    this.amount_out = null;
    this.balance = null;
    this.currency = null;
    this.decimal = null;
    this.dayfirst = true;
    this.notes = [];
  }

  get has_amount() {
    return this.amount !== null || this.amount_in !== null || this.amount_out !== null;
  }

  is_usable() {
    return this.date !== null && this.has_amount;
  }

  describe(table) {
    var out = [];
    if (this.date !== null) out.push("Dato: " + table.header_name(this.date));
    if (this.text.length) {
      out.push("Tekst: " + this.text.map(function (i) { return table.header_name(i); }).join(" + "));
    }
    if (this.amount !== null) out.push("Beløb: " + table.header_name(this.amount));
    if (this.amount_in !== null) out.push("Indbetalt: " + table.header_name(this.amount_in));
    if (this.amount_out !== null) out.push("Hævet: " + table.header_name(this.amount_out));
    if (this.balance !== null) out.push("Saldo (ignoreres): " + table.header_name(this.balance));
    out.push("Decimaltegn: " + (this.decimal === "," ? "komma" : "punktum"));
    out.push("Datoformat: " + (this.dayfirst ? "dag før måned" : "måned før dag"));
    return out;
  }
}

/** Compute ColumnStats for every column of table. */
function analyseColumns(table, limit) {
  limit = limit === undefined ? 300 : limit;
  var out = [];
  for (var i = 0; i < table.n_columns; i++) {
    out.push(new ColumnStats(i, table.header_name(i), table.column(i, limit)));
  }
  return out;
}

function _pickDate(stats) {
  var candidates = stats.filter(function (s) { return s.is_dateish(); });
  if (!candidates.length) candidates = stats.filter(function (s) { return s.date_ratio >= 0.3; });
  if (!candidates.length) return null;
  var named = candidates.filter(function (s) { return _hasWord(s.header, DATE_WORDS_PRIMARY); });
  if (named.length) return named[0].index;
  named = candidates.filter(function (s) { return _hasWord(s.header, DATE_WORDS); });
  if (named.length) return named[0].index;
  candidates.sort(function (a, b) { return (b.date_ratio - a.date_ratio) || (a.index - b.index); });
  return candidates[0].index;
}

/** True when balanceCol behaves like a running balance of amountCol. */
function _looksLikeBalance(amountCol, balanceCol, decimal) {
  var hits = 0, tested = 0, prev = null;
  var n = Math.min(amountCol.values.length, balanceCol.values.length);
  for (var i = 0; i < n; i++) {
    var balance = parseAmount(balanceCol.values[i], decimal);
    var amount = parseAmount(amountCol.values[i], decimal);
    if (balance === null || amount === null) {
      prev = balance !== null ? balance : prev;
      continue;
    }
    if (prev !== null) {
      tested += 1;
      if (Math.abs((balance - prev) - amount) < 0.02 || Math.abs((prev - balance) - amount) < 0.02) hits += 1;
    }
    prev = balance;
  }
  return tested >= 3 && hits >= tested * 0.7;
}

function _pickAmount(stats, dateIndex, mapping) {
  var numeric = stats.filter(function (s) {
    return s.is_numeric() && s.index !== dateIndex && s.long_int_ratio < 0.8;
  });
  if (!numeric.length) {
    mapping.notes.push("Fandt ingen talkolonne - vælg beløbskolonnen manuelt.");
    return;
  }

  // 1) An explicit balance column is never the amount.
  var balances = numeric.filter(function (s) { return _hasWord(s.header, BALANCE_WORDS); });
  if (balances.length) {
    mapping.balance = balances[0].index;
    numeric = numeric.filter(function (s) { return s.index !== mapping.balance; });
  }

  // 2) Separate in/out columns?
  var ins = numeric.filter(function (s) { return _hasWord(s.header, IN_WORDS); });
  var outs = numeric.filter(function (s) { return _hasWord(s.header, OUT_WORDS); });
  if (ins.length && outs.length && ins[0].index !== outs[0].index) {
    mapping.amount_in = ins[0].index;
    mapping.amount_out = outs[0].index;
    return;
  }
  if (numeric.length === 2 && !numeric.some(function (s) { return _hasWord(s.header, AMOUNT_WORDS); })) {
    var first = numeric[0], second = numeric[1];
    var exclusive = 0, tested = 0;
    var n = Math.min(first.values.length, second.values.length);
    for (var i = 0; i < n; i++) {
      var aFilled = !!String(first.values[i]).trim();
      var bFilled = !!String(second.values[i]).trim();
      if (aFilled || bFilled) {
        tested += 1;
        if (aFilled !== bFilled) exclusive += 1;
      }
    }
    if (tested >= 3 && exclusive >= tested * 0.8) {
      mapping.amount_in = first.index;
      mapping.amount_out = second.index;
      mapping.notes.push("To kolonner udfyldes skiftevis - tolkes som ind- og udbetaling. " +
        "Byt om i dialogen hvis det er omvendt.");
      return;
    }
  }

  if (numeric.length === 1) {
    mapping.amount = numeric[0].index;
    return;
  }

  // 3) Named amount column wins.
  var named = numeric.filter(function (s) { return _hasWord(s.header, AMOUNT_WORDS); });
  if (named.length) {
    mapping.amount = named[0].index;
    var rest = numeric.filter(function (s) { return s.index !== mapping.amount; });
    if (mapping.balance === null && rest.length) {
      for (var r = 0; r < rest.length; r++) {
        if (_looksLikeBalance(named[0], rest[r], mapping.decimal)) {
          mapping.balance = rest[r].index;
          break;
        }
      }
    }
    return;
  }

  // 4) Otherwise look for the running balance pattern.
  for (var c = 0; c < numeric.length; c++) {
    var candidate = numeric[c];
    for (var o = 0; o < numeric.length; o++) {
      var other = numeric[o];
      if (other.index === candidate.index) continue;
      if (_looksLikeBalance(candidate, other, mapping.decimal)) {
        mapping.amount = candidate.index;
        mapping.balance = other.index;
        mapping.notes.push("Kolonnen \"" + other.header +
          "\" ser ud til at være en løbende saldo og springes over.");
        return;
      }
    }
  }

  // 5) Give up gracefully: prefer decimals, signs and a late position.
  function score(s) {
    var values = s.filled.map(function (v) { return parseAmount(v, mapping.decimal); })
      .filter(function (v) { return v !== null; });
    var mixed = (values.some(function (v) { return v < 0; }) && values.some(function (v) { return v > 0; })) ? 1 : 0;
    return [mixed, s.decimal_ratio, -s.index];
  }
  numeric.sort(function (a, b) {
    var sa = score(a), sb = score(b);
    for (var k = 0; k < sa.length; k++) { if (sa[k] !== sb[k]) return sb[k] - sa[k]; }
    return 0;
  });
  mapping.amount = numeric[0].index;
  mapping.notes.push("Flere talkolonner - valgte \"" + numeric[0].header +
    "\" som beløb. Ret det i dialogen hvis det er forkert.");
}

function _pickText(stats, used) {
  var candidates = stats.filter(function (s) {
    return used.indexOf(s.index) === -1 && s.numeric_ratio < 0.5 && s.date_ratio < 0.5 &&
      s.fill_ratio > 0.2 && !_hasWord(s.header, IGNORE_WORDS);
  });
  if (!candidates.length) return [];

  function score(s) {
    var named = _hasWord(s.header, TEXT_WORDS) ? 1 : 0;
    return [named, s.distinct_ratio, Math.min(s.avg_length, 40), -s.index];
  }
  candidates.sort(function (a, b) {
    var sa = score(a), sb = score(b);
    for (var k = 0; k < sa.length; k++) { if (sa[k] !== sb[k]) return sb[k] - sa[k]; }
    return 0;
  });
  var best = candidates[0];
  var chosen = [best.index];

  // Several named text columns (fx "Navn" + "Titel") describe one and the
  // same transaction - keep both, so the keyword rules have more to work on.
  if (_hasWord(best.header, TEXT_WORDS)) {
    for (var i = 1; i < candidates.length; i++) {
      if (chosen.length >= 2) break;
      var other = candidates[i];
      if (_hasWord(other.header, TEXT_WORDS) && other.fill_ratio > 0.5 && other.avg_length >= 4) {
        chosen.push(other.index);
      }
    }
  } else if (best.distinct_ratio < 0.25 && candidates.length > 1) {
    // A "type" column (few distinct values) alone is not descriptive enough.
    chosen.push(candidates[1].index);
  }
  return chosen.sort(function (a, b) { return a - b; });
}

/** Detect the roles of the columns of table. */
function detectMapping(table, limit) {
  var mapping = new ColumnMapping();
  var stats = analyseColumns(table, limit);
  if (!stats.length) {
    mapping.notes.push("Tabellen er tom.");
    return mapping;
  }

  mapping.date = _pickDate(stats);
  if (mapping.date !== null) mapping.dayfirst = detectDayfirst(stats[mapping.date].filled);

  var numericValues = [];
  stats.forEach(function (s) {
    if (s.is_numeric() && s.index !== mapping.date) numericValues = numericValues.concat(s.filled);
  });
  mapping.decimal = detectDecimalSeparator(numericValues) || ",";

  _pickAmount(stats, mapping.date, mapping);

  var used = [mapping.date, mapping.amount, mapping.amount_in, mapping.amount_out, mapping.balance]
    .filter(function (i) { return i !== null; });
  for (var i = 0; i < stats.length; i++) {
    var s = stats[i];
    if (used.indexOf(s.index) === -1 && _hasWord(s.header, CURRENCY_WORDS) && s.avg_length <= 6) {
      mapping.currency = s.index;
      used.push(s.index);
      break;
    }
  }

  mapping.text = _pickText(stats, used);
  if (!mapping.text.length) mapping.notes.push("Fandt ingen tekstkolonne - alle poster får samme beskrivelse.");
  if (mapping.date === null) mapping.notes.push("Fandt ingen datokolonne - vælg den manuelt.");
  return mapping;
}

// ---- Budget.gs ----------------------------------------------------------------

// Turn a detected table into categorised transactions and a budget.
//
// Port of budget_core/budget.py. The output of buildTransactions is
// deliberately plain (an array of Transaction) so it can be written to a
// sheet, a CSV file or anything else.

var KIND_INCOME = "Indtægt";
var KIND_EXPENSE = "Udgift";

var SIGN_AUTO = "auto";
var SIGN_NEGATIVE_IS_EXPENSE = "negativ";
var SIGN_POSITIVE_IS_EXPENSE = "positiv";

/** One line in the Transaktioner sheet. */
class Transaction {
  constructor(date, text, amount, category, currency, sourceRow) {
    this.date = date;
    this.text = text;
    this.amount = amount; // negative = expense
    this.category = category;
    this.kind = amount < 0 ? KIND_EXPENSE : KIND_INCOME;
    this.month = monthKey(date);
    this.currency = currency || "";
    this.source_row = sourceRow || 0;
  }

  get month_name() {
    return monthLabel(this.month);
  }
}

/** User adjustable settings for buildTransactions. */
class BuildOptions {
  constructor(opts) {
    opts = opts || {};
    this.decimal = opts.decimal !== undefined ? opts.decimal : null;
    this.dayfirst = opts.dayfirst !== undefined ? opts.dayfirst : null;
    this.sign = opts.sign || SIGN_AUTO;
    this.skip_zero = opts.skip_zero !== undefined ? opts.skip_zero : true;
    this.date_from = opts.date_from !== undefined ? opts.date_from : null;
    this.date_to = opts.date_to !== undefined ? opts.date_to : null;
  }
}

class BuildResult {
  constructor(transactions, warnings, skippedNoDate, skippedNoAmount, uncategorised) {
    this.transactions = transactions;
    this.warnings = warnings;
    this.skipped_no_date = skippedNoDate;
    this.skipped_no_amount = skippedNoAmount;
    // [[shop, count, total, ruleKeyword]] sorted by |total| descending
    this.uncategorised = uncategorised;
  }

  get ok() {
    return this.transactions.length > 0;
  }

  summary_lines() {
    var lines = [this.transactions.length + " posteringer indlæst."];
    if (this.transactions.length) {
      var first = this.transactions.reduce(function (a, b) { return a.date < b.date ? a : b; }).date;
      var last = this.transactions.reduce(function (a, b) { return a.date > b.date ? a : b; }).date;
      lines.push("Periode: " + _isoDate(first) + " - " + _isoDate(last));
    }
    if (this.skipped_no_date) lines.push(this.skipped_no_date + " række(r) uden gyldig dato blev sprunget over.");
    if (this.skipped_no_amount) lines.push(this.skipped_no_amount + " række(r) uden gyldigt beløb blev sprunget over.");
    var uncategorised = this.transactions.filter(function (t) { return t.category === DEFAULT_CATEGORY; }).length;
    if (uncategorised) lines.push(uncategorised + " postering(er) er ikke kategoriseret.");
    var ignored = this.transactions.filter(function (t) { return t.category === IGNORED_CATEGORY; }).length;
    if (ignored) lines.push(ignored + " postering(er) er markeret som ignoreret og tæller ikke med i budgettet.");
    return lines.concat(this.warnings);
  }
}

function _isoDate(d) {
  var y = d.getUTCFullYear();
  var m = d.getUTCMonth() + 1;
  var day = d.getUTCDate();
  return y + "-" + (m < 10 ? "0" + m : m) + "-" + (day < 10 ? "0" + day : day);
}

function _cell(row, index) {
  if (index === null || index === undefined || index < 0 || index >= row.length) return "";
  var value = row[index];
  return value === null || value === undefined ? "" : String(value);
}

/** Return true when the amounts must be negated to make expenses negative. */
function _decideSign(values, mode) {
  if (mode === SIGN_POSITIVE_IS_EXPENSE) return true;
  if (mode === SIGN_NEGATIVE_IS_EXPENSE) return false;
  var negatives = values.filter(function (v) { return v < 0; }).length;
  return negatives === 0 && values.length > 0;
}

/** Amount of one row, honouring single or split amount columns. */
function _rowAmount(row, mapping, decimal) {
  if (mapping.amount_in !== null || mapping.amount_out !== null) {
    var valueIn = mapping.amount_in !== null ? parseAmount(_cell(row, mapping.amount_in), decimal) : null;
    var valueOut = mapping.amount_out !== null ? parseAmount(_cell(row, mapping.amount_out), decimal) : null;
    if (valueIn === null && valueOut === null) return null;
    return Math.abs(valueIn || 0.0) - Math.abs(valueOut || 0.0);
  }
  return parseAmount(_cell(row, mapping.amount), decimal);
}

/** Parse every row of table according to mapping. */
function buildTransactions(table, mapping, ruleset, options) {
  ruleset = ruleset || RuleSet.defaults();
  options = options || new BuildOptions();
  var warnings = (mapping.notes || []).slice();
  var decimal = options.decimal || mapping.decimal;
  var dayfirst = options.dayfirst === null || options.dayfirst === undefined ? mapping.dayfirst : options.dayfirst;

  if (mapping.date === null) {
    return new BuildResult([], warnings.concat(["Ingen datokolonne valgt."]), 0, 0, []);
  }
  if (!mapping.has_amount) {
    return new BuildResult([], warnings.concat(["Ingen beløbskolonne valgt."]), 0, 0, []);
  }

  var parsed = [];
  var skippedNoDate = 0;
  var skippedNoAmount = 0;

  for (var index = 0; index < table.rows.length; index++) {
    var row = table.rows[index];
    var date = parseDate(_cell(row, mapping.date), dayfirst);
    var amount = _rowAmount(row, mapping, decimal);
    if (date === null) {
      if (amount !== null || row.some(function (c) { return String(c).trim(); })) skippedNoDate += 1;
      continue;
    }
    if (amount === null) { skippedNoAmount += 1; continue; }
    if (options.skip_zero && Math.abs(amount) < 0.0000001) continue;
    if (options.date_from && date < options.date_from) continue;
    if (options.date_to && date > options.date_to) continue;
    var textParts = mapping.text.map(function (i) { return squeeze(_cell(row, i)); }).filter(function (p) { return p; });
    var text = textParts.join(" – ");
    if (!text) text = "(ingen tekst)";
    var currency = mapping.currency !== null ? squeeze(_cell(row, mapping.currency)) : "";
    parsed.push([date, text, amount, currency, index]);
  }

  if (!parsed.length) {
    return new BuildResult([], warnings.concat(["Ingen brugbare rækker. Kontrollér kolonnevalget i dialogen."]),
      skippedNoDate, skippedNoAmount, []);
  }

  var flip = false;
  if (mapping.amount !== null && mapping.amount_in === null) {
    flip = _decideSign(parsed.map(function (p) { return p[2]; }), options.sign);
    if (flip && options.sign === SIGN_AUTO) {
      warnings.push("Beløbskolonnen indeholder ingen negative tal - alle poster behandles som " +
        "udgifter. Skift fortegnsregel i dialogen hvis filen også indeholder indtægter.");
    }
  }

  var transactions = parsed.map(function (p) {
    var date = p[0], text = p[1], amount = p[2], currency = p[3], index = p[4];
    var value = flip ? -amount : amount;
    return new Transaction(date, text, value, ruleset.categorise(text), currency, index);
  });

  transactions.sort(function (a, b) {
    if (a.date < b.date) return -1;
    if (a.date > b.date) return 1;
    return a.source_row - b.source_row;
  });

  // Group by shop, not by the raw text: receipt numbers, dates and amounts
  // inside the text would otherwise make every single line its own group.
  var counterOrder = [];
  var counter = {};
  transactions.forEach(function (transaction) {
    if (transaction.category === DEFAULT_CATEGORY) {
      var key = merchantKey(transaction.text);
      if (!counter.hasOwnProperty(key)) {
        counter[key] = [merchantName(transaction.text), 0, 0.0, ruleKeyword(transaction.text)];
        counterOrder.push(key);
      }
      counter[key][1] += 1;
      counter[key][2] += transaction.amount;
    }
  });
  var uncategorised = counterOrder.map(function (key) { return counter[key]; })
    .sort(function (a, b) { return Math.abs(b[2]) - Math.abs(a[2]); });

  return new BuildResult(transactions, warnings, skippedNoDate, skippedNoAmount, uncategorised);
}

// --------------------------------------------------------------------------
// Summary
// --------------------------------------------------------------------------

/** Monthly totals per category - the actual budget. */
class Summary {
  constructor(transactions) {
    var counted = transactions.filter(function (t) { return t.category !== IGNORED_CATEGORY; });
    var monthSet = {};
    counted.forEach(function (t) { monthSet[t.month] = true; });
    this.months = Object.keys(monthSet).sort();
    this.cells = {}; // "kind|category|month" -> number
    this.counts = {}; // "kind|category" -> number
    var self = this;
    counted.forEach(function (t) {
      var key = t.kind + "|" + t.category + "|" + t.month;
      self.cells[key] = (self.cells[key] || 0.0) + t.amount;
      var ckey = t.kind + "|" + t.category;
      self.counts[ckey] = (self.counts[ckey] || 0) + 1;
    });
    this.expense_categories = this._categories(KIND_EXPENSE);
    this.income_categories = this._categories(KIND_INCOME);
    this.transaction_count = counted.length;
    this.ignored_count = transactions.length - counted.length;
  }

  _categories(kind) {
    var totals = {};
    var order = [];
    for (var key in this.cells) {
      var parts = key.split("|");
      var k = parts[0], category = parts[1];
      if (k !== kind) continue;
      if (!totals.hasOwnProperty(category)) { totals[category] = 0.0; order.push(category); }
      totals[category] += this.cells[key];
    }
    return order.sort(function (a, b) {
      var diff = Math.abs(totals[b]) - Math.abs(totals[a]);
      if (diff !== 0) return diff;
      return a < b ? -1 : a > b ? 1 : 0;
    });
  }

  value(kind, category, month) {
    var v = this.cells[kind + "|" + category + "|" + month];
    return v === undefined ? 0.0 : v;
  }

  category_total(kind, category) {
    var self = this;
    return this.months.reduce(function (sum, m) { return sum + self.value(kind, category, m); }, 0.0);
  }

  month_total(kind) {
    var self = this;
    var cats = this._categories(kind);
    var out = {};
    this.months.forEach(function (month) {
      out[month] = cats.reduce(function (sum, c) { return sum + self.value(kind, c, month); }, 0.0);
    });
    return out;
  }

  total(kind) {
    var sum = 0.0;
    for (var key in this.cells) {
      if (key.split("|")[0] === kind) sum += this.cells[key];
    }
    return sum;
  }

  get net() {
    return this.total(KIND_INCOME) + this.total(KIND_EXPENSE);
  }

  month_count() {
    return Math.max(1, this.months.length);
  }

  top_expenses(count) {
    count = count === undefined ? 5 : count;
    var self = this;
    return this.expense_categories.slice(0, count).map(function (c) { return [c, self.category_total(KIND_EXPENSE, c)]; });
  }
}

function summarise(transactions) {
  return new Summary(transactions);
}

// ---- Plan.gs ------------------------------------------------------------------

// Turn a year of transactions into a draft budget.
//
// Port of budget_core/plan.py. Every category is measured across the
// months of the file and sorted into three groups:
//
// fast       the same amount almost every month (husleje, forsikring, lån)
// variabel   there every month, but the amount moves (dagligvarer, restaurant)
// periodisk  only a few months a year (el-afregning, ferie, jul) - budgeted
//            as a monthly set-aside of the yearly total
//
// From that we suggest an amount per category: a careful figure for income
// (the typical month, not the best one), the average for the fixed costs
// and the median month for the variable ones, so a single expensive
// December does not blow up the grocery budget.

// How many contributing transactions to remember per category (for the
// on-hover comment in the spreadsheet).
var SAMPLE_SIZE = 8;

var GROUP_FIXED = "fast";
var GROUP_VARIABLE = "variabel";
var GROUP_PERIODIC = "periodisk";

var GROUP_LABELS = {};
GROUP_LABELS[GROUP_FIXED] = "FASTE UDGIFTER";
GROUP_LABELS[GROUP_VARIABLE] = "VARIABLE UDGIFTER";
GROUP_LABELS[GROUP_PERIODIC] = "PERIODISKE UDGIFTER";

// Categories that are fixed costs by nature. Billed every month they are
// "faste"; billed quarterly or yearly they become a monthly set-aside.
var FIXED_BY_NATURE = ["bolig", "el, vand og varme", "forsikring og pension",
  "telefon og internet", "abonnementer", "lan og afdrag",
  "born og uddannelse", "gebyrer og renter"];

// Categories that are spending decisions, not bills - they belong under
// "variable" even in a year where the monthly total happened to be steady.
var VARIABLE_BY_NATURE = ["dagligvarer", "restaurant", "shopping", "transport",
  "fritid og sport", "sundhed", "rejser", "mobilepay",
  "overforsel", "ukategoriseret"];

// How the groups are detected.
var MONTHLY_COVERAGE = 0.90; // a bill that really does turn up every month
var FIXED_COVERAGE = 0.70; // present in at least this share of the months
var FIXED_VARIATION = 0.30; // and varies less than this (std/mean)
var STEADY_VARIATION = 0.10; // so steady that it must be a subscription
var PERIODIC_COVERAGE = 0.40; // present in at most this share of the months
var UNCERTAIN_MONTHS = 6; // fewer months than this -> warn the user

function _median(values) {
  if (!values.length) return 0.0;
  var ordered = values.slice().sort(function (a, b) { return a - b; });
  var middle = Math.floor(ordered.length / 2);
  if (ordered.length % 2) return ordered[middle];
  return (ordered[middle - 1] + ordered[middle]) / 2.0;
}

function _stdev(values) {
  if (values.length < 2) return 0.0;
  var mean = values.reduce(function (s, v) { return s + v; }, 0) / values.length;
  var variance = values.reduce(function (s, v) { return s + (v - mean) * (v - mean); }, 0) / (values.length - 1);
  return Math.sqrt(variance);
}

function _roundTo(value, step, mode) {
  if (step <= 0 || !value) return Math.round(value * 100) / 100;
  var quotient = value / step;
  if (mode === "up") quotient = Math.ceil(quotient);
  else if (mode === "down") quotient = Math.floor(quotient);
  else quotient = Math.floor(quotient + 0.5);
  return quotient * step;
}

/** What one category costs per month, and what we suggest budgeting. */
class CategoryPlan {
  constructor(kind, category, values, transactions) {
    transactions = transactions || [];
    this.kind = kind;
    this.category = category;
    this.values = values.map(function (v) { return Math.abs(v); });
    var active = this.values.filter(function (v) { return v > 0.005; });
    var months = this.values.length || 1;

    this.months_present = active.length;
    this.total = this.values.reduce(function (s, v) { return s + v; }, 0.0);
    this.mean_all = this.total / months;
    this.mean_active = active.length ? active.reduce(function (s, v) { return s + v; }, 0) / active.length : 0.0;
    this.median = _median(active);
    this.low = active.length ? Math.min.apply(null, active) : 0.0;
    this.high = active.length ? Math.max.apply(null, active) : 0.0;
    this.variation = this.mean_active ? (_stdev(active) / this.mean_active) : 0.0;
    this.group = GROUP_VARIABLE;
    this.suggestion = 0.0;
    this.account = "";

    var ordered = transactions.slice().sort(function (a, b) { return Math.abs(b.amount) - Math.abs(a.amount); });
    this.sample = ordered.slice(0, SAMPLE_SIZE);
    this.transaction_count = transactions.length;
  }

  get coverage() {
    return this.months_present / (this.values.length || 1);
  }
}

function _classify(plan) {
  if (plan.kind === KIND_INCOME) return GROUP_FIXED;
  var name = fold(plan.category);
  if (FIXED_BY_NATURE.indexOf(name) !== -1) {
    // El, vand og varme is a fixed cost - but if it is only billed every
    // quarter, budgeting the bill itself would be wrong. Spread it out.
    return plan.coverage >= MONTHLY_COVERAGE ? GROUP_FIXED : GROUP_PERIODIC;
  }
  if (plan.coverage <= PERIODIC_COVERAGE) return GROUP_PERIODIC;
  var steady = plan.coverage >= FIXED_COVERAGE && plan.variation <= FIXED_VARIATION;
  if (steady && !(VARIABLE_BY_NATURE.indexOf(name) !== -1 && plan.variation > STEADY_VARIATION)) {
    return GROUP_FIXED;
  }
  return GROUP_VARIABLE;
}

/** The amount we propose budgeting per month. */
function _suggest(plan) {
  if (plan.kind === KIND_INCOME) {
    // Budget on the typical month, not on the month with the bonus.
    var base = plan.coverage >= FIXED_COVERAGE ? plan.median : plan.mean_all;
    return _roundTo(base, 100, "down");
  }
  if (plan.group === GROUP_FIXED) return _roundTo(plan.mean_active, 10, "up");
  if (plan.group === GROUP_PERIODIC) return _roundTo(plan.mean_all, 50, "up");
  var variableBase = plan.median || plan.mean_active;
  return _roundTo(variableBase, 50);
}

/** A draft budget for a normal month. */
class BudgetPlan {
  constructor(months, categories) {
    this.months = months.slice();
    this.categories = categories.slice();
    this.income = this.categories.filter(function (c) { return c.kind === KIND_INCOME; });
    this.expenses = this.categories.filter(function (c) { return c.kind === KIND_EXPENSE; });
    this.groups = {};
    var self = this;
    [GROUP_FIXED, GROUP_VARIABLE, GROUP_PERIODIC].forEach(function (g) {
      self.groups[g] = self.expenses.filter(function (c) { return c.group === g; });
    });
  }

  get month_count() { return Math.max(1, this.months.length); }
  get income_mean() { return this.income.reduce(function (s, c) { return s + c.mean_all; }, 0); }
  get expense_mean() { return this.expenses.reduce(function (s, c) { return s + c.mean_all; }, 0); }
  get savings_mean() { return this.income_mean - this.expense_mean; }
  get income_suggested() { return this.income.reduce(function (s, c) { return s + c.suggestion; }, 0); }
  get expense_suggested() { return this.expenses.reduce(function (s, c) { return s + c.suggestion; }, 0); }
  get savings_suggested() { return this.income_suggested - this.expense_suggested; }
  get uncertain() { return this.months.length < UNCERTAIN_MONTHS; }

  group_mean(group) {
    return (this.groups[group] || []).reduce(function (s, c) { return s + c.mean_all; }, 0);
  }

  summary_lines() {
    var lines = [
      "Gennemsnitlig indtægt: " + Math.round(this.income_mean) + " kr./md",
      "Faste udgifter: " + Math.round(this.group_mean(GROUP_FIXED)) + " kr./md",
      "Variable udgifter: " + Math.round(this.group_mean(GROUP_VARIABLE)) + " kr./md",
      "Periodiske udgifter: " + Math.round(this.group_mean(GROUP_PERIODIC)) + " kr./md",
      "Til opsparing i gennemsnit: " + Math.round(this.savings_mean) + " kr./md",
      "Forslaget giver plads til " + Math.round(this.savings_suggested) + " kr./md i opsparing",
    ];
    if (this.uncertain) {
      lines.push("Bemærk: forslaget bygger kun på " + this.months.length +
        " måned(er) - tallene er usikre.");
    }
    return lines;
  }
}

/**
 * Build a draft budget from a Summary. transactions (the same list used to
 * build summary) lets each category remember its largest contributing
 * transactions, for the on-hover comment shown in the spreadsheet.
 */
function buildPlan(summary, transactions, months) {
  transactions = transactions || [];
  months = months || summary.months;
  var byCategory = {};
  transactions.forEach(function (t) {
    if (months.indexOf(t.month) !== -1) {
      var key = t.kind + "|" + t.category;
      if (!byCategory[key]) byCategory[key] = [];
      byCategory[key].push(t);
    }
  });

  var categories = [];
  [[KIND_INCOME, summary.income_categories], [KIND_EXPENSE, summary.expense_categories]].forEach(function (pair) {
    var kind = pair[0], names = pair[1];
    names.forEach(function (name) {
      var values = months.map(function (month) { return summary.value(kind, name, month); });
      var plan = new CategoryPlan(kind, name, values, byCategory[kind + "|" + name] || []);
      plan.group = _classify(plan);
      plan.suggestion = _suggest(plan);
      categories.push(plan);
    });
  });

  // Biggest first inside each group.
  categories.sort(function (a, b) {
    var diff = Math.abs(b.mean_all) - Math.abs(a.mean_all);
    if (diff !== 0) return diff;
    return a.category < b.category ? -1 : a.category > b.category ? 1 : 0;
  });
  return new BudgetPlan(months, categories);
}

// ---- SheetWriter.gs -----------------------------------------------------------

// Write the budget straight into the active Google Sheets spreadsheet.
//
// Mirrors BudgetScript/xlsx_writer.py in spirit and sheet layout, but
// targets the SpreadsheetApp API instead of openpyxl - and writes into
// the spreadsheet the user already has open instead of a separate file,
// since that is the whole point of doing this in Sheets. All the actual
// parsing/categorisation/classification logic is reused as-is from the
// Budget/Plan/Rules/Transfers modules - this file only writes cells.
//
// This file cannot be unit tested outside of Apps Script (it needs a real
// SpreadsheetApp), so it is deliberately written to stay close to the
// already-tested xlsx_writer.py it mirrors, one section at a time.

// ---------------------------------------------------------------------------
// Design tokens (same palette as the LibreOffice/Excel/Script versions)
// ---------------------------------------------------------------------------
var NAVY = "#334960";
var ORANGE = "#F46524";
var TEXT_GREY = "#576475";
var MUTED = "#687887";
var DARK = "#434343";
var PEACH = "#FFF2ED";

var FONT_BODY = "Calibri";
var FONT_TITLE = "Calibri";

var FMT_CURRENCY = '#,##0 "kr."';
var FMT_CURRENCY2 = '#,##0.00 "kr."';
var FMT_SIGNED = '+#,##0 "kr.";-#,##0 "kr."';
var FMT_DATE = "dd-mm-yyyy";
var FMT_MONTH = "mmmm yyyy";
var FMT_MONTH_SHORT = "mmm yyyy";

var SHEET_TX = "Transaktioner";
var SHEET_RULES = "Kategorier";
var SHEET_PLAN = "Budgetforslag";
var SHEET_FORECAST = "Prognose";
var SHEET_ACCOUNTS = "Konti";
var SHEET_MONTHS = "Alle måneder";
var SHEET_ORDER = [SHEET_PLAN, SHEET_FORECAST, SHEET_TX, SHEET_MONTHS, SHEET_RULES, SHEET_ACCOUNTS];

var TX_HEADER_ROW = 4;
var TX_FIRST_ROW = 5;
var TX_MAX_ROW = 5000;
var COL_EXP_DATE = 2, COL_EXP_AMOUNT = 3, COL_EXP_TEXT = 4, COL_EXP_CAT = 5;
var COL_INC_DATE = 7, COL_INC_AMOUNT = 8, COL_INC_TEXT = 9, COL_INC_CAT = 10;

var ACC_HEADER_ROW = 4;
var ACC_FIRST_ROW = 5;
var ACC_COL_NAME = 2, ACC_COL_TYPE = 3, ACC_COL_BALANCE = 4, ACC_COL_NUMBER = 5;

var PLAN_HEADER_ROW = 5;
var FORECAST_MONTHS = 24;
var FORECAST_HEADER_ROW = 9;

function _col(index) {
  var letters = "";
  while (index > 0) {
    var rem = (index - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    index = Math.floor((index - 1) / 26);
  }
  return letters;
}

function _ref(col, row) { return _col(col) + row; }
function _range(col1, row1, col2, row2) { return _ref(col1, row1) + ":" + _ref(col2, row2); }

// Sheets/Apps Script reads a date-formatted cell back as a Date at local
// midnight in the SCRIPT's timezone, and writes a Date the same way - the
// reverse of the UTC-midnight Date objects Parsing.gs hands back. Anchor
// every written date at local noon (never local midnight) so no timezone
// offset on earth can push the calendar day itself backward or forward,
// and reconstruct a UTC-midnight Date from the LOCAL calendar fields Sheets
// gives back, so dates round-trip through a cell without drifting a day.
function _toSheetDate(date) {
  return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), 12, 0, 0);
}

function _fromSheetDate(date) {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
}

function _nextMonth(monthKeyStr) {
  if (!monthKeyStr) {
    var today = new Date();
    return new Date(Date.UTC(today.getFullYear(), today.getMonth(), 1));
  }
  var parts = monthKeyStr.split("-");
  return _addMonths(new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, 1)), 1);
}

function _addMonths(date, count) {
  var total = date.getUTCMonth() + count;
  var year = date.getUTCFullYear() + Math.floor(total / 12);
  var month = ((total % 12) + 12) % 12;
  return new Date(Date.UTC(year, month, 1));
}

function _monthStart(monthKeyStr) {
  var parts = monthKeyStr.split("-");
  return new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, 1));
}

/** Small styling helper wrapping one sheet. */
class Pen {
  constructor(sheet) {
    this.sheet = sheet;
  }

  cell(ref) { return this.sheet.getRange(ref); }

  text(ref, value, opts) {
    opts = opts || {};
    var range = this.sheet.getRange(ref);
    range.setValue(value === null || value === undefined ? "" : value);
    this._style(range, opts);
    return range;
  }

  number(ref, value, opts) {
    opts = opts || {};
    var range = this.sheet.getRange(ref);
    range.setValue(value);
    range.setNumberFormat(opts.fmt || FMT_CURRENCY);
    this._style(range, Object.assign({ align: opts.align || "right" }, opts));
    return range;
  }

  date(ref, value, opts) {
    opts = opts || {};
    var range = this.sheet.getRange(ref);
    range.setValue(_toSheetDate(value));
    range.setNumberFormat(opts.fmt || FMT_DATE);
    this._style(range, opts);
    return range;
  }

  formula(ref, formulaText, opts) {
    opts = opts || {};
    var range = this.sheet.getRange(ref);
    range.setFormula(formulaText);
    if (opts.fmt) range.setNumberFormat(opts.fmt);
    this._style(range, Object.assign({ align: opts.align || "right" }, opts));
    return range;
  }

  /** Re-style an existing range without touching its values. */
  style(ref, opts) {
    opts = opts || {};
    var range = this.sheet.getRange(ref);
    this._style(range, opts);
    if (opts.fmt) range.setNumberFormat(opts.fmt);
  }

  _style(range, opts) {
    range.setFontFamily(opts.font || FONT_BODY);
    if (opts.size !== undefined) range.setFontSize(opts.size);
    if (opts.bold !== undefined) range.setFontWeight(opts.bold ? "bold" : "normal");
    if (opts.italic !== undefined) range.setFontStyle(opts.italic ? "italic" : "normal");
    if (opts.color) range.setFontColor(opts.color);
    if (opts.align) range.setHorizontalAlignment(opts.align);
    if (opts.valign) range.setVerticalAlignment(opts.valign);
    if (opts.wrap) range.setWrap(true);
    if (opts.bg) range.setBackground(opts.bg);
  }

  merge(ref) { this.sheet.getRange(ref).merge(); }

  columnWidth(colIndex, width) { this.sheet.setColumnWidth(colIndex, Math.round(width * 7)); }

  rowHeight(row, height) { this.sheet.setRowHeight(row, Math.round(height * 4 / 3)); }

  note(ref, text) {
    if (!text) return;
    this.sheet.getRange(ref).setNote(text);
  }

  /** ref: A1 range on this sheet. sourceRange: a Range object elsewhere. */
  dropdown(ref, sourceRange) {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(true)
      .build();
    this.sheet.getRange(ref).setDataValidation(rule);
  }

  freeze(rows, cols) {
    this.sheet.setFrozenRows(rows);
    this.sheet.setFrozenColumns(cols);
  }
}

/** A transaction read back from the spreadsheet. */
class SheetTransaction {
  constructor(date, text, amount, category, sourceRow, income) {
    this.date = date;
    this.text = text;
    this.amount = amount;
    this.category = category;
    this.source_row = sourceRow;
    this.income = income;
  }

  get month() { return monthKey(this.date); }
  get kind() { return this.income ? KIND_INCOME : KIND_EXPENSE; }
}

/** Reads and writes one budget spreadsheet. */
class BudgetBook {
  constructor(ss) {
    this.ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  }

  // -- sheet helpers ------------------------------------------------
  sheet(name) {
    var existing = this.ss.getSheetByName(name);
    if (existing) this.ss.deleteSheet(existing);
    return this.ss.insertSheet(name);
  }

  has_sheet(name) { return this.ss.getSheetByName(name) !== null; }

  pen(sheet) { return new Pen(sheet); }

  used_row_count(sheet) { return sheet.getLastRow() || 1; }

  // -- Transaktioner --------------------------------------------------
  write_transactions(sheet, transactions) {
    var pen = this.pen(sheet);
    var widths = { 1: 30, 2: 90, 3: 110, 4: 220, 5: 130, 6: 30, 7: 90, 8: 110, 9: 220, 10: 130 };
    for (var col in widths) sheet.setColumnWidth(parseInt(col, 10), widths[col]);

    pen.merge("B1:J1");
    pen.text("B1", "Skift eller tilføj kategorier i kolonnerne nedenfor - og i arket " +
      "\"Kategorier\", hvis de skal sættes automatisk.",
      { size: 10, color: "#CCCCCC", bg: NAVY, italic: true, align: "left", valign: "middle" });
    sheet.setRowHeight(1, 30);

    pen.text("B2", "Udgifter", { font: FONT_TITLE, size: 18, bold: true, color: ORANGE });
    pen.text("G2", "Indtægter", { font: FONT_TITLE, size: 18, bold: true, color: ORANGE });
    sheet.setRowHeight(2, 34);

    var headers = ["Dato", "Beløb", "Beskrivelse", "Kategori"];
    for (var offset = 0; offset < headers.length; offset++) {
      pen.text(_ref(COL_EXP_DATE + offset, TX_HEADER_ROW), headers[offset], { bold: true, size: 11, color: NAVY });
      pen.text(_ref(COL_INC_DATE + offset, TX_HEADER_ROW), headers[offset], { bold: true, size: 11, color: NAVY });
    }

    var expenses = transactions.filter(function (t) { return t.amount < 0; });
    var income = transactions.filter(function (t) { return t.amount >= 0; });
    this._writeBlock(sheet, expenses, COL_EXP_DATE, TX_FIRST_ROW);
    this._writeBlock(sheet, income, COL_INC_DATE, TX_FIRST_ROW);

    var last = TX_FIRST_ROW + Math.max(expenses.length, income.length, 1) - 1;
    this._styleTxRange(pen, last);
    this._dimIgnoredRows(pen, expenses, COL_EXP_DATE, TX_FIRST_ROW);
    this._dimIgnoredRows(pen, income, COL_INC_DATE, TX_FIRST_ROW);
    this._addCategoryDropdown(sheet, COL_EXP_CAT, last);
    this._addCategoryDropdown(sheet, COL_INC_CAT, last);
    pen.freeze(TX_FIRST_ROW - 1, 1);
  }

  _styleTxRange(pen, last) {
    [COL_EXP_DATE, COL_INC_DATE].forEach(function (firstCol) {
      pen.style(_range(firstCol, TX_FIRST_ROW, firstCol, last), { fmt: FMT_DATE, color: MUTED, align: "left" });
      pen.style(_range(firstCol + 1, TX_FIRST_ROW, firstCol + 1, last),
        { fmt: FMT_CURRENCY2, color: TEXT_GREY, bold: true, align: "left" });
      pen.style(_range(firstCol + 2, TX_FIRST_ROW, firstCol + 3, last), { color: TEXT_GREY, align: "left" });
    });
  }

  /** Bulk-write one date/amount/text/category block with a single setValues call. */
  _writeBlock(sheet, transactions, firstCol, startRow) {
    if (!transactions.length) return;
    var rows = transactions.map(function (t) {
      return [_toSheetDate(t.date), Math.abs(t.amount), t.text, t.category];
    });
    sheet.getRange(startRow, firstCol, rows.length, 4).setValues(rows);
  }

  _dimIgnoredRows(pen, transactions, firstCol, startRow) {
    for (var offset = 0; offset < transactions.length; offset++) {
      if (transactions[offset].category === IGNORED_CATEGORY) {
        var row = startRow + offset;
        pen.style(_range(firstCol, row, firstCol + 3, row), { color: MUTED, italic: true });
      }
    }
  }

  _addCategoryDropdown(sheet, col, lastRow) {
    if (lastRow < TX_FIRST_ROW) return;
    var pen = this.pen(sheet);
    var rulesSheet = this.ss.getSheetByName(SHEET_RULES);
    if (!rulesSheet) return;
    pen.dropdown(_range(col, TX_FIRST_ROW, col, lastRow), rulesSheet.getRange("E5:E80"));
  }

  /** Bulk-read one date/amount/text/category block with a single getValues call. */
  read_transactions() {
    var sheet = this.ss.getSheetByName(SHEET_TX);
    if (!sheet) return [];
    var lastRow = this.used_row_count(sheet);
    if (lastRow < TX_FIRST_ROW) return [];
    var nRows = lastRow - TX_FIRST_ROW + 1;
    var out = [];
    var blocks = [
      [COL_EXP_DATE, -1.0, false],
      [COL_INC_DATE, 1.0, true],
    ];
    blocks.forEach(function (block) {
      var firstCol = block[0], sign = block[1], income = block[2];
      var values = sheet.getRange(TX_FIRST_ROW, firstCol, nRows, 4).getValues();
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var dateVal = row[0], amtVal = row[1], textVal = row[2], catVal = row[3];
        if (!(dateVal instanceof Date)) continue;
        if (typeof amtVal !== "number" || amtVal === 0) continue;
        out.push(new SheetTransaction(_fromSheetDate(dateVal), String(textVal || ""),
          sign * Math.abs(amtVal), String(catVal || "").trim(), TX_FIRST_ROW + i, income));
      }
    });
    out.sort(function (a, b) {
      if (a.date < b.date) return -1;
      if (a.date > b.date) return 1;
      return a.source_row - b.source_row;
    });
    return out;
  }

  has_budget() { return this.has_sheet(SHEET_TX); }

  // -- Kategorier -------------------------------------------------------
  write_rules(sheet, ruleset, accounts) {
    var pen = this.pen(sheet);
    var widths = { 1: 30, 2: 190, 3: 160, 4: 30, 5: 160, 6: 140 };
    for (var col in widths) sheet.setColumnWidth(parseInt(col, 10), widths[col]);

    pen.merge("B2:F2");
    pen.text("B2", "Kategorier", { font: FONT_TITLE, size: 18, bold: true, color: ORANGE, align: "left" });
    pen.merge("B3:F3");
    pen.text("B3", "Tilføj, ret eller slet regler i venstre tabel (nøgleord -> kategori). " +
      "Højre tabel viser alle kategorier og hvilken konto, de trækkes fra.",
      { size: 9, italic: true, color: MUTED, align: "left", wrap: true, valign: "middle" });
    sheet.setRowHeight(3, 46);

    pen.text("B4", "Nøgleord", { bold: true, size: 11, color: NAVY, align: "left" });
    pen.text("C4", "Kategori", { bold: true, size: 11, color: NAVY, align: "left" });
    pen.text("E4", "Kategori", { bold: true, size: 11, color: NAVY, align: "left" });
    pen.text("F4", "Konto", { bold: true, size: 11, color: NAVY, align: "left" });

    var ruleRows = ruleset.toRows();
    if (ruleRows.length) {
      sheet.getRange(5, 2, ruleRows.length, 2).setValues(ruleRows);
    }

    accounts = accounts || {};
    var categories = ruleset.categories();
    if (categories.length) {
      var catRows = categories.map(function (c) { return [c, accounts[c] || ""]; });
      sheet.getRange(5, 5, catRows.length, 2).setValues(catRows);
    }
    var lastCatRow = 4 + Math.max(categories.length, 1);
    var accSheet = this.ss.getSheetByName(SHEET_ACCOUNTS);
    if (accSheet) pen.dropdown(_range(6, 5, 6, lastCatRow), accSheet.getRange("B5:B50"));

    pen.freeze(4, 0);
  }

  read_rules() {
    var sheet = this.ss.getSheetByName(SHEET_RULES);
    if (!sheet) return [];
    var lastRow = this.used_row_count(sheet);
    if (lastRow < 5) return [];
    var values = sheet.getRange(5, 2, lastRow - 4, 2).getValues();
    var out = [];
    values.forEach(function (row) {
      var keyword = String(row[0] || "").trim();
      var category = String(row[1] || "").trim();
      if (keyword && category) out.push([keyword, category]);
    });
    return out;
  }

  read_category_accounts() {
    var sheet = this.ss.getSheetByName(SHEET_RULES);
    if (!sheet) return {};
    var lastRow = this.used_row_count(sheet);
    if (lastRow < 5) return {};
    var values = sheet.getRange(5, 5, lastRow - 4, 2).getValues();
    var out = {};
    values.forEach(function (row) {
      var category = String(row[0] || "").trim();
      var account = String(row[1] || "").trim();
      if (category && account) out[category] = account;
    });
    return out;
  }

  // -- Konti ------------------------------------------------------------
  write_accounts(sheet, rows) {
    var pen = this.pen(sheet);
    var widths = { 1: 30, 2: 160, 3: 130, 4: 110, 5: 170 };
    for (var col in widths) sheet.setColumnWidth(parseInt(col, 10), widths[col]);

    pen.merge("B2:F2");
    pen.text("B2", "Konti", { font: FONT_TITLE, size: 18, bold: true, color: ORANGE, align: "left" });
    pen.merge("B3:F3");
    pen.text("B3", "Nuværende saldo - Prognosen bruger summen som startsaldo. Sæt dit eget " +
      "kontonummer under \"Kontonummer(e)\" og kør \"Opdatér kategorier i budget\" igen for at " +
      "få overførsler mellem dine egne konti sat til \"" + IGNORED_CATEGORY + "\" i stedet for at " +
      "tælle som indtægt/udgift - adskil flere numre med komma.",
      { size: 9, italic: true, color: MUTED, align: "left", wrap: true, valign: "middle" });
    sheet.setRowHeight(3, 80);

    pen.text("B4", "Konto", { bold: true, size: 11, color: NAVY, align: "left" });
    pen.text("C4", "Type", { bold: true, size: 11, color: NAVY, align: "left" });
    pen.text("D4", "Saldo", { bold: true, size: 11, color: NAVY, align: "right" });
    pen.text("E4", "Kontonummer(e)", { bold: true, size: 11, color: NAVY, align: "left" });

    rows = rows.length ? rows : [["", "", 0.0, ""]];
    for (var r = 0; r < rows.length; r++) {
      var name = rows[r][0], kind = rows[r][1], balance = rows[r][2], number = rows[r][3];
      var rr = ACC_HEADER_ROW + r + 1;
      pen.text(_ref(ACC_COL_NAME, rr), name, { bold: true, size: 10, color: DARK, align: "left" });
      pen.text(_ref(ACC_COL_TYPE, rr), kind, { size: 10, color: TEXT_GREY, align: "left" });
      pen.number(_ref(ACC_COL_BALANCE, rr), balance, { fmt: FMT_CURRENCY, size: 10, color: DARK, align: "right", bg: PEACH });
      pen.text(_ref(ACC_COL_NUMBER, rr), number, { size: 10, color: DARK, align: "left", bg: PEACH });
    }

    var lastRow = ACC_HEADER_ROW + rows.length;
    var totalRow = lastRow + 2;
    pen.text(_ref(ACC_COL_NAME, totalRow), "I alt", { bold: true, size: 10, color: NAVY, align: "left" });
    pen.formula(_ref(ACC_COL_BALANCE, totalRow),
      "=SUM(" + _range(ACC_COL_BALANCE, ACC_FIRST_ROW, ACC_COL_BALANCE, lastRow) + ")",
      { fmt: FMT_CURRENCY, bold: true, color: NAVY, align: "right" });
    pen.freeze(ACC_HEADER_ROW, 0);
  }

  ensure_accounts_sheet() {
    if (this.has_sheet(SHEET_ACCOUNTS)) return;
    var sheet = this.sheet(SHEET_ACCOUNTS);
    this.write_accounts(sheet, [
      ["Lønkonto", "Lønkonto", 0.0, ""],
      ["Budgetkonto", "Budgetkonto", 0.0, ""],
      ["Opsparingskonto", "Opsparingskonto", 0.0, ""],
    ]);
  }

  scan_accounts() {
    var sheet = this.ss.getSheetByName(SHEET_ACCOUNTS);
    if (!sheet) return [];
    var lastRow = this.used_row_count(sheet);
    if (lastRow < ACC_FIRST_ROW) return [];
    var values = sheet.getRange(ACC_FIRST_ROW, ACC_COL_NAME, lastRow - ACC_FIRST_ROW + 1, 4).getValues();
    var out = [];
    for (var i = 0; i < values.length; i++) {
      var name = String(values[i][0] || "").trim();
      if (!name || fold(name) === "i alt") continue;
      var kind = String(values[i][1] || "").trim();
      var balanceVal = values[i][2];
      var balance = typeof balanceVal === "number" ? balanceVal : 0.0;
      var number = String(values[i][3] || "").trim();
      out.push([name, kind, balance, number, ACC_FIRST_ROW + i]);
    }
    return out;
  }

  read_account_numbers() {
    var out = [];
    this.scan_accounts().forEach(function (row) { out = out.concat(parseAccountNumbers(row[3])); });
    return out;
  }

  accounts_total() {
    return this.scan_accounts().reduce(function (sum, row) { return sum + row[2]; }, 0.0);
  }

  // -- Budgetforslag ------------------------------------------------------
  write_plan(sheet, plan, targets, accounts, planAccounts) {
    var pen = this.pen(sheet);
    var widths = { 1: 30, 2: 190, 3: 115, 4: 115, 5: 100, 6: 130 };
    for (var col in widths) sheet.setColumnWidth(parseInt(col, 10), widths[col]);

    pen.merge("B2:F2");
    pen.text("B2", "Budgetforslag", { font: FONT_TITLE, size: 18, bold: true, color: ORANGE, align: "left" });
    pen.merge("B3:F3");
    var intro;
    if (plan.uncertain) {
      intro = "Udkast til et fast månedsbudget. Bemærk: bygger kun på " + plan.months.length +
        " måned(er) - tallene er usikre. Ret Mål-kolonnen som du vil.";
    } else {
      intro = "Udkast til et fast månedsbudget ud fra hele perioden. Ret Mål-kolonnen som du vil - " +
        "resten genberegnes automatisk, når du kører \"Opdatér kategorier i budget\" igen.";
    }
    pen.text("B3", intro, { size: 9, italic: true, color: MUTED, align: "left", wrap: true, valign: "middle" });
    sheet.setRowHeight(3, 46);

    var header = PLAN_HEADER_ROW;
    pen.text(_ref(2, header), "Kategori", { bold: true, size: 11, color: NAVY, align: "left" });
    pen.text(_ref(3, header), "Gennemsnit/md", { bold: true, size: 10, color: NAVY, align: "right" });
    pen.text(_ref(4, header), "Mål", { bold: true, size: 11, color: NAVY, align: "right" });
    pen.text(_ref(5, header), "Forskel", { bold: true, size: 10, color: NAVY, align: "right" });
    pen.text(_ref(6, header), "Konto", { bold: true, size: 10, color: NAVY, align: "left" });

    targets = targets || {};
    accounts = accounts || {};
    planAccounts = planAccounts || {};
    var accountRows = {};

    var row = header + 1;
    pen.text(_ref(2, row), "INDTÆGTER", { bold: true, size: 10, color: ORANGE, align: "left" });
    row += 1;
    var incomeFirst = row;
    plan.income.forEach((c) => {
      this._writePlanRow(pen, row, c, targets, accounts, planAccounts, accountRows, 1);
      row += 1;
    });
    var incomeLast = row - 1;
    row += 1;
    var incomeTotalRow = row;
    pen.text(_ref(2, row), "INDTÆGTER I ALT", { bold: true, size: 10, color: NAVY, align: "left" });
    if (incomeLast >= incomeFirst) {
      pen.formula(_ref(4, row), "=SUM(" + _range(4, incomeFirst, 4, incomeLast) + ")",
        { fmt: FMT_CURRENCY, bold: true, color: NAVY, align: "right" });
    } else {
      pen.number(_ref(4, row), 0, { fmt: FMT_CURRENCY, bold: true, color: NAVY, align: "right" });
    }
    row += 2;

    var expenseFirst = row;
    [GROUP_FIXED, GROUP_VARIABLE, GROUP_PERIODIC].forEach((group) => {
      var members = plan.groups[group] || [];
      if (!members.length) return;
      pen.text(_ref(2, row), GROUP_LABELS[group], { bold: true, size: 10, color: ORANGE, align: "left" });
      row += 1;
      members.forEach((c) => {
        this._writePlanRow(pen, row, c, targets, accounts, planAccounts, accountRows, -1);
        row += 1;
      });
      row += 1;
    });
    var expenseLast = row - 1;

    var expenseTotalRow = row;
    pen.text(_ref(2, row), "UDGIFTER I ALT", { bold: true, size: 10, color: NAVY, align: "left" });
    pen.formula(_ref(4, row), "=SUM(" + _range(4, expenseFirst, 4, expenseLast) + ")",
      { fmt: FMT_CURRENCY, bold: true, color: NAVY, align: "right" });
    row += 2;

    var savingsRow = row;
    pen.text(_ref(2, row), "TIL OPSPARING", { bold: true, size: 11, color: ORANGE, align: "left" });
    pen.formula(_ref(4, row), "=" + _ref(4, incomeTotalRow) + "-" + _ref(4, expenseTotalRow),
      { fmt: FMT_SIGNED, bold: true, color: ORANGE, align: "right" });

    pen.freeze(header, 0);
    return { income: incomeTotalRow, expense: expenseTotalRow, savings: savingsRow, accounts: accountRows };
  }

  _writePlanRow(pen, row, c, targets, accounts, planAccounts, accountRows, sign) {
    pen.text(_ref(2, row), c.category, { size: 10, color: DARK, align: "left" });
    pen.number(_ref(3, row), c.mean_all, { fmt: FMT_CURRENCY, size: 10, color: MUTED, align: "right" });

    var key = c.kind + "|" + c.category;
    var target = targets.hasOwnProperty(key) ? targets[key] : c.suggestion;
    pen.number(_ref(4, row), target, { fmt: FMT_CURRENCY, size: 10, bold: true, color: DARK, align: "right", bg: PEACH });
    pen.formula(_ref(5, row), "=" + _ref(4, row) + "-" + _ref(3, row), { fmt: FMT_SIGNED, size: 10, color: TEXT_GREY, align: "right" });

    var account = planAccounts[key] || accounts[c.category] || "";
    pen.text(_ref(6, row), account, { size: 10, color: DARK, align: "left" });
    if (account) {
      if (!accountRows[account]) accountRows[account] = [];
      accountRows[account].push([row, sign]);
    }

    if (c.sample && c.sample.length) {
      var lines = ["Gennemsnit: " + Math.round(c.mean_all).toLocaleString("da-DK") + " kr./md",
        c.transaction_count + " postering(er) i perioden."];
      pen.note(_ref(2, row), lines.join("\n"));
    }
  }

  read_targets() {
    var sheet = this.ss.getSheetByName(SHEET_PLAN);
    if (!sheet) return {};
    var lastRow = this.used_row_count(sheet);
    var labels = {};
    ["INDTÆGTER", "INDTÆGTER I ALT", "UDGIFTER I ALT", "TIL OPSPARING"].forEach(function (l) { labels[fold(l)] = true; });
    Object.keys(GROUP_LABELS).forEach(function (g) { labels[fold(GROUP_LABELS[g])] = true; });
    var out = {};
    if (lastRow <= PLAN_HEADER_ROW) return out;
    var values = sheet.getRange(PLAN_HEADER_ROW + 1, 2, lastRow - PLAN_HEADER_ROW, 3).getValues();
    values.forEach(function (row) {
      var cat = String(row[0] || "").trim();
      if (!cat || labels[fold(cat)]) return;
      var val = row[2];
      if (typeof val !== "number") return;
      out[KIND_EXPENSE + "|" + cat] = val;
      out[KIND_INCOME + "|" + cat] = val;
    });
    return out;
  }

  read_plan_accounts() {
    var sheet = this.ss.getSheetByName(SHEET_PLAN);
    if (!sheet) return {};
    var lastRow = this.used_row_count(sheet);
    var labels = {};
    ["INDTÆGTER", "INDTÆGTER I ALT", "UDGIFTER I ALT", "TIL OPSPARING"].forEach(function (l) { labels[fold(l)] = true; });
    Object.keys(GROUP_LABELS).forEach(function (g) { labels[fold(GROUP_LABELS[g])] = true; });
    var out = {};
    if (lastRow <= PLAN_HEADER_ROW) return out;
    var values = sheet.getRange(PLAN_HEADER_ROW + 1, 2, lastRow - PLAN_HEADER_ROW, 5).getValues();
    values.forEach(function (row) {
      var cat = String(row[0] || "").trim();
      if (!cat || labels[fold(cat)]) return;
      var account = String(row[4] || "").trim();
      if (account) {
        out[KIND_EXPENSE + "|" + cat] = account;
        out[KIND_INCOME + "|" + cat] = account;
      }
    });
    return out;
  }

  // -- Prognose -----------------------------------------------------------
  write_forecast(sheet, planRows, lastMonthKey, fallbackStart) {
    var pen = this.pen(sheet);
    var widths = { 1: 30, 2: 110, 3: 105, 4: 105 };
    for (var col in widths) sheet.setColumnWidth(parseInt(col, 10), widths[col]);

    pen.merge("B2:D2");
    pen.text("B2", "Prognose", { font: FONT_TITLE, size: 18, bold: true, color: ORANGE, align: "left" });
    pen.merge("B3:D3");
    pen.text("B3", "Sådan udvikler saldoen sig de næste " + FORECAST_MONTHS + " måneder, hvis du " +
      "rammer målene i \"Budgetforslag\".",
      { size: 9, italic: true, color: MUTED, align: "left", wrap: true, valign: "middle" });
    sheet.setRowHeight(3, 40);

    var accountsTotal = this.accounts_total();
    pen.text("B5", "Startsaldo", { bold: true, size: 10, color: NAVY, align: "left" });
    if (accountsTotal) {
      pen.formula("D5", "=SUM('" + SHEET_ACCOUNTS + "'!" +
        _range(ACC_COL_BALANCE, ACC_FIRST_ROW, ACC_COL_BALANCE, ACC_FIRST_ROW + 20) + ")",
        { fmt: FMT_CURRENCY, size: 10, color: DARK, align: "right" });
    } else {
      pen.number("D5", fallbackStart, { fmt: FMT_CURRENCY, size: 10, color: DARK, align: "right", bg: PEACH });
    }
    pen.text("B6", "Netto pr. måned med dine mål", { size: 10, color: TEXT_GREY, align: "left" });
    pen.formula("D6", "='" + SHEET_PLAN + "'!" + _ref(4, planRows.savings),
      { fmt: FMT_SIGNED, size: 10, bold: true, color: ORANGE, align: "right" });

    var head = FORECAST_HEADER_ROW;
    pen.text(_ref(2, head), "Måned", { bold: true, size: 11, color: NAVY, align: "left" });
    pen.text(_ref(3, head), "Netto/md", { bold: true, size: 10, color: NAVY, align: "right" });
    pen.text(_ref(4, head), "Saldo", { bold: true, size: 11, color: NAVY, align: "right" });

    var startMonth = _nextMonth(lastMonthKey);
    for (var i = 0; i < FORECAST_MONTHS; i++) {
      var row = head + 1 + i;
      var month = _addMonths(startMonth, i);
      pen.date(_ref(2, row), month, { fmt: FMT_MONTH, size: 10, color: DARK, align: "left" });
      pen.formula(_ref(3, row), "=$D$6", { fmt: FMT_SIGNED, size: 10, color: ORANGE, align: "right" });
      var prevRef = i === 0 ? "$D$5" : _ref(4, row - 1);
      pen.formula(_ref(4, row), "=" + prevRef + "+" + _ref(3, row),
        { fmt: FMT_CURRENCY, size: 10, bold: true, color: ORANGE, align: "right" });
    }

    var lastRow = head + FORECAST_MONTHS;
    this._addForecastChart(sheet, head, lastRow);
    pen.freeze(head, 0);
  }

  _addForecastChart(sheet, headerRow, lastRow) {
    var dataRange = sheet.getRange(headerRow, 4, lastRow - headerRow + 1, 1);
    var catRange = sheet.getRange(headerRow + 1, 2, lastRow - headerRow, 1);
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(catRange)
      .addRange(dataRange)
      .setOption("title", "Saldoudvikling")
      .setOption("legend", { position: "none" })
      .setOption("width", 620)
      .setOption("height", 320)
      .setPosition(2, 6, 0, 0)
      .build();
    sheet.insertChart(chart);
  }

  read_forecast_start() {
    var sheet = this.ss.getSheetByName(SHEET_FORECAST);
    if (!sheet) return null;
    var val = sheet.getRange("D5").getValue();
    return typeof val === "number" ? val : null;
  }

  // -- Alle måneder ---------------------------------------------------
  write_months(sheet, summary) {
    var pen = this.pen(sheet);
    sheet.setColumnWidth(1, 30);
    sheet.setColumnWidth(2, 190);

    pen.merge("B2:D2");
    pen.text("B2", SHEET_MONTHS, { font: FONT_TITLE, size: 18, bold: true, color: ORANGE, align: "left" });

    var head = 4;
    pen.text(_ref(2, head), "Kategori", { bold: true, size: 11, color: NAVY, align: "left" });
    summary.months.forEach(function (month, i) {
      var col = i + 3;
      sheet.setColumnWidth(col, 95);
      pen.date(_ref(col, head), _monthStart(month), { fmt: FMT_MONTH_SHORT, bold: true, size: 10, color: NAVY, align: "right" });
    });

    var row = head + 1;
    summary.expense_categories.forEach((cat) => {
      pen.text(_ref(2, row), cat, { size: 10, color: DARK, align: "left" });
      summary.months.forEach((month, i) => {
        pen.number(_ref(i + 3, row), summary.value(KIND_EXPENSE, cat, month), { fmt: FMT_CURRENCY, size: 10, color: TEXT_GREY, align: "right" });
      });
      row += 1;
    });
    row += 1;
    summary.income_categories.forEach((cat) => {
      pen.text(_ref(2, row), cat, { size: 10, color: DARK, align: "left" });
      summary.months.forEach((month, i) => {
        pen.number(_ref(i + 3, row), summary.value(KIND_INCOME, cat, month), { fmt: FMT_CURRENCY, size: 10, color: NAVY, align: "right" });
      });
      row += 1;
    });

    pen.freeze(head, 1);
  }

  // -- orchestration ----------------------------------------------------
  build(transactions, ruleset, startBalance) {
    startBalance = startBalance || 0.0;
    var summary = new Summary(transactions);

    // Unlike openpyxl's lazy string-formula dropdowns, Apps Script data
    // validation needs a live Range object at build time - so the sheet a
    // dropdown reads from must already exist. Konti before Kategorier
    // before Transaktioner satisfies both the Kategorier account-dropdown
    // (reads Konti) and the Transaktioner category-dropdown (reads
    // Kategorier). Final tab order is fixed separately by _orderSheets().
    this.ensure_accounts_sheet();
    this.write_rules(this.sheet(SHEET_RULES), ruleset);
    this.write_transactions(this.sheet(SHEET_TX), transactions);

    if (summary.months.length > 1) {
      var plan = buildPlan(summary, transactions);
      var planRows = this.write_plan(this.sheet(SHEET_PLAN), plan);
      var accountsTotal = this.accounts_total();
      var fallback = accountsTotal ? accountsTotal : startBalance + summary.net;
      this.write_forecast(this.sheet(SHEET_FORECAST), planRows, summary.months[summary.months.length - 1], fallback);
      this.write_months(this.sheet(SHEET_MONTHS), summary);
    }

    this._removeBlankExtraSheets();
    this._orderSheets();
    return summary;
  }

  refresh(ruleset) {
    var transactions = this.read_transactions();
    if (!transactions.length) return [null, 0];
    var sheet = this.ss.getSheetByName(SHEET_TX);
    var accountNumbers = this.read_account_numbers();
    var changed = 0;
    transactions.forEach(function (t) {
      var category;
      if (accountNumbers.length && textMentionsAccount(t.text, accountNumbers)) {
        category = IGNORED_CATEGORY;
      } else {
        category = ruleset.categorise(t.text);
        if (category === DEFAULT_CATEGORY && t.category) category = t.category;
      }
      if (category !== t.category) {
        var col = t.income ? COL_INC_CAT : COL_EXP_CAT;
        sheet.getRange(t.source_row, col).setValue(category);
        t.category = category;
        changed += 1;
      }
    });
    this._applyIgnoredStyling(sheet, transactions);
    var summary = this.rebuild_summaries(transactions);
    return [summary, changed];
  }

  _applyIgnoredStyling(sheet, transactions) {
    var pen = this.pen(sheet);
    transactions.forEach(function (t) {
      var firstCol = t.income ? COL_INC_DATE : COL_EXP_DATE;
      var row = t.source_row;
      if (t.category === IGNORED_CATEGORY) {
        pen.style(_range(firstCol, row, firstCol + 3, row), { color: MUTED, italic: true });
      } else {
        pen.style(_range(firstCol, row, firstCol, row), { color: MUTED, italic: false });
        pen.style(_range(firstCol + 1, row, firstCol + 3, row), { color: TEXT_GREY, italic: false });
      }
    });
  }

  /** Add more transactions without touching what is already there.
   * Returns [addedCount, skippedCount, truncated, summaryOrNull]. */
  append_transactions(newTransactions) {
    var existing = this.read_transactions();
    var existingKeys = {};
    existing.forEach(function (t) {
      var key = _dedupDateKey(t.date) + "|" + t.text + "|" + Math.round(t.amount * 100) / 100;
      existingKeys[key] = (existingKeys[key] || 0) + 1;
    });

    var added = [];
    var skipped = 0;
    newTransactions.forEach(function (t) {
      var key = _dedupDateKey(t.date) + "|" + t.text + "|" + Math.round(t.amount * 100) / 100;
      if (existingKeys[key] > 0) {
        existingKeys[key] -= 1;
        skipped += 1;
      } else {
        added.push(t);
      }
    });

    if (!added.length) return [0, skipped, false, null];

    var sheet = this.ss.getSheetByName(SHEET_TX);
    var existingExpenses = existing.filter(function (t) { return !t.income; });
    var existingIncome = existing.filter(function (t) { return t.income; });
    var newExpenses = added.filter(function (t) { return t.amount < 0; });
    var newIncome = added.filter(function (t) { return t.amount >= 0; });

    var roomExp = Math.max(TX_MAX_ROW - TX_FIRST_ROW + 1 - existingExpenses.length, 0);
    var roomInc = Math.max(TX_MAX_ROW - TX_FIRST_ROW + 1 - existingIncome.length, 0);
    var truncated = false;
    if (newExpenses.length > roomExp) { newExpenses = newExpenses.slice(0, roomExp); truncated = true; }
    if (newIncome.length > roomInc) { newIncome = newIncome.slice(0, roomInc); truncated = true; }

    this._writeBlock(sheet, newExpenses, COL_EXP_DATE, TX_FIRST_ROW + existingExpenses.length);
    this._writeBlock(sheet, newIncome, COL_INC_DATE, TX_FIRST_ROW + existingIncome.length);

    var newExpCount = existingExpenses.length + newExpenses.length;
    var newIncCount = existingIncome.length + newIncome.length;
    var last = TX_FIRST_ROW + Math.max(newExpCount, newIncCount, 1) - 1;
    var pen = this.pen(sheet);
    this._styleTxRange(pen, last);

    var allTransactions = this.read_transactions();
    this._applyIgnoredStyling(sheet, allTransactions);
    this._addCategoryDropdown(sheet, COL_EXP_CAT, last);
    this._addCategoryDropdown(sheet, COL_INC_CAT, last);

    var summary = this.rebuild_summaries(allTransactions);
    return [added.length, skipped, truncated, summary];
  }

  /** Recreate Budgetforslag/Prognose/Alle måneder from a transaction list. */
  rebuild_summaries(transactions) {
    var summary = new Summary(transactions);

    var targets = this.read_targets();
    var planAccounts = this.read_plan_accounts();
    var categoryAccounts = this.read_category_accounts();
    var forecastStart = this.read_forecast_start();

    var self = this;
    [SHEET_PLAN, SHEET_FORECAST, SHEET_MONTHS].forEach(function (name) {
      if (self.has_sheet(name)) self.ss.deleteSheet(self.ss.getSheetByName(name));
    });

    this.ensure_accounts_sheet();

    if (summary.months.length > 1) {
      var plan = buildPlan(summary, transactions);
      var planRows = this.write_plan(this.sheet(SHEET_PLAN), plan, targets, categoryAccounts, planAccounts);
      var fallback;
      if (forecastStart !== null) {
        fallback = forecastStart;
      } else {
        var accountsTotal = this.accounts_total();
        fallback = accountsTotal ? accountsTotal : summary.net;
      }
      this.write_forecast(this.sheet(SHEET_FORECAST), planRows, summary.months[summary.months.length - 1], fallback);
      this.write_months(this.sheet(SHEET_MONTHS), summary);
    }

    this._orderSheets();
    return summary;
  }

  _orderSheets() {
    var self = this;
    SHEET_ORDER.slice().reverse().forEach(function (name) {
      var sheet = self.ss.getSheetByName(name);
      if (!sheet) return;
      self.ss.setActiveSheet(sheet);
      self.ss.moveActiveSheet(1);
    });
  }

  /** Remove a still-blank leftover default sheet (e.g. "Sheet1"/"Ark1"). */
  _removeBlankExtraSheets() {
    var ours = {};
    SHEET_ORDER.forEach(function (n) { ours[n] = true; });
    var sheets = this.ss.getSheets();
    if (sheets.length <= SHEET_ORDER.filter(this.has_sheet.bind(this)).length) return;
    var self = this;
    sheets.forEach(function (sheet) {
      if (ours[sheet.getName()]) return;
      if (self.ss.getSheets().length <= 1) return; // never delete the last sheet
      if (sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) {
        self.ss.deleteSheet(sheet);
      }
    });
  }
}

function _dedupDateKey(d) {
  return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0") + "-" +
    String(d.getUTCDate()).padStart(2, "0");
}

// ---- Code.gs ------------------------------------------------------------------

// Menu wiring and the server-side functions the import dialog calls.
//
// This is the only file that reacts to the Sheets UI directly - everything
// it calls (Budget.gs, Plan.gs, Rules.gs, SheetWriter.gs, ...) is plain
// logic with no dependency on the Sheets UI itself.

var APP_TITLE = "Budget fra CSV";

function onOpen() {
  SpreadsheetApp.getUi().createMenu(APP_TITLE)
    .addItem("Nyt budget fra CSV-fil…", "menuNewBudget")
    .addItem("Tilføj flere posteringer (CSV) til budget…", "menuAppendBudget")
    .addSeparator()
    .addItem("Opdatér kategorier i budget…", "menuRefreshBudget")
    .addToUi();
}

function menuNewBudget() {
  var ui = SpreadsheetApp.getUi();
  var book = new BudgetBook();
  if (book.has_budget()) {
    var resp = ui.alert(APP_TITLE,
      "Dette regneark har allerede et budget (arket \"Transaktioner\" findes). " +
      "Vil du overskrive det med et nyt?", ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }
  _showImportDialog("new");
}

function menuAppendBudget() {
  var ui = SpreadsheetApp.getUi();
  var book = new BudgetBook();
  if (!book.has_budget()) {
    ui.alert(APP_TITLE, "Dette regneark har ikke noget budget endnu. Brug " +
      "\"Nyt budget fra CSV-fil…\" først.", ui.ButtonSet.OK);
    return;
  }
  _showImportDialog("append");
}

function menuRefreshBudget() {
  var ui = SpreadsheetApp.getUi();
  var book = new BudgetBook();
  if (!book.has_budget()) {
    ui.alert(APP_TITLE, "Dette regneark har ikke noget budget endnu.", ui.ButtonSet.OK);
    return;
  }
  var rows = book.read_rules();
  var ruleset = rows.length ? RuleSet.fromRows(rows) : RuleSet.defaults();
  var result = book.refresh(ruleset);
  var summary = result[0], changed = result[1];
  if (!summary) {
    ui.alert(APP_TITLE, "Fandt ingen posteringer at opdatere.", ui.ButtonSet.OK);
    return;
  }
  ui.alert(APP_TITLE, "Budgettet er opdateret.\n\n" + changed + " postering(er) fik ny kategori.\n" +
    summary.months.length + " måned(er) i budgettet.", ui.ButtonSet.OK);
}

function _showImportDialog(mode) {
  var template = HtmlService.createTemplateFromFile("ImportDialog");
  template.mode = mode;
  template.title = mode === "new" ? "Nyt budget fra CSV-fil" : "Tilføj posteringer til budget";
  template.confirmLabel = mode === "new" ? "Opret budget" : "Tilføj posteringer";
  var html = template.evaluate().setWidth(520).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, APP_TITLE);
}

// -- called from ImportDialog.html via google.script.run -------------------

/**
 * Step 1: parse the (already decoded, in-browser) CSV text and detect the
 * column mapping, for the confirmation form. Nothing is written yet.
 */
function detectCsvMapping(csvText, fileName, encodingLabel) {
  var table = readTableFromText(csvText, { source: fileName, encoding: encodingLabel });
  if (!table.rows.length) {
    return { error: "Fandt ingen datarækker i filen." };
  }
  var mapping = detectMapping(table);
  return {
    headerNames: table.header_names(),
    describe: table.describe(),
    mapping: {
      date: mapping.date,
      text_index: mapping.text.length ? mapping.text[0] : null,
      amount: mapping.amount,
      amount_in: mapping.amount_in,
      amount_out: mapping.amount_out,
      balance: mapping.balance,
      decimal: mapping.decimal,
      dayfirst: mapping.dayfirst,
      notes: mapping.notes,
    },
  };
}

/** Step 2: build (or append) the budget with the user-confirmed mapping. */
function buildBudget(csvText, fileName, mappingObj, mode) {
  var table = readTableFromText(csvText, { source: fileName });

  var mapping = new ColumnMapping();
  mapping.date = mappingObj.date;
  mapping.text = mappingObj.text_index !== null && mappingObj.text_index !== undefined ? [mappingObj.text_index] : [];
  mapping.amount = mappingObj.amount;
  mapping.amount_in = mappingObj.amount_in;
  mapping.amount_out = mappingObj.amount_out;
  mapping.balance = mappingObj.balance;
  mapping.decimal = mappingObj.decimal;
  mapping.dayfirst = mappingObj.dayfirst;

  var book = new BudgetBook();
  var rows = book.read_rules();
  var ruleset = rows.length ? RuleSet.fromRows(rows) : RuleSet.defaults();
  var result = buildTransactions(table, mapping, ruleset,
    new BuildOptions({ decimal: mapping.decimal, dayfirst: mapping.dayfirst }));
  if (!result.ok) {
    return { error: "Der kunne ikke laves et budget.\n\n" + result.summary_lines().join("\n") };
  }

  if (mode === "new") {
    var startBalance = _startBalance(table, mapping, result);
    var summary = book.build(result.transactions, ruleset, startBalance);
    return { message: result.summary_lines().join("\n") + "\n\n" + summary.months.length + " måned(er) i budgettet." };
  }

  var appendResult = book.append_transactions(result.transactions);
  var added = appendResult[0], skipped = appendResult[1], truncated = appendResult[2], appendSummary = appendResult[3];
  if (!appendSummary) {
    return { message: "Ingen nye posteringer - alle " + skipped + " postering(er) fandtes i forvejen i budgettet." };
  }
  var lines = [added + " ny(e) postering(er) tilføjet."];
  if (skipped) lines.push(skipped + " postering(er) fandtes allerede og blev sprunget over.");
  if (truncated) {
    lines.push("Bemærk: budgettet kan højst rumme 4996 posteringer pr. type (udgift/indtægt) - " +
      "nogle af de nyeste blev ikke tilføjet.");
  }
  lines.push(appendSummary.months.length + " måned(er) i budgettet nu.");
  return { message: lines.join("\n") };
}

/** Balance before the first transaction, when the file has a balance column. */
function _startBalance(table, mapping, result) {
  if (mapping.balance === null || mapping.balance === undefined || !result.transactions.length) return 0.0;
  var first = result.transactions[0];
  var row = table.rows[first.source_row];
  if (!row) return 0.0;
  var balance = mapping.balance < row.length ? parseAmount(row[mapping.balance], mapping.decimal) : null;
  if (balance === null) return 0.0;
  return Math.round((balance - first.amount) * 100) / 100;
}
