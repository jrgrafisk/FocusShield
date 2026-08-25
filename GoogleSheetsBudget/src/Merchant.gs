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

if (typeof module !== "undefined") {
  module.exports = { merchantName: merchantName, merchantKey: merchantKey, ruleKeyword: ruleKeyword };
  var fold = require("./TextUtils.gs").fold;
  var squeeze = require("./TextUtils.gs").squeeze;
  var looksLikeAmount = require("./Parsing.gs").looksLikeAmount;
}
