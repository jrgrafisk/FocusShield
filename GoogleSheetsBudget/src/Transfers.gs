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

if (typeof module !== "undefined") {
  module.exports = {
    MIN_ACCOUNT_DIGITS: MIN_ACCOUNT_DIGITS,
    normalizeAccountNumber: normalizeAccountNumber,
    parseAccountNumbers: parseAccountNumbers,
    textMentionsAccount: textMentionsAccount,
  };
  var squeeze = require("./TextUtils.gs").squeeze;
}
