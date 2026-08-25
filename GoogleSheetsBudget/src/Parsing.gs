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

if (typeof module !== "undefined") {
  module.exports = {
    parseAmount: parseAmount, looksLikeAmount: looksLikeAmount,
    detectDecimalSeparator: detectDecimalSeparator, parseDate: parseDate,
    looksLikeDate: looksLikeDate, detectDayfirst: detectDayfirst,
    monthKey: monthKey, monthLabel: monthLabel, MONTH_NAMES_DA: MONTH_NAMES_DA,
  };
  var fold = require("./TextUtils.gs").fold;
}
