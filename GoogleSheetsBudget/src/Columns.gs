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

if (typeof module !== "undefined") {
  module.exports = {
    ColumnStats: ColumnStats, ColumnMapping: ColumnMapping,
    analyseColumns: analyseColumns, detectMapping: detectMapping,
  };
  var fold = require("./TextUtils.gs").fold;
  var P = require("./Parsing.gs");
  var looksLikeDate = P.looksLikeDate, looksLikeAmount = P.looksLikeAmount,
    parseAmount = P.parseAmount, detectDecimalSeparator = P.detectDecimalSeparator,
    detectDayfirst = P.detectDayfirst;
}
