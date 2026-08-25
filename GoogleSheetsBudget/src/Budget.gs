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

if (typeof module !== "undefined") {
  module.exports = {
    KIND_INCOME: KIND_INCOME, KIND_EXPENSE: KIND_EXPENSE, Transaction: Transaction,
    BuildOptions: BuildOptions, BuildResult: BuildResult, buildTransactions: buildTransactions,
    Summary: Summary, summarise: summarise, SIGN_AUTO: SIGN_AUTO,
    SIGN_NEGATIVE_IS_EXPENSE: SIGN_NEGATIVE_IS_EXPENSE, SIGN_POSITIVE_IS_EXPENSE: SIGN_POSITIVE_IS_EXPENSE,
  };
  var merchantMod = require("./Merchant.gs");
  var merchantKey = merchantMod.merchantKey, merchantName = merchantMod.merchantName, ruleKeyword = merchantMod.ruleKeyword;
  var parsingMod = require("./Parsing.gs");
  var monthKey = parsingMod.monthKey, monthLabel = parsingMod.monthLabel, parseAmount = parsingMod.parseAmount,
    parseDate = parsingMod.parseDate;
  var rulesMod = require("./Rules.gs");
  var DEFAULT_CATEGORY = rulesMod.DEFAULT_CATEGORY, IGNORED_CATEGORY = rulesMod.IGNORED_CATEGORY, RuleSet = rulesMod.RuleSet;
  var squeeze = require("./TextUtils.gs").squeeze;
}
