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

if (typeof module !== "undefined") {
  module.exports = {
    CategoryPlan: CategoryPlan, BudgetPlan: BudgetPlan, buildPlan: buildPlan,
    GROUP_FIXED: GROUP_FIXED, GROUP_VARIABLE: GROUP_VARIABLE, GROUP_PERIODIC: GROUP_PERIODIC,
    GROUP_LABELS: GROUP_LABELS, SAMPLE_SIZE: SAMPLE_SIZE,
  };
  var budgetMod = require("./Budget.gs");
  var KIND_EXPENSE = budgetMod.KIND_EXPENSE, KIND_INCOME = budgetMod.KIND_INCOME;
  var fold = require("./TextUtils.gs").fold;
}
