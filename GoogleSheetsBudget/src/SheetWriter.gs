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

if (typeof module !== "undefined") {
  module.exports = {
    Pen: Pen, SheetTransaction: SheetTransaction, BudgetBook: BudgetBook,
    SHEET_TX: SHEET_TX, SHEET_RULES: SHEET_RULES, SHEET_PLAN: SHEET_PLAN,
    SHEET_FORECAST: SHEET_FORECAST, SHEET_ACCOUNTS: SHEET_ACCOUNTS, SHEET_MONTHS: SHEET_MONTHS,
    ACC_HEADER_ROW: ACC_HEADER_ROW, ACC_FIRST_ROW: ACC_FIRST_ROW, ACC_COL_NUMBER: ACC_COL_NUMBER,
    _toSheetDate: _toSheetDate, _fromSheetDate: _fromSheetDate, _col: _col,
  };
  var budgetMod = require("./Budget.gs");
  var KIND_EXPENSE = budgetMod.KIND_EXPENSE, KIND_INCOME = budgetMod.KIND_INCOME, Summary = budgetMod.Summary;
  var planMod = require("./Plan.gs");
  var GROUP_FIXED = planMod.GROUP_FIXED, GROUP_VARIABLE = planMod.GROUP_VARIABLE,
    GROUP_PERIODIC = planMod.GROUP_PERIODIC, GROUP_LABELS = planMod.GROUP_LABELS, buildPlan = planMod.buildPlan;
  var rulesMod = require("./Rules.gs");
  var DEFAULT_CATEGORY = rulesMod.DEFAULT_CATEGORY, IGNORED_CATEGORY = rulesMod.IGNORED_CATEGORY;
  var fold = require("./TextUtils.gs").fold;
  var monthKey = require("./Parsing.gs").monthKey;
  var transfersMod = require("./Transfers.gs");
  var parseAccountNumbers = transfersMod.parseAccountNumbers, textMentionsAccount = transfersMod.textMentionsAccount;
}
