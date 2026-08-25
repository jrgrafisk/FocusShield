// Exercises SheetWriter.gs's BudgetBook end to end against the mock
// SpreadsheetApp in mock_spreadsheet_app.js. This is the one file that
// cannot be checked against a real Google Sheet from this environment, so
// this suite is the closest thing to a regression test it gets: it runs
// build() / append_transactions() / refresh() for real and checks what
// actually landed in the mock sheet cells, mirroring
// BudgetScript/tests/test_xlsx_writer.py's test_full_lifecycle. It cannot
// verify the *formulas* (no live Sheets engine here) - only that every
// call succeeds and that values written via setValues() round-trip
// correctly through the read_* methods.

const path = require("node:path");
const assert = require("node:assert/strict");
const { test, beforeEach } = require("node:test");

const { installGlobals } = require("./mock_spreadsheet_app.js");

const SRC = path.join(__dirname, "..", "src");
const CsvSniff = require(path.join(SRC, "CsvSniff.gs"));
const Columns = require(path.join(SRC, "Columns.gs"));
const Budget = require(path.join(SRC, "Budget.gs"));
const Rules = require(path.join(SRC, "Rules.gs"));

function freshSheetWriter() {
  installGlobals();
  delete require.cache[require.resolve(path.join(SRC, "SheetWriter.gs"))];
  return require(path.join(SRC, "SheetWriter.gs"));
}

function build(rows, ruleset) {
  const table = CsvSniff.readTableFromText(rows);
  const mapping = Columns.detectMapping(table);
  ruleset = ruleset || Rules.RuleSet.defaults();
  const result = Budget.buildTransactions(table, mapping, ruleset,
    new Budget.BuildOptions({ decimal: mapping.decimal, dayfirst: mapping.dayfirst }));
  return { table, mapping, result };
}

const MAY_CSV = "Dato;Tekst;Beløb\n01-05-2025;Løn maj;28000,00\n" +
  "02-05-2025;Netto Amager;-345,60\n10-05-2025;Circle K Amager;-450,00\n";

test("build() creates a Transaktioner sheet that round-trips", () => {
  const SW = freshSheetWriter();
  const { result } = build(MAY_CSV);
  const book = new SW.BudgetBook();
  book.build(result.transactions, Rules.RuleSet.defaults(), 10000.0);

  assert.ok(book.has_sheet(SW.SHEET_TX));
  assert.ok(book.has_sheet(SW.SHEET_RULES));
  assert.ok(book.has_sheet(SW.SHEET_ACCOUNTS));
  // Only one month of data - no plan/forecast/months sheets yet.
  assert.equal(book.has_sheet(SW.SHEET_PLAN), false);

  const readBack = book.read_transactions();
  assert.equal(readBack.length, 3);
  assert.deepEqual(readBack.map((t) => t.text).sort(), ["Circle K Amager", "Løn maj", "Netto Amager"]);
  const byText = Object.fromEntries(readBack.map((t) => [t.text, t]));
  assert.equal(byText["Netto Amager"].category, "Dagligvarer");
  assert.ok(Math.abs(byText["Netto Amager"].amount + 345.6) < 1e-9);
});

test("full lifecycle: build, hand-pick category, append with overlap, transfer detection, refresh", () => {
  const SW = freshSheetWriter();
  const { result: result1 } = build(MAY_CSV);
  const book = new SW.BudgetBook();
  book.build(result1.transactions, Rules.RuleSet.defaults(), 10000.0);

  // Hand-pick a category on an existing row (find it via the sheet API,
  // like a user editing the cell directly).
  const txSheet = book.ss.getSheetByName(SW.SHEET_TX);
  for (const t of book.read_transactions()) {
    if (t.text === "Circle K Amager") {
      txSheet.getRange(t.source_row, 5).setValue("Arbejdskørsel");
      break;
    }
  }

  // Append June: 1 day overlaps May (duplicate), 1 new transaction, plus a
  // new month - this should also create the Budgetforslag/Prognose/Alle
  // måneder sheets since there are now 2 months.
  const juneCsv = "Dato;Tekst;Beløb\n02-05-2025;Netto Amager;-345,60\n" +
    "01-06-2025;Løn juni;28000,00\n03-06-2025;Netto Amager;-289,00\n";
  const rulesRows = book.read_rules();
  const ruleset2 = rulesRows.length ? Rules.RuleSet.fromRows(rulesRows) : Rules.RuleSet.defaults();
  const { result: result2 } = build(juneCsv, ruleset2);
  const [added, skipped, truncated, summary2] = book.append_transactions(result2.transactions);
  assert.equal(added, 2);
  assert.equal(skipped, 1);
  assert.equal(truncated, false);
  assert.ok(summary2);
  assert.equal(summary2.months.length, 2);
  assert.ok(book.has_sheet(SW.SHEET_PLAN));
  assert.ok(book.has_sheet(SW.SHEET_FORECAST));
  assert.ok(book.has_sheet(SW.SHEET_MONTHS));

  let allTx = book.read_transactions();
  assert.equal(allTx.length, 5);
  assert.ok(allTx.some((t) => t.text === "Circle K Amager" && t.category === "Arbejdskørsel"));

  // Re-appending the same file adds nothing.
  const { result: result3 } = build(juneCsv, ruleset2);
  const [added2, skipped2] = book.append_transactions(result3.transactions);
  assert.equal(added2, 0);
  assert.equal(skipped2, 3);

  // Internal transfer: append a transfer-looking transaction, then
  // register the account number and refresh - it should get ignored.
  const { result: result4 } = build("Dato;Tekst;Beløb\n05-06-2025;Overført fra 1234567890;5000,00\n", ruleset2);
  const [added4] = book.append_transactions(result4.transactions);
  assert.equal(added4, 1);

  const kontiSheet = book.ss.getSheetByName(SW.SHEET_ACCOUNTS);
  kontiSheet.getRange(SW.ACC_HEADER_ROW + 1, SW.ACC_COL_NUMBER).setValue("1234567890");
  const [summary5, changed] = book.refresh(ruleset2);
  assert.ok(changed >= 1);
  allTx = book.read_transactions();
  const ignoredCount = allTx.filter((t) => t.category === "Ignoreret").length;
  assert.equal(ignoredCount, 1);
  assert.ok(allTx.some((t) => t.category === "Løn"));
});

test("write_rules and write_accounts round-trip through read_rules/scan_accounts", () => {
  const SW = freshSheetWriter();
  const book = new SW.BudgetBook();
  const ruleset = Rules.RuleSet.defaults();
  book.ensure_accounts_sheet();
  book.write_rules(book.sheet(SW.SHEET_RULES), ruleset, { Dagligvarer: "Budgetkonto" });

  const rows = book.read_rules();
  assert.equal(rows.length, ruleset.length);
  assert.equal(book.read_category_accounts().Dagligvarer, "Budgetkonto");

  const konti = book.ss.getSheetByName(SW.SHEET_ACCOUNTS);
  konti.getRange(SW.ACC_HEADER_ROW + 2, SW.ACC_COL_NUMBER).setValue("1234567890, 5301-1234567");
  assert.deepEqual(book.read_account_numbers(), ["1234567890", "53011234567"]);
});
