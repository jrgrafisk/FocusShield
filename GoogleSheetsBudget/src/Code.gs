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
