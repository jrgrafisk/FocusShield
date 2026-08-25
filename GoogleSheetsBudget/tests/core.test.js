// Tests for the pure-logic modules in src/ (no Google Sheets needed).
//
// Run with:  node --test tests/core.test.js

const path = require("node:path");
const assert = require("node:assert/strict");
const { test } = require("node:test");

const SRC = path.join(__dirname, "..", "src");
const TextUtils = require(path.join(SRC, "TextUtils.gs"));
const Parsing = require(path.join(SRC, "Parsing.gs"));
const Transfers = require(path.join(SRC, "Transfers.gs"));
const Merchant = require(path.join(SRC, "Merchant.gs"));
const Rules = require(path.join(SRC, "Rules.gs"));
const CsvSniff = require(path.join(SRC, "CsvSniff.gs"));
const Columns = require(path.join(SRC, "Columns.gs"));
const Budget = require(path.join(SRC, "Budget.gs"));
const Plan = require(path.join(SRC, "Plan.gs"));

test("fold normalises case, accents and Nordic letters", () => {
  assert.equal(TextUtils.fold("Café Øst"), TextUtils.fold("CAFE OST"));
  assert.equal(TextUtils.fold("cafe øst"), TextUtils.fold("CAFE OST"));
});

test("isBlankRow", () => {
  assert.equal(TextUtils.isBlankRow(["", " ", "", ""]), true);
  assert.equal(TextUtils.isBlankRow(["", "x"]), false);
});

test("parseAmount handles Danish and international number formats", () => {
  assert.equal(Parsing.parseAmount("1.234,56"), 1234.56);
  assert.equal(Parsing.parseAmount("1,234.56"), 1234.56);
  assert.equal(Parsing.parseAmount("-345,60"), -345.6);
  assert.equal(Parsing.parseAmount("(99,00)"), -99);
  assert.equal(Parsing.parseAmount("DKK 1.500,00"), 1500);
  assert.equal(Parsing.parseAmount("abc"), null);
});

test("parseDate understands the common bank export layouts", () => {
  const iso = (d) => d.toISOString().slice(0, 10);
  assert.equal(iso(Parsing.parseDate("2024-05-31")), "2024-05-31");
  assert.equal(iso(Parsing.parseDate("31-05-2024")), "2024-05-31");
  assert.equal(iso(Parsing.parseDate("31/05/24")), "2024-05-31");
  assert.equal(iso(Parsing.parseDate("20240531")), "2024-05-31");
  assert.equal(iso(Parsing.parseDate("31. maj 2024")), "2024-05-31");
  assert.equal(iso(Parsing.parseDate("May 31, 2024")), "2024-05-31");
  assert.equal(Parsing.parseDate("31-02-2024"), null); // no such day
});

test("account number matching for internal-transfer detection", () => {
  assert.deepEqual(Transfers.parseAccountNumbers("1234567890, 5301-1234567"),
    ["1234567890", "53011234567"]);
  assert.equal(Transfers.textMentionsAccount("Overført fra 1234567890", ["1234567890"]), true);
  assert.equal(Transfers.textMentionsAccount("Til konto 5301 1234567890", ["1234567890"]), true);
  assert.equal(Transfers.textMentionsAccount("Løn maj", ["1234567890"]), false);
});

test("merchant name strips card/receipt noise but keeps Danish letters", () => {
  assert.equal(Merchant.merchantName("Dankort-nota 4711 NETTO 8021 KØBENHAVN 21.05 kl. 17.42"),
    "NETTO KØBENHAVN");
  assert.equal(Merchant.merchantName("Løvbjerg Østerbro 123,45"), "Løvbjerg Østerbro");
  assert.equal(Merchant.ruleKeyword("JOE & THE JUICE FISKETORVET 4200,00"), "JOE & THE JUICE");
});

test("RuleSet categorises with default rules and user additions", () => {
  const rs = Rules.RuleSet.defaults();
  assert.equal(rs.categorise("Netto Amager"), "Dagligvarer");
  assert.equal(rs.categorise("Circle K Amager"), "Transport");
  assert.equal(rs.categorise("Firma ApS betaling"), "Ukategoriseret"); // short keyword "irma" must not match inside "firma"
  assert.equal(rs.categorise("IRMA 1234"), "Dagligvarer");
  rs.add("hyggekælder", "Restaurant");
  assert.equal(rs.categorise("Hyggekælder København"), "Restaurant");
});

test("CsvSniff parses semicolon/comma, sep= hints and strips preambles", () => {
  const semi = CsvSniff.readTableFromText("Dato;Tekst;Beløb\n01-05-2025;Løn maj;28000,00\n");
  assert.deepEqual(semi.header, ["Dato", "Tekst", "Beløb"]);
  assert.equal(semi.delimiter, ";");

  const sepHint = CsvSniff.readTableFromText("sep=;\nDato;Tekst;Beløb\n01-05-2025;Løn maj;28000,00\n");
  assert.equal(sepHint.rows.length, 1);

  const withPreamble = CsvSniff.readTableFromText(
    "Bank Statement\nAccount: 1234\n\nDato;Tekst;Beløb\n01-05-2025;Løn maj;28000,00\n");
  assert.equal(withPreamble.notes.length, 1);
  assert.equal(withPreamble.rows.length, 1);
});

test("detectMapping finds date/text/amount columns for a typical export", () => {
  const table = CsvSniff.readTableFromText(
    "Dato;Tekst;Beløb;Saldo\n01-05-2025;Løn maj;28000,00;30000,00\n" +
    "02-05-2025;Netto Amager;-345,60;29654,40\n");
  const mapping = Columns.detectMapping(table);
  assert.equal(mapping.date, 0);
  assert.deepEqual(mapping.text, [1]);
  assert.equal(mapping.amount, 2);
  assert.equal(mapping.balance, 3);
});

test("buildTransactions end-to-end produces categorised transactions", () => {
  const csv = "Dato;Tekst;Beløb\n01-05-2025;Løn maj;28000,00\n" +
    "02-05-2025;Netto Amager;-345,60\n10-05-2025;Circle K Amager;-450,00\n";
  const table = CsvSniff.readTableFromText(csv);
  const mapping = Columns.detectMapping(table);
  const ruleset = Rules.RuleSet.defaults();
  const result = Budget.buildTransactions(table, mapping, ruleset,
    new Budget.BuildOptions({ decimal: mapping.decimal, dayfirst: mapping.dayfirst }));
  assert.equal(result.transactions.length, 3);
  const byText = Object.fromEntries(result.transactions.map((t) => [t.text, t]));
  assert.equal(byText["Løn maj"].category, "Løn");
  assert.equal(byText["Netto Amager"].category, "Dagligvarer");
  assert.equal(byText["Circle K Amager"].category, "Transport");

  const summary = new Budget.Summary(result.transactions);
  assert.deepEqual(summary.months, ["2025-05"]);
  assert.ok(Math.abs(summary.net - 27204.4) < 1e-9);
});

test("buildPlan classifies fixed vs periodic costs across several months", () => {
  const rows = ["Dato;Tekst;Beløb"];
  ["01", "02", "03", "04", "05", "06", "07"].forEach((mm) => {
    rows.push(`01-${mm}-2025;Løn;30000,00`);
    rows.push(`05-${mm}-2025;Husleje;-9000,00`);
  });
  rows.push("15-01-2025;El-afregning;-1200,00");
  const csv = rows.join("\n") + "\n";
  const table = CsvSniff.readTableFromText(csv);
  const mapping = Columns.detectMapping(table);
  const ruleset = Rules.RuleSet.defaults();
  const result = Budget.buildTransactions(table, mapping, ruleset,
    new Budget.BuildOptions({ decimal: mapping.decimal, dayfirst: mapping.dayfirst }));
  const summary = new Budget.Summary(result.transactions);
  const plan = Plan.buildPlan(summary, result.transactions);

  const bolig = plan.categories.find((c) => c.category === "Bolig");
  assert.equal(bolig.group, Plan.GROUP_FIXED);
  const uncategorised = plan.categories.find((c) => c.category === "Ukategoriseret");
  assert.equal(uncategorised.group, Plan.GROUP_PERIODIC);
});
