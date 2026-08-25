# -*- coding: utf-8 -*-
"""Tests for xlsx_writer.py (no Excel/LibreOffice needed - pure openpyxl).

    python3 -m unittest discover -s tests -v
"""

from __future__ import annotations

import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "vendor"))

from budget_core import csvsniff
from budget_core.budget import BuildOptions, Summary, build_transactions
from budget_core.columns import detect_mapping
from budget_core.rules import IGNORED_CATEGORY, RuleSet
from xlsx_writer import ACC_COL_NUMBER, ACC_HEADER_ROW, BudgetBook, SHEET_TX


def _table(rows):
    table = csvsniff.Table.from_rows(rows)
    return table, detect_mapping(table)


def _build(rows, ruleset=None):
    table, mapping = _table(rows)
    ruleset = ruleset or RuleSet.defaults()
    result = build_transactions(table, mapping, ruleset,
                                BuildOptions(decimal=mapping.decimal, dayfirst=mapping.dayfirst))
    return table, mapping, result


MAY_ROWS = [
    ["Dato", "Tekst", "Beløb"],
    ["01-05-2025", "Løn maj", "28000,00"],
    ["02-05-2025", "Netto Amager", "-345,60"],
    ["10-05-2025", "Circle K Amager", "-450,00"],
]


class TransactionSheetTests(unittest.TestCase):
    def test_round_trip(self):
        _table_, _mapping, result = _build(MAY_ROWS)
        book = BudgetBook()
        book.write_transactions(book.sheet(SHEET_TX), result.transactions)
        readback = book.read_transactions()
        self.assertEqual(len(readback), 3)
        self.assertEqual(sorted(t.text for t in readback),
                         sorted(t.text for t in result.transactions))

    def test_ignored_rows_are_excluded_from_summary(self):
        _table_, _mapping, result = _build(MAY_ROWS)
        for t in result.transactions:
            if t.text == "Circle K Amager":
                t.category = IGNORED_CATEGORY
        book = BudgetBook()
        book.write_transactions(book.sheet(SHEET_TX), result.transactions)
        summary = Summary(book.read_transactions())
        self.assertEqual(summary.ignored_count, 1)
        self.assertEqual(summary.transaction_count, 2)


class RulesAndAccountsTests(unittest.TestCase):
    def test_rules_round_trip(self):
        book = BudgetBook()
        ruleset = RuleSet.defaults()
        book.write_rules(book.sheet("Kategorier"), ruleset, accounts={"Dagligvarer": "Budgetkonto"})
        rows = book.read_rules()
        self.assertEqual(len(rows), len(ruleset))
        self.assertEqual(book.read_category_accounts().get("Dagligvarer"), "Budgetkonto")

    def test_accounts_and_transfer_numbers(self):
        book = BudgetBook()
        book.ensure_accounts_sheet()
        ws = book.wb["Konti"]
        ws.cell(ACC_HEADER_ROW + 2, ACC_COL_NUMBER, "1234567890, 5301-1234567")
        numbers = book.read_account_numbers()
        self.assertEqual(numbers, ["1234567890", "53011234567"])


class BuildRefreshAppendTests(unittest.TestCase):
    """The end-to-end scenarios: build once, append with overlap, refresh
    with a new rule, catch an internal transfer, save/load."""

    def test_full_lifecycle(self):
        _table_, _mapping, result1 = _build(MAY_ROWS)
        book = BudgetBook()
        book.build(result1.transactions, RuleSet.defaults(), start_balance=10000.0)
        self.assertEqual(len(book.read_transactions()), 3)
        self.assertFalse(book.has_sheet("Budgetforslag"))  # only 1 month so far

        # Hand-pick a category on an existing row.
        ws = book.wb[SHEET_TX]
        for row in range(5, 8):
            if ws.cell(row, 4).value == "Circle K Amager":
                ws.cell(row, 5, "Arbejdskørsel")
                break

        # Append June: 1 day overlaps May (duplicate), 1 new transaction,
        # plus a new month.
        june_rows = [
            ["Dato", "Tekst", "Beløb"],
            ["02-05-2025", "Netto Amager", "-345,60"],   # duplicate of May
            ["01-06-2025", "Løn juni", "28000,00"],
            ["03-06-2025", "Netto Amager", "-289,00"],
        ]
        rows_saved = book.read_rules()
        ruleset2 = RuleSet(rows_saved) if rows_saved else RuleSet.defaults()
        _t2, _m2, result2 = _build(june_rows, ruleset2)
        added, skipped, truncated, summary2 = book.append_transactions(result2.transactions)
        self.assertEqual(added, 2)
        self.assertEqual(skipped, 1)
        self.assertFalse(truncated)
        self.assertIsNotNone(summary2)
        self.assertEqual(len(summary2.months), 2)
        self.assertTrue(book.has_sheet("Budgetforslag"))

        all_tx = book.read_transactions()
        self.assertEqual(len(all_tx), 5)
        self.assertTrue(any(t.text == "Circle K Amager" and t.category == "Arbejdskørsel"
                            for t in all_tx))

        # Re-appending the same file adds nothing.
        _t3, _m3, result3 = _build(june_rows, ruleset2)
        added2, skipped2, _t, summary3 = book.append_transactions(result3.transactions)
        self.assertEqual(added2, 0)
        self.assertEqual(skipped2, 3)
        self.assertIsNone(summary3)

        # Internal transfer: append a transfer-looking transaction, then
        # register the account number and refresh - it should get ignored.
        transfer_rows = [
            ["Dato", "Tekst", "Beløb"],
            ["05-06-2025", "Overført fra 1234567890", "5000,00"],
        ]
        _t4, _m4, result4 = _build(transfer_rows, ruleset2)
        added4, _s4, _tr4, _summary4 = book.append_transactions(result4.transactions)
        self.assertEqual(added4, 1)

        konti = book.wb["Konti"]
        konti.cell(ACC_HEADER_ROW + 1, ACC_COL_NUMBER, "1234567890")
        _summary5, changed = book.refresh(ruleset2)
        self.assertGreaterEqual(changed, 1)
        after = Summary(book.read_transactions())
        self.assertEqual(after.ignored_count, 1)
        self.assertTrue(any(t.category == "Løn" for t in book.read_transactions()))

        # Save/load round trip.
        path = "/tmp/_test_xlsx_writer_lifecycle.xlsx"
        book.save(path)
        try:
            book2 = BudgetBook.load(path)
            self.assertEqual(len(book2.read_transactions()), len(book.read_transactions()))
        finally:
            if os.path.exists(path):
                os.remove(path)


if __name__ == "__main__":
    unittest.main()
