# -*- coding: utf-8 -*-
"""Tests for the parsing/detection engine (no LibreOffice needed).

    python3 -m unittest discover -s tests -v
"""

from __future__ import annotations

import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from budget_core import csvsniff
from budget_core.budget import (BuildOptions, KIND_EXPENSE, KIND_INCOME,
                                SIGN_POSITIVE_IS_EXPENSE, build_transactions,
                                summarise)
from budget_core.columns import detect_mapping
from budget_core.parsing import (detect_dayfirst, detect_decimal_separator,
                                 month_key, parse_amount, parse_date)
from budget_core.merchant import merchant_key, merchant_name, rule_keyword
from budget_core.plan import (GROUP_FIXED, GROUP_PERIODIC, GROUP_VARIABLE,
                              build_plan)
from budget_core.rules import (DEFAULT_CATEGORY, IGNORED_CATEGORY, RuleSet,
                               default_categories)
from budget_core.textutils import fold

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")


def sample(name):
    return os.path.join(DATA, name)


def load(name, **kwargs):
    table = csvsniff.read_table(path=sample(name), **kwargs)
    mapping = detect_mapping(table)
    return table, mapping


def build(name, **kwargs):
    table, mapping = load(name)
    options = BuildOptions(decimal=mapping.decimal, dayfirst=mapping.dayfirst,
                           **kwargs)
    return table, mapping, build_transactions(table, mapping, RuleSet.defaults(),
                                              options)


class AmountTests(unittest.TestCase):

    def test_danish_and_english(self):
        self.assertEqual(parse_amount("1.234,56", ","), 1234.56)
        self.assertEqual(parse_amount("1,234.56", "."), 1234.56)
        self.assertEqual(parse_amount("-8.500,00", ","), -8500.0)
        self.assertEqual(parse_amount("123", ","), 123.0)

    def test_currency_and_spaces(self):
        self.assertEqual(parse_amount("1 234,56 kr.", ","), 1234.56)
        self.assertEqual(parse_amount("kr. 1.234,56", ","), 1234.56)
        self.assertEqual(parse_amount("$1,234.56", "."), 1234.56)
        self.assertEqual(parse_amount("€ 99,50", ","), 99.5)
        self.assertEqual(parse_amount("1 234,50", ","), 1234.5)

    def test_negative_notations(self):
        self.assertEqual(parse_amount("(120,45)", ","), -120.45)
        self.assertEqual(parse_amount("198,50-", ","), -198.5)
        self.assertEqual(parse_amount("−450,00", ","), -450.0)
        self.assertEqual(parse_amount("1.234,56 DR", ","), -1234.56)

    def test_rejects_junk(self):
        for value in ("", "   ", "abc", "12-05-2025x", None):
            self.assertIsNone(parse_amount(value))

    def test_separator_detection(self):
        self.assertEqual(detect_decimal_separator(
            ["-345,60", "28.500,00", "12,45"]), ",")
        self.assertEqual(detect_decimal_separator(
            ["-345.60", "28,500.00", "12.45"]), ".")
        self.assertEqual(detect_decimal_separator(["1.234", "5.678"]), ",")

    def test_guessing_without_column_context(self):
        self.assertEqual(parse_amount("1.234,56"), 1234.56)
        self.assertEqual(parse_amount("1,234.56"), 1234.56)
        self.assertEqual(parse_amount("1.234"), 1234.0)
        self.assertEqual(parse_amount("12,50"), 12.5)


class DateTests(unittest.TestCase):

    def test_layouts(self):
        expected = datetime.date(2025, 5, 31)
        for text in ("31-05-2025", "31/05/2025", "31.05.2025", "2025-05-31",
                     "20250531", "31052025", "31. maj 2025", "31 May 2025",
                     "2025-05-31 14:22:01", "31-05-2025 kl. 14:22"):
            self.assertEqual(parse_date(text), expected, text)

    def test_month_first(self):
        self.assertEqual(parse_date("05/31/2025", dayfirst=True),
                         datetime.date(2025, 5, 31))
        self.assertEqual(parse_date("02/03/2025", dayfirst=False),
                         datetime.date(2025, 2, 3))
        self.assertEqual(parse_date("02/03/2025", dayfirst=True),
                         datetime.date(2025, 3, 2))

    def test_two_digit_year(self):
        self.assertEqual(parse_date("31-05-25"), datetime.date(2025, 5, 31))
        self.assertEqual(parse_date("31-05-99"), datetime.date(1999, 5, 31))

    def test_dayfirst_detection(self):
        self.assertTrue(detect_dayfirst(["31/05/2025", "02/03/2025"]))
        self.assertFalse(detect_dayfirst(["05/31/2025", "12/25/2025"]))

    def test_rejects_junk(self):
        for value in ("", "Tekst", "12", "1.234,56", None):
            self.assertIsNone(parse_date(value))

    def test_month_key(self):
        self.assertEqual(month_key(datetime.date(2025, 5, 31)), "2025-05")


class SniffTests(unittest.TestCase):

    def test_semicolon_cp1252(self):
        table, _ = load("danskebank.csv")
        self.assertEqual(table.delimiter, ";")
        self.assertEqual(len(table.rows), 10)
        self.assertEqual(table.header[0], "Dato")
        self.assertIn("København", table.rows[0][1])

    def test_preamble_is_skipped(self):
        table, _ = load("nordea_preamble.csv")
        self.assertEqual(len(table.rows), 4)
        self.assertEqual(table.header[0], "Bogføringsdato")
        self.assertTrue(table.preamble)

    def test_comma_delimiter(self):
        table, _ = load("us_style.csv")
        self.assertEqual(table.delimiter, ",")
        self.assertEqual(len(table.rows), 5)

    def test_utf16_tab(self):
        table, _ = load("utf16_tab.csv")
        self.assertEqual(table.delimiter, "\t")
        self.assertTrue(table.encoding.startswith("utf-16"))
        self.assertEqual(len(table.rows), 4)
        self.assertIn("Ørsted", table.rows[0][1])

    def test_missing_header(self):
        table, _ = load("no_header.csv")
        self.assertIsNone(table.header)
        self.assertEqual(len(table.rows), 3)

    def test_sep_hint(self):
        table, _ = load("sep_hint.csv")
        self.assertEqual(table.delimiter, ";")
        self.assertEqual(len(table.rows), 3)

    def test_messy_file(self):
        table, _ = load("messy.csv")
        self.assertEqual(table.header[0], "Dato")
        # The ragged row is kept (padded), blank rows are dropped.
        self.assertTrue(len(table.rows) >= 5)


class MappingTests(unittest.TestCase):

    def test_balance_column_is_not_the_amount(self):
        table, mapping = load("danskebank.csv")
        self.assertEqual(table.header_name(mapping.date), "Dato")
        self.assertEqual(table.header_name(mapping.amount), "Beløb")
        self.assertEqual(table.header_name(mapping.balance), "Saldo")
        self.assertEqual(mapping.decimal, ",")
        self.assertTrue(mapping.dayfirst)

    def test_debit_credit_columns(self):
        table, mapping = load("haevet_indsat.csv")
        self.assertEqual(table.header_name(mapping.amount_out), "Hævet")
        self.assertEqual(table.header_name(mapping.amount_in), "Indsat")
        self.assertIsNone(mapping.amount)

    def test_english_file(self):
        table, mapping = load("us_style.csv")
        self.assertEqual(table.header_name(mapping.date), "Date")
        self.assertEqual(table.header_name(mapping.amount), "Amount")
        self.assertEqual(mapping.decimal, ".")
        self.assertFalse(mapping.dayfirst)

    def test_text_column(self):
        table, mapping = load("danskebank.csv")
        self.assertEqual([table.header_name(i) for i in mapping.text], ["Tekst"])

    def test_headerless_file(self):
        table, mapping = load("no_header.csv")
        self.assertEqual(mapping.date, 0)
        self.assertEqual(mapping.amount, 2)
        self.assertEqual(mapping.text, [1])


class BuildTests(unittest.TestCase):

    def test_danske_bank(self):
        _table, _mapping, result = build("danskebank.csv")
        self.assertEqual(len(result.transactions), 10)
        first = result.transactions[0]
        self.assertEqual(first.date, datetime.date(2025, 5, 2))
        self.assertAlmostEqual(first.amount, -345.60)
        self.assertEqual(first.category, "Dagligvarer")
        self.assertEqual(first.kind, KIND_EXPENSE)
        salary = [t for t in result.transactions if t.category == "Løn"]
        self.assertEqual(len(salary), 2)
        self.assertEqual(salary[0].kind, KIND_INCOME)

    def test_debit_credit_signs(self):
        _table, _mapping, result = build("haevet_indsat.csv")
        amounts = {t.text: t.amount for t in result.transactions}
        self.assertAlmostEqual(amounts["Ejerforening maj"], -2100.0)
        self.assertAlmostEqual(amounts["Løn"], 27400.0)

    def test_english_file(self):
        _table, _mapping, result = build("us_style.csv")
        self.assertEqual(len(result.transactions), 5)
        amazon = result.transactions[0]
        self.assertEqual(amazon.date, datetime.date(2025, 5, 2))
        self.assertAlmostEqual(amazon.amount, -120.45)
        self.assertEqual(amazon.category, "Shopping")

    def test_footer_and_ragged_rows_are_skipped(self):
        _table, _mapping, result = build("messy.csv")
        texts = [t.text for t in result.transactions]
        self.assertIn("MENY SLAGELSE", texts)
        self.assertNotIn("I alt", texts)
        self.assertEqual(result.skipped_no_date, 1)   # the "I alt" footer row
        self.assertEqual(result.skipped_no_amount, 1)  # the row without an amount

    def test_unsigned_amounts_become_expenses(self):
        """A column without a single minus sign is all spending."""
        table = csvsniff.Table.from_rows([
            ["Dato", "Tekst", "Beløb"],
            ["01-05-2025", "Netto", "198,45"],
            ["02-05-2025", "Circle K", "520,00"],
        ])
        mapping = detect_mapping(table)
        result = build_transactions(table, mapping, RuleSet.defaults())
        self.assertTrue(all(t.amount < 0 for t in result.transactions))
        self.assertTrue(any("fortegn" in w.lower() or "udgifter" in w.lower()
                            for w in result.warnings))

    def test_sign_rule_can_be_forced(self):
        table = csvsniff.Table.from_rows([
            ["Dato", "Tekst", "Beløb"],
            ["01-05-2025", "Netto", "198,45"],
            ["02-05-2025", "Løn", "-24.000,00"],
        ])
        mapping = detect_mapping(table)
        options = BuildOptions(sign=SIGN_POSITIVE_IS_EXPENSE)
        result = build_transactions(table, mapping, RuleSet.defaults(), options)
        amounts = {t.text: t.amount for t in result.transactions}
        self.assertAlmostEqual(amounts["Netto"], -198.45)
        self.assertAlmostEqual(amounts["Løn"], 24000.0)

    def test_uncategorised_are_reported(self):
        table = csvsniff.Table.from_rows([
            ["Dato", "Tekst", "Beløb"],
            ["01-05-2025", "Ukendt firma ApS", "-100,00"],
            ["02-05-2025", "Ukendt firma ApS", "-50,00"],
        ])
        mapping = detect_mapping(table)
        result = build_transactions(table, mapping, RuleSet.defaults())
        self.assertEqual(result.uncategorised[0][0], "Ukendt firma ApS")
        self.assertEqual(result.uncategorised[0][1], 2)
        self.assertAlmostEqual(result.uncategorised[0][2], -150.0)


class SummaryTests(unittest.TestCase):

    def test_monthly_totals(self):
        _table, _mapping, result = build("danskebank.csv")
        summary = summarise(result.transactions)
        self.assertEqual(summary.months, ["2025-05", "2025-06"])
        self.assertAlmostEqual(summary.value(KIND_INCOME, "Løn", "2025-05"), 28500.0)
        groceries = summary.value(KIND_EXPENSE, "Dagligvarer", "2025-05")
        self.assertAlmostEqual(groceries, -345.60)
        self.assertAlmostEqual(summary.total(KIND_INCOME), 57000.0)
        self.assertLess(summary.total(KIND_EXPENSE), 0)

    def test_category_ordering(self):
        _table, _mapping, result = build("danskebank.csv")
        summary = summarise(result.transactions)
        totals = [abs(summary.category_total(KIND_EXPENSE, c))
                  for c in summary.expense_categories]
        self.assertEqual(totals, sorted(totals, reverse=True))

    def test_ignored_postings_do_not_count(self):
        """A posting set to "Ignoreret" must vanish from every total, but
        stay out of the way rather than raise or get silently dropped."""
        table = csvsniff.Table.from_rows([
            ["Dato", "Tekst", "Beløb"],
            ["01-05-2025", "Netto", "-200,00"],
            ["02-05-2025", "Overførsel til opsparing", "-5000,00"],
            ["03-05-2025", "Løn", "20000,00"],
        ])
        mapping = detect_mapping(table)
        transactions = build_transactions(table, mapping, RuleSet.defaults(),
                                          BuildOptions(decimal=",")).transactions
        for t in transactions:
            if t.text == "Overførsel til opsparing":
                t.category = IGNORED_CATEGORY
        summary = summarise(transactions)

        self.assertNotIn(IGNORED_CATEGORY, summary.expense_categories)
        self.assertNotIn(IGNORED_CATEGORY, summary.income_categories)
        self.assertAlmostEqual(summary.total(KIND_EXPENSE), -200.0)
        self.assertAlmostEqual(summary.total(KIND_INCOME), 20000.0)
        self.assertAlmostEqual(summary.net, 19800.0)
        self.assertEqual(summary.transaction_count, 2)
        self.assertEqual(summary.ignored_count, 1)

    def test_ignored_only_month_drops_out(self):
        """A month with nothing but ignored postings is not a real month."""
        table = csvsniff.Table.from_rows([
            ["Dato", "Tekst", "Beløb"],
            ["01-05-2025", "Netto", "-200,00"],
            ["01-06-2025", "Test postering", "-50,00"],
        ])
        mapping = detect_mapping(table)
        transactions = build_transactions(table, mapping, RuleSet.defaults(),
                                          BuildOptions(decimal=",")).transactions
        for t in transactions:
            if t.month == "2025-06":
                t.category = IGNORED_CATEGORY
        summary = summarise(transactions)
        self.assertEqual(summary.months, ["2025-05"])


class PlanTests(unittest.TestCase):
    """The draft budget built from a full year of transactions."""

    @classmethod
    def setUpClass(cls):
        _table, _mapping, result = build("aar_2024.csv")
        cls.result = result
        cls.summary = summarise(result.transactions)
        cls.plan = build_plan(cls.summary, cls.result.transactions)
        cls.by_category = dict(((c.kind, c.category), c)
                               for c in cls.plan.categories)

    def entry(self, category, kind=KIND_EXPENSE):
        return self.by_category[(kind, category)]

    def test_twelve_months(self):
        self.assertEqual(len(self.plan.months), 12)
        self.assertFalse(self.plan.uncertain)

    def test_rent_is_fixed(self):
        rent = self.entry("Bolig")
        self.assertEqual(rent.group, GROUP_FIXED)
        self.assertAlmostEqual(rent.mean_all, 9200.0, places=2)
        self.assertEqual(rent.suggestion, 9200.0)

    def test_loan_and_insurance_are_fixed(self):
        for category in ("Lån og afdrag", "Forsikring og pension",
                         "Telefon og internet", "Abonnementer"):
            self.assertEqual(self.entry(category).group, GROUP_FIXED, category)

    def test_groceries_are_variable_and_use_the_median(self):
        """A steady monthly total is still a spending decision, not a bill."""
        groceries = self.entry("Dagligvarer")
        self.assertEqual(groceries.group, GROUP_VARIABLE)
        self.assertLess(groceries.variation, 0.3)      # steady, but still variable
        self.assertAlmostEqual(groceries.suggestion,
                               round(groceries.median / 50.0) * 50, places=2)

    def test_quarterly_bill_is_spread_over_the_year(self):
        power = self.entry("El, vand og varme")
        self.assertEqual(power.group, GROUP_PERIODIC)
        self.assertLess(power.coverage, 0.9)
        # Budget the yearly total per month, not the size of one bill.
        self.assertLess(power.suggestion, power.median)
        self.assertAlmostEqual(power.suggestion, power.mean_all, delta=50)

    def test_holiday_is_periodic(self):
        holiday = self.entry("Rejser")
        self.assertEqual(holiday.group, GROUP_PERIODIC)
        self.assertEqual(holiday.months_present, 1)
        self.assertAlmostEqual(holiday.suggestion, holiday.total / 12.0, delta=50)

    def test_income_is_conservative(self):
        """Bonus months must not inflate the budgeted income."""
        salary = self.entry("Løn", KIND_INCOME)
        self.assertLess(salary.suggestion, salary.mean_all)
        self.assertAlmostEqual(salary.suggestion, 31400.0, places=2)

    def test_quarterly_income_is_spread(self):
        benefit = self.entry("Offentlige ydelser", KIND_INCOME)
        self.assertEqual(benefit.months_present, 4)
        self.assertAlmostEqual(benefit.suggestion, benefit.mean_all, delta=100)

    def test_totals_add_up(self):
        expenses = sum(c.mean_all for c in self.plan.expenses)
        self.assertAlmostEqual(self.plan.expense_mean, expenses, places=2)
        self.assertAlmostEqual(self.plan.savings_mean,
                               self.plan.income_mean - self.plan.expense_mean,
                               places=2)
        # The suggestion is deliberately more careful than the average.
        self.assertLess(self.plan.savings_suggested, self.plan.savings_mean)
        self.assertGreater(self.plan.savings_suggested, 0)

    def test_every_category_is_grouped(self):
        grouped = sum(len(v) for v in self.plan.groups.values())
        self.assertEqual(grouped, len(self.plan.expenses))

    def test_short_period_is_flagged(self):
        _table, _mapping, result = build("danskebank.csv")
        plan = build_plan(summarise(result.transactions), result.transactions)
        self.assertTrue(plan.uncertain)

    def test_category_remembers_its_transactions(self):
        """So the spreadsheet can show a hover comment per row."""
        groceries = self.entry("Dagligvarer")
        self.assertGreater(groceries.transaction_count, 0)
        self.assertTrue(groceries.sample)
        self.assertLessEqual(len(groceries.sample), 8)
        # Largest first, so the comment leads with what actually matters.
        amounts = [abs(t.amount) for t in groceries.sample]
        self.assertEqual(amounts, sorted(amounts, reverse=True))
        for t in groceries.sample:
            self.assertEqual(t.category, "Dagligvarer")

    def test_sample_is_capped_but_count_is_not(self):
        holiday = self.entry("Rejser")
        self.assertLessEqual(len(holiday.sample), 8)
        self.assertGreaterEqual(holiday.transaction_count, len(holiday.sample))


class YearFileTests(unittest.TestCase):
    """The generated year of Danish transactions parses cleanly."""

    def test_everything_is_categorised(self):
        _table, _mapping, result = build("aar_2024.csv")
        self.assertGreater(len(result.transactions), 350)
        self.assertEqual(result.uncategorised, [])
        self.assertEqual(result.skipped_no_date, 0)
        self.assertEqual(result.skipped_no_amount, 0)

    def test_salary_beats_transfer(self):
        rules = RuleSet.defaults()
        self.assertEqual(rules.categorise("LØN OVERFØRSEL DANSK FIRMA A/S"), "Løn")
        self.assertEqual(rules.categorise("AFDRAG BILLÅN SANTANDER CONSUMER"),
                         "Lån og afdrag")


class MerchantTests(unittest.TestCase):
    """Grouping must ignore amounts, dates and receipt numbers."""

    def test_noise_is_stripped(self):
        self.assertEqual(
            merchant_name("Dankort-nota 4711 NETTO 8021 KØBENHAVN 21.05 kl. 17.42"),
            "NETTO KØBENHAVN")
        self.assertEqual(merchant_name("Visa kortkøb WOLT DANMARK 245,00 DKK"),
                         "WOLT DANMARK")
        self.assertEqual(merchant_name("FØTEX 1234 – Dankort-nota 998877"), "FØTEX")

    def test_card_type_is_never_the_rule(self):
        """A rule almost never makes sense based on which card was used."""
        cases = [
            ("Debitkort NETTO 8021 KØBENHAVN", "NETTO", "NETTO KØBENHAVN"),
            ("NETTO 8021 KØBENHAVN Debitkort-nota 4711", "NETTO",
             "NETTO KØBENHAVN"),
            ("DEBITKORT-NOTA WOLT DANMARK 245,00", "WOLT DANMARK",
             "WOLT DANMARK"),
            ("Visa/Dankort CIRCLE K AMAGER", "CIRCLE K AMAGER",
             "CIRCLE K AMAGER"),
            ("Kortkøbsnota 12345 FØTEX NØRREBRO", "FØTEX NØRREBRO",
             "FØTEX NØRREBRO"),
            ("Betaling m/ debitkort - MENY SLAGELSE", "MENY SLAGELSE",
             "MENY SLAGELSE"),
            ("Hævekort BILKA FIELDS", "BILKA FIELDS", "BILKA FIELDS"),
            ("Debitcard GUF KUGLER FORRETNING 21.05", "GUF KUGLER FORRETNING",
             "GUF KUGLER FORRETNING"),
            ("Kreditcard NOTA 4711 SPOTIFY AB", "SPOTIFY AB", "SPOTIFY AB"),
        ]
        for text, keyword, name in cases:
            self.assertEqual(rule_keyword(text), keyword, text)
            self.assertEqual(merchant_name(text), name, text)
            for card_word in ("kort", "card", "debit", "visa", "dankort",
                               "hæve", "kredit"):
                self.assertNotIn(card_word, fold(rule_keyword(text)), text)
                self.assertNotIn(card_word, fold(merchant_name(text)), text)

    def test_single_letter_shop_name_is_not_noise(self):
        """"K" in "Circle K" must survive - it is part of the actual name."""
        self.assertEqual(rule_keyword("Kortkøb 21.05 CIRCLE K AMAGER"),
                         "CIRCLE K AMAGER")

    def test_same_shop_same_group(self):
        first = "Dankort-nota 4711 NETTO 8021 KØBENHAVN 21.05 kl. 17.42"
        second = "Dankort-nota 9987 NETTO 8021 KØBENHAVN 03.06 kl. 09.11"
        self.assertEqual(merchant_key(first), merchant_key(second))

    def test_amount_in_the_text_is_ignored(self):
        self.assertEqual(merchant_key("MENY SLAGELSE 198,50"),
                         merchant_key("MENY SLAGELSE 1.204,00"))

    def test_keyword_is_a_substring_of_the_text(self):
        """Rules match as plain substrings, so the keyword must be one."""
        for text in ("JOE & THE JUICE FISKETORVET",
                     "BS BETALING TRYG FORSIKRING A/S POLICE 88123",
                     "Kortkøb 21.05 CIRCLE K AMAGER",
                     "PAYPAL *SPOTIFY 119,00"):
            keyword = rule_keyword(text)
            self.assertIn(keyword.lower(), text.lower(), text)

    def test_keyword_stops_at_a_shop_number(self):
        self.assertEqual(rule_keyword("NETTO 8021 KØBENHAVN"), "NETTO")
        self.assertEqual(rule_keyword("REMA 1000 VALBY"), "REMA")

    def test_keyword_matches_other_receipts_from_the_same_shop(self):
        rules = RuleSet([(rule_keyword("Dankort-nota 4711 NETTO 8021 KØBENHAVN"),
                          "Mad")])
        self.assertEqual(rules.categorise("Dankort-nota 5522 NETTO 1234 VALBY"),
                         "Mad")

    def test_grouping_in_the_result(self):
        table = csvsniff.Table.from_rows([
            ["Dato", "Tekst", "Beløb"],
            ["01-05-2025", "Dankort-nota 111 UKENDT BUTIK 4711 01.05", "-100,00"],
            ["02-05-2025", "Dankort-nota 222 UKENDT BUTIK 4711 02.05", "-50,00"],
            ["03-05-2025", "Dankort-nota 333 UKENDT BUTIK 4711 03.05", "-25,00"],
        ])
        mapping = detect_mapping(table)
        result = build_transactions(table, mapping, RuleSet.defaults())
        self.assertEqual(len(result.uncategorised), 1)
        name, count, total, keyword = result.uncategorised[0]
        self.assertEqual(name, "UKENDT BUTIK")
        self.assertEqual(count, 3)
        self.assertAlmostEqual(total, -175.0)
        self.assertEqual(keyword, "UKENDT BUTIK")

    def test_unreadable_text_falls_back_to_itself(self):
        self.assertEqual(merchant_name("1234567890"), "1234567890")
        self.assertTrue(rule_keyword("1234567890"))


class RuleTests(unittest.TestCase):

    def test_case_and_accents(self):
        rules = RuleSet.defaults()
        self.assertEqual(rules.categorise("NETTO 8021 KØBENHAVN"), "Dagligvarer")
        self.assertEqual(rules.categorise("cafe sværtegade"), "Restaurant")
        self.assertEqual(rules.categorise("Café Sværtegade"), "Restaurant")

    def test_longest_keyword_wins(self):
        rules = RuleSet([("k", "Kort"), ("circle k", "Transport")])
        self.assertEqual(rules.categorise("CIRCLE K AMAGER"), "Transport")

    def test_regex_rule(self):
        rules = RuleSet([("re:faktura\\s*\\d+", "Regninger")])
        self.assertEqual(rules.categorise("Faktura 12345"), "Regninger")
        self.assertEqual(rules.categorise("Faktura uden nummer"), DEFAULT_CATEGORY)

    def test_add_and_roundtrip(self):
        rules = RuleSet([("netto", "Dagligvarer")])
        rules.add("Slagter Hansen", "Dagligvarer")
        self.assertEqual(rules.categorise("SLAGTER HANSEN"), "Dagligvarer")
        copy = RuleSet.from_rows(rules.to_rows())
        self.assertEqual(copy.categorise("Slagter Hansen"), "Dagligvarer")

    def test_header_row_is_ignored(self):
        rules = RuleSet.from_rows([["Nøgleord", "Kategori"], ["netto", "Mad"]])
        self.assertEqual(len(rules), 1)

    def test_ignored_category_is_always_offered(self):
        """"Ignoreret" must be selectable even before anything uses it, so
        it is available up front in dropdowns and the categorise dialog."""
        self.assertIn(IGNORED_CATEGORY, RuleSet.defaults().categories())
        self.assertIn(IGNORED_CATEGORY, RuleSet([]).categories())
        self.assertIn(IGNORED_CATEGORY, default_categories())

    def test_ignored_category_not_duplicated(self):
        """A rule that already points to "Ignoreret" must not double up."""
        rules = RuleSet([("MobilePay Anders", IGNORED_CATEGORY)])
        names = rules.categories()
        self.assertEqual(names.count(IGNORED_CATEGORY), 1)

    def test_ignore_rule_works_like_any_other(self):
        rules = RuleSet([("Overførsel til opsparing", IGNORED_CATEGORY)])
        self.assertEqual(rules.categorise("Overførsel til opsparing 05-2025"),
                         IGNORED_CATEGORY)


if __name__ == "__main__":
    unittest.main()
