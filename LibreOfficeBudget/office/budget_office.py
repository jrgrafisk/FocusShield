# -*- coding: utf-8 -*-
"""Write the budget into a LibreOffice Calc document.

The layout follows the classic "Monthly Budget" template: a Summary sheet
with an Expenses table and an Income table (Planned / Actual / Diff.) and a
Transactions sheet where expenses live in columns B:E and income in G:J.
Everything is in Danish and in DKK, and the Actual columns are live
SUMIF/SUMIFS formulas, so editing a category on the Transactions sheet
updates the summary immediately.

This module needs to run inside LibreOffice (it imports ``uno``), but it
never talks to the user - dialogs live in ``budget_extension.py``.
"""

from __future__ import unicode_literals

import datetime

import uno

from budget_core.budget import KIND_EXPENSE, KIND_INCOME, Summary
from budget_core.merchant import merchant_key, merchant_name, rule_keyword
from budget_core.parsing import MONTH_NAMES_DA
from budget_core.plan import (GROUP_FIXED, GROUP_HELP, GROUP_LABELS,
                              GROUP_PERIODIC, GROUP_VARIABLE, build_plan)
from budget_core.rules import DEFAULT_CATEGORY
from budget_core.textutils import fold

# ---------------------------------------------------------------------------
# Design tokens taken from the template
# ---------------------------------------------------------------------------
NAVY = 0x334960
ORANGE = 0xF46524
TEXT_GREY = 0x576475
MUTED = 0x687887
DARK = 0x434343
PANEL = 0xEBEDEF
PEACH = 0xFFF2ED
WHITE = 0xFFFFFF
BAR_GREY = 0xAEB7C0
LIGHT_TEXT = 0xCCCCCC

FONT_TITLE = "Raleway"
FONT_BODY = "Lato"

BOLD = 150.0
NORMAL_WEIGHT = 100.0

SHEET_TX = "Transaktioner"
SHEET_RULES = "Kategorier"
SHEET_MONTHS = "Alle måneder"
SHEET_PLAN = "Budgetforslag"
SHEET_FORECAST = "Prognose"
SHEET_ACCOUNTS = "Konti"
SHEET_BANK = "Bankbudget"
SUMMARY_PREFIX = "Oversigt"

# Konti geometry
ACC_HEADER_ROW = 4
ACC_COL_NAME, ACC_COL_TYPE, ACC_COL_BALANCE = 1, 2, 3   # B, C, D (0 based)
ACC_TYPES = ("Lønkonto", "Budgetkonto", "Opsparingskonto", "Andet")

# Budgetforslag geometry
PLAN_HEADER_ROW = 5
PLAN_COL_NAME, PLAN_COL_MEAN, PLAN_COL_TARGET, PLAN_COL_DIFF = "B", "D", "H", "I"
PLAN_COL_ACCOUNT = "J"

# Prognose
FORECAST_MONTHS = 24
FORECAST_HEADER_ROW = 9

# Transactions sheet geometry (0 based columns)
TX_FIRST_ROW = 4           # row 5 in the UI
COL_EXP_DATE, COL_EXP_AMOUNT, COL_EXP_TEXT, COL_EXP_CAT = 1, 2, 3, 4
COL_INC_DATE, COL_INC_AMOUNT, COL_INC_TEXT, COL_INC_CAT = 6, 7, 8, 9
TX_MAX_ROW = 5000          # how far the summary formulas look

# Summary sheet geometry
SUM_FIRST_CAT_ROW = 27     # 1 based, as in the template
SUM_MIN_CATEGORY_ROWS = 12      # empty rows are room for your own categories
SUM_SPARE_CATEGORY_ROWS = 4

MAX_SHEET_NAME = 31


# ---------------------------------------------------------------------------
# Small UNO helpers
# ---------------------------------------------------------------------------

def _enum(type_name, value):
    return uno.Enum(type_name, value)


HORI = {
    "left": _enum("com.sun.star.table.CellHoriJustify", "LEFT"),
    "right": _enum("com.sun.star.table.CellHoriJustify", "RIGHT"),
    "center": _enum("com.sun.star.table.CellHoriJustify", "CENTER"),
    "standard": _enum("com.sun.star.table.CellHoriJustify", "STANDARD"),
}


def px_width(chars):
    """Excel column width (in characters) -> 1/100 mm."""
    return int(round((chars * 7 + 5) * 26.46))


class Formats(object):
    """Number formats, created once per document."""

    def __init__(self, doc):
        self.doc = doc
        self._cache = {}
        self.locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        self.locale.Language = "da"
        self.locale.Country = "DK"
        self.fallback = uno.createUnoStruct("com.sun.star.lang.Locale")
        self.fallback.Language = "en"
        self.fallback.Country = "US"

        self.currency = self.get('#.##0 "kr.";-#.##0 "kr."',
                                 '#,##0 "kr.";-#,##0 "kr."')
        self.currency2 = self.get('#.##0,00 "kr.";-#.##0,00 "kr."',
                                  '#,##0.00 "kr.";-#,##0.00 "kr."')
        self.signed = self.get('+#.##0 "kr.";-#.##0 "kr.";0 "kr."',
                               '+#,##0 "kr.";-#,##0 "kr.";0 "kr."')
        self.percent = self.get("+0%;-0%;0%", "+0%;-0%;0%")
        # The date format comes from the locale itself - format code keywords
        # are not translated the same way in every LibreOffice build.
        self.date = self.get("DD-MM-YYYY") or self.standard(2)
        self.month = self.get("MMMM YYYY")

    def standard(self, format_type):
        """The locale's own format for dates, times, currency, ..."""
        try:
            return self.doc.getNumberFormats().getStandardFormat(format_type,
                                                                 self.locale)
        except Exception:
            return 0

    def get(self, code, fallback_code=None):
        """Return a format key for ``code`` (Danish format code)."""
        if code in self._cache:
            return self._cache[code]
        formats = self.doc.getNumberFormats()
        key = -1
        for text, locale in ((code, self.locale),
                             (fallback_code or code, self.fallback)):
            try:
                key = formats.queryKey(text, locale, False)
                if key == -1:
                    key = formats.addNew(text, locale)
                if key != -1:
                    break
            except Exception:
                key = -1
        if key == -1:
            key = 0
        self._cache[code] = key
        return key


class SheetPen(object):
    """Tiny wrapper that makes the sheet writing code readable."""

    def __init__(self, sheet, formats, null_date):
        self.sheet = sheet
        self.formats = formats
        self.null_date = null_date

    # -- addressing -------------------------------------------------------
    def cell(self, ref):
        return self.sheet.getCellRangeByName(ref).getCellByPosition(0, 0)

    def rng(self, ref):
        return self.sheet.getCellRangeByName(ref)

    # -- writing ----------------------------------------------------------
    def text(self, ref, value, **style):
        cell = self.cell(ref)
        cell.setString(value if value is not None else "")
        self.style(ref, **style)
        return cell

    def number(self, ref, value, **style):
        cell = self.cell(ref)
        cell.setValue(float(value))
        self.style(ref, **style)
        return cell

    def formula(self, ref, value, **style):
        cell = self.cell(ref)
        cell.setFormula(value)
        self.style(ref, **style)
        return cell

    def date(self, ref, value, **style):
        style.setdefault("fmt", self.formats.date)
        return self.number(ref, (value - self.null_date).days, **style)

    def merge(self, ref):
        self.rng(ref).merge(True)

    def note(self, ref, text):
        """A cell comment shown on hover (not pinned open).

        ``cell.Annotation`` only returns a live, persisted annotation for a
        cell that already has one - setting text through a freshly fetched
        reference on a cell without one silently does nothing.  Creating it
        through ``sheet.Annotations.insertNew()`` is what actually sticks.
        """
        if not text:
            return
        try:
            cell = self.cell(ref)
            self.sheet.Annotations.insertNew(cell.CellAddress, text)
            cell.Annotation.IsVisible = False
        except Exception:
            pass

    # -- styling ----------------------------------------------------------
    def style(self, ref, bg=None, color=None, size=None, bold=None,
              italic=None, font=None, align=None, valign=None, wrap=None,
              fmt=None, indent=None):
        target = self.rng(ref)
        if bg is not None:
            target.CellBackColor = bg
        if color is not None:
            target.CharColor = color
        if size is not None:
            target.CharHeight = float(size)
        if bold is not None:
            target.CharWeight = BOLD if bold else NORMAL_WEIGHT
        if italic is not None:
            target.CharPosture = _enum("com.sun.star.awt.FontSlant",
                                       "ITALIC" if italic else "NONE")
        if font is not None:
            target.CharFontName = font
        if align is not None:
            target.HoriJustify = HORI[align]
        if valign is not None:
            try:
                target.VertJustify = _enum("com.sun.star.table.CellVertJustify",
                                           valign.upper())
            except Exception:
                pass
        if wrap is not None:
            target.IsTextWrapped = bool(wrap)
        if fmt is not None:
            target.NumberFormat = fmt
        if indent is not None:
            try:
                target.ParaIndent = int(indent)
            except Exception:
                pass
        return target

    def column_widths(self, widths):
        """``widths`` maps a 0 based column index to an Excel character width."""
        columns = self.sheet.Columns
        for index, chars in widths.items():
            try:
                columns.getByIndex(index).Width = px_width(chars)
            except Exception:
                pass

    def row_height(self, row, points):
        try:
            self.sheet.Rows.getByIndex(row - 1).Height = int(points * 35.28)
        except Exception:
            pass


def sheet_null_date(doc):
    try:
        settings = doc.getPropertyValue("NullDate")
        return datetime.date(settings.Year, settings.Month, settings.Day)
    except Exception:
        return datetime.date(1899, 12, 30)


def month_title(month_key):
    """``"2025-05"`` -> ``"Maj 2025"``."""
    try:
        year, month = month_key.split("-")
        name = MONTH_NAMES_DA[int(month) - 1]
        return "%s%s %s" % (name[0].upper(), name[1:], year)
    except Exception:
        return month_key


def month_bounds(month_key):
    year, month = (int(part) for part in month_key.split("-"))
    first = datetime.date(year, month, 1)
    if month == 12:
        last = datetime.date(year, 12, 31)
    else:
        last = datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)
    return first, last


# ---------------------------------------------------------------------------
# The workbook builder
# ---------------------------------------------------------------------------

class BudgetWorkbook(object):
    """Builds (and rebuilds) the whole document."""

    def __init__(self, doc):
        self.doc = doc
        self.formats = Formats(doc)
        self.null_date = sheet_null_date(doc)

    # -- infrastructure ---------------------------------------------------
    def serial(self, value):
        return (value - self.null_date).days

    def sheet(self, name, index=None):
        """Return an empty sheet called ``name`` (created if necessary)."""
        sheets = self.doc.Sheets
        if sheets.hasByName(name):
            sheets.removeByName(name)
        sheets.insertNewByName(name, sheets.Count if index is None else index)
        return sheets.getByName(name)

    def pen(self, sheet):
        return SheetPen(sheet, self.formats, self.null_date)

    # -- public API -------------------------------------------------------
    def build(self, result, ruleset, start_balance=None, suggest_planned=True):
        """Create every sheet from a :class:`BuildResult`."""
        transactions = result.transactions
        summary = Summary(transactions)
        months = summary.months or [""]

        self.doc.lockControllers()
        try:
            self._reset_sheets()
            tx_sheet = self.sheet(SHEET_TX)
            self.write_transactions(tx_sheet, transactions)
            self.write_rules(self.sheet(SHEET_RULES), ruleset, summary)

            self._ensure_accounts_sheet()
            category_accounts = self.read_category_accounts()
            plan = build_plan(summary, transactions) if len(summary.months) > 1 else None
            if plan is not None:
                plan_rows = self.write_plan(self.sheet(SHEET_PLAN), plan,
                                            accounts=category_accounts)
                accounts_total = self._accounts_total()
                fallback = (start_balance or 0.0) + summary.net
                self.write_forecast(self.sheet(SHEET_FORECAST), plan, plan_rows,
                                    accounts_total if accounts_total else fallback)

            planned = self._planned_values(summary) if suggest_planned else {}
            previous = None
            for index, month in enumerate(months):
                name = self.summary_name(months, month)
                sheet = self.sheet(name, index)
                self.write_summary(sheet, summary, month,
                                   multi_month=len(months) > 1,
                                   planned=planned,
                                   start_balance=start_balance if index == 0 else None,
                                   previous_sheet=previous)
                previous = name
            if len(months) > 1:
                self.write_months(self.sheet(SHEET_MONTHS), summary)
            self._order_sheets(months)
        finally:
            self.doc.unlockControllers()
        self.doc.calculateAll()
        self._activate(SHEET_PLAN if self.doc.Sheets.hasByName(SHEET_PLAN)
                       else self.summary_name(months, months[0]))
        return summary

    def summary_name(self, months, month):
        if len(months) <= 1 or not month:
            return SUMMARY_PREFIX
        return ("%s %s" % (SUMMARY_PREFIX, month_title(month)))[:MAX_SHEET_NAME]

    def _reset_sheets(self):
        sheets = self.doc.Sheets
        keep = sheets.getByIndex(0).Name
        for index in range(sheets.Count - 1, 0, -1):
            sheets.removeByName(sheets.getByIndex(index).Name)
        sheets.getByIndex(0).Name = "__tmp__" if keep != "__tmp__" else "__tmp2__"

    def _order_sheets(self, months):
        """Summaries first, then transactions, months and rules."""
        order = [SHEET_PLAN, SHEET_FORECAST]
        order += [self.summary_name(months, m) for m in months]
        order += [SHEET_TX, SHEET_MONTHS, SHEET_RULES, SHEET_ACCOUNTS, SHEET_BANK]
        sheets = self.doc.Sheets
        position = 0
        for name in order:
            if sheets.hasByName(name):
                sheets.moveByName(name, position)
                position += 1
        # Drop the placeholder sheet created with the document.
        for index in range(sheets.Count - 1, -1, -1):
            name = sheets.getByIndex(index).Name
            if name.startswith("__tmp"):
                sheets.removeByName(name)

    def _activate(self, name):
        try:
            controller = self.doc.CurrentController
            controller.setActiveSheet(self.doc.Sheets.getByName(name))
            controller.ShowGrid = False
        except Exception:
            pass

    def _planned_values(self, summary):
        """Suggested budget per category: the average across the months."""
        planned = {}
        months = max(1, len(summary.months))
        if len(summary.months) < 2:
            return planned
        for kind, categories in ((KIND_EXPENSE, summary.expense_categories),
                                 (KIND_INCOME, summary.income_categories)):
            for category in categories:
                total = abs(summary.category_total(kind, category))
                planned[(kind, category)] = round(total / months / 10.0) * 10
        return planned

    # -- Transactions sheet ----------------------------------------------
    def write_transactions(self, sheet, transactions):
        pen = self.pen(sheet)
        pen.column_widths({0: 4.0, 1: 11.0, 2: 15.0, 3: 30.0, 4: 16.0,
                           5: 4.0, 6: 11.0, 7: 15.0, 8: 30.0, 9: 16.0,
                           10: 4.0})

        pen.merge("B1:J1")
        pen.text("B1", "Skift eller tilføj kategorier i tabellerne på oversigtsarket "
                       "- og i arket \"Kategorier\", hvis de skal sættes automatisk.",
                 font=FONT_BODY, size=10, color=LIGHT_TEXT, bg=NAVY, italic=True,
                 align="left", valign="center")
        pen.row_height(1, 22)

        pen.text("B2", "Udgifter", font=FONT_TITLE, size=18, bold=True, color=ORANGE)
        pen.text("G2", "Indtægter", font=FONT_TITLE, size=18, bold=True, color=ORANGE)
        pen.row_height(2, 26)

        headers = ("Dato", "Beløb", "Beskrivelse", "Kategori")
        for column, title in zip("BCDE", headers):
            pen.text("%s4" % column, title, font=FONT_BODY, size=11, bold=True,
                     color=NAVY)
        for column, title in zip("GHIJ", headers):
            pen.text("%s4" % column, title, font=FONT_BODY, size=11, bold=True,
                     color=NAVY)

        expenses = [t for t in transactions if t.amount < 0]
        income = [t for t in transactions if t.amount >= 0]
        self._write_block(sheet, expenses, COL_EXP_DATE)
        self._write_block(sheet, income, COL_INC_DATE)

        rows = max(len(expenses), len(income), 1)
        last = TX_FIRST_ROW + rows
        for first_col in (COL_EXP_DATE, COL_INC_DATE):
            date_ref = _ref(first_col, TX_FIRST_ROW + 1, first_col, last)
            amount_ref = _ref(first_col + 1, TX_FIRST_ROW + 1, first_col + 1, last)
            text_ref = _ref(first_col + 2, TX_FIRST_ROW + 1, first_col + 3, last)
            pen.style(date_ref, fmt=self.formats.date, font=FONT_BODY, size=10,
                      color=MUTED, align="left")
            pen.style(amount_ref, fmt=self.formats.currency2, font=FONT_BODY,
                      size=10, color=TEXT_GREY, bold=True, align="left")
            pen.style(text_ref, font=FONT_BODY, size=10, color=TEXT_GREY,
                      align="left")

        self._add_category_dropdown(sheet, COL_EXP_CAT, last)
        self._add_category_dropdown(sheet, COL_INC_CAT, last)
        self._print_setup(sheet, COL_INC_CAT, last)
        self._freeze(sheet, 0, TX_FIRST_ROW)

    def _write_block(self, sheet, transactions, first_col):
        if not transactions:
            return
        data = []
        for t in transactions:
            data.append((float(self.serial(t.date)), abs(t.amount), t.text,
                         t.category))
        target = sheet.getCellRangeByPosition(
            first_col, TX_FIRST_ROW, first_col + 3, TX_FIRST_ROW + len(data) - 1)
        target.setDataArray(tuple(data))

    def _add_category_dropdown(self, sheet, column, last_row):
        """Cell validity that lists the categories from the rules sheet."""
        self._add_list_dropdown(sheet, column, TX_FIRST_ROW + 1, last_row,
                                "$%s.$E$5:$E$80" % SHEET_RULES)

    def _add_account_dropdown(self, sheet, column, first_row, last_row):
        """Cell validity that lists the account names from the Konti sheet."""
        self._add_list_dropdown(sheet, column, first_row, last_row,
                                "$%s.$B$%d:$B$50" % (SHEET_ACCOUNTS,
                                                     ACC_HEADER_ROW + 1))

    def _add_list_dropdown(self, sheet, column, first_row, last_row, source):
        try:
            ref = _ref(column, first_row, column, max(last_row, first_row))
            target = sheet.getCellRangeByName(ref)
            validation = target.Validation
            validation.Type = _enum("com.sun.star.sheet.ValidationType", "LIST")
            validation.ShowErrorMessage = False
            validation.ShowList = 1
            validation.setFormula1(source)
            target.Validation = validation
        except Exception:
            pass

    def _print_setup(self, sheet, last_column, last_row):
        """Fit the sheet on one page width when printing or exporting."""
        try:
            area = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
            area.Sheet = sheet.RangeAddress.Sheet
            area.StartColumn, area.EndColumn = 1, last_column
            area.StartRow, area.EndRow = 0, last_row
            sheet.setPrintAreas((area,))
        except Exception:
            pass
        try:
            style = self.doc.StyleFamilies.getByName("PageStyles").getByName(
                sheet.PageStyle)
            style.ScaleToPagesX = 1
        except Exception:
            pass

    def _freeze(self, sheet, column, row):
        try:
            controller = self.doc.CurrentController
            controller.setActiveSheet(sheet)
            controller.freezeAtPosition(column, row)
        except Exception:
            pass

    # -- Summary sheet ----------------------------------------------------
    def write_summary(self, sheet, summary, month, multi_month, planned,
                      start_balance=None, previous_sheet=None):
        pen = self.pen(sheet)
        widths = {0: 4.0, 12: 4.0}
        for index in range(1, 12):
            widths[index] = 11.5
        widths[13] = 8.0
        widths[14] = 12.0
        pen.column_widths(widths)

        expense_rows = self._category_rows(summary, KIND_EXPENSE, month)
        income_rows = self._category_rows(summary, KIND_INCOME, month)
        row_count = max(len(expense_rows) + SUM_SPARE_CATEGORY_ROWS,
                        len(income_rows) + SUM_SPARE_CATEGORY_ROWS,
                        SUM_MIN_CATEGORY_ROWS)
        first = SUM_FIRST_CAT_ROW
        last = first + row_count - 1

        self._summary_intro(pen, month, multi_month)
        self._summary_period(pen, month, multi_month)
        self._summary_balance(pen, start_balance, previous_sheet)
        self._summary_totals(pen, first, last)
        self._summary_table(pen, "expense", first, last, expense_rows, planned,
                            month, multi_month)
        self._summary_table(pen, "income", first, last, income_rows, planned,
                            month, multi_month)
        self._print_setup(sheet, 11, last + 1)

    def _category_rows(self, summary, kind, month):
        categories = (summary.expense_categories if kind == KIND_EXPENSE
                      else summary.income_categories)
        if not month:
            return list(categories)
        used = [c for c in categories if abs(summary.value(kind, c, month)) > 0.0001]
        return used or list(categories)

    def _summary_intro(self, pen, month, multi_month):
        pen.merge("B2:H2")
        pen.text("B2", "SÅDAN KOMMER DU I GANG", bg=NAVY, color=WHITE, bold=True,
                 size=9, font=FONT_BODY, align="left", valign="center")
        pen.merge("B3:G4")
        pen.text("B3", "Skriv din startsaldo i celle L8 og tilpas kategorier og "
                       "budgetterede beløb i tabellerne \"Udgifter\" og \"Indtægter\" "
                       "nedenfor.", bg=NAVY, color=LIGHT_TEXT, size=10,
                 font=FONT_BODY, align="left", valign="center", wrap=True)
        pen.merge("B5:G6")
        pen.text("B5", "Når du retter en kategori i arket \"Transaktioner\", "
                       "opdateres denne oversigt automatisk.", bg=NAVY,
                 color=LIGHT_TEXT, size=10, font=FONT_BODY, align="left",
                 valign="center", wrap=True)

        pen.merge("I2:L2")
        pen.text("I2", "BEMÆRK", bg=NAVY, color=WHITE, bold=True, size=9,
                 font=FONT_BODY, align="left", valign="center")
        pen.merge("I3:L3")
        pen.text("I3", "Redigér kun de fremhævede felter.", bg=PEACH, color=NAVY,
                 italic=True, size=10, font=FONT_BODY, align="left", valign="center")
        pen.merge("I4:M5")
        pen.text("I4", "Undgå at ændre celler med formler.", bg=NAVY,
                 color=LIGHT_TEXT, italic=True, size=10, font=FONT_BODY,
                 align="left", valign="center")

        title = "Månedsbudget"
        if month:
            title = "Budget - %s" % month_title(month)
        pen.merge("B8:F9")
        pen.text("B8", title, font=FONT_TITLE, size=25, bold=True, color=ORANGE,
                 bg=WHITE, align="left", valign="center")
        pen.row_height(8, 26)
        pen.row_height(9, 22)

    def _summary_period(self, pen, month, multi_month):
        """Period filter used by the SUMIFS formulas (visible, editable)."""
        if not month:
            return
        first, last = month_bounds(month)
        pen.text("N2", "Perioden", font=FONT_BODY, size=9, bold=True, color=MUTED,
                 align="left")
        pen.text("N3", "Fra", font=FONT_BODY, size=9, color=MUTED, align="left")
        pen.date("O3", first, font=FONT_BODY, size=9, color=MUTED, align="left")
        pen.text("N4", "Til", font=FONT_BODY, size=9, color=MUTED, align="left")
        pen.date("O4", last, font=FONT_BODY, size=9, color=MUTED, align="left")

    def _summary_balance(self, pen, start_balance, previous_sheet):
        pen.merge("J8:K8")
        pen.text("J8", "Startsaldo: ", font=FONT_BODY, size=10, bold=True,
                 color=NAVY, align="right", valign="center")
        if previous_sheet:
            pen.formula("L8", "=$'%s'.E17" % previous_sheet, bg=PEACH,
                        color=NAVY, font=FONT_BODY, fmt=self.formats.currency)
        else:
            pen.number("L8", start_balance or 0.0, bg=PEACH, color=NAVY,
                       font=FONT_BODY, fmt=self.formats.currency)

        pen.merge("I13:K13")
        pen.formula("I13", "=IFERROR(E17/D17-1;\"\")", bg=PANEL, color=NAVY,
                    size=24, font=FONT_BODY, align="center", valign="center",
                    fmt=self.formats.percent)
        pen.merge("I14:K14")
        pen.formula("I14", "=IF(I13<0;\"Mindre opsparing\";\"Mere opsparing\")",
                    bg=PANEL, color=TEXT_GREY, italic=True, size=10,
                    font=FONT_BODY, align="center", valign="center")
        pen.merge("I15:K15")
        pen.formula("I15", "=IFERROR(E17-D17;0)", bg=PANEL, color=NAVY, size=24,
                    font=FONT_BODY, align="center", valign="center",
                    fmt=self.formats.currency)
        pen.merge("I16:K16")
        pen.formula("I16", "=IF(I15<0;\"Brugt denne måned\";\"Sparet denne måned\")",
                    bg=PANEL, color=TEXT_GREY, italic=True, size=10,
                    font=FONT_BODY, align="center", valign="center")

        pen.text("D16", "STARTSALDO", font=FONT_BODY, size=14, bold=True,
                 color=NAVY, align="right")
        pen.text("E16", " SLUTSALDO", font=FONT_BODY, size=14, bold=True,
                 color=ORANGE, align="left")
        pen.formula("D17", "=IF(ISBLANK(L8);0;L8)", font=FONT_BODY, size=10,
                    italic=True, color=TEXT_GREY, align="center",
                    fmt=self.formats.currency)
        pen.formula("E17", "=D17+(I22-C22)", font=FONT_BODY, size=10, italic=True,
                    color=ORANGE, align="center", fmt=self.formats.currency)

    def _summary_totals(self, pen, first, last):
        pen.merge("B20:F20")
        pen.text("B20", "Udgifter", font=FONT_BODY, size=14, bold=True, color=NAVY,
                 align="left")
        pen.merge("H20:L20")
        pen.text("H20", "Indtægter", font=FONT_BODY, size=14, bold=True, color=NAVY,
                 align="left")

        for label_ref, value_ref, bar_ref, source, colour, bold in (
                ("B21", "C21", "D21:F21", "D26", BAR_GREY, False),
                ("B22", "C22", "D22:F22", "E26", NAVY, True)):
            pen.text(label_ref, "Budget" if source == "D26" else "Faktisk",
                     font=FONT_BODY, size=10, bold=True,
                     color=TEXT_GREY if not bold else NAVY, align="left")
            pen.formula(value_ref, "=%s" % source, font=FONT_BODY, size=10,
                        color=TEXT_GREY if not bold else NAVY, align="right",
                        fmt=self.formats.currency)
            pen.merge(bar_ref)
            pen.formula(bar_ref.split(":")[0],
                        '=IFERROR(REPT("▉";ROUND(%s/MAX($C$21:$C$22)*18;0));"")'
                        % value_ref, color=colour, size=10, font=FONT_BODY,
                        align="left", valign="center")

        for label_ref, value_ref, bar_ref, source, colour, bold in (
                ("H21", "I21", "J21:L21", "J26", BAR_GREY, False),
                ("H22", "I22", "J22:L22", "K26", NAVY, True)):
            pen.text(label_ref, "Budget" if source == "J26" else "Faktisk",
                     font=FONT_BODY, size=10, bold=True,
                     color=TEXT_GREY if not bold else NAVY, align="left")
            pen.formula(value_ref, "=%s" % source, font=FONT_BODY, size=10,
                        color=TEXT_GREY if not bold else NAVY, align="right",
                        fmt=self.formats.currency)
            pen.merge(bar_ref)
            pen.formula(bar_ref.split(":")[0],
                        '=IFERROR(REPT("▉";ROUND(%s/MAX($I$21:$I$22)*18;0));"")'
                        % value_ref, color=colour, size=10, font=FONT_BODY,
                        align="left", valign="center")

    def _summary_table(self, pen, kind, first, last, categories, planned, month,
                       multi_month):
        """Write one of the two category tables."""
        expense = kind == "expense"
        name_col = "B" if expense else "H"
        merge_to = "C" if expense else "I"
        plan_col = "D" if expense else "J"
        actual_col = "E" if expense else "K"
        diff_col = "F" if expense else "L"
        heading_row = 24
        header_row = 25
        totals_row = 26

        pen.merge("%s%d:%s%d" % (name_col, heading_row, merge_to, heading_row))
        pen.text("%s%d" % (name_col, heading_row),
                 "Udgifter" if expense else "Indtægter", font=FONT_TITLE,
                 size=18, bold=True, color=ORANGE, align="left")

        for column, title in ((plan_col, "Budget"), (actual_col, "Faktisk"),
                              (diff_col, "Diff.")):
            pen.text("%s%d" % (column, header_row), title, font=FONT_BODY,
                     size=11, bold=True, color=NAVY, align="right")

        pen.merge("%s%d:%s%d" % (name_col, totals_row, merge_to, totals_row))
        pen.text("%s%d" % (name_col, totals_row), "I alt", font=FONT_BODY, size=9,
                 italic=True, color=MUTED, align="left")
        for column in (plan_col, actual_col):
            pen.formula("%s%d" % (column, totals_row),
                        "=SUM(%s%d:%s%d)" % (column, first, column, last),
                        font=FONT_BODY, size=9, italic=True, color=MUTED,
                        align="right", fmt=self.formats.currency)
        pen.formula("%s%d" % (diff_col, totals_row),
                    "=SUM(%s%d:%s%d)" % (diff_col, first, diff_col, last),
                    font=FONT_BODY, size=9, italic=True, color=MUTED,
                    align="right", fmt=self.formats.signed)

        kind_key = KIND_EXPENSE if expense else KIND_INCOME
        for offset in range(last - first + 1):
            row = first + offset
            category = categories[offset] if offset < len(categories) else ""
            pen.merge("%s%d:%s%d" % (name_col, row, merge_to, row))
            pen.text("%s%d" % (name_col, row), category, font=FONT_BODY, size=10,
                     bold=True, color=DARK, align="left")
            plan_ref = "%s%d" % (plan_col, row)
            if category:
                pen.number(plan_ref, planned.get((kind_key, category), 0.0),
                           font=FONT_BODY, size=10, color=DARK, align="right",
                           bg=PEACH, fmt=self.formats.currency)
            else:
                pen.text(plan_ref, "", font=FONT_BODY, size=10, color=DARK,
                         align="right", bg=PEACH, fmt=self.formats.currency)
            pen.formula("%s%d" % (actual_col, row),
                        self._actual_formula(name_col, row, expense, month,
                                             multi_month),
                        font=FONT_BODY, size=10, color=DARK, align="right",
                        fmt=self.formats.currency)
            if expense:
                diff = '=IF(ISBLANK($B%d);"";D%d-E%d)' % (row, row, row)
            else:
                diff = '=IF(ISBLANK($H%d);"";K%d-J%d)' % (row, row, row)
            pen.formula("%s%d" % (diff_col, row), diff, font=FONT_BODY, size=10,
                        color=MUTED, align="right", fmt=self.formats.signed)

    def _actual_formula(self, name_col, row, expense, month, multi_month):
        """SUMIF over the transaction sheet, filtered by month when needed."""
        if expense:
            amount = "$%s.$C$%d:$C$%d" % (SHEET_TX, TX_FIRST_ROW + 1, TX_MAX_ROW)
            category = "$%s.$E$%d:$E$%d" % (SHEET_TX, TX_FIRST_ROW + 1, TX_MAX_ROW)
            date = "$%s.$B$%d:$B$%d" % (SHEET_TX, TX_FIRST_ROW + 1, TX_MAX_ROW)
        else:
            amount = "$%s.$H$%d:$H$%d" % (SHEET_TX, TX_FIRST_ROW + 1, TX_MAX_ROW)
            category = "$%s.$J$%d:$J$%d" % (SHEET_TX, TX_FIRST_ROW + 1, TX_MAX_ROW)
            date = "$%s.$G$%d:$G$%d" % (SHEET_TX, TX_FIRST_ROW + 1, TX_MAX_ROW)
        criterion = "$%s%d" % (name_col, row)
        if month and multi_month:
            body = ('SUMIFS(%s;%s;%s;%s;">="&$O$3;%s;"<="&$O$4)'
                    % (amount, category, criterion, date, date))
        else:
            body = "SUMIF(%s;%s;%s)" % (category, criterion, amount)
        return '=IF(ISBLANK(%s);"";%s)' % (criterion, body)

    # -- Rules sheet ------------------------------------------------------
    def write_rules(self, sheet, ruleset, summary=None, accounts=None):
        pen = self.pen(sheet)
        pen.column_widths({0: 5.13, 1: 24.0, 2: 20.0, 3: 5.13, 4: 20.0, 5: 16.0})
        accounts = accounts or {}

        pen.merge("B2:F2")
        pen.text("B2", "Kategoriregler", font=FONT_TITLE, size=18, bold=True,
                 color=ORANGE, align="left")
        pen.merge("B3:F3")
        pen.text("B3", "Et nøgleord matcher, når det indgår i teksten på posteringen "
                       "(uden hensyn til store/små bogstaver). Vælg \"Budget ▸ "
                       "Opdatér kategorier og budget\" for at bruge ændringerne. "
                       "Skriv \"re:\" foran for et regulært udtryk. Konto (kolonne "
                       "F) sætter, hvilken konto der som udgangspunkt betaler for "
                       "kategorien - se arket \"%s\".""" % SHEET_ACCOUNTS,
                 font=FONT_BODY, size=9, italic=True, color=MUTED, align="left",
                 wrap=True, valign="center")
        pen.row_height(3, 34)

        pen.text("B4", "Nøgleord", font=FONT_BODY, size=11, bold=True, color=NAVY)
        pen.text("C4", "Kategori", font=FONT_BODY, size=11, bold=True, color=NAVY)
        pen.text("E4", "Kategoriliste", font=FONT_BODY, size=11, bold=True,
                 color=NAVY)
        pen.text("F4", "Konto", font=FONT_BODY, size=11, bold=True, color=NAVY)

        rows = ruleset.to_rows()
        if rows:
            target = sheet.getCellRangeByPosition(1, 4, 2, 4 + len(rows) - 1)
            target.setDataArray(tuple(tuple(row) for row in rows))
            pen.style(_ref(1, 5, 2, 4 + len(rows)), font=FONT_BODY, size=10,
                      color=DARK, align="left")

        categories = list(ruleset.categories())
        if summary is not None:
            for extra in summary.expense_categories + summary.income_categories:
                if extra not in categories:
                    categories.append(extra)
        if DEFAULT_CATEGORY not in categories:
            categories.append(DEFAULT_CATEGORY)
        last_row = 4 + len(categories) - 1
        target = sheet.getCellRangeByPosition(4, 4, 4, last_row)
        target.setDataArray(tuple((name,) for name in categories))
        pen.style(_ref(4, 5, 4, last_row + 1), font=FONT_BODY, size=10,
                  color=DARK, align="left")
        account_values = tuple((accounts.get(name, ""),) for name in categories)
        sheet.getCellRangeByPosition(5, 4, 5, last_row).setDataArray(account_values)
        pen.style(_ref(5, 5, 5, last_row + 1), font=FONT_BODY, size=10, color=DARK,
                  align="left")
        self._add_account_dropdown(sheet, 5, 5, last_row + 1)
        self._freeze(sheet, 0, 4)

    # -- Konti --------------------------------------------------------------
    def _ensure_accounts_sheet(self):
        """Create the Konti sheet with a starter template, unless it exists."""
        if self.doc.Sheets.hasByName(SHEET_ACCOUNTS):
            return
        self.write_accounts(self.sheet(SHEET_ACCOUNTS),
                            [("Lønkonto", "Lønkonto", 0.0),
                             ("Budgetkonto", "Budgetkonto", 0.0),
                             ("Opsparingskonto", "Opsparingskonto", 0.0)])

    def write_accounts(self, sheet, rows):
        """Write the whole Konti sheet. ``rows``: [(navn, type, saldo), ...]."""
        pen = self.pen(sheet)
        pen.column_widths({0: 4.0, 1: 22.0, 2: 18.0, 3: 16.0})

        pen.merge("B2:D2")
        pen.text("B2", "Konti", font=FONT_TITLE, size=18, bold=True, color=ORANGE,
                 align="left")
        pen.merge("B3:D3")
        pen.text("B3", "Nuværende saldo. Prognosen bruger summen som "
                       "startsaldo. Sæt en konto pr. kategori i \"%s\" eller "
                       "\"%s\" for at se, hvordan hver konto udvikler sig."
                       % (SHEET_RULES, SHEET_PLAN),
                 font=FONT_BODY, size=9, italic=True, color=MUTED, align="left",
                 wrap=True, valign="center")
        pen.row_height(3, 46)

        row = ACC_HEADER_ROW
        pen.text("B%d" % row, "Konto", font=FONT_BODY, size=11, bold=True,
                 color=NAVY, align="left")
        pen.text("C%d" % row, "Type", font=FONT_BODY, size=11, bold=True,
                 color=NAVY, align="left")
        pen.text("D%d" % row, "Saldo", font=FONT_BODY, size=11, bold=True,
                 color=NAVY, align="right")

        first = row + 1
        rows = list(rows) or [("", "", 0.0)]
        for offset, (name, kind, balance) in enumerate(rows):
            current = first + offset
            pen.text("B%d" % current, name, font=FONT_BODY, size=10, bold=True,
                     color=DARK, align="left")
            pen.text("C%d" % current, kind, font=FONT_BODY, size=10,
                     color=TEXT_GREY, align="left")
            pen.number("D%d" % current, balance, font=FONT_BODY, size=10,
                       color=DARK, align="right", bg=PEACH,
                       fmt=self.formats.currency)

        last = first + len(rows) - 1
        total_row = last + 2
        pen.text("B%d" % total_row, "I alt", font=FONT_BODY, size=10, bold=True,
                 color=NAVY, align="left")
        pen.formula("D%d" % total_row,
                    "=SUM(D%d:D%d)" % (first, last), font=FONT_BODY, size=10,
                    bold=True, color=NAVY, align="right", fmt=self.formats.currency)

        self._print_setup(sheet, 3, total_row + 1)
        self._freeze(sheet, 0, row)

    def _scan_accounts(self):
        """[(navn, type, saldo, ui_row)] from the Konti sheet, or []."""
        if not self.doc.Sheets.hasByName(SHEET_ACCOUNTS):
            return []
        sheet = self.doc.Sheets.getByName(SHEET_ACCOUNTS)
        last_row = used_row_count(sheet)
        accounts = []
        for row in range(ACC_HEADER_ROW, last_row + 1):
            name = sheet.getCellByPosition(ACC_COL_NAME, row).getString().strip()
            if not name or fold(name) == "i alt":
                continue
            kind = sheet.getCellByPosition(ACC_COL_TYPE, row).getString().strip()
            balance = sheet.getCellByPosition(ACC_COL_BALANCE, row).getValue()
            accounts.append((name, kind, balance, row + 1))
        return accounts

    def read_accounts(self):
        """[(navn, type, saldo)] from the Konti sheet, or []."""
        return [(name, kind, balance)
                for name, kind, balance, _row in self._scan_accounts()]

    def _accounts_total(self):
        return sum(balance for _name, _kind, balance in self.read_accounts())

    def read_category_accounts(self):
        """{kategori: konto} from column F of the Kategorier sheet."""
        accounts = {}
        if not self.doc.Sheets.hasByName(SHEET_RULES):
            return accounts
        sheet = self.doc.Sheets.getByName(SHEET_RULES)
        last_row = used_row_count(sheet)
        for row in range(4, last_row + 1):
            category = sheet.getCellByPosition(4, row).getString().strip()
            account = sheet.getCellByPosition(5, row).getString().strip()
            if category and account:
                accounts[category] = account
        return accounts

    def read_plan_accounts(self):
        """{(kind, kategori): konto} read back from the Budgetforslag sheet."""
        accounts = {}
        if not self.doc.Sheets.hasByName(SHEET_PLAN):
            return accounts
        sheet = self.doc.Sheets.getByName(SHEET_PLAN)
        last_row = used_row_count(sheet)
        kind = KIND_INCOME
        for row in range(PLAN_HEADER_ROW, last_row + 1):
            name = sheet.getCellByPosition(1, row).getString().strip()
            account = sheet.getCellByPosition(9, row).getString().strip()
            if not name:
                continue
            upper = name.upper()
            if upper.startswith("INDTÆGT"):
                kind = KIND_INCOME
                continue
            if "UDGIFTER" in upper:
                kind = KIND_EXPENSE
                continue
            if account and not name.lower().startswith("i alt"):
                accounts[(kind, name)] = account
        return accounts

    # -- All months sheet -------------------------------------------------
    def write_months(self, sheet, summary):
        pen = self.pen(sheet)
        months = summary.months
        widths = {0: 4.0, 1: 26.0}
        for index in range(len(months) + 2):
            widths[2 + index] = 13.5
        pen.column_widths(widths)

        pen.merge("B2:E2")
        pen.text("B2", "Alle måneder", font=FONT_TITLE, size=18, bold=True,
                 color=ORANGE, align="left")

        header_row = 4
        pen.text("B%d" % header_row, "Kategori", font=FONT_BODY, size=11,
                 bold=True, color=NAVY, align="left")
        for index, month in enumerate(months):
            ref = "%s%d" % (_col_name(2 + index), header_row)
            first, _last = month_bounds(month)
            pen.date(ref, first, font=FONT_BODY, size=11, bold=True, color=NAVY,
                     align="right", fmt=self.formats.month)
        total_col = _col_name(2 + len(months))
        avg_col = _col_name(3 + len(months))
        pen.text("%s%d" % (total_col, header_row), "I alt", font=FONT_BODY,
                 size=11, bold=True, color=NAVY, align="right")
        pen.text("%s%d" % (avg_col, header_row), "Gns./md", font=FONT_BODY,
                 size=11, bold=True, color=NAVY, align="right")

        row = header_row + 1
        row = self._months_block(pen, "UDGIFTER", summary, KIND_EXPENSE, months,
                                 row, total_col, avg_col)
        expense_total_row = row - 1
        row += 1
        row = self._months_block(pen, "INDTÆGTER", summary, KIND_INCOME, months,
                                 row, total_col, avg_col)
        income_total_row = row - 1

        pen.text("B%d" % (row + 1), "NETTO", font=FONT_BODY, size=11, bold=True,
                 color=WHITE, bg=NAVY, align="left")
        for index in range(len(months) + 2):
            column = _col_name(2 + index)
            pen.formula("%s%d" % (column, row + 1),
                        "=%s%d-%s%d" % (column, income_total_row, column,
                                        expense_total_row),
                        font=FONT_BODY, size=11, bold=True, color=WHITE, bg=NAVY,
                        align="right", fmt=self.formats.signed)
        self._print_setup(sheet, 3 + len(months), row + 1)
        self._freeze(sheet, 2, header_row)

    def _months_block(self, pen, title, summary, kind, months, row, total_col,
                      avg_col):
        categories = (summary.expense_categories if kind == KIND_EXPENSE
                      else summary.income_categories)
        pen.text("B%d" % row, title, font=FONT_BODY, size=11, bold=True,
                 color=WHITE, bg=NAVY, align="left")
        for index in range(len(months) + 2):
            pen.style("%s%d" % (_col_name(2 + index), row), bg=NAVY)
        first_row = row + 1
        for offset, category in enumerate(categories):
            current = first_row + offset
            pen.text("B%d" % current, category, font=FONT_BODY, size=10,
                     color=DARK, align="left")
            for index, month in enumerate(months):
                column = _col_name(2 + index)
                pen.number("%s%d" % (column, current),
                           abs(summary.value(kind, category, month)),
                           font=FONT_BODY, size=10, color=DARK, align="right",
                           fmt=self.formats.currency)
            span = "%s%d:%s%d" % (_col_name(2), current,
                                  _col_name(1 + len(months)), current)
            pen.formula("%s%d" % (total_col, current), "=SUM(%s)" % span,
                        font=FONT_BODY, size=10, bold=True, color=NAVY,
                        align="right", fmt=self.formats.currency)
            pen.formula("%s%d" % (avg_col, current), "=IFERROR(AVERAGE(%s);0)" % span,
                        font=FONT_BODY, size=10, color=MUTED, align="right",
                        fmt=self.formats.currency)
        total_row = first_row + len(categories)
        pen.text("B%d" % total_row, "I alt", font=FONT_BODY, size=10, bold=True,
                 color=NAVY, align="left")
        for index in range(len(months) + 2):
            column = _col_name(2 + index)
            pen.formula("%s%d" % (column, total_row),
                        "=SUM(%s%d:%s%d)" % (column, first_row, column,
                                             total_row - 1),
                        font=FONT_BODY, size=10, bold=True, color=NAVY,
                        align="right", fmt=self.formats.currency)
        return total_row + 1


    # -- Budgetforslag ----------------------------------------------------
    def write_plan(self, sheet, plan, targets=None, accounts=None,
                   plan_accounts=None):
        """A draft budget for a normal month, with an editable target column."""
        pen = self.pen(sheet)
        targets = targets or {}
        accounts = accounts or {}
        plan_accounts = plan_accounts or {}
        account_rows = {}
        pen.column_widths({0: 4.0, 1: 26.0, 2: 6.0, 3: 12.5, 4: 12.5, 5: 12.5,
                           6: 12.5, 7: 13.5, 8: 13.5, 9: 16.0})

        pen.merge("B2:E2")
        pen.text("B2", "Budgetforslag", font=FONT_TITLE, size=18, bold=True,
                 color=ORANGE, align="left")
        first, last = plan.months[0], plan.months[-1]
        note = ("Bygger på %d måneder (%s - %s). Skriv dine egne tal i kolonnen "
                "\"Mål\" - \"Forskel\" viser, hvor meget du skal spare (minus) "
                "eller har i luft (plus) i forhold til dit nuværende forbrug. "
                "\"Konto\" sætter, hvilken konto der betaler - se arket \"%s\". "
                "Hold musen over en kategori for at se, hvilke posteringer "
                "der ligger bag." % (len(plan.months), month_title(first).lower(),
                                     month_title(last).lower(), SHEET_ACCOUNTS))
        if plan.uncertain:
            note += " Bemærk: få måneders data - tallene er usikre."
        pen.merge("B3:J3")
        pen.text("B3", note, font=FONT_BODY, size=9, italic=True, color=MUTED,
                 align="left", wrap=True, valign="center")
        pen.row_height(3, 34)

        row = PLAN_HEADER_ROW
        pen.merge("B%d:C%d" % (row, row))
        pen.text("B%d" % row, "Kategori", font=FONT_BODY, size=11, bold=True,
                 color=NAVY, align="left")
        for column, title in (("D", "Gns./md"), ("E", "Typisk md"),
                              ("F", "Laveste"), ("G", "Højeste"),
                              ("H", "Mål"), ("I", "Forskel")):
            pen.text("%s%d" % (column, row), title, font=FONT_BODY, size=11,
                     bold=True, color=NAVY, align="right")
        pen.text("J%d" % row, "Konto", font=FONT_BODY, size=11, bold=True,
                 color=NAVY, align="left")

        rows = {}
        row += 1
        row, rows["income"] = self._plan_block(
            pen, row, "INDTÆGTER", "Forslaget er en typisk måned - ikke den bedste.",
            plan.income, targets, "Indtægt i alt", accounts, plan_accounts,
            account_rows)

        group_rows = []
        for group in (GROUP_FIXED, GROUP_VARIABLE, GROUP_PERIODIC):
            entries = plan.groups.get(group, [])
            if not entries:
                continue
            row += 1
            row, total_row = self._plan_block(
                pen, row, GROUP_LABELS[group], GROUP_HELP[group], entries,
                targets, "I alt", accounts, plan_accounts, account_rows)
            group_rows.append(total_row)

        row += 1
        rows["expense"] = row
        pen.merge("B%d:C%d" % (row, row))
        pen.text("B%d" % row, "UDGIFTER I ALT", font=FONT_BODY, size=11,
                 bold=True, color=WHITE, bg=NAVY, align="left")
        for column in ("D", "H", "I"):
            reference = "+".join("%s%d" % (column, r) for r in group_rows) or "0"
            pen.formula("%s%d" % (column, row), "=%s" % reference,
                        font=FONT_BODY, size=11, bold=True, color=WHITE, bg=NAVY,
                        align="right",
                        fmt=self.formats.signed if column == "I"
                        else self.formats.currency)
        for column in ("E", "F", "G", "J"):
            pen.style("%s%d" % (column, row), bg=NAVY)

        row += 1
        rows["savings"] = row
        pen.merge("B%d:C%d" % (row, row))
        pen.text("B%d" % row, "TIL OPSPARING", font=FONT_BODY, size=11, bold=True,
                 color=WHITE, bg=ORANGE, align="left")
        for column in ("D", "H"):
            pen.formula("%s%d" % (column, row),
                        "=%s%d-%s%d" % (column, rows["income"], column,
                                        rows["expense"]),
                        font=FONT_BODY, size=11, bold=True, color=WHITE,
                        bg=ORANGE, align="right", fmt=self.formats.currency)
        for column in ("E", "F", "G", "J"):
            pen.style("%s%d" % (column, row), bg=ORANGE)
        pen.formula("I%d" % row, "=H%d-D%d" % (row, row), font=FONT_BODY, size=11,
                    bold=True, color=WHITE, bg=ORANGE, align="right",
                    fmt=self.formats.signed)

        row += 2
        rows["goal"] = row
        pen.merge("B%d:C%d" % (row, row))
        pen.text("B%d" % row, "Opsparingsmål pr. måned", font=FONT_BODY, size=10,
                 bold=True, color=NAVY, align="left")
        goal = targets.get(("MÅL", "Opsparingsmål"),
                           max(0.0, _round_500(plan.savings_suggested)))
        pen.number("H%d" % row, goal, font=FONT_BODY, size=10, color=DARK,
                   align="right", bg=PEACH, fmt=self.formats.currency)
        row += 1
        rows["gap"] = row
        pen.merge("B%d:C%d" % (row, row))
        pen.text("B%d" % row, "Forskel til opsparingsmålet", font=FONT_BODY,
                 size=10, bold=True, color=NAVY, align="left")
        pen.formula("H%d" % row, "=H%d-H%d" % (rows["savings"], rows["goal"]),
                    font=FONT_BODY, size=10, bold=True, color=DARK, align="right",
                    fmt=self.formats.signed)
        pen.merge("I%d:I%d" % (row, row))
        pen.formula("D%d" % row,
                    '=IF(H%d<0;"Du mangler at spare et beløb for at nå målet";'
                    '"Målet er inden for rækkevidde")' % row,
                    font=FONT_BODY, size=9, italic=True, color=MUTED, align="left")

        self._add_account_dropdown(sheet, 9, PLAN_HEADER_ROW + 1, row + 1)
        self._print_setup(sheet, 9, row + 1)
        self._freeze(sheet, 0, PLAN_HEADER_ROW)
        rows["accounts"] = account_rows
        return rows

    def _plan_block(self, pen, row, title, help_text, entries, targets, total_label,
                    accounts, plan_accounts, account_rows):
        """One section of the draft budget; returns (next_row, total_row)."""
        pen.merge("B%d:C%d" % (row, row))
        pen.text("B%d" % row, title, font=FONT_BODY, size=11, bold=True,
                 color=WHITE, bg=NAVY, align="left")
        pen.merge("D%d:I%d" % (row, row))
        pen.text("D%d" % row, help_text, font=FONT_BODY, size=9, italic=True,
                 color=LIGHT_TEXT, bg=NAVY, align="left", valign="center")
        pen.style("J%d" % row, bg=NAVY)

        first = row + 1
        for offset, entry in enumerate(entries):
            current = first + offset
            pen.merge("B%d:C%d" % (current, current))
            pen.text("B%d" % current, entry.category, font=FONT_BODY, size=10,
                     bold=True, color=DARK, align="left")
            pen.note("B%d" % current, _plan_comment(entry))
            for column, value in (("D", entry.mean_all), ("E", entry.median),
                                  ("F", entry.low), ("G", entry.high)):
                pen.number("%s%d" % (column, current), value, font=FONT_BODY,
                           size=10, color=DARK if column == "D" else MUTED,
                           align="right", fmt=self.formats.currency)
            target = targets.get((entry.kind, entry.category), entry.suggestion)
            pen.number("H%d" % current, target, font=FONT_BODY, size=10,
                       color=DARK, align="right", bg=PEACH,
                       fmt=self.formats.currency)
            pen.formula("I%d" % current, "=H%d-D%d" % (current, current),
                        font=FONT_BODY, size=10, color=MUTED, align="right",
                        fmt=self.formats.signed)
            account = (plan_accounts.get((entry.kind, entry.category)) or
                      accounts.get(entry.category, ""))
            pen.text("J%d" % current, account, font=FONT_BODY, size=10,
                     color=DARK, align="left")
            if account:
                sign = 1 if entry.kind == KIND_INCOME else -1
                account_rows.setdefault(account, []).append((current, sign))

        total_row = first + len(entries)
        pen.merge("B%d:C%d" % (total_row, total_row))
        pen.text("B%d" % total_row, total_label, font=FONT_BODY, size=10,
                 bold=True, color=NAVY, align="left")
        # Only the average, the target and the difference add up; a sum of
        # medians would be nonsense.
        for column in ("D", "H", "I"):
            pen.formula("%s%d" % (column, total_row),
                        "=SUM(%s%d:%s%d)" % (column, first, column, total_row - 1),
                        font=FONT_BODY, size=10, bold=True, color=NAVY,
                        align="right",
                        fmt=self.formats.signed if column == "I"
                        else self.formats.currency)
        return total_row + 1, total_row

    def read_targets(self):
        """The target amounts the user typed, so a rebuild keeps them."""
        targets = {}
        if not self.doc.Sheets.hasByName(SHEET_PLAN):
            return targets
        sheet = self.doc.Sheets.getByName(SHEET_PLAN)
        last_row = used_row_count(sheet)
        if last_row < PLAN_HEADER_ROW:
            return targets
        data = sheet.getCellRangeByPosition(1, PLAN_HEADER_ROW, 7,
                                            last_row).getDataArray()
        kind = KIND_INCOME
        for row in data:
            name = str(row[0]).strip()
            value = row[6]
            if not name:
                continue
            upper = name.upper()
            if upper.startswith("INDTÆGT"):
                kind = KIND_INCOME
                continue
            if upper.endswith("UDGIFTER") or "UDGIFTER" in upper:
                kind = KIND_EXPENSE
                continue
            if name == "Opsparingsmål pr. måned" and isinstance(value, float):
                targets[("MÅL", "Opsparingsmål")] = value
                continue
            if name.lower().startswith("i alt") or upper.startswith("TIL OPSPARING") \
                    or upper.startswith("FORSKEL"):
                continue
            if isinstance(value, float):
                targets[(kind, name)] = value
        return targets

    # -- Prognose ---------------------------------------------------------
    def write_forecast(self, sheet, plan, plan_rows, start_balance):
        """Where the balance ends up if nothing changes - and if you hit the goals."""
        pen = self.pen(sheet)
        pen.column_widths({0: 4.0, 1: 16.0, 2: 13.0, 3: 13.0, 4: 13.5, 5: 14.5,
                           6: 15.0, 7: 15.5})

        pen.merge("B2:E2")
        pen.text("B2", "Prognose", font=FONT_TITLE, size=18, bold=True,
                 color=ORANGE, align="left")
        pen.merge("B3:H3")
        pen.text("B3", "Sådan udvikler saldoen sig de næste %d måneder, hvis alt "
                       "fortsætter som nu - og hvis du rammer målene i "
                       "\"Budgetforslag\"." % FORECAST_MONTHS,
                 font=FONT_BODY, size=9, italic=True, color=MUTED, align="left",
                 wrap=True, valign="center")

        pen.merge("B5:C5")
        accounts = self._scan_accounts()
        # A freshly seeded Konti sheet has accounts but every balance is
        # still 0 - keep the manual field until the user has actually
        # filled one in, instead of quietly showing a startsaldo of 0.
        if any(balance for _n, _k, balance, _r in accounts):
            pen.text("B5", "Startsaldo (sum af konti)", font=FONT_BODY, size=10,
                     bold=True, color=NAVY, align="left")
            reference = "+".join("$%s.D%d" % (SHEET_ACCOUNTS, row)
                                 for _n, _k, _b, row in accounts)
            pen.formula("D5", "=%s" % reference, font=FONT_BODY, size=10,
                       color=DARK, align="right", fmt=self.formats.currency)
        else:
            pen.text("B5", "Startsaldo", font=FONT_BODY, size=10, bold=True,
                     color=NAVY, align="left")
            pen.number("D5", start_balance or 0.0, font=FONT_BODY, size=10,
                       color=DARK, align="right", bg=PEACH, fmt=self.formats.currency)
        pen.merge("B6:C6")
        pen.text("B6", "Netto pr. måned som nu", font=FONT_BODY, size=10,
                 color=TEXT_GREY, align="left")
        pen.formula("D6", "=$%s.$D$%d" % (SHEET_PLAN, plan_rows["savings"]),
                    font=FONT_BODY, size=10, color=NAVY, bold=True, align="right",
                    fmt=self.formats.signed)
        pen.merge("B7:C7")
        pen.text("B7", "Netto pr. måned med dine mål", font=FONT_BODY, size=10,
                 color=TEXT_GREY, align="left")
        pen.formula("D7", "=$%s.$H$%d" % (SHEET_PLAN, plan_rows["savings"]),
                    font=FONT_BODY, size=10, color=ORANGE, bold=True,
                    align="right", fmt=self.formats.signed)

        head = FORECAST_HEADER_ROW
        for column, title in (("B", "Måned"), ("C", "Indtægt"), ("D", "Udgifter"),
                              ("E", "Netto nu"), ("F", "Netto m. mål"),
                              ("G", "Saldo nu"), ("H", "Saldo m. mål")):
            pen.text("%s%d" % (column, head), title, font=FONT_BODY, size=11,
                     bold=True, color=NAVY,
                     align="left" if column == "B" else "right")

        month = _next_month(plan.months[-1])
        for index in range(FORECAST_MONTHS):
            row = head + 1 + index
            pen.date("B%d" % row, month, font=FONT_BODY, size=10, color=DARK,
                     align="left", fmt=self.formats.month)
            pen.formula("C%d" % row, "=$%s.$D$%d" % (SHEET_PLAN, plan_rows["income"]),
                        font=FONT_BODY, size=10, color=MUTED, align="right",
                        fmt=self.formats.currency)
            pen.formula("D%d" % row, "=$%s.$D$%d" % (SHEET_PLAN, plan_rows["expense"]),
                        font=FONT_BODY, size=10, color=MUTED, align="right",
                        fmt=self.formats.currency)
            pen.formula("E%d" % row, "=C%d-D%d" % (row, row), font=FONT_BODY,
                        size=10, color=NAVY, align="right", fmt=self.formats.signed)
            pen.formula("F%d" % row, "=$D$7", font=FONT_BODY, size=10,
                        color=ORANGE, align="right", fmt=self.formats.signed)
            previous = "$D$5" if index == 0 else "G%d" % (row - 1)
            pen.formula("G%d" % row, "=%s+E%d" % (previous, row), font=FONT_BODY,
                        size=10, bold=True, color=NAVY, align="right",
                        fmt=self.formats.currency)
            previous = "$D$5" if index == 0 else "H%d" % (row - 1)
            pen.formula("H%d" % row, "=%s+F%d" % (previous, row), font=FONT_BODY,
                        size=10, bold=True, color=ORANGE, align="right",
                        fmt=self.formats.currency)
            month = _next_month("%04d-%02d" % (month.year, month.month))

        last_row = head + FORECAST_MONTHS
        self._add_forecast_chart(sheet, head, last_row)
        account_end = self._write_account_forecast(pen, last_row + 26, plan_rows)
        self._print_setup(sheet, 7, max(last_row + 24, account_end))
        self._freeze(sheet, 0, head)
        return last_row

    def _write_account_forecast(self, pen, start_row, plan_rows):
        """A small table: each account's balance in 12/24 months."""
        account_rows = plan_rows.get("accounts") or {}
        accounts = [a for a in self._scan_accounts() if a[0] in account_rows]
        if not accounts:
            return start_row

        pen.merge("B%d:F%d" % (start_row, start_row))
        pen.text("B%d" % start_row, "Konti - forventet udvikling", font=FONT_TITLE,
                 size=14, bold=True, color=ORANGE, align="left")
        pen.merge("B%d:F%d" % (start_row + 1, start_row + 1))
        pen.text("B%d" % (start_row + 1),
                 "Netto/md kommer fra kategoriernes Konto-kolonne i "
                 "\"%s\"." % SHEET_PLAN, font=FONT_BODY, size=9, italic=True,
                 color=MUTED, align="left")

        header = start_row + 2
        for column, title in (("B", "Konto"), ("C", "Nu"), ("D", "Netto/md"),
                              ("E", "Om 12 md"), ("F", "Om 24 md")):
            pen.text("%s%d" % (column, header), title, font=FONT_BODY, size=10,
                     bold=True, color=NAVY,
                     align="left" if column == "B" else "right")

        row = header
        for name, _kind, _balance, account_row in accounts:
            row += 1
            pen.text("B%d" % row, name, font=FONT_BODY, size=10, color=DARK,
                     align="left")
            pen.formula("C%d" % row, "=$%s.D%d" % (SHEET_ACCOUNTS, account_row),
                        font=FONT_BODY, size=10, color=DARK, align="right",
                        fmt=self.formats.currency)
            pen.formula("D%d" % row, _account_net_formula(account_rows[name]),
                        font=FONT_BODY, size=10, bold=True, color=NAVY,
                        align="right", fmt=self.formats.signed)
            pen.formula("E%d" % row, "=C%d+D%d*12" % (row, row), font=FONT_BODY,
                        size=10, color=MUTED, align="right",
                        fmt=self.formats.currency)
            pen.formula("F%d" % row, "=C%d+D%d*24" % (row, row), font=FONT_BODY,
                        size=10, color=MUTED, align="right",
                        fmt=self.formats.currency)
        return row + 1

    def _add_forecast_chart(self, sheet, header_row, last_row):
        """A line chart of the two balance columns."""
        try:
            charts = sheet.Charts
            name = "Saldoudvikling"
            if charts.hasByName(name):
                charts.removeByName(name)
            rect = uno.createUnoStruct("com.sun.star.awt.Rectangle")
            rect.X, rect.Y = 500, 1200 + int(last_row * 450)
            rect.Width, rect.Height = 24000, 10000
            index = sheet.RangeAddress.Sheet
            ranges = []
            for start, end in ((1, 1), (6, 7)):          # column B, columns G:H
                address = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
                address.Sheet = index
                address.StartColumn, address.EndColumn = start, end
                address.StartRow, address.EndRow = header_row - 1, last_row - 1
                ranges.append(address)
            charts.addNewByName(name, rect, tuple(ranges), True, True)
            chart = charts.getByName(name).EmbeddedObject
            chart.setDiagram(chart.createInstance("com.sun.star.chart.LineDiagram"))
            chart.Diagram.SymbolType = -3
            chart.Diagram.Lines = True
            chart.HasMainTitle = True
            chart.Title.String = "Udvikling i saldo"
            chart.HasLegend = True
            try:
                chart.Area.FillStyle = _enum("com.sun.star.drawing.FillStyle",
                                             "NONE")
                wall = chart.Diagram.getWall()
                wall.FillStyle = _enum("com.sun.star.drawing.FillStyle", "SOLID")
                wall.FillColor = WHITE
                wall.LineColor = PANEL
            except Exception:
                pass
            for series, colour in ((0, NAVY), (1, ORANGE)):
                try:
                    properties = chart.Diagram.getDataRowProperties(series)
                    properties.LineColor = colour
                    properties.LineWidth = 60
                except Exception:
                    pass
        except Exception:
            pass


    # -- categories -------------------------------------------------------
    def rebuild(self):
        """Rebuild every summary from what is on the transactions sheet."""
        transactions = self.read_transactions()
        if not transactions:
            return None
        return self.rebuild_summaries(transactions, self.read_planned(),
                                      self.read_start_balance())

    def category_list(self):
        """The categories offered in the drop-down, plus the ones in use."""
        names = []
        if self.doc.Sheets.hasByName(SHEET_RULES):
            sheet = self.doc.Sheets.getByName(SHEET_RULES)
            last_row = max(used_row_count(sheet), 4)
            for row in sheet.getCellRangeByPosition(4, 4, 4, last_row).getDataArray():
                name = str(row[0]).strip()
                if name and name not in names:
                    names.append(name)
        for transaction in self.read_transactions():
            if transaction.category and transaction.category not in names:
                names.append(transaction.category)
        return names

    def set_category_list(self, names):
        """Rewrite the category list on the rules sheet."""
        if not self.doc.Sheets.hasByName(SHEET_RULES):
            return
        sheet = self.doc.Sheets.getByName(SHEET_RULES)
        last_row = max(used_row_count(sheet), 4 + len(names))
        empty = tuple(("",) for _ in range(last_row - 4 + 1))
        sheet.getCellRangeByPosition(4, 4, 4, last_row).setDataArray(empty)
        if names:
            sheet.getCellRangeByPosition(4, 4, 4, 4 + len(names) - 1).setDataArray(
                tuple((name,) for name in names))

    def uncategorised_groups(self):
        """The uncategorised rows, grouped by shop.

        Returns ``[(name, count, total, [(column, row), ...], keyword)]``.
        The grouping ignores amounts, dates and receipt numbers inside the
        text - otherwise every single line would be its own group.
        """
        groups = {}
        order = []
        for transaction in self.read_transactions():
            category = (transaction.category or "").strip()
            if category and category != DEFAULT_CATEGORY:
                continue
            key = merchant_key(transaction.text)
            entry = groups.get(key)
            if entry is None:
                entry = [merchant_name(transaction.text), 0, 0.0, [],
                         rule_keyword(transaction.text)]
                groups[key] = entry
                order.append(entry)
            entry[1] += 1
            entry[2] += transaction.amount
            column = COL_INC_CAT if transaction.income else COL_EXP_CAT
            entry[3].append((column, transaction.source_row))
        order.sort(key=lambda item: abs(item[2]), reverse=True)
        return [tuple(entry) for entry in order]

    def assign_category(self, cells, category):
        """Write ``category`` into the given (column, row) cells."""
        sheet = self.doc.Sheets.getByName(SHEET_TX)
        for column, row in cells:
            sheet.getCellByPosition(column, row).setString(category)

    def rename_category(self, old, new):
        """Rename a category everywhere it is used."""
        if not old or not new or old == new:
            return 0
        changed = 0
        sheet = self.doc.Sheets.getByName(SHEET_TX)
        for transaction in self.read_transactions():
            if transaction.category == old:
                column = COL_INC_CAT if transaction.income else COL_EXP_CAT
                sheet.getCellByPosition(column, transaction.source_row).setString(new)
                changed += 1
        if self.doc.Sheets.hasByName(SHEET_RULES):
            rules_sheet = self.doc.Sheets.getByName(SHEET_RULES)
            last_row = used_row_count(rules_sheet)
            for row in range(4, last_row + 1):
                cell = rules_sheet.getCellByPosition(2, row)
                if cell.getString().strip() == old:
                    cell.setString(new)
        names = [new if name == old else name for name in self.category_list()]
        self.set_category_list(_unique(names))
        return changed

    def delete_category(self, name):
        """Remove a category; its transactions become uncategorised."""
        changed = 0
        sheet = self.doc.Sheets.getByName(SHEET_TX)
        for transaction in self.read_transactions():
            if transaction.category == name:
                column = COL_INC_CAT if transaction.income else COL_EXP_CAT
                sheet.getCellByPosition(
                    column, transaction.source_row).setString(DEFAULT_CATEGORY)
                changed += 1
        if self.doc.Sheets.hasByName(SHEET_RULES):
            rules_sheet = self.doc.Sheets.getByName(SHEET_RULES)
            last_row = used_row_count(rules_sheet)
            keep = []
            for row in range(4, last_row + 1):
                keyword = rules_sheet.getCellByPosition(1, row).getString().strip()
                category = rules_sheet.getCellByPosition(2, row).getString().strip()
                if keyword and category and category != name:
                    keep.append((keyword, category))
            self.set_rules(keep)
        self.set_category_list([n for n in self.category_list() if n != name])
        return changed

    def set_rules(self, pairs):
        """Replace the keyword table on the rules sheet."""
        if not self.doc.Sheets.hasByName(SHEET_RULES):
            return
        sheet = self.doc.Sheets.getByName(SHEET_RULES)
        last_row = max(used_row_count(sheet), 4 + len(pairs))
        empty = tuple(("", "") for _ in range(last_row - 4 + 1))
        sheet.getCellRangeByPosition(1, 4, 2, last_row).setDataArray(empty)
        if pairs:
            sheet.getCellRangeByPosition(1, 4, 2, 4 + len(pairs) - 1).setDataArray(
                tuple((keyword, category) for keyword, category in pairs))

    def add_rules(self, pairs):
        """Append keyword rules, replacing any rule with the same keyword."""
        existing = list(self.read_rules() or [])
        lowered = dict((keyword.strip().lower(), index)
                       for index, (keyword, _c) in enumerate(existing))
        for keyword, category in pairs:
            key = keyword.strip().lower()
            if key in lowered:
                existing[lowered[key]] = (keyword, category)
            else:
                existing.append((keyword, category))
        self.set_rules(existing)
        return existing

    # -- Bankbudget (a printable snapshot) ---------------------------------
    def write_bank_budget(self, sheet, targets):
        """A clean, printable snapshot of the current Mål column - for the bank.

        Values are written as plain numbers, not formulas: this sheet is a
        frozen copy of "Budgetforslag" at the moment you asked for it, so it
        can be handed to someone else without changing under them.
        """
        pen = self.pen(sheet)
        pen.column_widths({0: 4.0, 1: 40.0, 2: 16.0, 3: 4.0})

        pen.merge("B2:C2")
        pen.text("B2", "Husstandsbudget", font=FONT_TITLE, size=20, bold=True,
                 color=ORANGE, align="left")
        pen.merge("B3:C3")
        pen.text("B3", "Øjebliksbillede af målene i \"%s\" pr. %s. Opdateres "
                       "ikke automatisk - kør \"Budget ▸ Opret bankbudget\" "
                       "igen, hvis du retter dine mål."
                       % (SHEET_PLAN, datetime.date.today().strftime("%d-%m-%Y")),
                 font=FONT_BODY, size=9, italic=True, color=MUTED, align="left",
                 wrap=True, valign="center")
        pen.row_height(3, 40)

        income = sorted(((name, value) for (kind, name), value in targets.items()
                         if kind == KIND_INCOME and value),
                        key=lambda item: item[1], reverse=True)
        expenses = sorted(((name, value) for (kind, name), value in targets.items()
                           if kind == KIND_EXPENSE and value),
                          key=lambda item: item[1], reverse=True)
        income_total = sum(value for _name, value in income)
        expense_total = sum(value for _name, value in expenses)

        row = 5
        row = self._bank_section(pen, row, "Indtægter pr. måned", income,
                                 income_total)
        row += 1
        row = self._bank_section(pen, row, "Udgifter pr. måned", expenses,
                                 expense_total)
        row += 1
        pen.text("B%d" % row, "Rådighedsbeløb / opsparing",
                 font=FONT_BODY, size=11, bold=True, color=WHITE, bg=NAVY,
                 align="left")
        pen.number("C%d" % row, income_total - expense_total, font=FONT_BODY,
                   size=12, bold=True, color=WHITE, bg=NAVY, align="right",
                   fmt=self.formats.signed)

        self._print_setup(sheet, 2, row + 2)
        self._freeze(sheet, 0, 4)
        return row

    def _bank_section(self, pen, row, title, items, total):
        pen.text("B%d" % row, title, font=FONT_BODY, size=13, bold=True,
                 color=NAVY, align="left")
        row += 1
        if not items:
            pen.text("B%d" % row, "(ingen)", font=FONT_BODY, size=10, italic=True,
                     color=MUTED, align="left")
            row += 1
        for name, value in items:
            pen.text("B%d" % row, name, font=FONT_BODY, size=10, color=DARK,
                     align="left")
            pen.number("C%d" % row, value, font=FONT_BODY, size=10, color=DARK,
                       align="right", fmt=self.formats.currency)
            row += 1
        pen.text("B%d" % row, "%s - i alt" % title, font=FONT_BODY, size=10,
                 bold=True, color=NAVY, align="left")
        pen.number("C%d" % row, total, font=FONT_BODY, size=10, bold=True,
                   color=NAVY, align="right", fmt=self.formats.currency)
        return row + 1

    # -- reading an existing document -------------------------------------
    def has_budget(self):
        return self.doc.Sheets.hasByName(SHEET_TX)

    def read_transactions(self):
        """Read the transactions back from the document."""
        sheet = self.doc.Sheets.getByName(SHEET_TX)
        last_row = used_row_count(sheet)
        transactions = []
        if last_row < TX_FIRST_ROW:
            return transactions
        data = sheet.getCellRangeByPosition(
            COL_EXP_DATE, TX_FIRST_ROW, COL_INC_CAT, last_row).getDataArray()
        for offset, row in enumerate(data):
            for base, sign in ((0, -1.0), (COL_INC_DATE - COL_EXP_DATE, 1.0)):
                serial, amount, text, category = row[base:base + 4]
                if not isinstance(serial, float) or not serial:
                    continue
                if not isinstance(amount, float):
                    continue
                date = self.null_date + datetime.timedelta(days=int(serial))
                transactions.append(_SheetTransaction(
                    date, str(text), sign * abs(amount), str(category).strip(),
                    TX_FIRST_ROW + offset, base != 0))
        transactions.sort(key=lambda t: (t.date, t.source_row))
        return transactions

    def read_rules(self):
        """Read the keyword rules from the rules sheet."""
        if not self.doc.Sheets.hasByName(SHEET_RULES):
            return None
        sheet = self.doc.Sheets.getByName(SHEET_RULES)
        last_row = used_row_count(sheet)
        if last_row < 4:
            return None
        data = sheet.getCellRangeByPosition(1, 4, 2, last_row).getDataArray()
        return [(str(row[0]).strip(), str(row[1]).strip()) for row in data
                if str(row[0]).strip() and str(row[1]).strip()]

    def read_planned(self):
        """Existing budget figures, so they survive a rebuild."""
        planned = {}
        for index in range(self.doc.Sheets.Count):
            sheet = self.doc.Sheets.getByIndex(index)
            if not sheet.Name.startswith(SUMMARY_PREFIX):
                continue
            last_row = used_row_count(sheet)
            if last_row < SUM_FIRST_CAT_ROW:
                continue
            data = sheet.getCellRangeByPosition(
                1, SUM_FIRST_CAT_ROW - 1, 11, last_row).getDataArray()
            for row in data:
                for name_index, value_index, kind in ((0, 2, KIND_EXPENSE),
                                                      (6, 8, KIND_INCOME)):
                    name = str(row[name_index]).strip()
                    value = row[value_index]
                    if name and isinstance(value, float) and value:
                        planned.setdefault((kind, name), value)
        return planned

    def read_forecast_start(self):
        """The start balance the user typed on the forecast sheet, if any."""
        if not self.doc.Sheets.hasByName(SHEET_FORECAST):
            return None
        try:
            return self.doc.Sheets.getByName(SHEET_FORECAST) \
                .getCellRangeByName("D5").getValue()
        except Exception:
            return None

    def read_start_balance(self):
        for index in range(self.doc.Sheets.Count):
            sheet = self.doc.Sheets.getByIndex(index)
            if sheet.Name.startswith(SUMMARY_PREFIX):
                value = sheet.getCellRangeByName("L8").getValue()
                return value
        return 0.0

    def refresh(self, ruleset):
        """Re-apply the rules to the sheet and rebuild every summary."""
        transactions = self.read_transactions()
        if not transactions:
            return None, 0
        changed = 0
        sheet = self.doc.Sheets.getByName(SHEET_TX)
        self.doc.lockControllers()
        try:
            for transaction in transactions:
                category = ruleset.categorise(transaction.text)
                if category == DEFAULT_CATEGORY and transaction.category:
                    # Keep a category the user picked by hand.
                    category = transaction.category
                if category != transaction.category:
                    column = COL_INC_CAT if transaction.income else COL_EXP_CAT
                    sheet.getCellByPosition(column,
                                            transaction.source_row).setString(category)
                    transaction.category = category
                    changed += 1
        finally:
            self.doc.unlockControllers()
        planned = self.read_planned()
        start_balance = self.read_start_balance()
        summary = self.rebuild_summaries(transactions, planned, start_balance)
        return summary, changed

    def rebuild_summaries(self, transactions, planned, start_balance):
        """Recreate the summary sheets (and the month overview) in place."""
        summary = Summary(transactions)
        months = summary.months or [""]
        targets = self.read_targets()
        forecast_start = self.read_forecast_start()
        # Read back before the old sheet is torn down below - the whole
        # point is to carry these across the rebuild.
        category_accounts = self.read_category_accounts()
        plan_accounts = self.read_plan_accounts()
        self.doc.lockControllers()
        try:
            for index in range(self.doc.Sheets.Count - 1, -1, -1):
                name = self.doc.Sheets.getByIndex(index).Name
                if name.startswith(SUMMARY_PREFIX) or name in (SHEET_MONTHS,
                                                               SHEET_PLAN,
                                                               SHEET_FORECAST):
                    self.doc.Sheets.removeByName(name)

            self._ensure_accounts_sheet()
            plan = build_plan(summary, transactions) if len(summary.months) > 1 else None
            if plan is not None:
                plan_rows = self.write_plan(self.sheet(SHEET_PLAN), plan, targets,
                                            accounts=category_accounts,
                                            plan_accounts=plan_accounts)
                if forecast_start is None:
                    accounts_total = self._accounts_total()
                    forecast_start = (accounts_total if accounts_total else
                                      (start_balance or 0.0) + summary.net)
                self.write_forecast(self.sheet(SHEET_FORECAST), plan, plan_rows,
                                    forecast_start)
            previous = None
            for index, month in enumerate(months):
                name = self.summary_name(months, month)
                sheet = self.sheet(name, index)
                self.write_summary(sheet, summary, month,
                                   multi_month=len(months) > 1,
                                   planned=planned,
                                   start_balance=start_balance if index == 0 else None,
                                   previous_sheet=previous)
                previous = name
            if len(months) > 1:
                self.write_months(self.sheet(SHEET_MONTHS), summary)
            self._order_sheets(months)
        finally:
            self.doc.unlockControllers()
        self.doc.calculateAll()
        return summary


class _SheetTransaction(object):
    """A transaction read back from the document."""

    __slots__ = ("date", "text", "amount", "category", "source_row", "income",
                 "kind", "month", "currency")

    def __init__(self, date, text, amount, category, source_row, income):
        self.date = date
        self.text = text
        self.amount = amount
        self.category = category
        self.source_row = source_row
        self.income = income
        self.kind = KIND_INCOME if amount >= 0 else KIND_EXPENSE
        self.month = "%04d-%02d" % (date.year, date.month)
        self.currency = ""


def used_row_count(sheet):
    """Index (0 based) of the last used row of ``sheet``."""
    try:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        return cursor.RangeAddress.EndRow
    except Exception:
        return 0


def read_sheet_as_table(doc, sheet):
    """Read any sheet as rows of strings, so the CSV detection can be reused."""
    null_date = sheet_null_date(doc)
    last_row = used_row_count(sheet)
    try:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        last_col = cursor.RangeAddress.EndColumn
    except Exception:
        last_col = 0
    if last_row < 0 or last_col < 0:
        return []
    area = sheet.getCellRangeByPosition(0, 0, last_col, last_row)
    values = area.getDataArray()

    # Which columns are formatted as dates?  Those floats are date serials.
    date_columns = set()
    formats = doc.getNumberFormats()
    for column in range(last_col + 1):
        for row in range(min(last_row + 1, 50)):
            cell = sheet.getCellByPosition(column, row)
            if cell.getType().value != "VALUE":
                continue
            try:
                properties = formats.getByKey(cell.NumberFormat)
                if properties.Type & 2:  # com.sun.star.util.NumberFormat.DATE
                    date_columns.add(column)
            except Exception:
                pass
            break

    rows = []
    for row in values:
        out = []
        for index, value in enumerate(row):
            if isinstance(value, float):
                if index in date_columns:
                    try:
                        out.append((null_date +
                                    datetime.timedelta(days=int(value))).isoformat())
                        continue
                    except Exception:
                        pass
                if value == int(value):
                    out.append(str(int(value)))
                else:
                    out.append(("%.10f" % value).rstrip("0"))
            else:
                out.append(str(value))
        rows.append(out)
    return rows


# ---------------------------------------------------------------------------
# Column helpers
# ---------------------------------------------------------------------------

def _unique(names):
    out = []
    for name in names:
        if name and name not in out:
            out.append(name)
    return out


def _round_500(value):
    return round(value / 500.0) * 500 if value else 0.0


def _format_kr(value):
    return "{:,.0f}".format(abs(value)).replace(",", ".") + " kr."


def _plan_comment(entry):
    """Hover text listing the transactions behind one Budgetforslag row."""
    if not entry.transaction_count:
        return ""
    lines = ["%d postering(er) i alt - gennemsnit %s/md."
             % (entry.transaction_count, _format_kr(entry.mean_all))]
    if entry.sample:
        lines.append("")
        lines.append("Største bidrag:")
        for t in entry.sample:
            lines.append("%s  %-26s %10s" % (t.date.strftime("%d-%m-%Y"),
                                             t.text[:26], _format_kr(t.amount)))
        extra = entry.transaction_count - len(entry.sample)
        if extra > 0:
            lines.append("... og %d flere" % extra)
    return "\n".join(lines)


def _account_net_formula(rows_signs):
    """Formula: sum of Budgetforslag Mål cells assigned to one account."""
    if not rows_signs:
        return "=0"
    parts = ["%s$%s.H%d" % ("+" if sign > 0 else "-", SHEET_PLAN, row)
             for row, sign in rows_signs]
    return "=" + "".join(parts)


def _next_month(month_key):
    """``"2025-12"`` -> ``date(2026, 1, 1)``."""
    year, month = (int(part) for part in month_key.split("-"))
    if month == 12:
        return datetime.date(year + 1, 1, 1)
    return datetime.date(year, month + 1, 1)


def _col_name(index):
    """0 -> A, 1 -> B, ..."""
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(ord("A") + remainder) + name
    return name


def _ref(col1, row1, col2, row2):
    """0 based column, 1 based row -> ``"B5:E20"``."""
    return "%s%d:%s%d" % (_col_name(col1), row1, _col_name(col2), row2)
