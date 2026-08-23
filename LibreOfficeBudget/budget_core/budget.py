"""Turn a detected table into categorised transactions and a budget.

The output of :func:`build_transactions` is deliberately plain (a list of
:class:`Transaction`) so it can be written to a CSV file, to a Calc
document or to anything else.
"""

from __future__ import annotations

import datetime as _dt
from collections import Counter, OrderedDict
from typing import Dict, List, Optional, Sequence, Tuple

from .parsing import month_key, month_label, parse_amount, parse_date
from .rules import DEFAULT_CATEGORY, RuleSet
from .textutils import squeeze

__all__ = ["Transaction", "BuildOptions", "BuildResult", "Summary",
           "build_transactions", "summarise", "KIND_INCOME", "KIND_EXPENSE"]

KIND_INCOME = "Indtægt"
KIND_EXPENSE = "Udgift"

SIGN_AUTO = "auto"
SIGN_NEGATIVE_IS_EXPENSE = "negativ"
SIGN_POSITIVE_IS_EXPENSE = "positiv"


class Transaction(object):
    """One line in the ``Transaktioner`` sheet."""

    __slots__ = ("date", "text", "amount", "category", "kind", "month",
                 "currency", "source_row")

    def __init__(self, date, text, amount, category, currency="", source_row=0):
        self.date: _dt.date = date
        self.text: str = text
        self.amount: float = amount          # negative = expense
        self.category: str = category
        self.kind: str = KIND_EXPENSE if amount < 0 else KIND_INCOME
        self.month: str = month_key(date)
        self.currency: str = currency
        self.source_row: int = source_row

    @property
    def month_name(self) -> str:
        return month_label(self.month)

    def as_row(self) -> List[object]:
        return [self.date, self.text, self.amount, self.category, self.kind,
                self.month]

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return "<Transaction %s %s %.2f %s>" % (self.date, self.text[:20],
                                                self.amount, self.category)


class BuildOptions(object):
    """User adjustable settings for :func:`build_transactions`."""

    def __init__(self, decimal=None, dayfirst=None, sign=SIGN_AUTO,
                 skip_zero=True, date_from=None, date_to=None):
        self.decimal: Optional[str] = decimal
        self.dayfirst: Optional[bool] = dayfirst
        self.sign: str = sign
        self.skip_zero: bool = skip_zero
        self.date_from: Optional[_dt.date] = date_from
        self.date_to: Optional[_dt.date] = date_to


class BuildResult(object):
    def __init__(self, transactions, warnings, skipped_no_date, skipped_no_amount,
                 uncategorised):
        self.transactions: List[Transaction] = transactions
        self.warnings: List[str] = warnings
        self.skipped_no_date = skipped_no_date
        self.skipped_no_amount = skipped_no_amount
        # [(text, count, total)] sorted by |total| descending
        self.uncategorised: List[Tuple[str, int, float]] = uncategorised

    @property
    def ok(self) -> bool:
        return bool(self.transactions)

    def summary_lines(self) -> List[str]:
        lines = ["%d posteringer indlæst." % len(self.transactions)]
        if self.transactions:
            first = min(t.date for t in self.transactions)
            last = max(t.date for t in self.transactions)
            lines.append("Periode: %s - %s" % (first.isoformat(), last.isoformat()))
        if self.skipped_no_date:
            lines.append("%d række(r) uden gyldig dato blev sprunget over."
                         % self.skipped_no_date)
        if self.skipped_no_amount:
            lines.append("%d række(r) uden gyldigt beløb blev sprunget over."
                         % self.skipped_no_amount)
        uncategorised = sum(1 for t in self.transactions
                            if t.category == DEFAULT_CATEGORY)
        if uncategorised:
            lines.append("%d postering(er) er ikke kategoriseret." % uncategorised)
        return lines + list(self.warnings)


def _cell(row: Sequence[str], index: Optional[int]) -> str:
    if index is None or index < 0 or index >= len(row):
        return ""
    value = row[index]
    return "" if value is None else str(value)


def _decide_sign(values: Sequence[float], mode: str) -> bool:
    """Return True when the amounts must be negated to make expenses negative."""
    if mode == SIGN_POSITIVE_IS_EXPENSE:
        return True
    if mode == SIGN_NEGATIVE_IS_EXPENSE:
        return False
    negatives = sum(1 for v in values if v < 0)
    return negatives == 0 and len(values) > 0


def build_transactions(table, mapping, ruleset: Optional[RuleSet] = None,
                       options: Optional[BuildOptions] = None) -> BuildResult:
    """Parse every row of ``table`` according to ``mapping``."""
    ruleset = ruleset or RuleSet.defaults()
    options = options or BuildOptions()
    warnings: List[str] = list(mapping.notes)
    decimal = options.decimal or mapping.decimal
    dayfirst = mapping.dayfirst if options.dayfirst is None else options.dayfirst

    if mapping.date is None:
        return BuildResult([], warnings + ["Ingen datokolonne valgt."], 0, 0, [])
    if not mapping.has_amount:
        return BuildResult([], warnings + ["Ingen beløbskolonne valgt."], 0, 0, [])

    parsed: List[Tuple[_dt.date, str, float, str, int]] = []
    skipped_no_date = 0
    skipped_no_amount = 0

    for index, row in enumerate(table.rows):
        date = parse_date(_cell(row, mapping.date), dayfirst)
        amount = _row_amount(row, mapping, decimal)
        if date is None:
            if amount is not None or any(str(c).strip() for c in row):
                skipped_no_date += 1
            continue
        if amount is None:
            skipped_no_amount += 1
            continue
        if options.skip_zero and abs(amount) < 0.0000001:
            continue
        if options.date_from and date < options.date_from:
            continue
        if options.date_to and date > options.date_to:
            continue
        text = " – ".join(
            part for part in (squeeze(_cell(row, i)) for i in mapping.text) if part)
        if not text:
            text = "(ingen tekst)"
        currency = squeeze(_cell(row, mapping.currency)) if mapping.currency is not None else ""
        parsed.append((date, text, amount, currency, index))

    if not parsed:
        return BuildResult([], warnings + [
            "Ingen brugbare rækker. Kontrollér kolonnevalget i dialogen."],
            skipped_no_date, skipped_no_amount, [])

    flip = False
    if mapping.amount is not None and mapping.amount_in is None:
        flip = _decide_sign([p[2] for p in parsed], options.sign)
        if flip and options.sign == SIGN_AUTO:
            warnings.append(
                "Beløbskolonnen indeholder ingen negative tal - alle poster "
                "behandles som udgifter. Skift fortegnsregel i dialogen hvis "
                "filen også indeholder indtægter.")

    transactions: List[Transaction] = []
    for date, text, amount, currency, index in parsed:
        value = -amount if flip else amount
        transactions.append(Transaction(date, text, value,
                                        ruleset.categorise(text), currency, index))

    transactions.sort(key=lambda t: (t.date, t.source_row))

    counter: Dict[str, List[float]] = OrderedDict()
    for transaction in transactions:
        if transaction.category == DEFAULT_CATEGORY:
            entry = counter.setdefault(transaction.text, [0, 0.0])
            entry[0] += 1
            entry[1] += transaction.amount
    uncategorised = sorted(((text, int(count), total)
                            for text, (count, total) in counter.items()),
                           key=lambda item: abs(item[2]), reverse=True)

    return BuildResult(transactions, warnings, skipped_no_date, skipped_no_amount,
                       uncategorised)


def _row_amount(row: Sequence[str], mapping, decimal: Optional[str]) -> Optional[float]:
    """Amount of one row, honouring single or split amount columns."""
    if mapping.amount_in is not None or mapping.amount_out is not None:
        value_in = parse_amount(_cell(row, mapping.amount_in), decimal) \
            if mapping.amount_in is not None else None
        value_out = parse_amount(_cell(row, mapping.amount_out), decimal) \
            if mapping.amount_out is not None else None
        if value_in is None and value_out is None:
            return None
        return abs(value_in or 0.0) - abs(value_out or 0.0)
    return parse_amount(_cell(row, mapping.amount), decimal)


# --------------------------------------------------------------------------
# Summary
# --------------------------------------------------------------------------

class Summary(object):
    """Monthly totals per category - the actual budget."""

    def __init__(self, transactions: Sequence[Transaction]):
        self.months: List[str] = sorted({t.month for t in transactions})
        self.cells: Dict[Tuple[str, str, str], float] = {}
        self.counts: Counter = Counter()
        for t in transactions:
            key = (t.kind, t.category, t.month)
            self.cells[key] = self.cells.get(key, 0.0) + t.amount
            self.counts[(t.kind, t.category)] += 1
        self.expense_categories = self._categories(KIND_EXPENSE)
        self.income_categories = self._categories(KIND_INCOME)
        self.transaction_count = len(transactions)

    def _categories(self, kind: str) -> List[str]:
        totals: Dict[str, float] = {}
        for (k, category, _month), value in self.cells.items():
            if k == kind:
                totals[category] = totals.get(category, 0.0) + value
        return sorted(totals, key=lambda c: (-abs(totals[c]), c))

    def value(self, kind: str, category: str, month: str) -> float:
        return self.cells.get((kind, category, month), 0.0)

    def category_total(self, kind: str, category: str) -> float:
        return sum(self.value(kind, category, m) for m in self.months)

    def month_total(self, kind: str) -> Dict[str, float]:
        out = {}
        for month in self.months:
            out[month] = sum(self.value(kind, c, month)
                             for c in self._categories(kind))
        return out

    def total(self, kind: str) -> float:
        return sum(v for (k, _c, _m), v in self.cells.items() if k == kind)

    @property
    def net(self) -> float:
        return self.total(KIND_INCOME) + self.total(KIND_EXPENSE)

    def month_count(self) -> int:
        return max(1, len(self.months))

    def top_expenses(self, count: int = 5) -> List[Tuple[str, float]]:
        return [(c, self.category_total(KIND_EXPENSE, c))
                for c in self.expense_categories[:count]]


def summarise(transactions: Sequence[Transaction]) -> Summary:
    return Summary(transactions)
