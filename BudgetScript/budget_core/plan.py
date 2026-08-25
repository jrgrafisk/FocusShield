"""Turn a year of transactions into a draft budget.

Every category is measured across the months of the file and sorted into
three groups:

``fast``       the same amount almost every month (husleje, forsikring, lån)
``variabel``   there every month, but the amount moves (dagligvarer, restaurant)
``periodisk``  only a few months a year (el-afregning, ferie, jul) - these are
               budgeted as a monthly set-aside of the yearly total

From that we suggest an amount per category: a careful figure for income
(the typical month, not the best one), the average for the fixed costs and
the *median* month for the variable ones, so a single expensive December
does not blow up the grocery budget.
"""

from __future__ import annotations

import math
from typing import Dict, List, Optional, Sequence

from .budget import KIND_EXPENSE, KIND_INCOME, Summary, Transaction
from .textutils import fold

__all__ = ["CategoryPlan", "BudgetPlan", "build_plan", "GROUP_FIXED",
           "GROUP_VARIABLE", "GROUP_PERIODIC", "GROUP_LABELS", "SAMPLE_SIZE"]

# How many contributing transactions to remember per category (for the
# on-hover comment in the spreadsheet).
SAMPLE_SIZE = 8

GROUP_FIXED = "fast"
GROUP_VARIABLE = "variabel"
GROUP_PERIODIC = "periodisk"

GROUP_LABELS = {
    GROUP_FIXED: "FASTE UDGIFTER",
    GROUP_VARIABLE: "VARIABLE UDGIFTER",
    GROUP_PERIODIC: "PERIODISKE UDGIFTER",
}

GROUP_HELP = {
    GROUP_FIXED: "Samme beløb hver måned - budgettet er gennemsnittet.",
    GROUP_VARIABLE: "Svinger fra måned til måned - forslaget er en typisk "
                    "måned (median), ikke gennemsnittet.",
    GROUP_PERIODIC: "Kommer få gange om året - her henlagt pr. måned, så "
                    "årets samlede beløb er delt ligeligt ud.",
}

# Categories that are fixed costs by nature.  Billed every month they are
# "faste"; billed quarterly or yearly they become a monthly set-aside.
FIXED_BY_NATURE = ("bolig", "el, vand og varme", "forsikring og pension",
                   "telefon og internet", "abonnementer", "lan og afdrag",
                   "born og uddannelse", "gebyrer og renter")

# Categories that are spending decisions, not bills - they belong under
# "variable" even in a year where the monthly total happened to be steady.
VARIABLE_BY_NATURE = ("dagligvarer", "restaurant", "shopping", "transport",
                      "fritid og sport", "sundhed", "rejser", "mobilepay",
                      "overforsel", "ukategoriseret")

# How the groups are detected.
MONTHLY_COVERAGE = 0.90        # a bill that really does turn up every month
FIXED_COVERAGE = 0.70          # present in at least this share of the months
FIXED_VARIATION = 0.30         # and varies less than this (std/mean)
STEADY_VARIATION = 0.10        # so steady that it must be a subscription
PERIODIC_COVERAGE = 0.40       # present in at most this share of the months
UNCERTAIN_MONTHS = 6           # fewer months than this -> warn the user


def _median(values: Sequence[float]) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2.0


def _stdev(values: Sequence[float]) -> float:
    if len(values) < 2:
        return 0.0
    mean = sum(values) / float(len(values))
    variance = sum((v - mean) ** 2 for v in values) / float(len(values) - 1)
    return math.sqrt(variance)


def _round_to(value: float, step: int, mode: str = "nearest") -> float:
    if step <= 0 or not value:
        return round(value, 2)
    quotient = value / float(step)
    if mode == "up":
        quotient = math.ceil(quotient)
    elif mode == "down":
        quotient = math.floor(quotient)
    else:
        quotient = math.floor(quotient + 0.5)
    return quotient * step


class CategoryPlan(object):
    """What one category costs per month, and what we suggest budgeting."""

    __slots__ = ("kind", "category", "group", "values", "months_present",
                 "total", "mean_all", "mean_active", "median", "low", "high",
                 "variation", "suggestion", "sample", "transaction_count",
                 "account")

    def __init__(self, kind: str, category: str, values: Sequence[float],
                 transactions: Sequence[Transaction] = ()):
        self.kind = kind
        self.category = category
        self.values = [abs(v) for v in values]
        active = [v for v in self.values if v > 0.005]
        months = len(self.values) or 1

        self.months_present = len(active)
        self.total = sum(self.values)
        self.mean_all = self.total / float(months)
        self.mean_active = (sum(active) / float(len(active))) if active else 0.0
        self.median = _median(active)
        self.low = min(active) if active else 0.0
        self.high = max(active) if active else 0.0
        self.variation = (_stdev(active) / self.mean_active) if self.mean_active else 0.0
        self.group = GROUP_VARIABLE
        self.suggestion = 0.0
        self.account = ""

        ordered = sorted(transactions, key=lambda t: abs(t.amount), reverse=True)
        self.sample = ordered[:SAMPLE_SIZE]
        self.transaction_count = len(transactions)

    @property
    def coverage(self) -> float:
        return self.months_present / float(len(self.values) or 1)

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return "<CategoryPlan %s %s %s %.0f>" % (self.kind, self.category,
                                                 self.group, self.suggestion)


def _classify(plan: CategoryPlan) -> str:
    if plan.kind == KIND_INCOME:
        return GROUP_FIXED
    name = fold(plan.category)
    if name in FIXED_BY_NATURE:
        # El, vand og varme is a fixed cost - but if it is only billed every
        # quarter, budgeting the bill itself would be wrong.  Spread it out.
        return (GROUP_FIXED if plan.coverage >= MONTHLY_COVERAGE
                else GROUP_PERIODIC)
    if plan.coverage <= PERIODIC_COVERAGE:
        return GROUP_PERIODIC
    steady = plan.coverage >= FIXED_COVERAGE and plan.variation <= FIXED_VARIATION
    if steady and not (name in VARIABLE_BY_NATURE and
                       plan.variation > STEADY_VARIATION):
        return GROUP_FIXED
    return GROUP_VARIABLE


def _suggest(plan: CategoryPlan) -> float:
    """The amount we propose budgeting per month."""
    if plan.kind == KIND_INCOME:
        # Budget on the typical month, not on the month with the bonus.
        base = plan.median if plan.coverage >= FIXED_COVERAGE else plan.mean_all
        return _round_to(base, 100, "down")
    if plan.group == GROUP_FIXED:
        return _round_to(plan.mean_active, 10, "up")
    if plan.group == GROUP_PERIODIC:
        return _round_to(plan.mean_all, 50, "up")
    base = plan.median or plan.mean_active
    return _round_to(base, 50)


class BudgetPlan(object):
    """A draft budget for a normal month."""

    def __init__(self, months: Sequence[str], categories: Sequence[CategoryPlan]):
        self.months = list(months)
        self.categories = list(categories)
        self.income = [c for c in self.categories if c.kind == KIND_INCOME]
        self.expenses = [c for c in self.categories if c.kind == KIND_EXPENSE]
        self.groups: Dict[str, List[CategoryPlan]] = {
            GROUP_FIXED: [c for c in self.expenses if c.group == GROUP_FIXED],
            GROUP_VARIABLE: [c for c in self.expenses if c.group == GROUP_VARIABLE],
            GROUP_PERIODIC: [c for c in self.expenses if c.group == GROUP_PERIODIC],
        }

    # -- totals -----------------------------------------------------------
    @property
    def month_count(self) -> int:
        return max(1, len(self.months))

    @property
    def income_mean(self) -> float:
        return sum(c.mean_all for c in self.income)

    @property
    def expense_mean(self) -> float:
        return sum(c.mean_all for c in self.expenses)

    @property
    def savings_mean(self) -> float:
        return self.income_mean - self.expense_mean

    @property
    def income_suggested(self) -> float:
        return sum(c.suggestion for c in self.income)

    @property
    def expense_suggested(self) -> float:
        return sum(c.suggestion for c in self.expenses)

    @property
    def savings_suggested(self) -> float:
        return self.income_suggested - self.expense_suggested

    @property
    def uncertain(self) -> bool:
        return len(self.months) < UNCERTAIN_MONTHS

    def group_mean(self, group: str) -> float:
        return sum(c.mean_all for c in self.groups.get(group, ()))

    def summary_lines(self) -> List[str]:
        lines = [
            "Gennemsnitlig indtægt: %.0f kr./md" % self.income_mean,
            "Faste udgifter: %.0f kr./md" % self.group_mean(GROUP_FIXED),
            "Variable udgifter: %.0f kr./md" % self.group_mean(GROUP_VARIABLE),
            "Periodiske udgifter: %.0f kr./md" % self.group_mean(GROUP_PERIODIC),
            "Til opsparing i gennemsnit: %.0f kr./md" % self.savings_mean,
            "Forslaget giver plads til %.0f kr./md i opsparing"
            % self.savings_suggested,
        ]
        if self.uncertain:
            lines.append("Bemærk: forslaget bygger kun på %d måned(er) - "
                         "tallene er usikre." % len(self.months))
        return lines


def build_plan(summary: Summary, transactions: Sequence[Transaction] = (),
              months: Optional[Sequence[str]] = None) -> BudgetPlan:
    """Build a draft budget from a :class:`~budget_core.budget.Summary`.

    ``transactions`` (the same list used to build ``summary``) lets each
    category remember its largest contributing transactions, for the
    on-hover comment shown in the spreadsheet.
    """
    months = list(months or summary.months)
    by_category: Dict[tuple, List[Transaction]] = {}
    for t in transactions:
        if t.month in months:
            by_category.setdefault((t.kind, t.category), []).append(t)

    categories: List[CategoryPlan] = []
    for kind, names in ((KIND_INCOME, summary.income_categories),
                        (KIND_EXPENSE, summary.expense_categories)):
        for name in names:
            values = [summary.value(kind, name, month) for month in months]
            plan = CategoryPlan(kind, name, values,
                                by_category.get((kind, name), ()))
            plan.group = _classify(plan)
            plan.suggestion = _suggest(plan)
            categories.append(plan)

    # Biggest first inside each group.
    categories.sort(key=lambda c: (-c.mean_all, c.category))
    return BudgetPlan(months, categories)
