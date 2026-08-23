#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Budget fra CSV - kommandolinjeudgaven.

Læser en CSV-fil fra banken, kategoriserer posteringerne og skriver en
kategoriseret CSV plus en månedsoversigt - samme motor som
LibreOffice-udvidelsen bruger.

    python3 budget_cli.py kontoudtog.csv
    python3 budget_cli.py kontoudtog.csv --interaktiv --regler mine.csv
    python3 budget_cli.py kontoudtog.csv --kun-rapport
"""

from __future__ import annotations

import argparse
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from budget_core import csvsniff                                  # noqa: E402
from budget_core.budget import (BuildOptions, KIND_EXPENSE,       # noqa: E402
                                KIND_INCOME, SIGN_AUTO,
                                SIGN_NEGATIVE_IS_EXPENSE,
                                SIGN_POSITIVE_IS_EXPENSE,
                                build_transactions, summarise)
from budget_core.columns import detect_mapping                    # noqa: E402
from budget_core.parsing import month_label                       # noqa: E402
from budget_core.rules import (DEFAULT_CATEGORY, RuleSet,         # noqa: E402
                               user_rules_path)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Lav et budget ud fra en CSV-fil med posteringer.")
    parser.add_argument("fil", help="CSV-filen med posteringer")
    parser.add_argument("-o", "--ud", help="mappe til resultatfilerne "
                                           "(standard: samme som inputfilen)")
    parser.add_argument("--regler", help="CSV/JSON-fil med kategoriregler "
                                         "(nøgleord;kategori)")
    parser.add_argument("--gem-regler", action="store_true",
                        help="gem reglerne (inkl. nye) til --regler eller "
                             "brugerens standardplacering")
    parser.add_argument("--tegnsaet", help="tving et tegnsæt, fx utf-8 eller cp1252")
    parser.add_argument("--skilletegn", help="tving et skilletegn, fx ';' eller ','")
    parser.add_argument("--decimal", choices=[",", "."],
                        help="tving decimaltegn")
    parser.add_argument("--datoformat", choices=["dmy", "mdy"],
                        help="tving dag/måned-rækkefølge")
    parser.add_argument("--fortegn", choices=["auto", "negativ", "positiv"],
                        default="auto",
                        help="hvilket fortegn der betyder udgift (standard: auto)")
    parser.add_argument("--interaktiv", action="store_true",
                        help="spørg om kategori for ukendte posteringer")
    parser.add_argument("--kun-rapport", action="store_true",
                        help="skriv ingen filer, vis kun oversigten")
    return parser.parse_args(argv)


def load_ruleset(path):
    if path and os.path.isfile(path):
        return RuleSet.load(path)
    default_path = user_rules_path()
    if not path and os.path.isfile(default_path):
        return RuleSet.load(default_path)
    return RuleSet.defaults()


def interactive_pass(result, ruleset):
    """Ask once per unknown text and remember the answer as a new rule."""
    if not result.uncategorised:
        return 0
    print("\n%d ukendte forretninger. Tryk Enter for at springe over, "
          "eller skriv et kategorinavn." % len(result.uncategorised))
    known = ruleset.categories()
    print("Kendte kategorier: %s" % ", ".join(known))
    added = 0
    for text, count, total, keyword in result.uncategorised:
        try:
            answer = input("  %-45s (%d stk., %.2f) > " % (text[:45], count, total))
        except (EOFError, KeyboardInterrupt):
            print()
            break
        answer = answer.strip()
        if not answer:
            continue
        ruleset.add(keyword or text, answer)
        added += 1
    if added:
        for transaction in result.transactions:
            if transaction.category == DEFAULT_CATEGORY:
                transaction.category = ruleset.categorise(transaction.text)
                transaction.kind = (KIND_EXPENSE if transaction.amount < 0
                                    else KIND_INCOME)
    return added


def write_transactions(path, transactions):
    with open(path, "w", newline="", encoding="utf-8-sig") as fp:
        writer = csv.writer(fp, delimiter=";")
        writer.writerow(["Dato", "Tekst", "Beløb", "Kategori", "Type", "Måned"])
        for t in transactions:
            writer.writerow([t.date.strftime("%d-%m-%Y"), t.text,
                             ("%.2f" % t.amount).replace(".", ","),
                             t.category, t.kind, t.month])


def write_summary(path, summary):
    with open(path, "w", newline="", encoding="utf-8-sig") as fp:
        writer = csv.writer(fp, delimiter=";")
        writer.writerow(["Type", "Kategori"] + [month_label(m) for m in summary.months]
                        + ["I alt", "Gns./md"])
        for kind, categories in ((KIND_EXPENSE, summary.expense_categories),
                                 (KIND_INCOME, summary.income_categories)):
            for category in categories:
                values = [abs(summary.value(kind, category, m)) for m in summary.months]
                total = sum(values)
                writer.writerow(
                    [kind, category] +
                    [("%.2f" % v).replace(".", ",") for v in values] +
                    [("%.2f" % total).replace(".", ","),
                     ("%.2f" % (total / summary.month_count())).replace(".", ",")])
        writer.writerow([])
        for kind in (KIND_EXPENSE, KIND_INCOME):
            totals = summary.month_total(kind)
            values = [abs(totals[m]) for m in summary.months]
            writer.writerow(["I alt", kind] +
                            [("%.2f" % v).replace(".", ",") for v in values] +
                            [("%.2f" % sum(values)).replace(".", ","), ""])


def print_report(result, summary, table, mapping):
    print("Fil:      %s" % table.describe())
    print("Kolonner: %s" % "; ".join(mapping.describe(table)))
    print()
    for line in result.summary_lines():
        print(line)
    if not summary.months:
        return
    print()
    width = max([len(c) for c in summary.expense_categories +
                 summary.income_categories] + [12])
    header = "%-*s" % (width, "Kategori")
    for month in summary.months:
        header += "%14s" % month
    header += "%14s%14s" % ("I alt", "Gns./md")
    print(header)
    print("-" * len(header))
    for kind, categories in ((KIND_EXPENSE, summary.expense_categories),
                             (KIND_INCOME, summary.income_categories)):
        if not categories:
            continue
        print(kind.upper())
        for category in categories:
            line = "%-*s" % (width, category)
            total = 0.0
            for month in summary.months:
                value = abs(summary.value(kind, category, month))
                total += value
                line += "%14s" % _kr(value)
            line += "%14s%14s" % (_kr(total), _kr(total / summary.month_count()))
            print(line)
    print("-" * len(header))
    print("%-*s%s" % (width, "Netto", "%14s" % _kr(summary.net)))
    if result.uncategorised:
        print("\nStørste ukategoriserede forretninger:")
        for text, count, total, _keyword in result.uncategorised[:10]:
            print("  %-45s %d stk. %12s" % (text[:45], count, _kr(total)))


def _kr(value):
    text = "{:,.2f}".format(value)
    return text.replace(",", " ").replace(".", ",") + " kr."


SIGNS = {"auto": SIGN_AUTO, "negativ": SIGN_NEGATIVE_IS_EXPENSE,
         "positiv": SIGN_POSITIVE_IS_EXPENSE}


def main(argv=None):
    args = parse_args(argv)
    if not os.path.isfile(args.fil):
        print("Filen findes ikke: %s" % args.fil, file=sys.stderr)
        return 2

    table = csvsniff.read_table(path=args.fil, encoding=args.tegnsaet,
                                delimiter=args.skilletegn)
    if not table.rows:
        print("Fandt ingen datarækker i filen.", file=sys.stderr)
        return 1

    mapping = detect_mapping(table)
    if args.decimal:
        mapping.decimal = args.decimal
    if args.datoformat:
        mapping.dayfirst = args.datoformat == "dmy"

    ruleset = load_ruleset(args.regler)
    options = BuildOptions(decimal=mapping.decimal, dayfirst=mapping.dayfirst,
                           sign=SIGNS[args.fortegn])
    result = build_transactions(table, mapping, ruleset, options)
    if not result.ok:
        print("Kunne ikke lave et budget:", file=sys.stderr)
        for line in result.summary_lines():
            print("  " + line, file=sys.stderr)
        return 1

    if args.interaktiv:
        if interactive_pass(result, ruleset):
            result.uncategorised = [
                item for item in result.uncategorised
                if ruleset.categorise(item[0]) == DEFAULT_CATEGORY]


    summary = summarise(result.transactions)
    print_report(result, summary, table, mapping)

    if args.gem_regler:
        path = args.regler or user_rules_path()
        ruleset.save(path)
        print("\nRegler gemt i %s" % path)

    if args.kun_rapport:
        return 0

    directory = args.ud or os.path.dirname(os.path.abspath(args.fil))
    if not os.path.isdir(directory):
        os.makedirs(directory)
    stem = os.path.splitext(os.path.basename(args.fil))[0]
    transactions_path = os.path.join(directory, "%s_kategoriseret.csv" % stem)
    summary_path = os.path.join(directory, "%s_budget.csv" % stem)
    write_transactions(transactions_path, result.transactions)
    write_summary(summary_path, summary)
    print("\nSkrev %s" % transactions_path)
    print("Skrev %s" % summary_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
