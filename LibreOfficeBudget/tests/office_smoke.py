# -*- coding: utf-8 -*-
"""End to end check of the Calc part against a running LibreOffice.

Start LibreOffice first:

    soffice --headless --norestore \
            --accept="socket,host=localhost,port=2002;urp;" &

then run:

    python3 tests/office_smoke.py [file.csv]

It builds the workbook exactly like the extension does, verifies the
formulas actually calculate, and saves the result next to the CSV file.
"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "office"))
sys.path.insert(0, os.path.join(ROOT, "oxt"))

import uno  # noqa: E402
import unohelper  # noqa: E402

import budget_extension  # noqa: E402
import budget_office  # noqa: E402
from budget_core import csvsniff  # noqa: E402
from budget_core.budget import (BuildOptions, KIND_EXPENSE,  # noqa: E402
                                build_transactions)
from budget_core.columns import detect_mapping  # noqa: E402
from budget_core.rules import RuleSet  # noqa: E402

FAILURES = []


def check(label, got, expected, tolerance=0.005):
    if isinstance(expected, float):
        ok = got is not None and abs(float(got) - expected) <= tolerance
    else:
        ok = got == expected
    print("%-4s %-52s %r%s" % ("ok" if ok else "FAIL", label, got,
                               "" if ok else " (forventet %r)" % (expected,)))
    if not ok:
        FAILURES.append(label)
    return ok


def connect(port=2002):
    context = uno.getComponentContext()
    resolver = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", context)
    url = ("uno:socket,host=localhost,port=%d;urp;"
           "StarOffice.ComponentContext" % port)
    return resolver.resolve(url)


def new_calc(context):
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context)
    return desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())


def main(argv):
    path = argv[1] if len(argv) > 1 else os.path.join(ROOT, "tests", "data",
                                                      "danskebank.csv")
    context = connect()
    table = csvsniff.read_table(path=path)
    mapping = detect_mapping(table)
    ruleset = RuleSet.defaults()
    result = build_transactions(table, mapping, ruleset,
                                BuildOptions(decimal=mapping.decimal,
                                             dayfirst=mapping.dayfirst))
    print("%d posteringer, kolonner: %s" % (len(result.transactions),
                                            "; ".join(mapping.describe(table))))

    doc = new_calc(context)
    workbook = budget_office.BudgetWorkbook(doc)
    summary = workbook.build(result, ruleset, start_balance=12795.85)

    sheets = [doc.Sheets.getByIndex(i).Name for i in range(doc.Sheets.Count)]
    print("Ark: %s" % ", ".join(sheets))
    check("første ark er en oversigt", sheets[0].startswith("Oversigt"), True)
    check("transaktionsarket findes", budget_office.SHEET_TX in sheets, True)
    check("kategoriarket findes", budget_office.SHEET_RULES in sheets, True)

    # -- transactions ----------------------------------------------------
    tx = doc.Sheets.getByName(budget_office.SHEET_TX)
    check("overskrift B4", tx.getCellRangeByName("B4").getString(), "Dato")
    check("første udgift, dato", tx.getCellRangeByName("B5").getString(),
          "02-05-2025")
    check("første udgift, beløb", tx.getCellRangeByName("C5").getValue(), 345.60)
    check("første udgift, kategori", tx.getCellRangeByName("E5").getString(),
          "Dagligvarer")
    check("første indtægt, beløb", tx.getCellRangeByName("H5").getValue(), 28500.0)
    check("beløb vises i kroner",
          tx.getCellRangeByName("C5").getString().endswith("kr."), True)

    # -- summary ---------------------------------------------------------
    first = doc.Sheets.getByName(sheets[0])
    check("titel", first.getCellRangeByName("B8").getString(), "Budget - Maj 2025")
    check("startsaldo", first.getCellRangeByName("L8").getValue(), 12795.85)

    groceries_row = None
    for row in range(26, 60):
        if first.getCellByPosition(1, row).getString() == "Dagligvarer":
            groceries_row = row + 1
            break
    check("Dagligvarer står i udgiftstabellen", groceries_row is not None, True)
    if groceries_row:
        actual = first.getCellRangeByName("E%d" % groceries_row).getValue()
        check("faktisk forbrug, Dagligvarer (maj)", actual, 345.60)
        planned = first.getCellRangeByName("D%d" % groceries_row).getValue()
        check("budgetforslag udfyldt", planned > 0, True)
        diff = first.getCellRangeByName("F%d" % groceries_row).getValue()
        check("difference = budget - faktisk", diff, planned - actual)

    expenses = first.getCellRangeByName("C22").getValue()
    income = first.getCellRangeByName("I22").getValue()
    check("udgifter i alt (maj)", expenses,
          abs(sum(summary.value(KIND_EXPENSE, c, "2025-05")
                  for c in summary.expense_categories)))
    check("indtægter i alt (maj)", income, 28500.0)
    check("slutsaldo", first.getCellRangeByName("E17").getValue(),
          12795.85 + income - expenses)
    check("søjle vises", "▉" in first.getCellRangeByName("D22").getString(), True)

    # -- second month ----------------------------------------------------
    second = doc.Sheets.getByIndex(1)
    check("andet ark er juni", second.Name, "Oversigt Juni 2025")
    check("startsaldo overføres", second.getCellRangeByName("L8").getValue(),
          first.getCellRangeByName("E17").getValue())
    check("juni filtreres på periode",
          second.getCellRangeByName("C22").getValue(),
          abs(sum(summary.value(KIND_EXPENSE, c, "2025-06")
                  for c in summary.expense_categories)))

    months = doc.Sheets.getByName(budget_office.SHEET_MONTHS)
    check("månedsark har overskrift", months.getCellRangeByName("B2").getString(),
          "Alle måneder")
    check("månedsark viser maj", months.getCellRangeByName("C4").getString(),
          "maj 2025")

    print("\nVisning: startsaldo=%r  udgifter=%r  slutsaldo=%r  diff=%r" % (
        first.getCellRangeByName("L8").getString(),
        first.getCellRangeByName("C22").getString(),
        first.getCellRangeByName("E17").getString(),
        first.getCellRangeByName("F%d" % groceries_row).getString()
        if groceries_row else ""))
    print("Transaktion: dato=%r beløb=%r" % (
        tx.getCellRangeByName("B5").getString(),
        tx.getCellRangeByName("C5").getString()))

    # -- refresh ---------------------------------------------------------
    rules_sheet = doc.Sheets.getByName(budget_office.SHEET_RULES)
    rules_sheet.getCellRangeByName("B3").setString("hotel skagen")
    rules_sheet.getCellRangeByName("C3").setString("Ferie")
    # (row 3 is above the table - use the first free row instead)
    rules_sheet.getCellRangeByName("B3").setString("")
    rules_sheet.getCellRangeByName("C3").setString("")
    last = budget_office.used_row_count(rules_sheet)
    rules_sheet.getCellByPosition(1, last + 1).setString("hotel skagen")
    rules_sheet.getCellByPosition(2, last + 1).setString("Ferie")
    rows = workbook.read_rules()
    summary2, changed = workbook.refresh(RuleSet(rows))
    check("opdatering ændrede en kategori", changed >= 1, True)
    june = doc.Sheets.getByName("Oversigt Juni 2025")
    ferie_row = None
    for row in range(26, 60):
        if june.getCellByPosition(1, row).getString() == "Ferie":
            ferie_row = row + 1
            break
    check("ny kategori er med i oversigten", ferie_row is not None, True)
    if ferie_row:
        check("beløb flyttet til Ferie",
              june.getCellRangeByName("E%d" % ferie_row).getValue(), 2450.0)

    # -- import dialog ---------------------------------------------------
    def create(name):
        return context.ServiceManager.createInstanceWithContext(name, context)

    model = budget_extension.build_options_model(create, table, mapping,
                                                 os.path.basename(path))
    check("dialogen har alle felter",
          all(model.hasByName(n) for n in ("date", "text1", "amount",
                                           "amount_in", "decimal", "dayfirst",
                                           "sign", "suggest", "ok", "cancel")),
          True)
    check("datokolonnen er valgt på forhånd",
          model.getByName("date").SelectedItems[0], mapping.date + 1)
    check("knappen er en OK-knap", model.getByName("ok").PushButtonType, 1)
    settings = budget_extension.read_options_model(model)
    check("dialogen giver samme kolonnevalg", settings["date"], mapping.date)
    check("dialogen giver samme beløbskolonne", settings["amount"],
          mapping.amount)
    check("dialogen giver samme decimaltegn", settings["decimal"],
          mapping.decimal)

    # -- save ------------------------------------------------------------
    out = os.path.join(os.path.dirname(path), "budget_smoke.ods")
    url = unohelper.systemPathToFileUrl(out)
    doc.storeToURL(url, ())
    print("\nGemte %s" % out)
    doc.close(False)

    if FAILURES:
        print("\n%d fejl: %s" % (len(FAILURES), ", ".join(FAILURES)))
        return 1
    print("\nAlt OK")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
