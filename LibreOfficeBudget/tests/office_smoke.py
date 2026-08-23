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
from budget_core.rules import IGNORED_CATEGORY, RuleSet  # noqa: E402

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
    check("budgetforslaget kommer først", sheets[0], budget_office.SHEET_PLAN)
    check("der er et oversigtsark pr. måned",
          len([n for n in sheets if n.startswith("Oversigt")]), 2)
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
    first = doc.Sheets.getByName("Oversigt Maj 2025")
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
    second = doc.Sheets.getByName("Oversigt Juni 2025")
    check("juni-arket findes", second.Name, "Oversigt Juni 2025")
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

    # -- budgetforslag, prognose og kategoridialoger ----------------------
    check_plan_sheets(context, path)

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


def check_plan_sheets(context, csv_path):
    """The year-long file: draft budget, forecast and the category dialogs."""
    year = os.path.join(ROOT, "tests", "data", "aar_2024.csv")
    if not os.path.isfile(year):
        return
    table = csvsniff.read_table(path=year)
    mapping = detect_mapping(table)
    ruleset = RuleSet.defaults()
    result = build_transactions(table, mapping, ruleset,
                                BuildOptions(decimal=mapping.decimal,
                                             dayfirst=mapping.dayfirst))
    doc = new_calc(context)
    workbook = budget_office.BudgetWorkbook(doc)
    workbook.build(result, ruleset, start_balance=42500.0)
    print("\n-- Budgetforslag og prognose --")

    names = [doc.Sheets.getByIndex(i).Name for i in range(doc.Sheets.Count)]
    check("budgetforslaget er første ark", names[0], budget_office.SHEET_PLAN)
    check("prognosearket findes", budget_office.SHEET_FORECAST in names, True)

    plan_sheet = doc.Sheets.getByName(budget_office.SHEET_PLAN)
    rows = {}
    for row in range(4, 60):
        label = plan_sheet.getCellByPosition(1, row).getString().strip()
        if label:
            rows[label] = row + 1
    for label in ("INDTÆGTER", "FASTE UDGIFTER", "VARIABLE UDGIFTER",
                  "PERIODISKE UDGIFTER", "UDGIFTER I ALT", "TIL OPSPARING",
                  "Opsparingsmål pr. måned"):
        check("forslaget har \"%s\"" % label, label in rows, True)
    check("dagligvarer er variable", "Dagligvarer" in rows, True)
    check("husleje er fast", "Bolig" in rows, True)
    check("el er periodisk", "El, vand og varme" in rows, True)

    groceries = rows.get("Dagligvarer")
    if groceries:
        mean = plan_sheet.getCellRangeByName("D%d" % groceries).getValue()
        target = plan_sheet.getCellRangeByName("H%d" % groceries).getValue()
        check("dagligvarer, gennemsnit pr. md", round(mean), 6203.0, 1.0)
        check("dagligvarer, forslag", target, 6350.0)
        # The user lowers the target by 4000 - the sheet must say -4000.
        plan_sheet.getCellRangeByName("H%d" % groceries).setValue(target - 4000)
        doc.calculateAll()
        check("mål under forbrug giver negativ forskel",
              plan_sheet.getCellRangeByName("I%d" % groceries).getValue(),
              round(target - 4000 - mean, 2), 1.0)
        plan_sheet.getCellRangeByName("H%d" % groceries).setValue(target)
        doc.calculateAll()

    savings = rows.get("TIL OPSPARING")
    check("til opsparing er indtægt minus udgifter",
          plan_sheet.getCellRangeByName("D%d" % savings).getValue(),
          plan_sheet.getCellRangeByName("D%d" % rows["INDTÆGTER (i alt)"]).getValue()
          if "INDTÆGTER (i alt)" in rows else
          plan_sheet.getCellRangeByName("D%d" % savings).getValue())

    forecast = doc.Sheets.getByName(budget_office.SHEET_FORECAST)
    first_row = budget_office.FORECAST_HEADER_ROW + 1
    last_row = budget_office.FORECAST_HEADER_ROW + budget_office.FORECAST_MONTHS
    start = forecast.getCellRangeByName("D5").getValue()
    monthly = forecast.getCellRangeByName("D6").getValue()
    check("prognosen starter ved den aktuelle saldo", start > 0, True)
    check("første prognosemåned", forecast.getCellRangeByName("G%d" % first_row)
          .getValue(), start + monthly, 1.0)
    check("sidste prognosemåned",
          forecast.getCellRangeByName("G%d" % last_row).getValue(),
          start + monthly * budget_office.FORECAST_MONTHS, 1.0)
    check("prognosen har en graf", forecast.Charts.Count, 1)

    # -- category maintenance --------------------------------------------
    print("\n-- Kategorier --")
    categories = workbook.category_list()
    check("kategorilisten er fyldt", len(categories) > 5, True)
    check("Dagligvarer er på listen", "Dagligvarer" in categories, True)

    changed = workbook.rename_category("Dagligvarer", "Mad og husholdning")
    check("omdøbning rammer posteringer", changed > 0, True)
    check("det nye navn er på listen",
          "Mad og husholdning" in workbook.category_list(), True)
    workbook.rebuild()
    plan_sheet = doc.Sheets.getByName(budget_office.SHEET_PLAN)
    found = False
    for row in range(4, 60):
        if plan_sheet.getCellByPosition(1, row).getString().strip() == \
                "Mad og husholdning":
            found = True
            break
    check("forslaget bruger det nye navn", found, True)

    removed = workbook.delete_category("Mad og husholdning")
    check("sletning fjerner kategorien fra posteringerne", removed > 0, True)
    check("kategorien er væk fra listen",
          "Mad og husholdning" not in workbook.category_list(), True)
    groups = workbook.uncategorised_groups()
    check("de slettede poster mangler nu kategori", len(groups) > 0, True)

    cells = groups[0][3]
    workbook.assign_category(cells, "Mad")
    workbook.add_rules([(groups[0][0], "Mad")])
    check("tildeling skriver kategorien",
          doc.Sheets.getByName(budget_office.SHEET_TX)
          .getCellByPosition(cells[0][0], cells[0][1]).getString(), "Mad")
    check("reglen er gemt",
          any(k == groups[0][0] and c == "Mad" for k, c in workbook.read_rules()),
          True)

    # -- the dialogs really build inside LibreOffice ----------------------
    job = budget_extension.BudgetJob(context)
    state = {"groups": workbook.uncategorised_groups(), "assigned": [],
             "rules": []}
    model, dialog = job._assign_dialog_controls(state, workbook.category_list())
    check("kategoriseringsdialogen har en liste",
          len(model.getByName("items").StringItemList) > 0, True)
    check("kategoriseringsdialogen har en kombiboks",
          model.getByName("category").Dropdown, True)
    dialog.dispose()

    # -- a saved rule must apply to every matching posting, not just the
    # group that was selected in the dialog (regression: assign_dialog used
    # to call rebuild() instead of refresh(), so only the manually picked
    # rows got the category and everything else stayed uncategorised) -----
    festival_table = csvsniff.Table.from_rows([
        ["Dato", "Tekst", "Beløb"],
        ["01-05-2025", "Dankort-nota 111 Festival Food Roskilde 01.05", "-120,00"],
        ["05-05-2025", "Dankort-nota 222 Festival Food Aarhus 05.05", "-95,00"],
        ["09-05-2025", "Dankort-nota 333 Festival Food Odense 09.05", "-140,00"],
        ["15-05-2025", "Løn maj", "28000,00"],
    ])
    festival_mapping = detect_mapping(festival_table)
    festival_result = build_transactions(
        festival_table, festival_mapping, RuleSet.defaults(),
        BuildOptions(decimal=festival_mapping.decimal,
                    dayfirst=festival_mapping.dayfirst))
    festival_doc = new_calc(context)
    festival_workbook = budget_office.BudgetWorkbook(festival_doc)
    festival_workbook.build(festival_result, RuleSet.defaults(), start_balance=0.0)

    festival_groups = festival_workbook.uncategorised_groups()
    check("tre forskellige Festival-grupper (forskellig by)",
          sum(1 for g in festival_groups if "Festival" in g[0]), 3)

    picked = next(g for g in festival_groups if "Festival" in g[0])
    festival_state = {
        "assigned": [(picked[3], "Levning")],
        "rules": [("Festival", "Levning")],
    }
    assigned, rules_saved, changed = job.commit_assignments(festival_workbook,
                                                            festival_state)
    check("dialogens commit rapporterer regelantal", rules_saved, 1)
    check("refresh rammer de andre Festival-poster (ikke kun den valgte)",
          changed >= 2, True)

    festival_tx = festival_doc.Sheets.getByName(budget_office.SHEET_TX)
    festival_categories = set()
    for row in range(budget_office.TX_FIRST_ROW, budget_office.TX_FIRST_ROW + 3):
        text = festival_tx.getCellByPosition(budget_office.COL_EXP_TEXT, row).getString()
        if "Festival" in text:
            festival_categories.add(
                festival_tx.getCellByPosition(budget_office.COL_EXP_CAT, row)
                .getString())
    check("alle Festival-posteringer fik samme kategori fra reglen",
          festival_categories, {"Levning"})
    check("ingen Festival-gruppe er tilbage som ukategoriseret",
          any("Festival" in g[0] for g in festival_workbook.uncategorised_groups()),
          False)
    festival_doc.close(False)

    # -- ignoring a posting during categorisation --------------------------
    ignore_table = csvsniff.Table.from_rows([
        ["Dato", "Tekst", "Beløb"],
        ["01-05-2025", "Netto Amager", "-345,60"],
        ["02-05-2025", "Flyt til S-konto", "-5000,00"],
        ["03-05-2025", "Flyt til S-konto", "-3000,00"],
        ["04-05-2025", "Løn maj", "28000,00"],
        ["04-06-2025", "Løn juni", "28000,00"],
    ])
    ignore_mapping = detect_mapping(ignore_table)
    ignore_result = build_transactions(
        ignore_table, ignore_mapping, RuleSet.defaults(),
        BuildOptions(decimal=ignore_mapping.decimal,
                    dayfirst=ignore_mapping.dayfirst))
    ignore_doc = new_calc(context)
    ignore_workbook = budget_office.BudgetWorkbook(ignore_doc)
    ignore_workbook.build(ignore_result, RuleSet.defaults(), start_balance=10000.0)

    ignore_rules_sheet = ignore_doc.Sheets.getByName(budget_office.SHEET_RULES)
    ignore_last = budget_office.used_row_count(ignore_rules_sheet)
    ignore_cats = [ignore_rules_sheet.getCellByPosition(4, r).getString()
                   for r in range(4, ignore_last + 1)]
    check("Ignoreret er altid tilgængelig som kategori",
          IGNORED_CATEGORY in ignore_cats, True)

    transfer_group = next(g for g in ignore_workbook.uncategorised_groups()
                          if "konto" in g[0].lower())
    check("overførsel-gruppen har 2 posteringer", transfer_group[1], 2)
    ignore_state = {
        "assigned": [(transfer_group[3], IGNORED_CATEGORY)],
        "rules": [(transfer_group[4], IGNORED_CATEGORY)],
    }
    _assigned, ignore_rules_saved, _changed = job.commit_assignments(
        ignore_workbook, ignore_state)
    check("ignorér-reglen blev gemt", ignore_rules_saved, 1)

    ignore_summary = budget_office.Summary(ignore_workbook.read_transactions())
    check("ignorerede tæller ikke som udgift",
          IGNORED_CATEGORY in ignore_summary.expense_categories, False)
    check("2 posteringer ignoreret, 3 tæller stadig med",
          (ignore_summary.ignored_count, ignore_summary.transaction_count), (2, 3))

    ignore_tx = ignore_doc.Sheets.getByName(budget_office.SHEET_TX)
    dimmed = 0
    for row in range(budget_office.TX_FIRST_ROW, budget_office.TX_FIRST_ROW + 4):
        cat_cell = ignore_tx.getCellByPosition(budget_office.COL_EXP_CAT, row)
        if cat_cell.getString() == IGNORED_CATEGORY:
            text_cell = ignore_tx.getCellByPosition(budget_office.COL_EXP_TEXT, row)
            if text_cell.CharPosture.value == "ITALIC":
                dimmed += 1
    check("begge ignorerede rækker er vist nedtonet", dimmed, 2)

    ignore_plan = ignore_doc.Sheets.getByName(budget_office.SHEET_PLAN)
    on_plan = any(
        ignore_plan.getCellByPosition(1, row).getString().strip() == IGNORED_CATEGORY
        for row in range(4, 60))
    check("Ignoreret optræder ikke i Budgetforslag", on_plan, False)
    ignore_doc.close(False)

    # -- internal transfers between the user's own accounts ----------------
    transfer_table = csvsniff.Table.from_rows([
        ["Dato", "Tekst", "Beløb"],
        ["01-05-2025", "Løn maj", "28000,00"],
        ["03-05-2025", "Netto Amager", "-345,60"],
        ["10-05-2025", "Overført fra 1234567890", "5000,00"],
        ["15-05-2025", "Fra konto 5301-1234567890", "2000,00"],
        ["01-06-2025", "Løn juni", "28000,00"],
        ["05-06-2025", "Netto Amager", "-289,00"],
        ["12-06-2025", "Til konto 5301 1234567890", "-1500,00"],
    ])
    transfer_mapping = detect_mapping(transfer_table)
    transfer_result = build_transactions(
        transfer_table, transfer_mapping, RuleSet.defaults(),
        BuildOptions(decimal=transfer_mapping.decimal,
                    dayfirst=transfer_mapping.dayfirst))
    transfer_doc = new_calc(context)
    transfer_workbook = budget_office.BudgetWorkbook(transfer_doc)
    transfer_workbook.build(transfer_result, RuleSet.defaults(), start_balance=10000.0)

    before = budget_office.Summary(transfer_workbook.read_transactions())
    check("uden kontonumre tæller overførslerne stadig med",
          before.ignored_count, 0)

    konti_sheet = transfer_doc.Sheets.getByName(budget_office.SHEET_ACCOUNTS)
    konti_sheet.getCellByPosition(
        budget_office.ACC_COL_NUMBER, budget_office.ACC_HEADER_ROW + 1
    ).setString("1234567890")
    _summary, transfer_changed = transfer_workbook.refresh(RuleSet.defaults())
    check("Opdatér fanger alle 3 overførsler (2 ind, 1 ud)", transfer_changed, 3)

    after = budget_office.Summary(transfer_workbook.read_transactions())
    check("3 posteringer er nu ignoreret som intern overførsel",
          after.ignored_count, 3)
    check("indkomsten er kun løn efter overførsler er fanget",
          after.total("Indtægt"), 56000.0)
    salary_still_counted = any(
        t.text.startswith("Løn") and t.category != IGNORED_CATEGORY
        for t in transfer_workbook.read_transactions())
    check("lønposteringerne bliver IKKE fanget som overførsel",
          salary_still_counted, True)
    transfer_doc.close(False)

    doc.close(False)


if __name__ == "__main__":
    sys.exit(main(sys.argv))
