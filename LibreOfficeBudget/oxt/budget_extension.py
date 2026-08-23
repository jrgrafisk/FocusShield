# -*- coding: utf-8 -*-
"""LibreOffice extension: Budget fra CSV.

Adds a "Budget" menu to LibreOffice with commands that turn a bank CSV
export into a finished budget workbook (Oversigt + Transaktioner), and that
re-apply the category rules afterwards.

The heavy lifting lives in ``budget_core`` (pure Python, no UNO) and
``budget_office`` (writes the Calc document).  This file is only the glue:
menu dispatch, file picker, options dialog and error reporting.
"""

from __future__ import unicode_literals

import os
import sys
import traceback

import uno
import unohelper

from com.sun.star.awt import XActionListener, XItemListener
from com.sun.star.task import XJobExecutor

# ---------------------------------------------------------------------------
# Make the bundled pythonpath importable even if the loader did not do it.
# ---------------------------------------------------------------------------
def _bootstrap_path():
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except NameError:  # pragma: no cover
        return
    candidate = os.path.join(here, "pythonpath")
    if os.path.isdir(candidate) and candidate not in sys.path:
        sys.path.insert(0, candidate)


_bootstrap_path()

from budget_core import csvsniff, columns as column_detect          # noqa: E402
from budget_core.budget import (BuildOptions, SIGN_AUTO,            # noqa: E402
                                SIGN_NEGATIVE_IS_EXPENSE,
                                SIGN_POSITIVE_IS_EXPENSE,
                                build_transactions)
from budget_core.parsing import parse_amount                        # noqa: E402
from budget_core.rules import RuleSet, user_rules_path              # noqa: E402
import budget_office                                                # noqa: E402

IMPLEMENTATION_NAME = "dk.jrgrafisk.budgetfracsv.BudgetJob"
SERVICE_NAME = "dk.jrgrafisk.budgetfracsv.BudgetService"

NONE_LABEL = "(ingen)"
SIGN_CHOICES = ((SIGN_AUTO, "Automatisk"),
                (SIGN_NEGATIVE_IS_EXPENSE, "Negative beløb er udgifter"),
                (SIGN_POSITIVE_IS_EXPENSE, "Positive beløb er udgifter"))


class BudgetJob(unohelper.Base, XJobExecutor):
    """Entry point for every menu command of the extension."""

    def __init__(self, ctx):
        self.ctx = ctx
        self.smgr = ctx.ServiceManager
        # UNO objects cannot hold Python attributes, so dialog listeners are
        # kept alive here for as long as the job lives.
        self._listeners = []

    # -- dispatch ---------------------------------------------------------
    def trigger(self, args):
        action = (args or "").split("?")[0].strip().lower()
        try:
            if action in ("", "import"):
                self.import_csv()
            elif action == "sheet":
                self.import_active_sheet()
            elif action in ("refresh", "recalc"):
                self.refresh()
            elif action == "assign":
                self.assign_dialog()
            elif action == "categories":
                self.categories_dialog()
            elif action == "rules":
                self.show_rules()
            elif action == "resetrules":
                self.reset_rules()
            elif action == "accounts":
                self.show_accounts()
            elif action == "bankbudget":
                self.bank_budget()
            elif action == "about":
                self.about()
            elif action.startswith("headless"):
                # Used by the test suite: import without any dialog.
                self.import_headless((args or "").split(":", 1)[1])
            else:
                self.message("Budget", "Ukendt kommando: %s" % action)
        except Exception:
            self.report_error(traceback.format_exc())

    # -- services ---------------------------------------------------------
    def create(self, name):
        return self.smgr.createInstanceWithContext(name, self.ctx)

    @property
    def desktop(self):
        return self.create("com.sun.star.frame.Desktop")

    def parent_window(self):
        try:
            frame = self.desktop.getCurrentFrame()
            if frame:
                return frame.getContainerWindow()
        except Exception:
            pass
        return self.create("com.sun.star.awt.Toolkit").getDesktopWindow()

    def message(self, title, text, kind="INFOBOX"):
        toolkit = self.create("com.sun.star.awt.Toolkit")
        box = toolkit.createMessageBox(
            self.parent_window(), uno.Enum("com.sun.star.awt.MessageBoxType", kind),
            1,  # BUTTONS_OK
            title, text)
        try:
            box.execute()
        finally:
            box.dispose()

    def ask(self, title, text):
        toolkit = self.create("com.sun.star.awt.Toolkit")
        box = toolkit.createMessageBox(
            self.parent_window(),
            uno.Enum("com.sun.star.awt.MessageBoxType", "QUERYBOX"),
            2,  # BUTTONS_OK_CANCEL
            title, text)
        try:
            return box.execute() == 1
        finally:
            box.dispose()

    def report_error(self, details):
        path = os.path.join(_log_directory(), "budget-fra-csv-fejl.txt")
        try:
            with open(path, "w") as handle:
                handle.write(details)
        except Exception:
            path = "(kunne ikke gemmes)"
        first = [line for line in details.strip().splitlines() if line.strip()]
        self.message("Budget fra CSV - fejl",
                     "Der opstod en fejl:\n\n%s\n\nDetaljer er gemt i:\n%s"
                     % (first[-1] if first else "Ukendt fejl", path),
                     kind="ERRORBOX")

    # -- commands ---------------------------------------------------------
    def import_csv(self):
        path = self.pick_file()
        if not path:
            return
        try:
            table = csvsniff.read_table(path=path)
        except Exception as exc:
            self.message("Budget fra CSV", "Kunne ikke læse filen:\n%s" % exc,
                         kind="ERRORBOX")
            return
        self._import_table(table, os.path.basename(path))

    def import_active_sheet(self):
        doc = self.desktop.getCurrentComponent()
        if not doc or not hasattr(doc, "Sheets"):
            self.message("Budget fra CSV",
                         "Åbn først et regneark med posteringerne.")
            return
        sheet = doc.CurrentController.getActiveSheet()
        rows = budget_office.read_sheet_as_table(doc, sheet)
        if not rows:
            self.message("Budget fra CSV", "Arket er tomt.")
            return
        table = csvsniff.Table.from_rows(rows, source=sheet.Name)
        self._import_table(table, "arket \"%s\"" % sheet.Name)

    def _import_table(self, table, source_name):
        if not table.rows:
            self.message("Budget fra CSV",
                         "Fandt ingen datarækker i %s." % source_name)
            return
        mapping = column_detect.detect_mapping(table)
        settings = self.options_dialog(table, mapping, source_name)
        if settings is None:
            return

        mapping.date = settings["date"]
        mapping.text = settings["text"]
        mapping.amount = settings["amount"]
        mapping.amount_in = settings["amount_in"]
        mapping.amount_out = settings["amount_out"]
        if mapping.amount_in is not None or mapping.amount_out is not None:
            mapping.amount = None
        mapping.decimal = settings["decimal"]
        mapping.dayfirst = settings["dayfirst"]

        ruleset = load_rules()
        options = BuildOptions(decimal=settings["decimal"],
                               dayfirst=settings["dayfirst"],
                               sign=settings["sign"])
        result = build_transactions(table, mapping, ruleset, options)
        if not result.ok:
            self.message("Budget fra CSV",
                         "Der kunne ikke laves et budget.\n\n%s"
                         % "\n".join(result.summary_lines()), kind="ERRORBOX")
            return

        start_balance = _start_balance(table, mapping, result)
        doc = self.new_calc()
        workbook = budget_office.BudgetWorkbook(doc)
        summary = workbook.build(result, ruleset, start_balance=start_balance,
                                 suggest_planned=settings["suggest"])
        self.message("Budget fra CSV", _done_text(result, summary, table))

    def import_headless(self, path):
        """Import a file with the detected settings and no dialogs at all."""
        table = csvsniff.read_table(path=path.strip())
        mapping = column_detect.detect_mapping(table)
        ruleset = load_rules()
        result = build_transactions(table, mapping, ruleset,
                                    BuildOptions(decimal=mapping.decimal,
                                                 dayfirst=mapping.dayfirst))
        if not result.ok:
            raise RuntimeError("; ".join(result.summary_lines()))
        doc = self.new_calc()
        budget_office.BudgetWorkbook(doc).build(
            result, ruleset, start_balance=_start_balance(table, mapping, result))
        return doc

    def refresh(self):
        doc = self.desktop.getCurrentComponent()
        if not doc or not hasattr(doc, "Sheets"):
            self.message("Budget fra CSV", "Åbn budgettet først.")
            return
        workbook = budget_office.BudgetWorkbook(doc)
        if not workbook.has_budget():
            self.message("Budget fra CSV",
                         "Dette dokument indeholder ikke et ark ved navn "
                         "\"%s\"." % budget_office.SHEET_TX)
            return
        rows = workbook.read_rules()
        ruleset = RuleSet(rows) if rows else load_rules()
        summary, changed = workbook.refresh(ruleset)
        if summary is None:
            self.message("Budget fra CSV", "Fandt ingen posteringer at opdatere.")
            return
        if rows:
            try:
                RuleSet(rows).save(user_rules_path())
            except Exception:
                pass
        self.message("Budget fra CSV",
                     "Budgettet er opdateret.\n\n"
                     "%d postering(er) fik ny kategori.\n"
                     "%d måned(er) i budgettet."
                     % (changed, len(summary.months)))

    def _workbook(self):
        """The BudgetWorkbook for the current document, or None."""
        doc = self.desktop.getCurrentComponent()
        if not doc or not hasattr(doc, "Sheets"):
            self.message("Budget fra CSV", "Åbn budgettet først.")
            return None
        workbook = budget_office.BudgetWorkbook(doc)
        if not workbook.has_budget():
            self.message("Budget fra CSV",
                         "Dette dokument indeholder ikke et ark ved navn "
                         "\"%s\"." % budget_office.SHEET_TX)
            return None
        return workbook

    def assign_dialog(self):
        """Give the uncategorised transactions a category, a few clicks."""
        workbook = self._workbook()
        if workbook is None:
            return
        groups = workbook.uncategorised_groups()
        if not groups:
            self.message("Kategorisér poster",
                         "Alle posteringer har allerede en kategori.")
            return

        categories = workbook.category_list()
        state = {"groups": list(groups), "assigned": [], "rules": []}

        model, dialog = self._assign_dialog_controls(state, categories)
        try:
            dialog.execute()
        finally:
            dialog.dispose()

        if not state["assigned"]:
            return
        doc = workbook.doc
        doc.lockControllers()
        try:
            for cells, category in state["assigned"]:
                workbook.assign_category(cells, category)
            if state["rules"]:
                workbook.add_rules(state["rules"])
            workbook.set_category_list(_merge(workbook.category_list(),
                                              [c for _cells, c in state["assigned"]]))
        finally:
            doc.unlockControllers()
        rows = workbook.read_rules()
        if rows:
            try:
                RuleSet(rows).save(user_rules_path())
            except Exception:
                pass
        workbook.rebuild()
        self.message("Kategorisér poster",
                     "%d posteringsgruppe(r) fik en kategori.\n"
                     "%d ny(e) regel(er) blev gemt."
                     % (len(state["assigned"]), len(state["rules"])))

    def _assign_dialog_controls(self, state, categories):
        model = self.create("com.sun.star.awt.UnoControlDialogModel")
        model.Width, model.Height = 300, 218
        model.Title = "Kategorisér poster"

        def add(kind, name, x, y, w, h, **props):
            control = model.createInstance("com.sun.star.awt.UnoControl%sModel"
                                           % kind)
            model.insertByName(name, control)
            control.PositionX, control.PositionY = x, y
            control.Width, control.Height = w, h
            for key, value in props.items():
                try:
                    setattr(control, key, value)
                except Exception:
                    pass
            return control

        add("FixedText", "help", 6, 6, 288, 20,
            Label="Posterne er samlet pr. forretning - beløb, datoer og "
                  "notanumre i teksten ignoreres. Vælg en eller flere, vælg "
                  "eller skriv en kategori, og tryk Tildel.",
            MultiLine=True)
        listbox = add("ListBox", "items", 6, 30, 288, 112, MultiSelection=True)
        listbox.StringItemList = tuple(_group_labels(state["groups"]))

        add("FixedText", "l_cat", 6, 148, 44, 10, Label="Kategori:")
        combo = add("ComboBox", "category", 52, 146, 140, 12, Dropdown=True,
                    LineCount=20, Text=categories[0] if categories else "")
        combo.StringItemList = tuple(categories)
        add("Button", "assign", 198, 145, 46, 14, Label="Tildel")
        add("Button", "close", 248, 145, 46, 14, Label="Luk", PushButtonType=1,
            DefaultButton=True)

        add("FixedText", "l_rule", 6, 166, 44, 10, Label="Regel:")
        add("Edit", "keyword", 52, 164, 140, 12,
            Text=state["groups"][0][4] if state["groups"] else "")
        add("FixedText", "rule_hint", 198, 166, 96, 10, Label="(tom = per post)")
        add("CheckBox", "rule", 52, 180, 242, 10, State=1,
            Label="Gem som regel, så forretningen kendes næste gang")
        add("FixedText", "status", 6, 194, 288, 20, Label="", MultiLine=True)

        dialog = self.create("com.sun.star.awt.UnoControlDialog")
        dialog.setModel(model)
        dialog.setVisible(False)
        dialog.createPeer(self.create("com.sun.star.awt.Toolkit"), None)

        def selected_indexes():
            return list(model.getByName("items").SelectedItems or ())

        def on_select():
            """Show the keyword that will be saved for a single selection."""
            selected = selected_indexes()
            if len(selected) == 1 and selected[0] < len(state["groups"]):
                model.getByName("keyword").Text = state["groups"][selected[0]][4]
            elif len(selected) > 1:
                model.getByName("keyword").Text = ""

        def on_assign():
            category = model.getByName("category").Text.strip()
            selected = selected_indexes()
            if not category or not selected:
                model.getByName("status").Label = (
                    "Vælg mindst én post og skriv en kategori.")
                return
            save_rule = bool(model.getByName("rule").State)
            override = model.getByName("keyword").Text.strip()
            remaining = []
            done = 0
            for index, group in enumerate(state["groups"]):
                if index not in selected:
                    remaining.append(group)
                    continue
                _name, _count, _total, cells, keyword = group
                state["assigned"].append((cells, category))
                if save_rule:
                    state["rules"].append((override or keyword, category))
                done += 1
            state["groups"] = remaining
            model.getByName("items").StringItemList = tuple(
                _group_labels(remaining))
            combo_model = model.getByName("category")
            if category not in combo_model.StringItemList:
                combo_model.StringItemList = tuple(
                    list(combo_model.StringItemList) + [category])
            model.getByName("keyword").Text = ""
            model.getByName("status").Label = (
                "%d gruppe(r) sat til \"%s\". %d tilbage."
                % (done, category, len(remaining)))

        listener = _ActionListener(on_assign)
        dialog.getControl("assign").addActionListener(listener)
        self._listeners.append(listener)
        selection = _ItemListener(on_select)
        try:
            dialog.getControl("items").addItemListener(selection)
            self._listeners.append(selection)
        except Exception:
            pass
        return model, dialog

    def categories_dialog(self):
        """Create, rename and delete categories."""
        workbook = self._workbook()
        if workbook is None:
            return
        state = {"names": workbook.category_list(), "dirty": False,
                 "renames": [], "deletes": [], "adds": []}

        model = self.create("com.sun.star.awt.UnoControlDialogModel")
        model.Width, model.Height = 240, 190
        model.Title = "Kategorier"

        def add(kind, name, x, y, w, h, **props):
            control = model.createInstance("com.sun.star.awt.UnoControl%sModel"
                                           % kind)
            model.insertByName(name, control)
            control.PositionX, control.PositionY = x, y
            control.Width, control.Height = w, h
            for key, value in props.items():
                try:
                    setattr(control, key, value)
                except Exception:
                    pass
            return control

        add("FixedText", "help", 6, 6, 228, 18,
            Label="Vælg en kategori for at omdøbe eller slette den, eller skriv "
                  "et nyt navn og tryk Tilføj.", MultiLine=True)
        listbox = add("ListBox", "names", 6, 26, 160, 120)
        listbox.StringItemList = tuple(state["names"])
        add("Edit", "name", 6, 150, 160, 12, Text="")
        add("Button", "add", 172, 26, 62, 14, Label="Tilføj")
        add("Button", "rename", 172, 44, 62, 14, Label="Omdøb")
        add("Button", "delete", 172, 62, 62, 14, Label="Slet")
        add("Button", "close", 172, 150, 62, 14, Label="Luk", PushButtonType=1,
            DefaultButton=True)
        add("FixedText", "status", 6, 166, 228, 18, Label="", MultiLine=True)

        dialog = self.create("com.sun.star.awt.UnoControlDialog")
        dialog.setModel(model)
        dialog.setVisible(False)
        dialog.createPeer(self.create("com.sun.star.awt.Toolkit"), None)

        def selected_name():
            items = model.getByName("names").SelectedItems
            if items and 0 <= items[0] < len(state["names"]):
                return state["names"][items[0]]
            return ""

        def refresh_list(status=""):
            model.getByName("names").StringItemList = tuple(state["names"])
            model.getByName("status").Label = status

        def on_add():
            name = model.getByName("name").Text.strip()
            if not name:
                refresh_list("Skriv et navn først.")
                return
            if name in state["names"]:
                refresh_list("\"%s\" findes allerede." % name)
                return
            state["names"].append(name)
            state["adds"].append(name)
            state["dirty"] = True
            model.getByName("name").Text = ""
            refresh_list("Tilføjede \"%s\"." % name)

        def on_rename():
            old = selected_name()
            new = model.getByName("name").Text.strip()
            if not old or not new:
                refresh_list("Vælg en kategori og skriv det nye navn.")
                return
            state["names"] = [new if n == old else n for n in state["names"]]
            state["renames"].append((old, new))
            state["dirty"] = True
            model.getByName("name").Text = ""
            refresh_list("Omdøbte \"%s\" til \"%s\"." % (old, new))

        def on_delete():
            name = selected_name()
            if not name:
                refresh_list("Vælg en kategori først.")
                return
            state["names"] = [n for n in state["names"] if n != name]
            state["deletes"].append(name)
            state["dirty"] = True
            refresh_list("Sletter \"%s\" - posteringerne bliver "
                         "ukategoriserede." % name)

        for control_name, callback in (("add", on_add), ("rename", on_rename),
                                       ("delete", on_delete)):
            listener = _ActionListener(callback)
            dialog.getControl(control_name).addActionListener(listener)
            self._listeners.append(listener)

        try:
            dialog.execute()
        finally:
            dialog.dispose()

        if not state["dirty"]:
            return
        doc = workbook.doc
        doc.lockControllers()
        try:
            for old, new in state["renames"]:
                workbook.rename_category(old, new)
            for name in state["deletes"]:
                workbook.delete_category(name)
            workbook.set_category_list(state["names"])
        finally:
            doc.unlockControllers()
        rows = workbook.read_rules()
        if rows:
            try:
                RuleSet(rows).save(user_rules_path())
            except Exception:
                pass
        workbook.rebuild()
        self.message("Kategorier",
                     "%d tilføjet, %d omdøbt, %d slettet."
                     % (len(state["adds"]), len(state["renames"]),
                        len(state["deletes"])))

    def show_rules(self):
        doc = self.desktop.getCurrentComponent()
        if not doc or not hasattr(doc, "Sheets"):
            self.message("Budget fra CSV", "Åbn budgettet først.")
            return
        name = budget_office.SHEET_RULES
        workbook = budget_office.BudgetWorkbook(doc)
        if not doc.Sheets.hasByName(name):
            workbook.write_rules(workbook.sheet(name), load_rules())
        doc.CurrentController.setActiveSheet(doc.Sheets.getByName(name))

    def reset_rules(self):
        doc = self.desktop.getCurrentComponent()
        if not doc or not hasattr(doc, "Sheets"):
            self.message("Budget fra CSV", "Åbn budgettet først.")
            return
        if not self.ask("Nulstil kategoriregler",
                        "Erstat reglerne i arket \"%s\" med standardreglerne? "
                        "Konto-kolonnen bevares." % budget_office.SHEET_RULES):
            return
        workbook = budget_office.BudgetWorkbook(doc)
        accounts = workbook.read_category_accounts()
        ruleset = RuleSet.defaults()
        workbook.write_rules(workbook.sheet(budget_office.SHEET_RULES), ruleset,
                             accounts=accounts)
        doc.CurrentController.setActiveSheet(
            doc.Sheets.getByName(budget_office.SHEET_RULES))

    def show_accounts(self):
        """Jump to (or create) the Konti sheet."""
        doc = self.desktop.getCurrentComponent()
        if not doc or not hasattr(doc, "Sheets"):
            self.message("Budget fra CSV", "Åbn budgettet først.")
            return
        workbook = budget_office.BudgetWorkbook(doc)
        workbook._ensure_accounts_sheet()
        doc.CurrentController.setActiveSheet(
            doc.Sheets.getByName(budget_office.SHEET_ACCOUNTS))

    def bank_budget(self):
        """Snapshot the current Mål column into a clean, printable sheet."""
        workbook = self._workbook()
        if workbook is None:
            return
        if not workbook.doc.Sheets.hasByName(budget_office.SHEET_PLAN):
            self.message("Budget fra CSV",
                         "Der er intet budgetforslag endnu - importér en fil "
                         "med mere end én måned først.")
            return
        targets = workbook.read_targets()
        if not targets:
            self.message("Budget fra CSV",
                         "Fandt ingen mål i \"%s\"." % budget_office.SHEET_PLAN)
            return
        doc = workbook.doc
        sheet = workbook.sheet(budget_office.SHEET_BANK)
        workbook.write_bank_budget(sheet, targets)
        doc.CurrentController.setActiveSheet(sheet)
        self.message("Bankbudget",
                     "Arket \"%s\" er oprettet ud fra de mål, du har sat lige "
                     "nu i \"%s\". Det opdateres ikke automatisk - kør "
                     "kommandoen igen, hvis du retter dine mål."
                     % (budget_office.SHEET_BANK, budget_office.SHEET_PLAN))

    def about(self):
        self.message(
            "Om Budget fra CSV",
            "Budget fra CSV\n\n"
            "Laver et månedsbudget ud fra en CSV-fil fra banken:\n"
            "posteringerne kategoriseres automatisk, og oversigten viser "
            "budget, faktisk forbrug og forskel pr. kategori i kroner.\n\n"
            "1. Budget ▸ Importér CSV-fil…\n"
            "2. Ret kategorier i arket \"Transaktioner\" eller reglerne i "
            "arket \"Kategorier\"\n"
            "3. Budget ▸ Opdatér kategorier og budget")

    # -- file picking -----------------------------------------------------
    def pick_file(self):
        picker = self.create("com.sun.star.ui.dialogs.FilePicker")
        picker.setTitle("Vælg CSV-fil med posteringer")
        picker.appendFilter("CSV- og tekstfiler (*.csv;*.txt;*.tsv)",
                            "*.csv;*.txt;*.tsv")
        picker.appendFilter("Alle filer (*.*)", "*.*")
        try:
            picker.setCurrentFilter("CSV- og tekstfiler (*.csv;*.txt;*.tsv)")
        except Exception:
            pass
        if picker.execute() != 1:
            return None
        files = picker.getSelectedFiles()
        if not files:
            return None
        return unohelper.fileUrlToSystemPath(files[0])

    def new_calc(self):
        return self.desktop.loadComponentFromURL(
            "private:factory/scalc", "_blank", 0, ())

    # -- options dialog ---------------------------------------------------
    def options_dialog(self, table, mapping, source_name):
        """Let the user confirm/adjust the detected columns.

        Returns a settings dict, or ``None`` when the user cancels.
        """
        model = build_options_model(self.create, table, mapping, source_name)
        dialog = self.create("com.sun.star.awt.UnoControlDialog")
        dialog.setModel(model)
        dialog.setVisible(False)
        dialog.createPeer(self.create("com.sun.star.awt.Toolkit"), None)
        try:
            if dialog.execute() != 1:
                return None
            settings = read_options_model(model)
        finally:
            dialog.dispose()

        if settings["date"] is None:
            self.message("Budget fra CSV", "Vælg en datokolonne.", kind="WARNINGBOX")
            return None
        if settings["amount"] is None and settings["amount_in"] is None and \
                settings["amount_out"] is None:
            self.message("Budget fra CSV", "Vælg en beløbskolonne.",
                         kind="WARNINGBOX")
            return None
        return settings


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

class _ItemListener(unohelper.Base, XItemListener):
    """Runs a callback when the selection in a list box changes."""

    def __init__(self, callback):
        self.callback = callback

    def itemStateChanged(self, event):
        self.callback()

    def disposing(self, source):
        pass


class _ActionListener(unohelper.Base, XActionListener):
    """Runs a callback when a dialog button is pressed."""

    def __init__(self, callback):
        self.callback = callback

    def actionPerformed(self, event):
        self.callback()

    def disposing(self, source):
        pass


def _group_labels(groups):
    return ["%-44s %3d stk.   %9.0f kr." % (name[:44], count, total)
            for name, count, total, _cells, _keyword in groups]


def _merge(names, extra):
    out = list(names)
    for name in extra:
        if name and name not in out:
            out.append(name)
    return out


def _index(value):
    """Column index -> list box position (0 is "(ingen)")."""
    return 0 if value is None else value + 1


def _or_none(value):
    return None if value < 0 else value


def _select(control, position):
    try:
        control.SelectedItems = (int(position),)
    except Exception:
        try:
            control.setPropertyValue("SelectedItems",
                                     uno.Any("[]short", (int(position),)))
        except Exception:
            pass


def _selected(control):
    try:
        items = control.SelectedItems
        if items:
            return int(items[0])
    except Exception:
        pass
    return 0


def _preview_text(table, mapping):
    lines = ["Eksempel på de første rækker:"]
    for row in table.sample(3):
        parts = []
        if mapping.date is not None and mapping.date < len(row):
            parts.append(row[mapping.date])
        for index in mapping.text:
            if index < len(row):
                parts.append(row[index])
        for index in (mapping.amount, mapping.amount_in, mapping.amount_out):
            if index is not None and index < len(row) and row[index].strip():
                parts.append(row[index])
        lines.append("  " + "   ".join(part.strip() for part in parts if part))
    return "\n".join(lines)


def build_options_model(create, table, mapping, source_name):
    """Build the import dialog (as a model, so it can be tested headlessly)."""
    choices = [NONE_LABEL] + table.header_names()

    model = create("com.sun.star.awt.UnoControlDialogModel")
    model.Width = 260
    model.Height = 214
    model.Title = "Importér til budget"

    def add(kind, name, x, y, w, h, **props):
        control = model.createInstance("com.sun.star.awt.UnoControl%sModel" % kind)
        model.insertByName(name, control)
        control.PositionX, control.PositionY = x, y
        control.Width, control.Height = w, h
        for key, value in props.items():
            try:
                setattr(control, key, value)
            except Exception:
                pass
        return control

    def listbox(name, x, y, w, items, selected):
        control = add("ListBox", name, x, y, w, 12, Dropdown=True)
        control.StringItemList = tuple(items)
        _select(control, selected)
        return control

    def label(name, x, y, w, h, text, **props):
        return add("FixedText", name, x, y, w, h, Label=text, **props)

    label("info", 6, 6, 248, 18, "%s: %s" % (source_name, table.describe()),
          MultiLine=True)

    label("l_date", 6, 30, 60, 10, "Datokolonne:")
    listbox("date", 70, 28, 90, choices, _index(mapping.date))
    label("l_dayfirst", 168, 30, 34, 10, "Format:")
    listbox("dayfirst", 204, 28, 50, ["DD-MM-ÅÅÅÅ", "MM-DD-ÅÅÅÅ"],
            0 if mapping.dayfirst else 1)

    label("l_text", 6, 48, 60, 10, "Tekstkolonne:")
    listbox("text1", 70, 46, 90, choices,
            _index(mapping.text[0] if mapping.text else None))
    label("l_text2", 168, 48, 34, 10, "+ tekst:")
    listbox("text2", 204, 46, 50, choices,
            _index(mapping.text[1] if len(mapping.text) > 1 else None))

    label("l_amount", 6, 66, 60, 10, "Beløbskolonne:")
    listbox("amount", 70, 64, 90, choices, _index(mapping.amount))
    label("l_decimal", 168, 66, 34, 10, "Decimal:")
    listbox("decimal", 204, 64, 50, ["Komma (1.234,56)", "Punktum (1,234.56)"],
            0 if (mapping.decimal or ",") == "," else 1)

    label("l_in", 6, 84, 60, 10, "Eller indsat:")
    listbox("amount_in", 70, 82, 90, choices, _index(mapping.amount_in))
    label("l_out", 168, 84, 34, 10, "hævet:")
    listbox("amount_out", 204, 82, 50, choices, _index(mapping.amount_out))

    label("l_sign", 6, 102, 60, 10, "Fortegn:")
    listbox("sign", 70, 100, 184, [text for _key, text in SIGN_CHOICES], 0)

    add("CheckBox", "suggest", 70, 118, 184, 10,
        Label="Foreslå budgettal ud fra gennemsnittet pr. måned", State=1)

    label("preview", 6, 134, 248, 46, _preview_text(table, mapping),
          MultiLine=True)
    label("notes", 6, 178, 248, 18,
          "\n".join(mapping.notes[:2]) if mapping.notes else "", MultiLine=True)

    add("Button", "ok", 146, 198, 52, 14, Label="Opret budget", PushButtonType=1,
        DefaultButton=True)
    add("Button", "cancel", 202, 198, 52, 14, Label="Annullér", PushButtonType=2)
    return model


def read_options_model(model):
    """Turn the dialog's state into the settings dict used by the import."""
    def get(name):
        return model.getByName(name)

    text_columns = []
    for name in ("text1", "text2"):
        index = _selected(get(name)) - 1
        if index >= 0 and index not in text_columns:
            text_columns.append(index)
    return {
        "date": _or_none(_selected(get("date")) - 1),
        "text": text_columns,
        "amount": _or_none(_selected(get("amount")) - 1),
        "amount_in": _or_none(_selected(get("amount_in")) - 1),
        "amount_out": _or_none(_selected(get("amount_out")) - 1),
        "decimal": "," if _selected(get("decimal")) == 0 else ".",
        "dayfirst": _selected(get("dayfirst")) == 0,
        "sign": SIGN_CHOICES[max(0, _selected(get("sign")))][0],
        "suggest": bool(get("suggest").State),
    }


def _start_balance(table, mapping, result):
    """Balance before the first transaction, when the file has a balance column."""
    if mapping.balance is None or not result.transactions:
        return 0.0
    first = result.transactions[0]
    try:
        row = table.rows[first.source_row]
    except IndexError:
        return 0.0
    balance = parse_amount(row[mapping.balance], mapping.decimal) \
        if mapping.balance < len(row) else None
    if balance is None:
        return 0.0
    return round(balance - first.amount, 2)


def _done_text(result, summary, table):
    lines = list(result.summary_lines())
    lines.append("")
    lines.append("Måneder: %d" % len(summary.months))
    if result.uncategorised:
        lines.append("")
        lines.append("Uden kategori (top 3) - brug \"Budget ▸ Kategorisér "
                     "poster uden kategori\":")
        for text, count, total, _keyword in result.uncategorised[:3]:
            lines.append("  %s (%d stk., %.0f kr.)" % (text[:40], count, total))
    return "\n".join(lines)


def load_rules():
    """The user's own rules if they exist, otherwise the defaults."""
    path = user_rules_path()
    if os.path.isfile(path):
        try:
            rules = RuleSet.load(path)
            if len(rules):
                return rules
        except Exception:
            pass
    return RuleSet.defaults()


def _log_directory():
    for name in ("TMPDIR", "TEMP", "TMP"):
        value = os.environ.get(name)
        if value and os.path.isdir(value):
            return value
    return os.path.expanduser("~")


# ---------------------------------------------------------------------------
# UNO registration
# ---------------------------------------------------------------------------
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(BudgetJob, IMPLEMENTATION_NAME,
                                         (SERVICE_NAME,))
