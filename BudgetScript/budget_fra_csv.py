#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Budget fra CSV - lav et Excel-budget fra et bank-CSV-udtræk.

Kør scriptet (dobbeltklik på start_budget.bat, eller
"python budget_fra_csv.py" fra en kommandoprompt) for at komme i gang.
Ingen makroer, ingen tilføjelsesprogrammer - filen, den laver, er et helt
almindeligt Excel-regneark.
"""

from __future__ import annotations

import os
import sys
import traceback

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from budget_core import csvsniff
from budget_core.budget import BuildOptions, build_transactions
from budget_core.columns import detect_mapping
from budget_core.parsing import parse_amount
from budget_core.rules import RuleSet, user_rules_path
from xlsx_writer import BudgetBook

APP_TITLE = "Budget fra CSV"


# ---------------------------------------------------------------------------
# Rules persistence
# ---------------------------------------------------------------------------

def load_rules() -> RuleSet:
    path = user_rules_path()
    if os.path.isfile(path):
        try:
            return RuleSet.load(path)
        except Exception:
            pass
    return RuleSet.defaults()


def save_rules_quietly(ruleset: RuleSet) -> None:
    try:
        ruleset.save(user_rules_path())
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Column mapping confirmation dialog
# ---------------------------------------------------------------------------

class MappingDialog(tk.Toplevel):
    """Confirm/correct the auto-detected column mapping before building."""

    def __init__(self, parent, table, mapping, source_name):
        super().__init__(parent)
        self.title("Bekræft kolonner")
        self.resizable(False, False)
        self.table = table
        self.mapping = mapping
        self.result = None
        self.transient(parent)
        self.grab_set()

        names = table.header_names()
        options = ["(ingen)"] + [f"{i}: {n}" for i, n in enumerate(names)]

        pad = {"padx": 10, "pady": 4}
        frame = ttk.Frame(self, padding=12)
        frame.grid()

        ttk.Label(frame, text=f"Fil: {source_name}", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", **pad)
        ttk.Label(frame, text=f"{table.describe()}").grid(
            row=1, column=0, columnspan=2, sticky="w", **pad)

        self.date_var = tk.StringVar(value=self._opt(options, mapping.date))
        self.text_var = tk.StringVar(value=self._opt(options, mapping.text[0] if mapping.text else None))
        self.amount_var = tk.StringVar(value=self._opt(options, mapping.amount))
        self.amount_in_var = tk.StringVar(value=self._opt(options, mapping.amount_in))
        self.amount_out_var = tk.StringVar(value=self._opt(options, mapping.amount_out))
        self.decimal_var = tk.StringVar(value="komma (,)" if mapping.decimal == "," else "punktum (.)")
        self.dayfirst_var = tk.StringVar(value="dag før måned" if mapping.dayfirst else "måned før dag")

        row = 2
        row = self._combo(frame, row, "Dato:", self.date_var, options)
        row = self._combo(frame, row, "Tekst:", self.text_var, options)
        row = self._combo(frame, row, "Beløb (én kolonne):", self.amount_var, options)
        row = self._combo(frame, row, "  - eller Indbetalt:", self.amount_in_var, options)
        row = self._combo(frame, row, "  - og Hævet:", self.amount_out_var, options)
        row = self._combo(frame, row, "Decimaltegn:", self.decimal_var,
                          ["komma (,)", "punktum (.)"])
        row = self._combo(frame, row, "Datoformat:", self.dayfirst_var,
                          ["dag før måned", "måned før dag"])

        if mapping.notes:
            note_text = "\n".join(mapping.notes)
            ttk.Label(frame, text=note_text, foreground="#966", wraplength=420,
                     justify="left").grid(row=row, column=0, columnspan=2, sticky="w", **pad)
            row += 1

        buttons = ttk.Frame(frame)
        buttons.grid(row=row, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(buttons, text="Opret budget", command=self._on_ok).pack(side="left", padx=6)
        ttk.Button(buttons, text="Annullér", command=self._on_cancel).pack(side="left", padx=6)

        self.bind("<Return>", lambda e: self._on_ok())
        self.bind("<Escape>", lambda e: self._on_cancel())
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    @staticmethod
    def _opt(options, index):
        if index is None:
            return options[0]
        for opt in options[1:]:
            if opt.startswith(f"{index}:"):
                return opt
        return options[0]

    def _combo(self, frame, row, label, var, options):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=4)
        combo = ttk.Combobox(frame, textvariable=var, values=options, state="readonly", width=32)
        combo.grid(row=row, column=1, sticky="w", padx=10, pady=4)
        return row + 1

    @staticmethod
    def _index_of(value: str):
        if value.startswith("(ingen)"):
            return None
        return int(value.split(":", 1)[0])

    def _on_ok(self):
        self.mapping.date = self._index_of(self.date_var.get())
        text_idx = self._index_of(self.text_var.get())
        self.mapping.text = [text_idx] if text_idx is not None else []
        self.mapping.amount = self._index_of(self.amount_var.get())
        self.mapping.amount_in = self._index_of(self.amount_in_var.get())
        self.mapping.amount_out = self._index_of(self.amount_out_var.get())
        if self.mapping.amount_in is not None or self.mapping.amount_out is not None:
            self.mapping.amount = None
        self.mapping.decimal = "," if self.decimal_var.get().startswith("komma") else "."
        self.mapping.dayfirst = self.dayfirst_var.get().startswith("dag")
        self.result = self.mapping
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(False, False)

        frame = ttk.Frame(self, padding=20)
        frame.grid()

        ttk.Label(frame, text=APP_TITLE, font=("", 16, "bold")).grid(
            row=0, column=0, pady=(0, 4), sticky="w")
        ttk.Label(frame, text="Importér et bank-CSV-udtræk og få et færdigt\n"
                              "Excel-budget - uden makroer.", justify="left").grid(
            row=1, column=0, pady=(0, 16), sticky="w")

        ttk.Button(frame, text="Nyt budget fra CSV-fil…", width=36,
                  command=self.new_budget).grid(row=2, column=0, pady=4, sticky="ew")
        ttk.Button(frame, text="Tilføj flere posteringer (CSV) til budget…", width=36,
                  command=self.append_budget).grid(row=3, column=0, pady=4, sticky="ew")
        ttk.Button(frame, text="Opdatér kategorier i budget…", width=36,
                  command=self.refresh_budget).grid(row=4, column=0, pady=4, sticky="ew")
        ttk.Button(frame, text="Luk", width=36, command=self.destroy).grid(
            row=5, column=0, pady=(16, 0), sticky="ew")

    # -- shared helpers -----------------------------------------------
    def _pick_csv(self) -> str:
        return filedialog.askopenfilename(
            title="Vælg CSV-fil", filetypes=[("CSV-filer", "*.csv *.txt"), ("Alle filer", "*.*")])

    def _pick_xlsx(self, title="Vælg budget (.xlsx)") -> str:
        return filedialog.askopenfilename(
            title=title, filetypes=[("Excel-filer", "*.xlsx"), ("Alle filer", "*.*")])

    def _pick_save(self, initial="budget.xlsx") -> str:
        return filedialog.asksaveasfilename(
            title="Gem budget som…", defaultextension=".xlsx",
            initialfile=initial, filetypes=[("Excel-filer", "*.xlsx")])

    def _read_csv(self, path: str):
        try:
            return csvsniff.read_table(path=path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Kunne ikke læse filen:\n{exc}")
            return None

    def _confirm_mapping(self, table, source_name):
        mapping = detect_mapping(table)
        dialog = MappingDialog(self, table, mapping, source_name)
        self.wait_window(dialog)
        return dialog.result

    def _offer_open(self, path: str) -> None:
        if messagebox.askyesno(APP_TITLE, f"Budgettet er gemt:\n{path}\n\nÅbn det nu?"):
            try:
                os.startfile(path)  # type: ignore[attr-defined]
            except AttributeError:
                import subprocess
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.Popen([opener, path])
            except Exception as exc:
                messagebox.showwarning(APP_TITLE, f"Kunne ikke åbne filen automatisk:\n{exc}")

    # -- commands -----------------------------------------------------
    def new_budget(self):
        path = self._pick_csv()
        if not path:
            return
        table = self._read_csv(path)
        if table is None or not table.rows:
            messagebox.showinfo(APP_TITLE, "Fandt ingen datarækker i filen.")
            return
        source_name = os.path.basename(path)
        mapping = self._confirm_mapping(table, source_name)
        if mapping is None:
            return

        ruleset = load_rules()
        result = build_transactions(table, mapping, ruleset,
                                    BuildOptions(decimal=mapping.decimal, dayfirst=mapping.dayfirst))
        if not result.ok:
            messagebox.showerror(APP_TITLE, "Der kunne ikke laves et budget.\n\n" +
                                 "\n".join(result.summary_lines()))
            return

        save_path = self._pick_save()
        if not save_path:
            return

        book = BudgetBook()
        start_balance = _start_balance(table, mapping, result)
        summary = book.build(result.transactions, ruleset, start_balance=start_balance)
        book.save(save_path)
        save_rules_quietly(ruleset)

        messagebox.showinfo(APP_TITLE, "\n".join(result.summary_lines()) +
                            f"\n\n{len(summary.months)} måned(er) i budgettet.")
        self._offer_open(save_path)

    def append_budget(self):
        budget_path = self._pick_xlsx("Vælg det budget, du vil tilføje til")
        if not budget_path:
            return
        try:
            book = BudgetBook.load(budget_path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Kunne ikke åbne budgetfilen:\n{exc}")
            return
        if not book.has_budget():
            messagebox.showerror(APP_TITLE, "Denne fil er ikke et budget lavet af dette script.")
            return

        csv_path = self._pick_csv()
        if not csv_path:
            return
        table = self._read_csv(csv_path)
        if table is None or not table.rows:
            messagebox.showinfo(APP_TITLE, "Fandt ingen datarækker i filen.")
            return
        source_name = os.path.basename(csv_path)
        mapping = self._confirm_mapping(table, source_name)
        if mapping is None:
            return

        rows = book.read_rules()
        ruleset = RuleSet(rows) if rows else load_rules()
        result = build_transactions(table, mapping, ruleset,
                                    BuildOptions(decimal=mapping.decimal, dayfirst=mapping.dayfirst))
        if not result.ok:
            messagebox.showerror(APP_TITLE, "Der kunne ikke tilføjes posteringer.\n\n" +
                                 "\n".join(result.summary_lines()))
            return

        added, skipped, truncated, summary = book.append_transactions(result.transactions)
        if summary is None:
            messagebox.showinfo(APP_TITLE, f"Ingen nye posteringer - alle {skipped} "
                                "postering(er) fandtes i forvejen i budgettet.")
            return

        save_path = self._pick_save(initial=os.path.basename(budget_path))
        if not save_path:
            return
        book.save(save_path)

        lines = [f"{added} ny(e) postering(er) tilføjet."]
        if skipped:
            lines.append(f"{skipped} postering(er) fandtes allerede og blev sprunget over.")
        if truncated:
            lines.append("Bemærk: budgettet kan højst rumme 4996 posteringer pr. type "
                         "(udgift/indtægt) - nogle af de nyeste blev ikke tilføjet.")
        lines.append(f"{len(summary.months)} måned(er) i budgettet nu.")
        messagebox.showinfo(APP_TITLE, "\n".join(lines))
        self._offer_open(save_path)

    def refresh_budget(self):
        budget_path = self._pick_xlsx()
        if not budget_path:
            return
        try:
            book = BudgetBook.load(budget_path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Kunne ikke åbne budgetfilen:\n{exc}")
            return
        if not book.has_budget():
            messagebox.showerror(APP_TITLE, "Denne fil er ikke et budget lavet af dette script.")
            return

        rows = book.read_rules()
        ruleset = RuleSet(rows) if rows else load_rules()
        summary, changed = book.refresh(ruleset)
        if summary is None:
            messagebox.showinfo(APP_TITLE, "Fandt ingen posteringer at opdatere.")
            return

        save_path = self._pick_save(initial=os.path.basename(budget_path))
        if not save_path:
            return
        book.save(save_path)
        if rows:
            save_rules_quietly(RuleSet(rows))

        messagebox.showinfo(APP_TITLE, f"Budgettet er opdateret.\n\n"
                            f"{changed} postering(er) fik ny kategori.\n"
                            f"{len(summary.months)} måned(er) i budgettet.")
        self._offer_open(save_path)


def _start_balance(table, mapping, result) -> float:
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


def main():
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        try:
            messagebox.showerror(APP_TITLE, "Der opstod en uventet fejl:\n\n" + traceback.format_exc())
        except Exception:
            pass
