#!/usr/bin/env python3
"""Regenerate BudgetFraCSV.gs from the individual files under src/.

The single combined file is what users paste into the Apps Script editor
(one file beats copy-pasting eleven). src/ stays the source of truth for
development and the Node test suite - run this after editing anything
under src/*.gs.
"""
import os

ROOT = os.path.dirname(os.path.abspath(__file__))
ORDER = ["TextUtils", "Parsing", "Transfers", "Merchant", "Rules", "CsvSniff",
         "Columns", "Budget", "Plan", "SheetWriter", "Code"]

BANNER = (
    "// ============================================================================\n"
    "// Budget fra CSV - til Google Sheets\n"
    "//\n"
    "// Dette er alle scriptets filer samlet i én, til nem installation: opret ét\n"
    "// script-fil i Apps Script-editoren (Extensions > Apps Script), slet den\n"
    "// forududfyldte kode, og indsæt hele denne fils indhold. Se README.md for\n"
    "// den fulde installationsvejledning (kun denne fil + ImportDialog.html skal\n"
    "// oprettes manuelt).\n"
    "//\n"
    "// Kildekoden ligger opdelt i enkeltfiler under src/, hvis du vil læse eller\n"
    "// ændre den - denne fil genereres derfra (build.py) og bør ikke redigeres\n"
    "// direkte.\n"
    "// ============================================================================\n"
)

# Files export a small module.exports block (guarded by "typeof module !==
# undefined", a no-op in real Apps Script) purely so the Node test suite can
# require() them - strip it from the combined file since it is dead weight
# there.
MARKER = 'if (typeof module !== "undefined") {'


def main():
    parts = [BANNER]
    for name in ORDER:
        path = os.path.join(ROOT, "src", name + ".gs")
        with open(path, encoding="utf-8") as f:
            content = f.read()
        idx = content.find(MARKER)
        if idx != -1:
            content = content[:idx].rstrip() + "\n"
        parts.append("\n// ---- %s.gs %s\n\n" % (name, "-" * (70 - len(name))))
        parts.append(content)

    combined = "".join(parts)
    out_path = os.path.join(ROOT, "BudgetFraCSV.gs")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(combined)
    print("wrote %s (%d bytes)" % (out_path, len(combined)))


if __name__ == "__main__":
    main()
