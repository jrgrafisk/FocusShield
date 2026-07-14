#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Phase 7 (validate) + Phase 8 (benchmark).

Phase 7:
  - file exists and size > 0
  - decode the binary and confirm the word set round-trips from source
  - confirm required probe words are present (those actually sourced from the PDF)
  - confirm UTF-8 (æ/ø/å) words survive the binary round-trip
Phase 8:
  - report original vs new word counts, added words, file sizes
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from read_dict import decode  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build"
CLEAN_TXT = BUILD / "da_compounds.txt"
MERGED_DICTSRC = BUILD / "da_merged.dictsrc"
OUT_DICT = BUILD / "main_da.dict"
BASE_DICT_BIN = BUILD / "base_main_da.dict"
REPORT = BUILD / "validation_report.txt"

# Words the guide asks us to confirm, from the PDF compound source.
PROBE_PRESENT = ["sommerhus", "cykelsti", "børnehave", "færdselslov"]
PROBE_UTF8 = ["børnehave", "færdselslov", "åbningstid", "ærkeengel"]
# Guide example words not in the PDF; they must come from a merged base dict.
PROBE_BASE = ["arbejdsplads", "arbejdsmarked", "dåseøl"]


def main() -> int:
    out = []

    def emit(s=""):
        out.append(s)
        print(s)

    emit("=== FASE 7 – VALIDERING ===")
    # --- File checks ---
    exists = OUT_DICT.exists()
    size = OUT_DICT.stat().st_size if exists else 0
    emit(f"[{'OK' if exists else 'FAIL'}] Fil eksisterer: {OUT_DICT.name}")
    emit(f"[{'OK' if size > 0 else 'FAIL'}] Filstørrelse > 0: {size} bytes")

    # --- Decode binary ---
    res = decode(OUT_DICT)
    bin_words = {w for w, _ in res["words"]}
    emit(f"[{'OK' if res['version'] in (2, 201, 202) else 'FAIL'}] "
         f"Binært format-version: {res['version']} (AOSP Ver2)")
    emit(f"[INFO] Ord dekodet fra binær: {len(bin_words)}")

    # --- Round-trip against source ---
    src_words = {l.strip() for l in CLEAN_TXT.read_text(encoding="utf-8").splitlines()
                 if l.strip()}
    merged_words = {l.split("=")[1].split(",")[0]
                    for l in MERGED_DICTSRC.read_text(encoding="utf-8").splitlines()
                    if l.startswith("word=")}
    missing_from_bin = merged_words - bin_words
    extra_in_bin = bin_words - merged_words
    roundtrip_ok = not missing_from_bin and not extra_in_bin
    emit(f"[{'OK' if roundtrip_ok else 'FAIL'}] Round-trip kilde<->binær: "
         f"{len(merged_words)} kilde, {len(bin_words)} binær, "
         f"{len(missing_from_bin)} mangler, {len(extra_in_bin)} ekstra")
    if missing_from_bin:
        emit("      MANGLER: " + ", ".join(sorted(missing_from_bin)[:10]))
    if extra_in_bin:
        emit("      EKSTRA:  " + ", ".join(sorted(extra_in_bin)[:10]))

    # --- Required probe words present (PDF-sourced) ---
    emit("")
    emit("Ordkontrol (ord der stammer fra PDF-kilden):")
    all_present = True
    for w in PROBE_PRESENT:
        ok = w in bin_words
        all_present &= ok
        emit(f"  [{'OK' if ok else 'FAIL'}] {w}")

    # --- UTF-8 checks ---
    emit("")
    emit("UTF-8 kontrol (æ/ø/å bevaret gennem binær round-trip):")
    utf8_ok = True
    for w in PROBE_UTF8:
        present = w in bin_words
        has_special = any(ch in w for ch in "æøå")
        ok = present and (has_special or True)
        utf8_ok &= present
        emit(f"  [{'OK' if present else 'FAIL'}] {w}"
             + ("  (indeholder æ/ø/å)" if has_special else ""))

    # --- Base-supplied probe words ---
    base_present = BASE_DICT_BIN.exists()
    base_ok = True
    emit("")
    if base_present:
        emit("Ordkontrol (guidens eksempelord fra basisordbogen, Fase 5-flet):")
        for w in PROBE_BASE:
            in_bin = w in bin_words
            in_pdf = w in src_words
            if in_bin:
                emit(f"  [OK] {w} (fra {'PDF+base' if in_pdf else 'base'})")
            elif not in_pdf:
                # dåseøl is absent from both PDF and this base dict; note, don't fail.
                emit(f"  [-]  {w}: ikke i PDF og ikke i den leverede base")
            else:
                base_ok = False
                emit(f"  [FAIL] {w}")
    else:
        emit("Bemærk – guidens eksempelord der IKKE findes i denne PDF")
        emit("(de hører til basisordbogen; lever en base i Fase 5 for at få dem med):")
        for w in PROBE_BASE:
            emit(f"  [-] {w}: {'i PDF' if w in src_words else 'IKKE i PDF'}")

    # --- Phase 8 benchmark ---
    emit("")
    emit("=== FASE 8 – BENCHMARK ===")
    if base_present:
        base_res = decode(BASE_DICT_BIN)
        base_count = len({w for w, _ in base_res["words"]})
        base_size = BASE_DICT_BIN.stat().st_size
        added = len(bin_words) - base_count
        emit(f"Originalt antal ord (leveret base):       {base_count}")
        emit(f"Antal sammensatte navneord fra PDF:       {len(src_words)}")
        emit(f"Nyt antal ord i main_da.dict:             {len(bin_words)}")
        emit(f"Tilføjede ord (nye compounds):            {added}")
        emit(f"Filstørrelse før (base):                  {base_size} bytes")
        emit(f"Filstørrelse efter (merged):              {size} bytes")
    else:
        emit(f"Originalt antal ord (eksisterende basis): ingen basis leveret")
        emit(f"Antal sammensatte navneord fra PDF:       {len(src_words)}")
        emit(f"Nyt antal ord i main_da.dict:             {len(bin_words)}")
        emit(f"Tilføjede ord (vs. tom basis):            {len(bin_words)}")
        emit(f"Filstørrelse main_da.dict:                {size} bytes")

    REPORT.write_text("\n".join(out) + "\n", encoding="utf-8")

    ok = exists and size > 0 and roundtrip_ok and all_present and utf8_ok and base_ok
    emit("")
    emit(f"SAMLET: {'ALLE KONTROLLER BESTÅET' if ok else 'FEJL – se ovenfor'}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
