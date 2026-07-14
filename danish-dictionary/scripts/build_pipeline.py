#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Reproducible pipeline: Danish compound-noun PDF -> AOSP/HeliBoard dictionary source.

Phases implemented here (all deterministic, UTF-8 in/out, originals never overwritten):
  Phase 1  Extract raw text from the PDF                -> da_compounds_raw.txt
  Phase 2  Clean the word list (regex/lowercase/sort)   -> da_compounds.txt
  Phase 3  Quality control + statistics report          -> qc_report.txt
  Phase 4  Assign frequency (f=128)                      -> da_compounds.dictsrc

Every rejected line and every error is logged to logs/pipeline.log so the run is
auditable and reproducible.
"""

import re
import sys
import logging
from pathlib import Path

import fitz  # PyMuPDF

# --- Paths -----------------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent          # danish-dictionary/
BUILD = ROOT / "build"
LOGS = ROOT / "logs"
BUILD.mkdir(exist_ok=True)
LOGS.mkdir(exist_ok=True)

PDF_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
    "/root/.claude/uploads/849ef4be-8006-50f7-90ac-bdcdc6367fa8/4d05ed66-sammensattenavneord.pdf"
)

RAW_TXT = BUILD / "da_compounds_raw.txt"
CLEAN_TXT = BUILD / "da_compounds.txt"
QC_REPORT = BUILD / "qc_report.txt"
DICTSRC = BUILD / "da_compounds.dictsrc"

FREQUENCY = 128            # Phase 4: default frequency for all new words
MIN_LEN = 3               # Phase 3: minimum length
MAX_LEN = 60              # Phase 3: maximum length
WORD_RE = re.compile(r"^[a-zæøå-]+$")   # Phase 2: allowed shape

# --- Logging ---------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler(LOGS / "pipeline.log", mode="w", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("pipeline")


def phase1_extract(pdf_path: Path) -> list[str]:
    """Extract UTF-8 text, dropping headers/footers/page numbers/empty lines.

    The PDF repeats four boilerplate header lines on every page plus a
    'Side N af M' page-number line. We drop those and keep the word lines,
    preserving the Danish characters æ/ø/å.
    """
    log.info("PHASE 1: extracting text from %s", pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    doc = fitz.open(pdf_path)
    log.info("PDF has %d pages", doc.page_count)

    # Boilerplate that must be stripped (matched case-insensitively / by prefix).
    boilerplate_substrings = (
        "SAMMENSATTE NAVNEORD",           # title / heading
        "RO 2012",                         # source/copyright note
        "alle ord er OK i RO",             # subtitle note
        "ligner ikke noget",               # subtitle note
    )
    page_re = re.compile(r"^\s*Side\s+\d+\s+af\s+\d+\s*$", re.IGNORECASE)

    kept: list[str] = []
    dropped_headers = 0
    dropped_empty = 0
    for page in doc:
        for raw_line in page.get_text().split("\n"):
            line = raw_line.strip()
            if not line:
                dropped_empty += 1
                continue
            if page_re.match(line):
                dropped_headers += 1
                continue
            if any(sub in line for sub in boilerplate_substrings):
                dropped_headers += 1
                continue
            kept.append(line)
    doc.close()

    RAW_TXT.write_text("\n".join(kept) + "\n", encoding="utf-8")
    log.info(
        "PHASE 1 done: %d word lines kept, %d header/footer lines dropped, "
        "%d empty lines dropped -> %s",
        len(kept), dropped_headers, dropped_empty, RAW_TXT.name,
    )
    return kept


def phase2_clean(raw_lines: list[str]) -> list[str]:
    """Lowercase, keep only ^[a-zæøå-]+$, drop dupes/numbers/parens/metadata.

    Returns a deterministically sorted, de-duplicated list.
    """
    log.info("PHASE 2: cleaning word list")
    seen: set[str] = set()
    rejected = 0
    for line in raw_lines:
        word = line.strip().lower()
        if not word:
            continue
        if not WORD_RE.match(word):
            rejected += 1
            log.debug("PHASE 2 reject (bad shape): %r", line)
            continue
        seen.add(word)

    words = sorted(seen)   # deterministic alphabetical sort
    CLEAN_TXT.write_text("\n".join(words) + "\n", encoding="utf-8")
    log.info(
        "PHASE 2 done: %d unique valid words, %d lines rejected -> %s",
        len(words), rejected, CLEAN_TXT.name,
    )
    return words


def phase3_qc(words: list[str]) -> list[str]:
    """Enforce length bounds / no spaces / valid chars, then write a stats report."""
    log.info("PHASE 3: quality control")
    kept: list[str] = []
    removed_short = removed_long = removed_space = removed_invalid = 0
    for w in words:
        if " " in w or "\t" in w:
            removed_space += 1
            log.debug("PHASE 3 reject (space): %r", w)
            continue
        if not WORD_RE.match(w):
            removed_invalid += 1
            log.debug("PHASE 3 reject (invalid char): %r", w)
            continue
        if len(w) < MIN_LEN:
            removed_short += 1
            log.debug("PHASE 3 reject (too short): %r", w)
            continue
        if len(w) > MAX_LEN:
            removed_long += 1
            log.debug("PHASE 3 reject (too long): %r", w)
            continue
        kept.append(w)

    # De-dupe defensively (should already be unique from phase 2).
    before = len(kept)
    kept = sorted(set(kept))
    dupes_removed = before - len(kept)

    shortest = min(kept, key=len) if kept else ""
    longest = max(kept, key=len) if kept else ""

    report = [
        "=== QUALITETSKONTROL / QC-RAPPORT ===",
        f"Antal ord (final):        {len(kept)}",
        f"Korteste ord:             {shortest} ({len(shortest)} tegn)",
        f"Længste ord:              {longest} ({len(longest)} tegn)",
        f"Antal dubletter fjernet:  {dupes_removed}",
        "",
        "--- Fjernet i QC ---",
        f"For korte (<{MIN_LEN}):          {removed_short}",
        f"For lange (>{MAX_LEN}):         {removed_long}",
        f"Indeholdt mellemrum:      {removed_space}",
        f"Ugyldige tegn:            {removed_invalid}",
        "",
    ]
    QC_REPORT.write_text("\n".join(report) + "\n", encoding="utf-8")
    # Rewrite the clean list so it reflects QC results.
    CLEAN_TXT.write_text("\n".join(kept) + "\n", encoding="utf-8")
    log.info("PHASE 3 done: %d words after QC -> %s", len(kept), QC_REPORT.name)
    for line in report:
        if line:
            log.info("  %s", line)
    return kept


def phase4_frequency(words: list[str]) -> None:
    """Write the dicttool source lines with a fixed frequency."""
    log.info("PHASE 4: assigning frequency f=%d", FREQUENCY)
    lines = [f"word={w},f={FREQUENCY}" for w in words]
    DICTSRC.write_text("\n".join(lines) + "\n", encoding="utf-8")
    log.info("PHASE 4 done: %d entries -> %s", len(lines), DICTSRC.name)


def main() -> int:
    try:
        raw = phase1_extract(PDF_PATH)
        cleaned = phase2_clean(raw)
        qc = phase3_qc(cleaned)
        phase4_frequency(qc)
    except Exception:
        log.exception("PIPELINE FAILED")
        return 1
    log.info("PIPELINE OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
