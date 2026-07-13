#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Phase 5 (merge) + Phase 6 (compile) of the pipeline.

Phase 5 - Merge with an existing Danish dictionary *if one is provided*:
    The guide says "Hvis source-format findes" (if the source format exists).
    A base dictionary is optional. When present (as original_da.txt or a
    dictsrc-style file with 'word=...,f=N' or 'word,f=N' lines) we merge it with
    da_compounds.dictsrc, dropping duplicates and keeping the HIGHEST frequency,
    with the existing (base) frequency winning on ties. When absent, the compound
    list stands alone and this is logged.

Phase 6 - Compile to the AOSP/HeliBoard binary with dicttool:
    Build the '.combined' source (header + 'word=..,f=..' lines) and run
    'dicttool_aosp.jar makedict' to produce main_da.dict.

Originals are never overwritten; all outputs go under build/. UTF-8 only.
"""

import re
import sys
import subprocess
import logging
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build"
LOGS = ROOT / "logs"
TOOLS = ROOT / "tools"

DICTSRC = BUILD / "da_compounds.dictsrc"          # Phase 4 output (compounds)
MERGED_DICTSRC = BUILD / "da_merged.dictsrc"       # Phase 5 output
COMBINED = BUILD / "da_wordlist.combined"          # Phase 6 dicttool input
OUT_DICT = BUILD / "main_da.dict"                  # Phase 6 output
JAR = TOOLS / "dicttool_aosp.jar"

# Optional base dictionary (Phase 5). Present only if the user supplies it.
BASE_CANDIDATES = [BUILD / "original_da.txt", ROOT / "original_da.txt"]

FREQUENCY = 128
# Fixed build date -> reproducible, byte-stable dictionary header.
BUILD_DATE = int(datetime(2026, 7, 13, tzinfo=timezone.utc).timestamp())
VERSION = 1
LOCALE = "da"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler(LOGS / "build_dict.log", mode="w", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("build_dict")

LINE_RE = re.compile(r"^\s*word=(?P<w>[^,]+),\s*f=(?P<f>\d+)")
PLAIN_RE = re.compile(r"^(?P<w>[a-zæøå-]+)\s*(?:,\s*f=(?P<f>\d+))?\s*$")


def load_dictsrc(path: Path) -> dict[str, int]:
    """Parse a dictsrc / wordlist file into {word: freq}."""
    out: dict[str, int] = {}
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line:
            continue
        m = LINE_RE.match(line)
        if m:
            out[m.group("w").strip().lower()] = int(m.group("f"))
            continue
        m = PLAIN_RE.match(line)
        if m:
            out[m.group("w").strip().lower()] = int(m.group("f") or FREQUENCY)
    return out


def phase5_merge() -> dict[str, int]:
    log.info("PHASE 5: merge with existing dictionary (if present)")
    compounds = load_dictsrc(DICTSRC)
    log.info("  compounds loaded: %d", len(compounds))

    base_path = next((p for p in BASE_CANDIDATES if p.exists()), None)
    if base_path is None:
        log.info("  no base dictionary found (%s) -> compounds stand alone",
                 " / ".join(str(p.name) for p in BASE_CANDIDATES))
        merged = dict(compounds)
    else:
        base = load_dictsrc(base_path)
        log.info("  base dictionary %s loaded: %d words", base_path.name, len(base))
        merged = dict(base)  # start from base; base freq wins on ties
        added = 0
        raised = 0
        for w, f in compounds.items():
            if w not in merged:
                merged[w] = f
                added += 1
            elif f > merged[w]:
                # keep highest frequency; existing wins on equality
                merged[w] = f
                raised += 1
        log.info("  merge: %d new words added, %d freqs raised, base freq kept on ties",
                 added, raised)

    # Deterministic order.
    merged = dict(sorted(merged.items()))
    lines = [f"word={w},f={f}" for w, f in merged.items()]
    MERGED_DICTSRC.write_text("\n".join(lines) + "\n", encoding="utf-8")
    log.info("PHASE 5 done: %d total entries -> %s", len(merged), MERGED_DICTSRC.name)
    return merged


def phase6_compile(entries: dict[str, int]) -> None:
    log.info("PHASE 6: compile to AOSP binary with dicttool")
    if not JAR.exists():
        raise FileNotFoundError(f"dicttool jar missing: {JAR}")

    desc = (f"Danish compound nouns (RO 2012) + base; {len(entries)} words"
            if any(p.exists() for p in BASE_CANDIDATES)
            else f"Danish compound nouns from RO 2012 wordlist; {len(entries)} words")
    header = (f"dictionary=main:{LOCALE},locale={LOCALE},"
              f"description={desc},date={BUILD_DATE},version={VERSION}")

    # dicttool 'combined' format: header line, then ' word=..,f=..' (leading space).
    lines = [header]
    for w, f in sorted(entries.items()):
        lines.append(f" word={w},f={f}")
    COMBINED.write_text("\n".join(lines) + "\n", encoding="utf-8")
    log.info("  wrote combined source (%d entries) -> %s", len(entries), COMBINED.name)

    cmd = ["java", "-jar", str(JAR), "makedict",
           "-s", str(COMBINED), "-d", str(OUT_DICT)]
    log.info("  running: %s", " ".join(cmd))
    proc = subprocess.run(cmd, capture_output=True, text=True)
    for line in (proc.stdout + proc.stderr).splitlines():
        if "Picked up JAVA_TOOL" in line:
            continue
        log.info("    dicttool | %s", line)
    if proc.returncode != 0 or not OUT_DICT.exists():
        raise RuntimeError(f"dicttool makedict failed (rc={proc.returncode})")
    log.info("PHASE 6 done: %s (%d bytes)", OUT_DICT.name, OUT_DICT.stat().st_size)


def main() -> int:
    try:
        entries = phase5_merge()
        phase6_compile(entries)
    except Exception:
        log.exception("BUILD FAILED")
        return 1
    log.info("BUILD OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
