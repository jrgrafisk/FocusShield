# Danish compound-noun dictionary for HeliBoard / AOSP

This directory converts a PDF word list of Danish compound nouns
(*sammensatte navneord*, from **RO 2012**) into a binary
[HeliBoard](https://github.com/Helium314/HeliBoard) / AOSP dictionary
(`main_da.dict`).

The whole thing is reproducible: one command rebuilds every artifact
byte-for-byte, keeps all intermediate results, logs every step, never
overwrites the source, and stays UTF-8 with the Danish letters æ/ø/å intact.

## Quick start

```bash
scripts/run_all.sh path/to/sammensattenavneord.pdf
# outputs land in build/ , logs in logs/
```

Requirements: Python 3 with **PyMuPDF** (`pip install pymupdf`) and a **JRE**
(Java 8+) for the AOSP `dicttool`.

## Pipeline (matches the 8-phase spec)

| Phase | Script | Output |
|------:|--------|--------|
| 1 – Extract text from PDF | `build_pipeline.py` | `build/da_compounds_raw.txt` |
| 2 – Clean (`^[a-zæøå-]+$`, lowercase, sort, de-dup) | `build_pipeline.py` | `build/da_compounds.txt` |
| 3 – Quality control + statistics | `build_pipeline.py` | `build/qc_report.txt` |
| 4 – Assign frequency `f=128` | `build_pipeline.py` | `build/da_compounds.dictsrc` |
| 5 – Merge with an existing dictionary *(if provided)* | `build_dict.py` | `build/da_merged.dictsrc` |
| 6 – Compile to AOSP binary | `build_dict.py` | `build/da_wordlist.combined`, `build/main_da.dict` |
| 7 – Validate (binary round-trip, probe words, UTF-8) | `validate.py` | `build/validation_report.txt` |
| 8 – Benchmark (counts, sizes) | `validate.py` | `build/validation_report.txt` |

`read_dict.py` is a small pure-Python decoder for the AOSP static (Ver2)
dictionary format. It walks the compiled trie and pulls every word back out so
Phase 7 can prove the binary actually contains the words we put in (not just
that the file exists).

## Results for the supplied PDF

- **17 773** unique Danish compound nouns extracted (76 pages).
- All match `^[a-zæøå-]+$`, lengths 6–15 characters, **0 duplicates**.
- Compiled `main_da.dict`: **173 096 bytes**, AOSP format version 2, 25 118 trie
  nodes.
- **Round-trip check passes**: decoding the binary yields exactly the 17 773
  source words, æ/ø/å preserved.
- Rebuild is **byte-identical** (fixed header date → deterministic output).

## Phase 5 — merging with the existing HeliBoard Danish dictionary

Phase 5 is conditional in the spec ("*Hvis source-format findes*"). No base
`main_da.dict` / `original_da.txt` was available in this environment
(HeliBoard's prebuilt dictionaries live on `codeberg.org`, which was not
reachable), so the compound list stands alone here.

The merge logic is implemented and ready: drop a base word list at
`build/original_da.txt` (either plain `word` / `word,f=N` lines or dicttool
`word=..,f=..` lines) and re-run — `build_dict.py` will merge, drop duplicates,
and keep the **highest** frequency (existing frequency wins on ties), then
recompile.

> Note: a few example words in the spec (`arbejdsplads`, `arbejdsmarked`,
> `dåseøl`) are **not** in this particular PDF — they are ordinary base
> vocabulary that would come from that existing dictionary via Phase 5, not from
> the compound-noun source. The validation report calls this out explicitly
> instead of inventing them.

## Installing into HeliBoard

`build/main_da.dict` can replace the Danish main dictionary: in HeliBoard,
*Settings → Languages → Danish → +*, and load the file, or push it to the app's
dictionary directory. Its header declares `locale=da`, `dictionary=main:da`.

## Third-party tool

`tools/dicttool_aosp.jar` is the AOSP dictionary compiler from
[remi0s/aosp-dictionary-tools](https://github.com/remi0s/aosp-dictionary-tools)
(`master`). SHA-256 `a8c5bd21f631ed0a92235d42d2fe83af5d70216172bf7e22781a9a946858237e`.
Invoked as `java -jar dicttool_aosp.jar makedict -s <wordlist.combined> -d main_da.dict`.
