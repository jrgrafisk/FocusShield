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

## Results (compounds merged with the supplied HeliBoard base dictionary)

- **17 773** unique Danish compound nouns extracted from the PDF (76 pages), all
  matching `^[a-zæøå-]+$`, lengths 6–15 chars, **0 duplicates**.
- Merged with the supplied HeliBoard base `main_da.dict` (**178 449** words):
  - **8 752** compounds already existed in the base → base frequency kept.
  - **9 021** genuinely-new compounds added at the `f=128` placeholder.
  - **187 470** words in the merged dictionary.
- Compiled `main_da.dict`: **1 375 831 bytes**, AOSP format version 2.
- **Round-trip check passes**: decoding the binary yields exactly the 187 470
  merged words, æ/ø/å preserved; base frequencies verified intact
  (`arbejdsplads`=110, `og`=214, …).
- Rebuild is **byte-identical** (fixed header date → deterministic output).

The compound-only dictionary (no base) is still produced by running
`build_dict.py` with no `base_main_da.dict` present.

## Phase 5 — merging with a base dictionary

Phase 5 is conditional in the spec ("*Hvis source-format findes*"). To merge,
drop a base dictionary in `build/` and re-run `build_dict.py`:

- **Binary HeliBoard/AOSP `.dict`** → `build/base_main_da.dict` (preferred; it is
  decoded with `read_dict.py`, skipping bigrams/shortcuts), or
- **Text / dictsrc** → `build/original_da.txt` (`word`, `word,f=N`, or
  `word=..,f=..` lines).

**Merge policy** (chosen for this build): the base is authoritative — every word
already in the base keeps its real corpus frequency, and only genuinely-new
compounds are added at the `f=128` placeholder, so a placeholder never
overwrites real data. Duplicates are dropped deterministically. (Set
`KEEP_BASE_FREQUENCY = False` in `build_dict.py` for strict "highest frequency
wins" instead.)

> Note: `dåseøl` — one of the spec's example words — is in neither the PDF nor
> the supplied base dict, so it is not in the result. The validation report notes
> this rather than inventing the word.

## Installing into HeliBoard

`build/main_da.dict` can replace the Danish main dictionary: in HeliBoard,
*Settings → Languages → Danish → +*, and load the file, or push it to the app's
dictionary directory. Its header declares `locale=da`, `dictionary=main:da`.

## Third-party tool

`tools/dicttool_aosp.jar` is the AOSP dictionary compiler from
[remi0s/aosp-dictionary-tools](https://github.com/remi0s/aosp-dictionary-tools)
(`master`). SHA-256 `a8c5bd21f631ed0a92235d42d2fe83af5d70216172bf7e22781a9a946858237e`.
Invoked as `java -jar dicttool_aosp.jar makedict -s <wordlist.combined> -d main_da.dict`.
