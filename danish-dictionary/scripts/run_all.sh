#!/usr/bin/env bash
# Reproducible end-to-end build of the Danish compound-noun dictionary.
#
# Usage:
#   scripts/run_all.sh [path/to/sammensattenavneord.pdf]
#
# Runs all phases and leaves every intermediate result under build/ and logs
# under logs/. Originals are never overwritten. UTF-8 throughout.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(dirname "$HERE")"
PDF="${1:-/root/.claude/uploads/849ef4be-8006-50f7-90ac-bdcdc6367fa8/4d05ed66-sammensattenavneord.pdf}"

echo ">>> Phase 1-4: extract, clean, QC, frequency"
python3 "$HERE/build_pipeline.py" "$PDF"

echo ">>> Phase 5-6: merge (if base present) + compile to main_da.dict"
python3 "$HERE/build_dict.py"

echo ">>> Phase 7-8: validate + benchmark"
python3 "$HERE/validate.py"

echo ">>> Done. Outputs in $ROOT/build/ , logs in $ROOT/logs/"
