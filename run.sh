#!/usr/bin/env bash
# Usage: run.sh MM YYYY PJEMAAT
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="${SCRIPT_DIR%/pythonScripts}"
PY="${PYTHON:-python3}"

MM="${1:-}"; YYYY="${2:-}"; PJ="${3:-3}"
if [[ -z "$MM" || -z "$YYYY" ]]; then
  echo "Usage: run.sh MM YYYY PJEMAAT" >&2
  exit 2
fi

OUT_DIR="$ROOT_DIR/output"
mkdir -p "$OUT_DIR"

# OUTPUT_PATH can be injected from Electron. Fallback to default name if missing.
if [[ -z "${OUTPUT_PATH:-}" ]]; then
  OUT_FILE="$OUT_DIR/Jadwal-Bulanan.xlsx"
else
  OUT_FILE="$OUTPUT_PATH"
fi

"$PY" "$ROOT_DIR/church_scheduler.py" \
  --master "$ROOT_DIR/Master.xlsx" \
  --year "$YYYY" --month "$((10#$MM))" \
  --pjemaat-count "$PJ" \
  --output "$OUT_FILE"
