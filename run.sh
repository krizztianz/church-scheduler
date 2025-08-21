#!/usr/bin/env bash
# Usage: run.sh MM YYYY PJEMAAT
# Robust wrapper:
#  - Discovers church_scheduler.py and Master.xlsx in common locations
#  - Creates/uses a venv (writable path), installs requirements.txt if present
#  - Runs the scheduler with provided args and OUTPUT_PATH env if set
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="${SCRIPT_DIR%/pythonScripts}"

APP_NAME="church-scheduler"
# Where to put the venv (prefer project root if writable; else XDG data dir)
if [[ -n "${VENV_DIR:-}" ]]; then
  VENV_DIR="${VENV_DIR}"
elif [[ -w "$ROOT_DIR" ]]; then
  VENV_DIR="$ROOT_DIR/.venv"
else
  VENV_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/$APP_NAME/venv"
fi
mkdir -p "$VENV_DIR"

# Boot python to make venv (system python); then use venv python for the rest
PY_BOOT="${PYTHON:-python3}"

# ---- locate files ----
CANDIDATE_SCRIPTS=(
  "$ROOT_DIR/church_scheduler.py"
  "$SCRIPT_DIR/../church_scheduler.py"
  "$SCRIPT_DIR/church_scheduler.py"
  "$ROOT_DIR/pythonScripts/church_scheduler.py"
)
SCRIPT_PATH=""
for p in "${CANDIDATE_SCRIPTS[@]}"; do
  if [[ -f "$p" ]]; then SCRIPT_PATH="$p"; break; fi
done
if [[ -z "$SCRIPT_PATH" ]]; then
  echo "ERROR: church_scheduler.py tidak ditemukan. Cek lokasi berikut:" >&2
  printf ' - %s\n' "${CANDIDATE_SCRIPTS[@]}" >&2
  exit 3
fi

CANDIDATE_MASTER=(
  "$ROOT_DIR/Master.xlsx"
  "$SCRIPT_DIR/../Master.xlsx"
  "$SCRIPT_DIR/Master.xlsx"
  "$ROOT_DIR/pythonScripts/Master.xlsx"
)
MASTER_PATH=""
for p in "${CANDIDATE_MASTER[@]}"; do
  if [[ -f "$p" ]]; then MASTER_PATH="$p"; break; fi
done
if [[ -z "$MASTER_PATH" ]]; then
  echo "ERROR: Master.xlsx tidak ditemukan. Cek lokasi berikut:" >&2
  printf ' - %s\n' "${CANDIDATE_MASTER[@]}" >&2
  exit 4
fi

# Optional requirements.txt
CANDIDATE_REQ=(
  "$ROOT_DIR/requirements.txt"
  "$SCRIPT_DIR/../requirements.txt"
  "$SCRIPT_DIR/requirements.txt"
  "$ROOT_DIR/pythonScripts/requirements.txt"
)
REQ_PATH=""
for p in "${CANDIDATE_REQ[@]}"; do
  if [[ -f "$p" ]]; then REQ_PATH="$p"; break; fi
done

# ---- create venv if missing ----
if [[ ! -x "$VENV_DIR/bin/python" ]]; then
  "$PY_BOOT" -m venv "$VENV_DIR"
fi

VENV_PY="$VENV_DIR/bin/python"

# Upgrade pip (best effort)
"$VENV_PY" -m pip install --upgrade pip >/dev/null 2>&1 || true

# Install requirements if present
if [[ -n "$REQ_PATH" ]]; then
  "$VENV_PY" -m pip install -r "$REQ_PATH"
fi

# ---- args ----
MM="${1:-}"; YYYY="${2:-}"; PJ="${3:-3}"
if [[ -z "$MM" || -z "$YYYY" ]]; then
  echo "Usage: run.sh MM YYYY PJEMAAT" >&2
  exit 2
fi

OUT_DIR="$ROOT_DIR/output"
mkdir -p "$OUT_DIR"

OUT_FILE="${OUTPUT_PATH:-$OUT_DIR/Jadwal-Bulanan.xlsx}"

exec "$VENV_PY" "$SCRIPT_PATH" \
  --master "$MASTER_PATH" \
  --year "$YYYY" --month "$((10#$MM))" \
  --pjemaat-count "$PJ" \
  --output "$OUT_FILE"
