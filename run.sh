#!/usr/bin/env bash
set -euo pipefail

MM="${1:-}"; YYYY="${2:-}"; PJ="${3:-3}"
if [[ -z "$MM" || -z "$YYYY" ]]; then
  echo "Usage: run.sh MM YYYY PJEMAAT" >&2
  exit 2
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

APP_NAME="church-scheduler"

if [[ -n "${VENV_DIR:-}" ]]; then
  VENV_DIR="${VENV_DIR}"
elif [[ -w "$ROOT_DIR" ]]; then
  VENV_DIR="$ROOT_DIR/.venv"
else
  VENV_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/$APP_NAME/venv"
fi
mkdir -p "$VENV_DIR"

PY_BOOT="${PYTHON:-python3}"

find_first() { for p in "$@"; do [[ -f "$p" ]] && { echo "$p"; return 0; }; done; return 1; }

SCRIPT_PATH="$(find_first \
  "$ROOT_DIR/church_scheduler.py" \
  "$SCRIPT_DIR/church_scheduler.py" \
  "$ROOT_DIR/pythonScripts/church_scheduler.py")" || {
  echo "ERROR: church_scheduler.py tidak ditemukan." >&2; exit 3; }

MASTER_PATH="$(find_first \
  "$ROOT_DIR/Master.xlsx" \
  "$SCRIPT_DIR/Master.xlsx" \
  "$ROOT_DIR/pythonScripts/Master.xlsx")" || {
  echo "ERROR: Master.xlsx tidak ditemukan." >&2; exit 4; }

REQ_PATH="$(find_first \
  "$ROOT_DIR/requirements.txt" \
  "$SCRIPT_DIR/requirements.txt" \
  "$ROOT_DIR/pythonScripts/requirements.txt")" || true

[[ -x "$VENV_DIR/bin/python" ]] || "$PY_BOOT" -m venv "$VENV_DIR"
VENV_PY="$VENV_DIR/bin/python"
"$VENV_PY" -m pip install --upgrade pip >/dev/null 2>&1 || true
[[ -n "${REQ_PATH:-}" ]] && "$VENV_PY" -m pip install -r "$REQ_PATH"

OUT_DIR="$ROOT_DIR/output"
mkdir -p "$OUT_DIR"
OUT_FILE="${OUTPUT_PATH:-$OUT_DIR/Jadwal-Bulanan.xlsx}"

if [[ -n "${WRAPPER_DEBUG:-}" ]]; then
  echo "[DEBUG] ROOT_DIR=$ROOT_DIR"
  echo "[DEBUG] SCRIPT_DIR=$SCRIPT_DIR"
  echo "[DEBUG] SCRIPT_PATH=$SCRIPT_PATH"
  echo "[DEBUG] MASTER_PATH=$MASTER_PATH"
  echo "[DEBUG] VENV_DIR=$VENV_DIR"
  echo "[DEBUG] OUT_FILE=$OUT_FILE"
fi

exec "$VENV_PY" "$SCRIPT_PATH" --master "$MASTER_PATH" --year "$YYYY" --month "$((10#$MM))" --pjemaat-count "$PJ" --output "$OUT_FILE"
