#!/usr/bin/env bash
set -euo pipefail

MM="${1:-}"; YYYY="${2:-}"; PJ="${3:-3}"
if [[ -z "$MM" || -z "$YYYY" ]]; then
  echo "Usage: run.sh MM YYYY PJEMAAT" >&2
  exit 2
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

DOCS="${XDG_DOCUMENTS_DIR:-}"
if [[ -z "$DOCS" ]]; then DOCS="$HOME/Documents"; fi

APPDIR="$DOCS/JadwalPetugas"
CONFDIR="$APPDIR/config"
OUTDIR_DEFAULT="$APPDIR/output"
mkdir -p "$CONFDIR" "$OUTDIR_DEFAULT"

find_first() { for p in "$@"; do [[ -f "$p" ]] && { echo "$p"; return 0; }; done; return 1; }

SCRIPT_PATH="$(find_first \
  "$ROOT_DIR/church_scheduler.py" \
  "$SCRIPT_DIR/church_scheduler.py" \
  "$ROOT_DIR/pythonScripts/church_scheduler.py")" || { echo "ERROR: church_scheduler.py tidak ditemukan." >&2; exit 3; }

MASTER_DOCS="$CONFDIR/Master.xlsx"
if [[ ! -f "$MASTER_DOCS" ]]; then
  MASTER_SRC="$(find_first \
    "$ROOT_DIR/Master.xlsx" \
    "$SCRIPT_DIR/Master.xlsx" \
    "$ROOT_DIR/pythonScripts/Master.xlsx")" || { echo "ERROR: Master.xlsx tidak ditemukan di resources untuk fallback copy." >&2; exit 4; }
  cp -f "$MASTER_SRC" "$MASTER_DOCS"
fi
MASTER_PATH="$MASTER_DOCS"

APP_NAME="jadwal-petugas"
if [[ -n "${VENV_DIR:-}" ]]; then
  VENV="$VENV_DIR"
elif [[ -w "$ROOT_DIR" ]]; then
  VENV="$ROOT_DIR/.venv"
else
  VENV="${XDG_DATA_HOME:-$HOME/.local/share}/$APP_NAME/venv"
fi
mkdir -p "$VENV"

PY_BOOT="${PYTHON:-python3}"
[[ -x "$VENV/bin/python" ]] || "$PY_BOOT" -m venv "$VENV"
VENV_PY="$VENV/bin/python"

"$VENV_PY" -m pip install --upgrade pip >/dev/null 2>&1 || true
if [[ -z "${NO_PIP_INSTALL:-}" ]]; then
  REQ="$(find_first \
    "$ROOT_DIR/requirements.txt" \
    "$SCRIPT_DIR/requirements.txt" \
    "$ROOT_DIR/pythonScripts/requirements.txt")" || true
  [[ -n "${REQ:-}" ]] && "$VENV_PY" -m pip install -r "$REQ" || true
fi

OUT_FILE="${OUTPUT_PATH:-$OUTDIR_DEFAULT/Jadwal-Bulanan.xlsx}"

if [[ -n "${WRAPPER_DEBUG:-}" ]]; then
  echo "[DEBUG] ROOT_DIR=$ROOT_DIR"
  echo "[DEBUG] SCRIPT_DIR=$SCRIPT_DIR"
  echo "[DEBUG] DOCUMENTS=$DOCS"
  echo "[DEBUG] CONFDIR=$CONFDIR"
  echo "[DEBUG] MASTER_PATH=$MASTER_PATH"
  echo "[DEBUG] VENV=$VENV"
  echo "[DEBUG] OUT_FILE=$OUT_FILE"
fi

if [[ -n "${NO_RUN:-}" ]]; then exit 0; fi

exec "$VENV_PY" "$SCRIPT_PATH" --master "$MASTER_PATH" --year "$YYYY" --month "$((10#$MM))" --pjemaat-count "$PJ" --output "$OUT_FILE"
