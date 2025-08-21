#!/usr/bin/env bash
# Universal Linux/macOS wrapper for JadwalPetugas
# - Works even if python3-venv is missing by falling back to vendor site-packages
# - Avoids system installs; installs into venv OR $DOCS/JadwalPetugas/vendor
# - Keeps Master.xlsx copy-prefer-Documents logic and output enforcement
# Exit codes:
# 2 usage, 3 no_script, 4 no_master_src, 5 no_pip/venv impossible, 6 deps missing, 7 python ok but no output, 8 no python

set -euo pipefail

MM="${1:-}"; YYYY="${2:-}"; PJ="${3:-3}"
if [[ -z "$MM" || -z "$YYYY" ]]; then
  echo "Usage: run.sh MM YYYY PJEMAAT" >&2
  exit 2
fi

# Resolve dirs
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
DOCS="${XDG_DOCUMENTS_DIR:-$HOME/Documents}"
APPDIR="$DOCS/JadwalPetugas"
CONFDIR="$APPDIR/config"
OUTDIR_DEFAULT="$APPDIR/output"
VENDORDIR="$APPDIR/vendor"
mkdir -p "$CONFDIR" "$OUTDIR_DEFAULT"

# Helpers
find_first() { for p in "$@"; do [[ -f "$p" ]] && { echo "$p"; return 0; }; done; return 1; }

# Discover script & master
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

# Pick python
if [[ -n "${PYTHON:-}" ]]; then
  PY="$PYTHON"
elif command -v python3 >/dev/null 2>&1; then
  PY="python3"
elif command -v python >/dev/null 2>&1; then
  PY="python"
else
  echo "ERROR: Python interpreter not found." >&2
  exit 8
fi

# Decide runtime: try venv first, else vendor site-packages
RUNTIME_MODE=""
VENV="${VENV_DIR:-}"
if [[ -z "$VENV" ]]; then
  if [[ -w "$ROOT_DIR" ]]; then
    VENV="$ROOT_DIR/.venv"
  else
    VENV="${XDG_DATA_HOME:-$HOME/.local/share}/jadwal-petugas/venv"
  fi
fi

VENV_PY="$VENV/bin/python"
set +e
"$PY" -m venv "$VENV" >/dev/null 2>&1
venv_rc=$?
set -e
if [[ $venv_rc -eq 0 && -x "$VENV_PY" ]]; then
  RUNTIME_MODE="venv"
else
  RUNTIME_MODE="vendor"
  mkdir -p "$VENDORDIR"
fi

# Ensure pip for chosen runtime
ensure_pip() {
  if [[ "$RUNTIME_MODE" == "venv" ]]; then
    "$VENV_PY" -m pip --version >/dev/null 2>&1 && return 0
    "$VENV_PY" -m ensurepip --upgrade >/dev/null 2>&1 || true
    "$VENV_PY" -m pip --version >/dev/null 2>&1
    return $?
  else
    "$PY" -m pip --version >/dev/null 2>&1 && return 0
    "$PY" -m ensurepip --upgrade >/dev/null 2>&1 || true
    "$PY" -m pip --version >/dev/null 2>&1
    return $?
  fi
}

# Install deps from requirements or fallback
install_deps() {
  local REQ=""
  REQ="$(find_first \
    "$ROOT_DIR/requirements.txt" \
    "$SCRIPT_DIR/requirements.txt" \
    "$ROOT_DIR/pythonScripts/requirements.txt")" || true

  if [[ "$RUNTIME_MODE" == "venv" ]]; then
    "$VENV_PY" -m pip install --upgrade pip >/dev/null 2>&1 || true
    if [[ -n "$REQ" ]]; then
      "$VENV_PY" -m pip install -r "$REQ"
    fi
    "$VENV_PY" - <<'PY' >/dev/null 2>&1 || "$VENV_PY" -m pip install --upgrade "pandas>=2.1" "openpyxl>=3.1" "python-dateutil>=2.8" "numpy>=1.26"
import pandas, openpyxl  # noqa: F401
PY
  else
    "$PY" -m pip install --upgrade pip >/dev/null 2>&1 || true
    if [[ -n "$REQ" ]]; then
      "$PY" -m pip install --target "$VENDORDIR" -r "$REQ"
    fi
    "$PY" - <<'PY' >/dev/null 2>&1 || "$PY" -m pip install --target "$VENDORDIR" --upgrade "pandas>=2.1" "openpyxl>=3.1" "python-dateutil>=2.8" "numpy>=1.26"
import pandas, openpyxl  # noqa: F401
PY
  fi
}

verify_deps() {
  if [[ "$RUNTIME_MODE" == "venv" ]]; then
    "$VENV_PY" - <<'PY'
import pandas, openpyxl  # noqa: F401
PY
  else
    PYTHONPATH="$VENDORDIR${PYTHONPATH:+:$PYTHONPATH}" "$PY" - <<'PY'
import sys, os
import pandas, openpyxl  # noqa: F401
PY
  fi
}

if ! ensure_pip; then
  # No pip at all. Try system packages first; if not present, error 5.
  if [[ "$RUNTIME_MODE" == "venv" ]]; then
    if ! "$VENV_PY" - <<'PY' >/dev/null 2>&1; then
      echo "ERROR: pip/ensurepip unavailable; install python3-venv or python3-pip." >&2
      exit 5
PY
    fi
  else
    if ! "$PY" - <<'PY' >/dev/null 2>&1; then
      echo "ERROR: pip/ensurepip unavailable; install python3-pip or enable pip." >&2
      exit 5
PY
    fi
  fi
fi

# Install and verify
install_deps || true
if ! verify_deps >/dev/null 2>&1; then
  echo "ERROR: Dependencies not installed (pandas/openpyxl missing)." >&2
  exit 6
fi

# Output path
OUT_FILE="${OUTPUT_PATH:-$OUTDIR_DEFAULT/Jadwal-Bulanan.xlsx}"
[[ -n "${WRAPPER_DEBUG:-}" ]] && {
  echo "[DEBUG] SCRIPT_DIR=$SCRIPT_DIR"
  echo "[DEBUG] ROOT_DIR=$ROOT_DIR"
  echo "[DEBUG] DOCUMENTS=$DOCS"
  echo "[DEBUG] CONFDIR=$CONFDIR"
  echo "[DEBUG] MASTER_PATH=$MASTER_PATH"
  echo "[DEBUG] RUNTIME_MODE=$RUNTIME_MODE"
  [[ "$RUNTIME_MODE" == "venv" ]] && echo "[DEBUG] VENV=$VENV" || echo "[DEBUG] VENDOR=$VENDORDIR"
  echo "[DEBUG] OUT_FILE=$OUT_FILE"
}

# Normalize month (strip leading zero)
if [[ "$MM" =~ ^0[0-9]$ ]]; then MM="${MM#0}"; fi

# Run python
set +e
if [[ "$RUNTIME_MODE" == "venv" ]]; then
  "$VENV_PY" "$SCRIPT_PATH" --master "$MASTER_PATH" --year "$YYYY" --month "$MM" --pjemaat-count "$PJ" --output "$OUT_FILE"
else
  PYTHONPATH="$VENDORDIR${PYTHONPATH:+:$PYTHONPATH}" "$PY" "$SCRIPT_PATH" --master "$MASTER_PATH" --year "$YYYY" --month "$MM" --pjemaat-count "$PJ" --output "$OUT_FILE"
fi
PY_RC=$?
set -e

# Enforce output presence
pick_newest() { local d="$1"; [[ -d "$d" ]] || return 1; ls -1t "$d"/*.xlsx 2>/dev/null | head -n1; }
if [[ ! -f "$OUT_FILE" ]]; then
  FOUND="$(pick_newest "$CONFDIR" || true)"
  [[ -z "$FOUND" ]] && FOUND="$(pick_newest "$SCRIPT_DIR" || true)"
  [[ -z "$FOUND" ]] && FOUND="$(pick_newest "$ROOT_DIR" || true)"
  [[ -z "$FOUND" ]] && FOUND="$(pick_newest "$ROOT_DIR/output" || true)"
  [[ -z "$FOUND" ]] && FOUND="$(pick_newest "$PWD" || true)"
  [[ -n "$FOUND" ]] && mv -f "$FOUND" "$OUT_FILE"
fi

if [[ ! -f "$OUT_FILE" ]]; then
  if [[ "$PY_RC" -eq 0 ]]; then
    echo "ERROR: Python exited OK but no output file was produced." >&2
    exit 7
  else
    exit "$PY_RC"
  fi
fi

exit "$PY_RC"
