#!/bin/bash
# Simple runner for church_scheduler.py
# Usage:
#   ./run.sh [month] [year] [pjemaat_count]

MONTH=${1:-$(date +%m)}
YEAR=${2:-$(date +%Y)}
PJEMAAT=${3:-3}

VENV=venv
PYTHON=$VENV/bin/python

if [ ! -d "$VENV" ]; then
  echo "Virtualenv not found, creating..."
  python3 -m venv $VENV
  source $VENV/bin/activate
  pip install -r requirements.txt
else
  source $VENV/bin/activate
fi

OUTFILE="Jadwal-Bulanan-${YEAR}-${MONTH}"
if [ "$PJEMAAT" -eq 4 ]; then
  OUTFILE="${OUTFILE}-4jemaat.xlsx"
  $PYTHON church_scheduler.py --master Master.xlsx --year $YEAR --month $MONTH --pjemaat-count 4 --output $OUTFILE
else
  OUTFILE="${OUTFILE}.xlsx"
  $PYTHON church_scheduler.py --master Master.xlsx --year $YEAR --month $MONTH --pjemaat-count $PJEMAAT --output $OUTFILE
fi

echo "Generated $OUTFILE"
