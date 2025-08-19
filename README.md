# Church Scheduler

Python script to generate **monthly church service schedule** from a master Excel file (`Master.xlsx`).  
It distributes tasks fairly among elders (*Penatua*) and members (*Jemaat*) following church rules.

## Features

- Reads `Master.xlsx` dynamically (names, tasks, Penatua flag).
- Supports tasks: **DP/PA, W/PB, Persembahan, Kolektan, P. Jemaat, Lektor, Prokantor, Pemusik, Multimedia**.
- Configurable **P. Jemaat count** (default 3, max 4) via CLI parameter.
- Ensures:
  - Elders for elder-only tasks (DP/PA, W/PB, Persembahan, Kolektan).
  - P. Jemaat is mix of elders & members (e.g. 2 members + 1 elder).
  - Lektor, Prokantor, Pemusik → members only.
  - Multimedia → anyone with ability.
  - A person cannot get **multiple tasks in the same week**.
  - Non-elders are avoided from serving on **consecutive weeks** if pool is sufficient.
- Output: Excel file with one sheet per month, tasks as rows, Sundays as columns.

## Requirements

- Python 3.8+
- Packages:
  - pandas
  - openpyxl

Install via:
```bash
pip install -r requirements.txt
```

## Usage

### CLI

```bash
python church_scheduler.py --master Master.xlsx --year 2025 --month 8 --output Jadwal-Bulanan-2025-08.xlsx
```

Options:
- `--pjemaat-count N` → Number of **P. Jemaat** per week (1–4, default 3).
- `--repeat-non-elder` → Allow members to serve on consecutive weeks if pool is small.

### With Makefile

```bash
# Setup venv + install deps
make setup

# Run for current month/year (default 3 P. Jemaat)
make run

# Run for specific month/year
make run MONTH=8 YEAR=2025

# Run with 4 P. Jemaat
make run4 MONTH=8 YEAR=2025
```

### With run.sh (Linux/macOS)

```bash
# Default (current month/year, 3 P. Jemaat)
./run.sh

# August 2025, 4 P. Jemaat
./run.sh 8 2025 4
```

### With run.bat (Windows)

```bat
:: Default
run.bat

:: August 2025, 4 P. Jemaat
run.bat 8 2025 4
```

## Project Structure

```
church-scheduler/
├── church_scheduler.py   # main script
├── Master.xlsx           # master data input (not in repo)
├── requirements.txt
├── Makefile
├── run.sh
├── run.bat
└── README.md
```

## Output Folder

All generated schedules are saved under the `output/` directory.  
The folder will be created automatically if it does not exist.

Example after running for August 2025:

```
output/
├── Jadwal-Bulanan-2025-08.xlsx
├── Jadwal-Bulanan-2025-08-4jemaat.xlsx
└── Jadwal-Bulanan.xlsx    # default run
```

## Notes

- Do not commit `Master.xlsx` (personal data) or generated `.xlsx` schedules.
- See `.gitignore` for ignored files.
- Adjust task rules easily in `church_scheduler.py` if requirements change.

---

© 2025 - Kristian Andi - Church Scheduler Tool
