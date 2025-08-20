PYTHON=python3
VENV=venv
ACTIVATE=. $(VENV)/bin/activate

# Allow override: make run MONTH=8 YEAR=2025
MONTH?=$(shell date +%m)
YEAR?=$(shell date +%Y)

setup:
	$(PYTHON) -m venv $(VENV)
	$(ACTIVATE) && pip install -r requirements.txt

run:
	mkdir -p output
	$(ACTIVATE) && $(PYTHON) church_scheduler.py --master Master.xlsx --year $(YEAR) --month $(MONTH) --output output/Jadwal-Bulanan-$(YEAR)-$(MONTH).xlsx

run4:
	mkdir -p output
	$(ACTIVATE) && $(PYTHON) church_scheduler.py --master Master.xlsx --year $(YEAR) --month $(MONTH) --pjemaat-count 4 --output output/Jadwal-Bulanan-$(YEAR)-$(MONTH)-4jemaat.xlsx

clean:
	rm -rf $(VENV) output/*.xlsx
