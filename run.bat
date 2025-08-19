@echo off
REM Simple runner for church_scheduler.py on Windows
REM Usage:
REM   run.bat [MONTH] [YEAR] [PJEMAAT_COUNT]
REM Examples:
REM   run.bat            (uses current month/year, 3 P. Jemaat)
REM   run.bat 8 2025     (August 2025, 3 P. Jemaat)
REM   run.bat 8 2025 4   (August 2025, 4 P. Jemaat)

setlocal ENABLEDELAYEDEXPANSION

REM Determine defaults using PowerShell for reliability across locales
for /f "usebackq delims=" %%M in (`powershell -NoProfile -Command "(Get-Date).ToString('MM')"`) do set CURMONTH=%%M
for /f "usebackq delims=" %%Y in (`powershell -NoProfile -Command "(Get-Date).ToString('yyyy')"`) do set CURYEAR=%%Y

REM Read args or fallback to defaults
set MONTH=%1
if "%MONTH%"=="" set MONTH=%CURMONTH%

set YEAR=%2
if "%YEAR%"=="" set YEAR=%CURYEAR%

set PJEMAAT=%3
if "%PJEMAAT%"=="" set PJEMAAT=3

REM Ensure venv exists and dependencies installed
if not exist "venv\Scripts\python.exe" (
  echo [INFO] Creating virtual environment...
  py -3 -m venv venv
  if errorlevel 1 (
    echo [ERROR] Failed to create venv. Ensure Python 3 is installed and 'py' launcher is available.
    goto :eof
  )
  echo [INFO] Installing requirements...
  call venv\Scripts\activate.bat
  pip install -r requirements.txt
) else (
  call venv\Scripts\activate.bat
)

REM Build output filename
set OUTFILE=Jadwal-Bulanan-%YEAR%-%MONTH%

if "%PJEMAAT%"=="4" (
  set OUTFILE=%OUTFILE%-4jemaat.xlsx
  venv\Scripts\python.exe church_scheduler.py --master Master.xlsx --year %YEAR% --month %MONTH% --pjemaat-count %PJEMAAT% --output "%OUTFILE%"
) else (
  set OUTFILE=%OUTFILE%.xlsx
  venv\Scripts\python.exe church_scheduler.py --master Master.xlsx --year %YEAR% --month %MONTH% --pjemaat-count %PJEMAAT% --output "%OUTFILE%"
)

if errorlevel 1 (
  echo [ERROR] Generation failed.
) else (
  echo [OK] Generated "%OUTFILE%"
)

endlocal
