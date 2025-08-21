@echo off
REM Usage: run.bat MM YYYY PJEMAAT
REM Robust wrapper:
REM  - Discovers church_scheduler.py and Master.xlsx
REM  - Creates/uses a venv (writable path), installs requirements.txt if present
REM  - Runs the scheduler; OUTPUT_PATH env can override output filename
setlocal ENABLEDELAYEDEXPANSION

set MM=%~1
set YYYY=%~2
set PJ=%~3

if "%MM%"=="" goto :usage
if "%YYYY%"=="" goto :usage
if "%PJ%"=="" set PJ=3

REM Determine ROOT (one up from pythonScripts directory)
set ROOT=%~dp0
set ROOT=%ROOT:~0,-14%

REM Decide venv location: prefer ROOT\.venv, else LOCALAPPDATA
if "%VENV_DIR%"=="" (
  set VENV_DIR=%ROOT%\.venv
  if not exist "%VENV_DIR%" mkdir "%VENV_DIR%" 2>nul
  if not exist "%VENV_DIR%" (
    if "%LOCALAPPDATA%"=="" (
      set VENV_DIR=%USERPROFILE%\.church-scheduler\venv
    ) else (
      set VENV_DIR=%LOCALAPPDATA%\church-scheduler\venv
    )
    if not exist "%VENV_DIR%" mkdir "%VENV_DIR%"
  )
) else (
  if not exist "%VENV_DIR%" mkdir "%VENV_DIR%"
)

REM Bootstrap python to create venv
set PY_BOOT=%PYTHON%
if "%PY_BOOT%"=="" set PY_BOOT=python

REM Discover church_scheduler.py
set SCRIPT_PATH=
if exist "%ROOT%\church_scheduler.py" set SCRIPT_PATH=%ROOT%\church_scheduler.py
if "%SCRIPT_PATH%"=="" if exist "%~dp0..\church_scheduler.py" set SCRIPT_PATH=%~dp0..\church_scheduler.py
if "%SCRIPT_PATH%"=="" if exist "%~dp0church_scheduler.py" set SCRIPT_PATH=%~dp0church_scheduler.py
if "%SCRIPT_PATH%"=="" if exist "%ROOT%\pythonScripts\church_scheduler.py" set SCRIPT_PATH=%ROOT%\pythonScripts\church_scheduler.py
if "%SCRIPT_PATH%"=="" (
  echo ERROR: church_scheduler.py tidak ditemukan. Lokasi yang dicek:>&2
  echo  - %ROOT%\church_scheduler.py>&2
  echo  - %~dp0..\church_scheduler.py>&2
  echo  - %~dp0church_scheduler.py>&2
  echo  - %ROOT%\pythonScripts\church_scheduler.py>&2
  exit /b 3
)

REM Discover Master.xlsx
set MASTER_PATH=
if exist "%ROOT%\Master.xlsx" set MASTER_PATH=%ROOT%\Master.xlsx
if "%MASTER_PATH%"=="" if exist "%~dp0..\Master.xlsx" set MASTER_PATH=%~dp0..\Master.xlsx
if "%MASTER_PATH%"=="" if exist "%~dp0Master.xlsx" set MASTER_PATH=%~dp0Master.xlsx
if "%MASTER_PATH%"=="" if exist "%ROOT%\pythonScripts\Master.xlsx" set MASTER_PATH=%ROOT%\pythonScripts\Master.xlsx
if "%MASTER_PATH%"=="" (
  echo ERROR: Master.xlsx tidak ditemukan. Lokasi yang dicek:>&2
  echo  - %ROOT%\Master.xlsx>&2
  echo  - %~dp0..\Master.xlsx>&2
  echo  - %~dp0Master.xlsx>&2
  echo  - %ROOT%\pythonScripts\Master.xlsx>&2
  exit /b 4
)

REM Optional requirements.txt
set REQ_PATH=
if exist "%ROOT%\requirements.txt" set REQ_PATH=%ROOT%\requirements.txt
if "%REQ_PATH%"=="" if exist "%~dp0..\requirements.txt" set REQ_PATH=%~dp0..\requirements.txt
if "%REQ_PATH%"=="" if exist "%~dp0requirements.txt" set REQ_PATH=%~dp0requirements.txt
if "%REQ_PATH%"=="" if exist "%ROOT%\pythonScripts\requirements.txt" set REQ_PATH=%ROOT%\pythonScripts\requirements.txt

REM Create venv if missing
if not exist "%VENV_DIR%\Scripts\python.exe" (
  "%PY_BOOT%" -m venv "%VENV_DIR%"
)

set VENV_PY=%VENV_DIR%\Scripts\python.exe

REM Upgrade pip (best effort)
"%VENV_PY%" -m pip install --upgrade pip >nul 2>nul

REM Install requirements if present
if not "%REQ_PATH%"=="" (
  "%VENV_PY%" -m pip install -r "%REQ_PATH%"
)

REM Output path fallback
set OUTDIR=%ROOT%\output
if not exist "%OUTDIR%" mkdir "%OUTDIR%"
if "%OUTPUT_PATH%"=="" (
  set OUTFILE=%OUTDIR%\Jadwal-Bulanan.xlsx
) else (
  set OUTFILE=%OUTPUT_PATH%
)

"%VENV_PY%" "%SCRIPT_PATH" --master "%MASTER_PATH" --year %YYYY% --month %MM% --pjemaat-count %PJ% --output "%OUTFILE%"
exit /b %errorlevel%

:usage
echo Usage: run.bat MM YYYY PJEMAAT
exit /b 2
