@echo off
setlocal

REM Args
set "MM=%~1"
set "YYYY=%~2"
set "PJ=%~3"

if "%MM%"=="" goto :usage
if "%YYYY%"=="" goto :usage
if "%PJ%"=="" set "PJ=3"

REM Determine directories
set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%\..") do set "ROOT=%%~fI"

REM Venv location
if "%VENV_DIR%"=="" (
  set "VENV_DIR=%ROOT%\.venv"
  if not exist "%VENV_DIR%" (
    if defined LOCALAPPDATA (
      set "VENV_DIR=%LOCALAPPDATA%\church-scheduler\venv"
    ) else (
      set "VENV_DIR=%USERPROFILE%\.church-scheduler\venv"
    )
  )
)
if not exist "%VENV_DIR%" mkdir "%VENV_DIR%"

set "PY_BOOT=%PYTHON%"
if "%PY_BOOT%"=="" set "PY_BOOT=python"

REM Discover church_scheduler.py
set "SCRIPT_PATH="
if exist "%ROOT%\church_scheduler.py" set "SCRIPT_PATH=%ROOT%\church_scheduler.py"
if "%SCRIPT_PATH%"=="" if exist "%SCRIPT_DIR%\church_scheduler.py" set "SCRIPT_PATH=%SCRIPT_DIR%\church_scheduler.py"
if "%SCRIPT_PATH%"=="" if exist "%ROOT%\pythonScripts\church_scheduler.py" set "SCRIPT_PATH=%ROOT%\pythonScripts\church_scheduler.py"
if "%SCRIPT_PATH%"=="" goto :no_script

REM Discover Master.xlsx
set "MASTER_PATH="
if exist "%ROOT%\Master.xlsx" set "MASTER_PATH=%ROOT%\Master.xlsx"
if "%MASTER_PATH%"=="" if exist "%SCRIPT_DIR%\Master.xlsx" set "MASTER_PATH=%SCRIPT_DIR%\Master.xlsx"
if "%MASTER_PATH%"=="" if exist "%ROOT%\pythonScripts\Master.xlsx" set "MASTER_PATH=%ROOT%\pythonScripts\Master.xlsx"
if "%MASTER_PATH%"=="" goto :no_master

REM Optional requirements.txt
set "REQ_PATH="
if exist "%ROOT%\requirements.txt" set "REQ_PATH=%ROOT%\requirements.txt"
if "%REQ_PATH%"=="" if exist "%SCRIPT_DIR%\requirements.txt" set "REQ_PATH=%SCRIPT_DIR%\requirements.txt"
if "%REQ_PATH%"=="" if exist "%ROOT%\pythonScripts\requirements.txt" set "REQ_PATH=%ROOT%\pythonScripts\requirements.txt"

REM Create venv if missing
if not exist "%VENV_DIR%\Scripts\python.exe" (
  "%PY_BOOT%" -m venv "%VENV_DIR%" || goto :venv_fail
)

set "VENV_PY=%VENV_DIR%\Scripts\python.exe"

REM Upgrade pip & install requirements
"%VENV_PY%" -m pip --version >nul 2>&1 || goto :venv_fail
"%VENV_PY%" -m pip install --upgrade pip >nul 2>&1
if not "%REQ_PATH%"=="" "%VENV_PY%" -m pip install -r "%REQ_PATH%"

REM Output path
set "OUTDIR=%ROOT%\output"
if not exist "%OUTDIR%" mkdir "%OUTDIR%"
if "%OUTPUT_PATH%"=="" (
  set "OUTFILE=%OUTDIR%\Jadwal-Bulanan.xlsx"
) else (
  set "OUTFILE=%OUTPUT_PATH%"
)

REM Debug dump
if defined WRAPPER_DEBUG (
  echo [DEBUG] ROOT=%ROOT%
  echo [DEBUG] SCRIPT_DIR=%SCRIPT_DIR%
  echo [DEBUG] SCRIPT_PATH=%SCRIPT_PATH%
  echo [DEBUG] MASTER_PATH=%MASTER_PATH%
  echo [DEBUG] VENV_DIR=%VENV_DIR%
  echo [DEBUG] OUTFILE=%OUTFILE%
)

"%VENV_PY%" "%SCRIPT_PATH%" --master "%MASTER_PATH%" --year %YYYY% --month %MM% --pjemaat-count %PJ% --output "%OUTFILE%"
exit /b %errorlevel%

:usage
echo Usage: run.bat MM YYYY PJEMAAT
exit /b 2

:no_script
echo ERROR: church_scheduler.py tidak ditemukan. >&2
exit /b 3

:no_master
echo ERROR: Master.xlsx tidak ditemukan. >&2
exit /b 4

:venv_fail
echo ERROR: Gagal membuat/menggunakan virtualenv. >&2
exit /b 5
