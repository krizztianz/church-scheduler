@echo off
setlocal ENABLEDELAYEDEXPANSION

set "MM=%~1"
set "YYYY=%~2"
set "PJ=%~3"

if "%MM%"=="" goto :usage
if "%YYYY%"=="" goto :usage
if "%PJ%"=="" set "PJ=3"

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%\..") do set "ROOT=%%~fI"

set "DOCS=%USERPROFILE%\Documents"
set "APPDIR=%DOCS%\JadwalPetugas"
set "CONFDIR=%APPDIR%\config"
set "OUTDIR_DEFAULT=%APPDIR%\output"
if not exist "%CONFDIR%" mkdir "%CONFDIR%"
if not exist "%OUTDIR_DEFAULT%" mkdir "%OUTDIR_DEFAULT%"

set "SCRIPT_PATH="
if exist "%ROOT%\church_scheduler.py" set "SCRIPT_PATH=%ROOT%\church_scheduler.py"
if "%SCRIPT_PATH%"=="" if exist "%SCRIPT_DIR%\church_scheduler.py" set "SCRIPT_PATH=%SCRIPT_DIR%\church_scheduler.py"
if "%SCRIPT_PATH%"=="" if exist "%ROOT%\pythonScripts\church_scheduler.py" set "SCRIPT_PATH=%ROOT%\pythonScripts\church_scheduler.py"
if "%SCRIPT_PATH%"=="" goto :no_script

set "MASTER_DOCS=%CONFDIR%\Master.xlsx"
if not exist "%MASTER_DOCS%" (
  set "MASTER_SRC="
  if exist "%ROOT%\Master.xlsx" set "MASTER_SRC=%ROOT%\Master.xlsx"
  if "%MASTER_SRC%"=="" if exist "%SCRIPT_DIR%\Master.xlsx" set "MASTER_SRC=%SCRIPT_DIR%\Master.xlsx"
  if "%MASTER_SRC%"=="" if exist "%ROOT%\pythonScripts\Master.xlsx" set "MASTER_SRC=%ROOT%\pythonScripts\Master.xlsx"
  if "%MASTER_SRC%"=="" goto :no_master_src
  copy /Y "%MASTER_SRC%" "%MASTER_DOCS%" >nul
)
set "MASTER_PATH=%MASTER_DOCS%"

if "%VENV_DIR%"=="" (
  set "VENV_DIR=%ROOT%\.venv"
  if not exist "%VENV_DIR%" (
    if defined LOCALAPPDATA (
      set "VENV_DIR=%LOCALAPPDATA%\jadwal-petugas\venv"
    ) else (
      set "VENV_DIR=%USERPROFILE%\.jadwal-petugas\venv"
    )
  )
)
if not exist "%VENV_DIR%" mkdir "%VENV_DIR%"

set "PY_BOOT=%PYTHON%"
if "%PY_BOOT%"=="" set "PY_BOOT=python"
if not exist "%VENV_DIR%\Scripts\python.exe" (
  "%PY_BOOT%" -m venv "%VENV_DIR%" || goto :venv_fail
)
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"

"%VENV_PY%" -m pip --version >nul 2>&1 || goto :venv_fail
if "%NO_PIP_INSTALL%"=="" "%VENV_PY%" -m pip install --upgrade pip >nul 2>&1
if "%NO_PIP_INSTALL%"=="" (
  set "REQ_PATH="
  if exist "%ROOT%\requirements.txt" set "REQ_PATH=%ROOT%\requirements.txt"
  if "%REQ_PATH%"=="" if exist "%SCRIPT_DIR%\requirements.txt" set "REQ_PATH=%SCRIPT_DIR%\requirements.txt"
  if "%REQ_PATH%"=="" if exist "%ROOT%\pythonScripts\requirements.txt" set "REQ_PATH=%ROOT%\pythonScripts\requirements.txt"
  if not "%REQ_PATH%"=="" "%VENV_PY%" -m pip install -r "%REQ_PATH%"
)

if "%OUTPUT_PATH%"=="" (
  set "OUTFILE=%OUTDIR_DEFAULT%\Jadwal-Bulanan.xlsx"
) else (
  set "OUTFILE=%OUTPUT_PATH%"
)

if defined WRAPPER_DEBUG (
  echo [DEBUG] ROOT=%ROOT%
  echo [DEBUG] SCRIPT_DIR=%SCRIPT_DIR%
  echo [DEBUG] DOCUMENTS=%DOCS%
  echo [DEBUG] CONFDIR=%CONFDIR%
  echo [DEBUG] MASTER_PATH=%MASTER_PATH%
  echo [DEBUG] VENV_DIR=%VENV_DIR%
  echo [DEBUG] OUTFILE=%OUTFILE%
)

if defined NO_RUN exit /b 0

"%VENV_PY%" "%SCRIPT_PATH%" --master "%MASTER_PATH%" --year %YYYY% --month %MM% --pjemaat-count %PJ% --output "%OUTFILE%"
exit /b %errorlevel%

:usage
echo Usage: run.bat MM YYYY PJEMAAT
exit /b 2

:no_script
echo ERROR: church_scheduler.py tidak ditemukan. >&2
exit /b 3

:no_master_src
echo ERROR: Master.xlsx tidak ditemukan di resources untuk fallback copy. >&2
exit /b 4

:venv_fail
echo ERROR: Gagal membuat/menggunakan virtualenv. >&2
exit /b 5
