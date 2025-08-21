@echo off
REM Bulletproof Windows wrapper for JadwalPetugas (no fragile () blocks)
REM Usage: run.bat MM YYYY PJEMAAT

setlocal ENABLEEXTENSIONS

REM ---- Args ----
set "MM=%~1"
set "YYYY=%~2"
set "PJ=%~3"
if "%MM%"=="" goto :usage
if "%YYYY%"=="" goto :usage
if "%PJ%"=="" set "PJ=3"

REM ---- Dirs (no FOR blocks) ----
set "SCRIPT_DIR=%~dp0"
set "ROOT=%SCRIPT_DIR%.."
set "DOCS=%USERPROFILE%\Documents"

if defined WRAPPER_DEBUG echo [DEBUG] START WRAPPER
if defined WRAPPER_DEBUG echo [DEBUG] SCRIPT_DIR=%SCRIPT_DIR%
if defined WRAPPER_DEBUG echo [DEBUG] ROOT=%ROOT%
if defined WRAPPER_DEBUG echo [DEBUG] DOCUMENTS=%DOCS%

set "APPDIR=%DOCS%\JadwalPetugas"
set "CONFDIR=%APPDIR%\config"
set "OUTDIR_DEFAULT=%APPDIR%\output"
if not exist "%CONFDIR%" mkdir "%CONFDIR%"
if not exist "%OUTDIR_DEFAULT%" mkdir "%OUTDIR_DEFAULT%"

REM ---- Discover church_scheduler.py ----
set "SCRIPT_PATH="
if exist "%ROOT%\church_scheduler.py" set "SCRIPT_PATH=%ROOT%\church_scheduler.py"
if not defined SCRIPT_PATH if exist "%SCRIPT_DIR%\church_scheduler.py" set "SCRIPT_PATH=%SCRIPT_DIR%\church_scheduler.py"
if not defined SCRIPT_PATH if exist "%ROOT%\pythonScripts\church_scheduler.py" set "SCRIPT_PATH=%ROOT%\pythonScripts\church_scheduler.py"
if not defined SCRIPT_PATH goto :no_script

REM ---- Master.xlsx (prefer Documents; copy from resources if missing) ----
set "MASTER_DOCS=%CONFDIR%\Master.xlsx"
if exist "%MASTER_DOCS%" goto :have_master
set "MASTER_SRC="
if exist "%ROOT%\Master.xlsx" set "MASTER_SRC=%ROOT%\Master.xlsx"
if not defined MASTER_SRC if exist "%SCRIPT_DIR%\Master.xlsx" set "MASTER_SRC=%SCRIPT_DIR%\Master.xlsx"
if not defined MASTER_SRC if exist "%ROOT%\pythonScripts\Master.xlsx" set "MASTER_SRC=%ROOT%\pythonScripts\Master.xlsx"
if not defined MASTER_SRC goto :no_master_src
copy /Y "%MASTER_SRC%" "%MASTER_DOCS%" >nul

:have_master
set "MASTER_PATH=%MASTER_DOCS%"

REM ---- Venv (fixed path to avoid special chars) ----
if not defined VENV_DIR set "VENV_DIR=%LOCALAPPDATA%\jadwal-petugas\venv"
if not exist "%VENV_DIR%" mkdir "%VENV_DIR%"
set "PY_BOOT=%PYTHON%"
if not defined PY_BOOT set "PY_BOOT=python"
if not exist "%VENV_DIR%\Scripts\python.exe" "%PY_BOOT%" -m venv "%VENV_DIR%" || goto :venv_fail
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
"%VENV_PY%" -m pip --version >nul 2>&1 || goto :venv_fail

REM ---- Deps: requirements (if present) -> fallback -> verify ----
if not defined NO_PIP_INSTALL "%VENV_PY%" -m pip install --upgrade pip >nul 2>&1

if not defined NO_PIP_INSTALL goto :deps_check
goto :deps_verify

:deps_check
set "REQ_PATH="
if exist "%ROOT%\requirements.txt" set "REQ_PATH=%ROOT%\requirements.txt"
if not defined REQ_PATH if exist "%SCRIPT_DIR%\requirements.txt" set "REQ_PATH=%SCRIPT_DIR%\requirements.txt"
if not defined REQ_PATH if exist "%ROOT%\pythonScripts\requirements.txt" set "REQ_PATH=%ROOT%\pythonScripts\requirements.txt"

if defined REQ_PATH goto :deps_install_req
goto :deps_verify

:deps_install_req
"%VENV_PY%" -m pip install -r "%REQ_PATH%"
goto :deps_verify

:deps_verify
"%VENV_PY%" -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 goto :deps_fallback
goto :deps_ok

:deps_fallback
if defined WRAPPER_DEBUG echo [DEBUG] Installing pandas/openpyxl fallback...
"%VENV_PY%" -m pip install --upgrade pandas openpyxl python-dateutil numpy
"%VENV_PY%" -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo ERROR: Dependencies not installed (pandas/openpyxl missing). >&2
  exit /b 6
)

:deps_ok
REM ---- Desired OUTFILE ----
if "%OUTPUT_PATH%"=="" (
  set "OUTFILE=%OUTDIR_DEFAULT%\Jadwal-Bulanan.xlsx"
) else (
  set "OUTFILE=%OUTPUT_PATH%"
)

if defined WRAPPER_DEBUG echo [DEBUG] MASTER_PATH=%MASTER_PATH%
if defined WRAPPER_DEBUG echo [DEBUG] OUTFILE=%OUTFILE%
if defined WRAPPER_DEBUG echo [DEBUG] VENV_DIR=%VENV_DIR%

REM ---- Run python ----
"%VENV_PY%" "%SCRIPT_PATH%" --master "%MASTER_PATH%" --year %YYYY% --month %MM% --pjemaat-count %PJ% --output "%OUTFILE%"
set "PY_RC=%ERRORLEVEL%"

if exist "%OUTFILE%" exit /b %PY_RC%

REM ---- Output enforcement: move newest .xlsx to OUTFILE if Python saved elsewhere ----
set "FOUND="
call :pickNewest "%CONFDIR%"
if not defined FOUND call :pickNewest "%SCRIPT_DIR%"
if not defined FOUND call :pickNewest "%ROOT%"
if not defined FOUND call :pickNewest "%ROOT%\output"
if not defined FOUND call :pickNewest "%CD%"
if defined WRAPPER_DEBUG if defined FOUND echo [DEBUG] PICKED="%FOUND%"
if defined FOUND (
  if not exist "%OUTDIR_DEFAULT%" mkdir "%OUTDIR_DEFAULT%"
  move /Y "%FOUND%" "%OUTFILE%" >nul
)

if exist "%OUTFILE%" exit /b %PY_RC%
if "%PY_RC%"=="0" (
  echo ERROR: Python exited OK but no output file was produced. >&2
  exit /b 7
) else (
  exit /b %PY_RC%
)

goto :eof

:pickNewest
set "dir=%~1"
if not exist "%dir%" goto :eof
set "list=%TEMP%\_xlsxlist.txt"
dir /b /a:-d /o:-d "%dir%\*.xlsx" > "%list%" 2>nul
if not exist "%list%" goto :eof
set /p firstline=<"%list%"
del /f /q "%list%" >nul 2>&1
if not "%firstline%"=="" set "FOUND=%dir%\%firstline%"
goto :eof

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
