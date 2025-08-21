@echo off
REM Usage: run.bat MM YYYY PJEMAAT
setlocal
set MM=%~1
set YYYY=%~2
set PJ=%~3

if "%MM%"=="" goto :usage
if "%YYYY%"=="" goto :usage
if "%PJ%"=="" set PJ=3

REM ROOT is ...\pythonScripts\, so go up one level
set ROOT=%~dp0
set ROOT=%ROOT:~0,-14%

set OUTDIR=%ROOT%\output
if not exist "%OUTDIR%" mkdir "%OUTDIR%"

set PY=%PYTHON%
if "%PY%"=="" set PY=python

REM OUTPUT_PATH may be provided by Electron. Use default if missing.
if "%OUTPUT_PATH%"=="" (
  set OUTFILE=%OUTDIR%\Jadwal-Bulanan.xlsx
) else (
  set OUTFILE=%OUTPUT_PATH%
)

"%PY%" "%ROOT%\church_scheduler.py" --master "%ROOT%\Master.xlsx" --year %YYYY% --month %MM% --pjemaat-count %PJ% --output "%OUTFILE%"
exit /b %errorlevel%

:usage
echo Usage: run.bat MM YYYY PJEMAAT
exit /b 2
