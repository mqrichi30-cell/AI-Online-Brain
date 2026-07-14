@echo off
setlocal enabledelayedexpansion
REM ============================================================================
REM  AWG - Ship with POs Report  ·  one-click draft generator (Windows)
REM  ---------------------------------------------------------------------------
REM  Generates one Outlook draft (.eml) per recipient from the Consolidations
REM  report. Matched Ship-tos go to their contact; unmatched go to
REM  pgcustservw2.im@pg.com with a "Please add contact" note at the top.
REM
REM  HOW TO USE
REM    * Drag the report .xlsx onto this file, OR
REM    * Double-click it and paste the report path when asked, OR
REM    * run_shipwith.cmd  "C:\path\report.xlsx"  "C:\path\contacts.xlsx"
REM
REM  Set CONTACTS below to your local/synced copy of
REM  "Regional Team - Contacts Data Base.xlsx" so you don't pass it every time.
REM ============================================================================

REM ---- EDIT THIS: path to your contacts workbook (or leave blank to be asked) --
set "CONTACTS=%OneDrive%\Regional Team - Contacts Data Base.xlsx"

cd /d "%~dp0"

REM ---- Report file: 1st argument, or prompt --------------------------------
set "REPORT=%~1"
if "%REPORT%"=="" (
    set /p "REPORT=Drag the report .xlsx here (or paste its path) and press Enter: "
)
REM strip surrounding quotes
set "REPORT=%REPORT:"=%"
if not exist "%REPORT%" (
    echo [ERROR] Report file not found: "%REPORT%"
    pause & exit /b 1
)

REM ---- Contacts file: 2nd argument overrides the default -------------------
if not "%~2"=="" set "CONTACTS=%~2"
set "CONTACTS=%CONTACTS:"=%"
if not exist "%CONTACTS%" (
    echo [WARN] Contacts workbook not found: "%CONTACTS%"
    echo        Continuing WITHOUT contacts - every draft will go to pgcustservw2.im@pg.com
    echo        with a "Please add contact" note.
    set "CONTACTS="
)

REM ---- Find Python --------------------------------------------------------
set "PY="
where py  >nul 2>&1 && set "PY=py -3"
if "%PY%"=="" ( where python >nul 2>&1 && set "PY=python" )
if "%PY%"=="" (
    echo [ERROR] Python is not installed or not on PATH.
    echo         Install it from https://www.python.org/downloads/ and re-run.
    pause & exit /b 1
)

REM ---- Ensure openpyxl -----------------------------------------------------
%PY% -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing required package openpyxl ...
    %PY% -m pip install --quiet --user openpyxl
)

REM ---- Run ----------------------------------------------------------------
set "OUT=%~dp0drafts_out"
if defined CONTACTS (
    %PY% "%~dp0shipwith_drafts.py" "%REPORT%" "%CONTACTS%" -o "%OUT%"
) else (
    %PY% "%~dp0shipwith_drafts.py" "%REPORT%" -o "%OUT%"
)
if errorlevel 1 ( echo [ERROR] Generation failed. & pause & exit /b 1 )

echo.
echo Done. Opening the drafts folder...
start "" "%OUT%"
echo Open each .eml in Outlook (they open as editable drafts), review, and Send.
pause
endlocal
