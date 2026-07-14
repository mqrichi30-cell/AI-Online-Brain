@echo off
setlocal enabledelayedexpansion
REM ============================================================================
REM  AWG - Ship with POs Report  ·  ONE-TIME setup of everything OUTSIDE Power
REM  Automate. Run this once, then import the .zip into Power Automate.
REM  ---------------------------------------------------------------------------
REM  It provisions the two prerequisites the flow needs:
REM    1. A fixed temp workbook  ->  <OneDrive>\AWG\ShipWith_Temp.xlsx
REM    2. The Office Script code copied to your clipboard, ready to paste into
REM       Excel on the web (Automate -> New Script -> paste -> save as
REM       "shipWithReport").
REM  Just double-click this file.
REM ============================================================================

cd /d "%~dp0"
set "SCRIPT_TS=%~dp0..\office-scripts\shipWithReport.ts"
set "TEMPLATE_XLSX=%~dp0..\power-automate\assets\ShipWith_Temp.xlsx"

echo ============================================================
echo  AWG - Ship with POs Report : setup (outside Power Automate)
echo ============================================================
echo.

REM ---- 1. Locate OneDrive --------------------------------------------------
set "OD="
if defined OneDriveCommercial if exist "%OneDriveCommercial%" set "OD=%OneDriveCommercial%"
if not defined OD if defined OneDrive if exist "%OneDrive%" set "OD=%OneDrive%"
if not defined OD if exist "%USERPROFILE%\OneDrive" set "OD=%USERPROFILE%\OneDrive"

if not defined OD (
    echo [WARN] Could not find your OneDrive folder automatically.
    set /p "OD=Paste the full path to your OneDrive folder and press Enter: "
    set "OD=!OD:"=!"
)

if not exist "%OD%" (
    echo [ERROR] OneDrive folder not found: "%OD%"
    echo         Create the temp file manually - see the printed steps below.
    goto :script_step
)

REM ---- 1a. Create <OneDrive>\AWG\ShipWith_Temp.xlsx ------------------------
set "AWGDIR=%OD%\AWG"
if not exist "%AWGDIR%" mkdir "%AWGDIR%"
if exist "%TEMPLATE_XLSX%" (
    copy /y "%TEMPLATE_XLSX%" "%AWGDIR%\ShipWith_Temp.xlsx" >nul
    echo [OK] Temp workbook created:  "%AWGDIR%\ShipWith_Temp.xlsx"
    echo      ^(OneDrive will sync it to the cloud - wait for the green check.^)
) else (
    echo [WARN] Template not found: "%TEMPLATE_XLSX%"
    echo        Put ANY empty .xlsx at "%AWGDIR%\ShipWith_Temp.xlsx" instead.
)
echo.

:script_step
REM ---- 2. Copy the Office Script to the clipboard --------------------------
if exist "%SCRIPT_TS%" (
    type "%SCRIPT_TS%" | clip
    echo [OK] Office Script copied to your CLIPBOARD (ready to paste).
) else (
    echo [WARN] Script file not found: "%SCRIPT_TS%"
    echo        Open office-scripts\shipWithReport.ts and copy it manually.
)
echo.

REM ---- 3. Open Excel on the web -------------------------------------------
echo Opening Excel on the web so you can paste the script...
start "" "https://excel.office.com/"
echo.

echo ============================================================
echo  NEXT STEPS
echo ============================================================
echo.
echo  A) CREATE THE OFFICE SCRIPT (one time)
echo     1. In Excel on the web, open ANY workbook.
echo     2. Automate  ^>  New Script.
echo     3. Select all, DELETE, then PASTE (Ctrl+V) - the script is on your
echo        clipboard.
echo     4. Rename the script to exactly:  shipWithReport   and Save.
echo.
echo  B) CONFIRM THE CONTACTS DATA BASE IS A TABLE (one time)
echo     - Open "Regional Team - Contacts Data Base.xlsx".
echo     - Make sure the contacts are inside a formatted Table (Insert ^> Table).
echo     - Note the table name and that it has the Ship-to name and Email columns.
echo.
echo  C) IMPORT THE FLOW
echo     - Power Automate ^> My flows ^> Import ^> Import Package (Legacy).
echo     - Upload:  power-automate\AWG-ShipWithPOs-Repot.zip
echo     - Pick connections: Office 365 Outlook (pgcustservw2), Excel Online
echo       (Business), OneDrive for Business.
echo     - Open the flow and bind the files/script:
echo         * Update temp report file  ^> File:  /AWG/ShipWith_Temp.xlsx
echo         * List contacts            ^> the Contacts Data Base + its table
echo         * Run script               ^> File: /AWG/ShipWith_Temp.xlsx,
echo                                      Script: shipWithReport
echo.
echo  Done. After that, every "AWG - Ship with POs Repot" email creates the
echo  drafts automatically in the pgcustservw2 mailbox.
echo.
echo  (Optional) To preview/generate drafts locally without waiting for an
echo  email, use  run_shipwith.cmd  instead.
echo.
pause
endlocal
