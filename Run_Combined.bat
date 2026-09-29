@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo     Recycle Report - Combined Data Generator
echo     Merges FTT + HFPA + Re-Inspection into one Excel file
echo ============================================================
echo.

REM Default output filename — edit here or pass as argument
set OUTPUT=Recycle_Report.xlsx

REM If a custom output filename is provided as first argument, use it
if not "%~1"=="" set OUTPUT=%~1

echo  Source folders:
echo    FTT            : %~dp0FTT
echo    HFPA           : %~dp0HFPA
echo    Re-Inspection  : %~dp0Re-Inspection
echo.
echo  Output file    : %OUTPUT%
echo.

python Combined.py --re-dir "Re-Inspection" --ftt-dir "FTT" --hfpa-dir "HFPA" --output "%OUTPUT%"

echo.
if %ERRORLEVEL% EQU 0 (
    echo [OK] Report generated successfully!
) else (
    echo [ERROR] Something went wrong. Check the messages above.
)
echo.
pause
