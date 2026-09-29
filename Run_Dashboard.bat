@echo off
title Quality Intelligence Suite - Executive Dashboard
color 0B
setlocal
cd /d "%~dp0"

echo ===============================================================================
echo     QUALITY INTELLIGENCE SUITE - EXECUTIVE DASHBOARD & REPORT GENERATOR
echo ===============================================================================
echo.
echo  [+] Checking Python environment...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  [!] Python is not found in PATH.
    echo  [*] Starting installer to configure environment...
    echo.
    call "%~dp0Install_Dependencies.bat"
    exit /b
)

echo  [+] Starting Dashboard Application...
echo.

python Run_Dashboard.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [!] Application exited with code %ERRORLEVEL%.
    echo  [Tip] If packages like pandas or openpyxl are missing, run Install_Dependencies.bat
    pause
)

