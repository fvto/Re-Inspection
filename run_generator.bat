@echo off
setlocal
cd /d "%~dp0"

echo ===================================================
echo     Re-Inspection Excel Database Auto-Generator
echo ===================================================
echo.

if "%~1"=="" (
    echo Launching Graphical User Interface...
    python generate_database.py --gui
) else (
    echo Running command-line mode...
    python generate_database.py %*
)

echo.
pause
