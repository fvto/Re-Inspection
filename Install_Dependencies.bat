@echo off
title Quality Intelligence Suite - Environment and Dependency Installer
color 0B
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ===============================================================================
echo     QUALITY INTELLIGENCE SUITE - ENVIRONMENT AND DEPENDENCY INSTALLER
echo ===============================================================================
echo.
echo  Checking system environment for Python and required libraries...
echo.

set "PYTHON_CMD="
set "PYTHON_DIR="

:: 1. Check if 'python' command is active and functional
python -c "import sys" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=python"
    goto :PYTHON_READY
)

:: 2. Check if Python Launcher ('py') is available
py -3 -c "import sys" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=py -3"
    goto :PYTHON_READY
)

:: 3. Scan common Windows installation directories (in case PATH is missing)
for %%P in (
    "%LocalAppData%\Programs\Python\Python313"
    "%LocalAppData%\Programs\Python\Python312"
    "%LocalAppData%\Programs\Python\Python311"
    "%LocalAppData%\Programs\Python\Python310"
    "C:\Program Files\Python313"
    "C:\Program Files\Python312"
    "C:\Program Files\Python311"
    "C:\Program Files\Python310"
) do (
    if exist "%%~P\python.exe" (
        "%%~P\python.exe" -c "import sys" >nul 2>&1
        if !ERRORLEVEL! EQU 0 (
            set "PYTHON_DIR=%%~P"
            set "PYTHON_CMD="%%~P\python.exe""
            set "PATH=%%~P;%%~P\Scripts;!PATH!"
            goto :PYTHON_READY
        )
    )
)

:: 4. Python not detected - Perform automated installation
echo  [!] Python 3 is not installed or not in PATH.
echo  [*] Starting automated setup of Python 3.12 (with pip, Tkinter, and PATH)...
echo.

:: Check winget first
winget --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo  [+] Installing Python 3.12 via Windows Package Manager (winget)...
    winget install --id Python.Python.3.12 -e --accept-package-agreements --accept-source-agreements
    if %ERRORLEVEL% EQU 0 (
        set "PYTHON_DIR=%LocalAppData%\Programs\Python\Python312"
        if exist "!PYTHON_DIR!\python.exe" (
            set "PYTHON_CMD="!PYTHON_DIR!\python.exe""
            set "PATH=!PYTHON_DIR!;!PYTHON_DIR!\Scripts;!PATH!"
            echo  [OK] Python installed successfully via winget.
            goto :PYTHON_READY
        )
    )
)

:: Download official installer if winget was unavailable or failed
echo  [+] Downloading official Python 3.12 installer from python.org...
set "INSTALLER=%TEMP%\python_312_installer.exe"

curl.exe -L -o "!INSTALLER!" "https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe" >nul 2>&1
if not exist "!INSTALLER!" (
    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe', '!INSTALLER!')" >nul 2>&1
)

if exist "!INSTALLER!" (
    echo  [+] Installing Python (quiet mode with Tkinter GUI, pip, and user PATH)...
    start /wait "" "!INSTALLER!" /passive InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1
    del /f /q "!INSTALLER!" >nul 2>&1
    
    set "PYTHON_DIR=%LocalAppData%\Programs\Python\Python312"
    if exist "!PYTHON_DIR!\python.exe" (
        set "PYTHON_CMD="!PYTHON_DIR!\python.exe""
        set "PATH=!PYTHON_DIR!;!PYTHON_DIR!\Scripts;!PATH!"
        echo  [OK] Python installed successfully.
        goto :PYTHON_READY
    )
)

:: If automated install could not proceed, provide clear manual instructions
echo.
echo  ===========================================================================
echo   [!] MANUAL ACTION REQUIRED:
echo       Could not automatically complete Python installation.
echo.
echo       1. Download Python from: https://www.python.org/downloads/
echo       2. Run the installer and CRITICAL: Check the box:
echo          [x] "Add python.exe to PATH"
echo       3. After installation finishes, re-run this script (Install_Dependencies.bat)
echo  ===========================================================================
echo.
pause
exit /b 1

:PYTHON_READY
echo  [OK] Python environment located:
!PYTHON_CMD! --version
echo.
echo  [*] Executing dependency installer (install_dependencies.py)...
echo.

!PYTHON_CMD! "%~dp0install_dependencies.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [!] Some dependencies could not be verified. Review the error details above.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo  ===========================================================================
echo   ALL PREREQUISITES AND LIBRARIES ARE INSTALLED AND READY!
echo  ===========================================================================
echo.
set /p RUN_NOW="  Launch Quality Intelligence Dashboard now? (Y/N) [Y]: "
if /i "!RUN_NOW!"=="" set "RUN_NOW=Y"
if /i "!RUN_NOW!"=="Y" (
    echo.
    echo  [+] Starting Dashboard Application...
    call "%~dp0Run_Dashboard.bat"
)

exit /b 0
