@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  KH Baektec Competitor Quarterly Report Generator
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python is not installed on this PC.
    echo Please install it from https://www.python.org/downloads/
    echo IMPORTANT: check "Add Python to PATH" during setup.
    echo.
    pause
    exit /b 1
)

python -c "import OpenDartReader" >nul 2>nul
if errorlevel 1 (
    echo Installing required libraries for the first time. Please wait...
    python -m pip install -r requirements.txt
)

python generate_report.py

pause
