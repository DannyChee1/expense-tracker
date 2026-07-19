@echo off
cd /d "%~dp0"

echo Expense Tracker
echo ===============

set "PYEXE="
where py >nul 2>nul && set "PYEXE=py -3"
if not defined PYEXE ( where python >nul 2>nul && set "PYEXE=python" )
if not defined PYEXE (
    echo.
    echo Python 3 is required but was not found.
    echo Install it from https://www.python.org/downloads/ and check "Add python.exe to PATH".
    echo Then double-click this again.
    echo.
    pause
    exit /b 1
)

set NEED=0
if not exist ".venv\Scripts\python.exe" set NEED=1
if "%NEED%"=="0" (
    .venv\Scripts\python.exe -c "import pandas, openpyxl, pdfplumber" 2>nul || set NEED=1
)

if "%NEED%"=="1" (
    echo First-time setup: installing ^(needs internet, ~1 minute^)...
    if exist ".venv" rmdir /s /q .venv
    %PYEXE% -m venv .venv || ( echo Could not create environment. & pause & exit /b 1 )
    .venv\Scripts\python.exe -m pip install --quiet --upgrade pip
    .venv\Scripts\python.exe -m pip install --quiet -r requirements.txt || ( echo Install failed. & pause & exit /b 1 )
)

.venv\Scripts\python.exe build_tracker.py

echo.
pause
