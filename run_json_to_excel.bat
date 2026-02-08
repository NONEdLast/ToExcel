@echo off

REM Set working directory to script location
cd /d "%~dp0"

REM Check if at least one argument is provided
if "%~1"=="" (
    echo Usage: %~n0.bat ^<json_file_path^>
    pause
    exit /b 1
)

REM Create lib directory if it doesn't exist
if not exist "lib" mkdir "lib"

REM Check if Python 3.14 is available
where py >nul 2>nul
if %errorlevel% equ 0 (
    echo Found Python launcher (py)
    REM Install/upgrade dependencies using Python 3.14
    echo Checking dependencies...
    py -3.14 -m pip install --upgrade --target="lib" pandas openpyxl
) else (
    echo Python launcher not found, using default python
    REM Install/upgrade dependencies using default Python
    echo Checking dependencies...
    pip install --upgrade --target="lib" pandas openpyxl
)

REM Add local lib directory to Python path
set PYTHONPATH=lib;%PYTHONPATH%

echo Running JSON to Excel conversion for file: %~1

REM Check if Python 3.14 is available
where py >nul 2>nul
if %errorlevel% equ 0 (
    REM Use Python 3.14 to run the script
    py -3.14 json_to_excel.py "%~1"
) else (
    REM Use default Python to run the script
    python json_to_excel.py "%~1"
)

pause