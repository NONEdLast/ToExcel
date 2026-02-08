@echo off

REM Set working directory to script location
cd /d "%~dp0"

REM Check if lib directory exists
if not exist "lib" (
    echo Error: lib directory not found!
    echo Please install dependencies first using:
    echo py -3.14 -m pip install --target="lib" pandas openpyxl gradio
    pause
    exit /b 1
)

REM Check if required dependencies are installed in lib directory
set "missing_deps=false"

if not exist "lib\pandas" (
    echo Missing dependency: pandas
    set "missing_deps=true"
)

if not exist "lib\openpyxl" (
    echo Missing dependency: openpyxl
    set "missing_deps=true"
)

if not exist "lib\gradio" (
    echo Missing dependency: gradio
    set "missing_deps=true"
)

if "%missing_deps%" equ "true" (
    echo Please install missing dependencies first using:
    echo py -3.14 -m pip install --target="lib" pandas openpyxl gradio
    pause
    exit /b 1
)

REM Add local lib directory to Python path
set PYTHONPATH=lib;%PYTHONPATH%

echo Dependencies found, starting Gradio app...

REM Check if Python 3.14 is available
where py >nul 2>nul
if %errorlevel% equ 0 (
    REM Use Python 3.14 to run the app
    py -3.14 gradio_app.py
) else (
    echo Python launcher not found, using default python
    python gradio_app.py
)