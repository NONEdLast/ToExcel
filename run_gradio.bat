@echo off

REM Set working directory to script location
cd /d "%~dp0"

REM Check if embedded Python exists
if not exist "runtime\python.exe" (
    echo Error: Embedded Python not found in runtime directory!
    pause
    exit /b 1
)

echo Starting Gradio app with embedded Python...

REM Run the app with embedded Python
runtime\python.exe gradio_app.py