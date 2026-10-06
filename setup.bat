@echo off
setlocal
cd /d "%~dp0"

py -3 -m venv .venv
if errorlevel 1 (
    echo Failed to create the virtual environment. Install Python 3.10+ with Tkinter first.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 (
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    pause
    exit /b 1
)

echo Setup complete. Run run_app.bat to start the application.
pause
