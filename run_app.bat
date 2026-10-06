@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m DailyFlowAnalyzer
) else (
    py -3 -m DailyFlowAnalyzer
)

if errorlevel 1 pause
