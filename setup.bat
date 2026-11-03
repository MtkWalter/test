@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    if errorlevel 1 goto error
)
".venv\Scripts\python.exe" -m pip install -e ".[dev]"
if errorlevel 1 goto error
echo Setup complete. Double-click run.bat to launch Excel Insights.
pause
exit /b 0
:error
echo Setup failed. Install Python 3.12 and check "Add python.exe to PATH", then retry.
pause
exit /b 1
