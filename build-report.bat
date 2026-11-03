@echo off
setlocal
set PYTHONUTF8=1
cd /d "%~dp0"
if not exist "build\report-tools\Scripts\python.exe" (
    python -m venv build\report-tools
    if errorlevel 1 goto error
)
"build\report-tools\Scripts\python.exe" -m pip install uv==0.12.23
if errorlevel 1 goto error
"build\report-tools\Scripts\python.exe" -m uv sync --locked --extra dev
if errorlevel 1 goto error
powershell -NoProfile -File scripts\publish-report-windows.ps1
if errorlevel 1 goto error
echo Ready: artifacts\MonthlyReport-Windows-x64.zip
pause
exit /b 0
:error
echo Build failed. Install Python 3.12 and see the error above.
pause
exit /b 1
