@echo off
cd /d "%~dp0"
echo Installing Python packages...
where py >nul 2>nul
if not errorlevel 1 (
    py -3 -m pip install -r requirements.txt
) else (
    python -m pip install -r requirements.txt
)
if errorlevel 1 (
    echo Installation failed. Please confirm Python 3 is installed.
) else (
    echo Installation completed. Double-click start.bat to launch.
)
pause

