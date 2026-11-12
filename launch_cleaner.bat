@echo off
cd /d "%~dp0"
python app/server.py --open
if errorlevel 1 (
    echo.
    echo Failed to start HIS File Cleaner. Ensure Python and requirements.txt are installed.
    pause
)

