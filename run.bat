@echo off
cd /d "%~dp0"
python cafe_manager.py
if errorlevel 1 (
    echo.
    echo ERROR: Could not start. Make sure Python and openpyxl are installed.
    echo Run:  pip install openpyxl
    pause
)
