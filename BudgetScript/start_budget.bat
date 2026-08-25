@echo off
cd /d "%~dp0"
py -3 budget_fra_csv.py
if errorlevel 1 (
    echo.
    echo Der opstod en fejl - se teksten ovenfor.
    pause
)
