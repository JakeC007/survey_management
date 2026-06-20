@echo off
REM ============================================================
REM  DRY-RUN. Prints the payment summary, the Pay list, and the
REM  Hold (flagged) list to the console, but writes NO files.
REM  Use this to sanity-check before updating the ledger/report.
REM ============================================================
cd /d "%~dp0"

if not exist "..\.venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup.bat in the parent folder first.
    pause
    exit /b 1
)

"..\.venv\Scripts\python.exe" manage_payments.py --dry-run
echo.
pause
