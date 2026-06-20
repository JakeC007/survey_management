@echo off
REM ============================================================
REM  DRY-RUN mode. Parses the response ZIP (if present), updates
REM  the tracker with any new completions, and prints the gift
REM  card list and reminder queue - but sends NO emails and
REM  makes no changes to follow-up timestamps.
REM
REM  Use this to sanity-check before a real or draft run.
REM ============================================================
cd /d "%~dp0"

if not exist "..\.venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup.bat in the parent folder first.
    pause
    exit /b 1
)

"..\.venv\Scripts\python.exe" manage_responses.py --dry-run
echo.
pause
