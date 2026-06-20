@echo off
REM ============================================================
REM  DRAFT mode. Stages each reminder email in Outlook's Drafts
REM  folder instead of sending immediately. Review the drafts
REM  before sending them from Outlook.
REM
REM  Note: once a reminder is staged as a Draft it is recorded
REM  in the tracker as handled. Do NOT follow a draft run with
REM  run_responses.bat - send the drafts from Outlook instead.
REM ============================================================
cd /d "%~dp0"

if not exist "..\.venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup.bat in the parent folder first.
    pause
    exit /b 1
)

"..\.venv\Scripts\python.exe" manage_responses.py --draft
echo.
pause
