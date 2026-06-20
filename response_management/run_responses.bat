@echo off
REM ============================================================
REM  SEND reminders for real. Emails every eligible participant
REM  who hasn't completed the survey and is past a reminder
REM  threshold, then updates the tracker.
REM
REM  Before running:
REM    1. Make sure classic Outlook is open and signed in.
REM    2. Optionally drop the response survey ZIP in ingest\.
REM ============================================================
cd /d "%~dp0"

if not exist "..\.venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup.bat in the parent folder first.
    pause
    exit /b 1
)

"..\.venv\Scripts\python.exe" manage_responses.py
echo.
pause
