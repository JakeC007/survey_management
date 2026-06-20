@echo off
REM ============================================================
REM  DRAFT mode (recommended for your first batch).
REM  Stages each invite in Outlook's Drafts folder instead of
REM  sending. No mail goes out until you open Drafts and send
REM  them yourself. The ingest ZIP is KEPT in draft mode, so you
REM  can re-run if review turns up a problem.
REM  Note: drafted people are recorded as handled, so do NOT
REM  follow a draft run with run.bat - send the drafts from
REM  Outlook instead.
REM ============================================================
cd /d "%~dp0"

if not exist "..\.venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup.bat in the parent folder first.
    pause
    exit /b 1
)

"..\.venv\Scripts\python.exe" send_survey_emails.py --draft
echo.
pause
