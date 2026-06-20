@echo off
REM ============================================================
REM  SEND for real. Emails every eligible person in the ingest
REM  folder's ZIP(s) that hasn't been emailed before, then
REM  deletes the ZIP. Make sure classic Outlook is open.
REM ============================================================
cd /d "%~dp0"

if not exist "..\.venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup.bat in the parent folder first.
    pause
    exit /b 1
)

"..\.venv\Scripts\python.exe" send_survey_emails.py
echo.
pause
