@echo off
REM ============================================================
REM  Build the payment ledger + unpaid report for real.
REM
REM  Reads the master tracker (data\participant_tracker_auto.xlsx)
REM  for who completed (consent + survey, merged on cid), runs the
REM  quality filter on the newest response-survey export in ingest\,
REM  upserts data\payment_tracker.xlsx (your hand-edited `paid`
REM  column is preserved), and writes data\payment_report_unpaid.xlsx
REM  split into Pay and Hold sheets.
REM
REM  Before running:
REM    1. Make sure manage_responses.py has stamped completions.
REM    2. Optionally drop the response survey ZIP in ingest\ so
REM       quality flags can be computed.
REM ============================================================
cd /d "%~dp0"

if not exist "..\.venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup.bat in the parent folder first.
    pause
    exit /b 1
)

"..\.venv\Scripts\python.exe" manage_payments.py
echo.
pause
