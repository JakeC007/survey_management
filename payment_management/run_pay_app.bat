@echo off
REM Payment console - double-click to launch.
REM Uses the shared venv created by survey_management\setup.bat.
REM Requires classic Outlook (NOT New Outlook) open + signed in for the email step.
cd /d "%~dp0"

if exist "..\.venv\Scripts\python.exe" (
    "..\.venv\Scripts\python.exe" pay_app.py
) else (
    echo Shared venv not found at ..\.venv
    echo Run setup.bat in the survey_management folder first.
    echo Falling back to system Python...
    python pay_app.py
)

pause
