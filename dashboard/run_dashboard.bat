@echo off
REM Survey pipeline console — double-click to launch.
REM Uses the shared venv created by survey_management\setup.bat.
cd /d "%~dp0"

if exist "..\.venv\Scripts\python.exe" (
    "..\.venv\Scripts\python.exe" app.py
) else (
    echo Shared venv not found at ..\.venv
    echo Run setup.bat in the survey_management folder first.
    echo Falling back to system Python...
    python app.py
)

pause
