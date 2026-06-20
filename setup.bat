@echo off
REM ============================================================
REM  One-time setup: creates a shared virtual environment (.venv)
REM  in this folder and installs packages for both scripts.
REM  Run this ONCE before using either consent_management/ or
REM  response_management/ runners.
REM ============================================================
cd /d "%~dp0"

echo.
echo Creating virtual environment in .venv ...
py -3 -m venv .venv
if errorlevel 1 (
    echo.
    echo Could not find the "py" launcher. Trying "python" instead...
    python -m venv .venv
)
if not exist ".venv\Scripts\python.exe" (
    echo.
    echo ERROR: virtual environment was not created.
    echo Make sure Python 3 is installed and on your PATH, then run setup.bat again.
    pause
    exit /b 1
)

echo.
echo Installing packages ...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: package install failed. See the messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Setup complete.
echo  Next: edit config.yaml with your survey link, sender info,
echo  and shared mailbox. Then use the runners in:
echo    consent_management\  -- for sending invitations
echo    response_management\ -- for reminders and completions
echo ============================================================
echo.
pause
