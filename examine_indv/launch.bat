@echo off
REM Launch the "examine a participant" web app.
REM
REM   launch.bat            (default port 8765)
REM   launch.bat 9000       (custom port)
REM
REM Windows equivalent of launch.sh.
setlocal

REM cd to this script's folder so relative paths work no matter where it's run.
set "HERE=%~dp0"
REM Strip trailing backslash.
if "%HERE:~-1%"=="\" set "HERE=%HERE:~0,-1%"
cd /d "%HERE%"

REM REPO is the parent directory of this script's folder.
for %%I in ("%HERE%\..") do set "REPO=%%~fI"

set "PORT=%~1"
if "%PORT%"=="" set "PORT=8765"

REM Find the project's virtualenv Python (Windows layout first, then Unix),
REM falling back to whatever python is on PATH.
if exist "%REPO%\.venv\Scripts\python.exe" (
  set "PY=%REPO%\.venv\Scripts\python.exe"
) else if exist "%REPO%\.venv\bin\python.exe" (
  set "PY=%REPO%\.venv\bin\python.exe"
) else (
  where python >nul 2>&1 && set "PY=python"
)

if not defined PY (
  echo ERROR: no Python found. Install Python or create the project .venv ^(run setup.bat^). 1>&2
  exit /b 1
)

echo Using Python: %PY%
echo Starting examine_indv on http://127.0.0.1:%PORT%  (Ctrl+C to stop)
"%PY%" "%HERE%\examine_app.py" "%PORT%"
exit /b %ERRORLEVEL%
