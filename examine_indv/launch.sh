#!/usr/bin/env bash
# Launch the "examine a participant" web app.
#
#   ./launch.sh            # default port 8765
#   ./launch.sh 9000       # custom port
#
# Works from any directory and on Git Bash/WSL (Windows .venv) or macOS/Linux.
set -euo pipefail

# cd to this script's folder so relative paths work no matter where it's run.
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$HERE"

REPO="$(dirname "$HERE")"
PORT="${1:-8765}"

# Find the project's virtualenv Python (Windows layout first, then Unix),
# falling back to whatever python is on PATH.
if [ -x "$REPO/.venv/Scripts/python.exe" ]; then
  PY="$REPO/.venv/Scripts/python.exe"
elif [ -x "$REPO/.venv/bin/python" ]; then
  PY="$REPO/.venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then
  PY="python3"
elif command -v python >/dev/null 2>&1; then
  PY="python"
else
  echo "ERROR: no Python found. Install Python or create the project .venv (run setup.bat)." >&2
  exit 1
fi

echo "Using Python: $PY"
echo "Starting examine_indv on http://127.0.0.1:$PORT  (Ctrl+C to stop)"
exec "$PY" "$HERE/examine_app.py" "$PORT"
