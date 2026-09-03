#!/usr/bin/env sh
# Start Gridlint on http://127.0.0.1:8000
set -e
cd "$(dirname "$0")"
[ -d .venv ] || python3 -m venv .venv
. .venv/bin/activate
python -m pip install -q -r requirements.txt
echo "Gridlint is starting on http://127.0.0.1:8000"
exec python -m gridlint serve --port 8000
