#!/bin/bash
set -e
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
  .venv/bin/pip install -r requirements.txt
fi

source .venv/bin/activate
PORT="${PORT:-8080}"
echo "RVtools Filter App: http://localhost:${PORT}"
exec python -c "from app import app; app.run(debug=True, host='0.0.0.0', port=${PORT})"
