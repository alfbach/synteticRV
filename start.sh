#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"

VENV="${VENV:-$ROOT/.venv}"
PY="${PY:-python3}"

if [[ ! -d "$VENV" ]]; then
  echo "Creating virtual environment: $VENV"
  "$PY" -m venv "$VENV"
fi

# shellcheck source=/dev/null
source "$VENV/bin/activate"

python -m pip install -q -r "$ROOT/requirements.txt"

echo "Starting synteticRV at http://127.0.0.1:8765 (Ctrl+C to stop)"
exec python "$ROOT/app.py"
