#!/usr/bin/env bash

set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$PROJECT_DIR"

python_is_supported() {
  "$1" -c 'import sys; raise SystemExit(sys.version_info < (3, 11))' \
    >/dev/null 2>&1
}

find_python() {
  local candidate
  local candidates=(
    "/opt/homebrew/bin/python3.12"
    "/opt/homebrew/bin/python3.11"
    "/opt/homebrew/bin/python3"
    "/usr/local/bin/python3.12"
    "/usr/local/bin/python3.11"
    "/usr/local/bin/python3"
    "python3.12"
    "python3.11"
    "python3"
  )

  for candidate in "${candidates[@]}"; do
    if command -v "$candidate" >/dev/null 2>&1 && python_is_supported "$candidate"; then
      command -v "$candidate"
      return 0
    fi
  done
  return 1
}

SYSTEM_PYTHON="$(find_python || true)"
if [[ -z "$SYSTEM_PYTHON" ]]; then
  echo "Python 3.11 or newer is required."
  echo "Install it with: brew install python@3.12"
  exit 1
fi

if [[ -x .venv/bin/python ]] && ! python_is_supported .venv/bin/python; then
  VENV_VERSION="$(.venv/bin/python -c 'import sys; print(".".join(map(str, sys.version_info[:3])))')"
  echo "Existing .venv uses unsupported Python $VENV_VERSION."
  echo "Remove .venv and run this script again."
  exit 1
fi

if [[ ! -x .venv/bin/python ]]; then
  "$SYSTEM_PYTHON" -m venv .venv
fi

PYTHON="$PROJECT_DIR/.venv/bin/python"

"$PYTHON" -m pip install --upgrade pip
"$PYTHON" -m pip install -r requirements.txt

if [[ ! -f .env ]]; then
  cp .env.example .env
fi

if command -v npm >/dev/null 2>&1; then
  if (
    cd frontend
    npm install
    npm run build
  ); then
    echo "Frontend build complete."
  else
    echo "Frontend setup failed. The bundled prebuilt UI will be used."
  fi
else
  echo "npm was not found. The bundled prebuilt UI will be used."
fi

"$PYTHON" -m backend.scripts.init_db

echo
echo "Setup complete."
echo "Edit .env, then run: .venv/bin/python launcher.py"
