#!/usr/bin/env bash

set -euo pipefail

# Creates a local virtual environment and installs Python packages needed for
# ingress-analysis.py.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_BIN="${PYTHON_BIN:-python3}"
VENV_DIR="${SCRIPT_DIR}/.venv"

if ! command -v "${PYTHON_BIN}" >/dev/null 2>&1; then
  echo "Python executable '${PYTHON_BIN}' not found. Set PYTHON_BIN to override." >&2
  exit 1
fi

if [ ! -d "${VENV_DIR}" ]; then
  echo "Creating virtual environment at ${VENV_DIR}"
  "${PYTHON_BIN}" -m venv "${VENV_DIR}"
fi

VENV_PY="${VENV_DIR}/bin/python"

echo "Upgrading pip..."
"${VENV_PY}" -m pip install --upgrade pip

echo "Installing required libraries..."
"${VENV_PY}" -m pip install matplotlib

echo "Verifying Tkinter availability (required for the GUI)..."
if ! "${VENV_PY}" - <<'PY'
try:
    import tkinter  # noqa: F401
except Exception as exc:  # noqa: BLE001
    raise SystemExit(f"Tkinter is missing: {exc}")
PY
then
  cat <<'MSG' >&2
Tkinter is missing from this Python installation.
On Debian/Ubuntu, install it with: sudo apt-get install python3-tk
MSG
  exit 1
fi

echo "Done. Activate with: source ${VENV_DIR}/bin/activate"
