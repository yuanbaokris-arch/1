#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
runtime_python=python3
if [ -x .venv/bin/python ]; then runtime_python=.venv/bin/python; fi
exec "$runtime_python" -m uvicorn app:app --host 0.0.0.0 --port "${PORT:-8000}"
