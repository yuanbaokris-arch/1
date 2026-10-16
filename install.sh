#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
setup_python="${BUS_BUDGET_PYTHON:-python3}"
"$setup_python" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 12) else "请先安装 Python 3.12 或更新版本。")'
if [ ! -x .venv/bin/python ]; then
  "$setup_python" -m venv --system-site-packages .venv
fi
PIP_NO_INPUT=1 PIP_DISABLE_PIP_VERSION_CHECK=1 PIP_CACHE_DIR="$PWD/data/pip-cache" \
  .venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -c 'import fastapi, uvicorn, multipart, lxml; print("Python依赖检查通过")'
.venv/bin/python - <<'PY'
from budget_engine import engine_status
status = engine_status()
if not status['available']:
    raise SystemExit('未找到可用的 LibreOffice。Mac 用户请将 LibreOffice.app 拖入“应用程序”并打开一次，再重试。')
print(status.get('version', 'LibreOffice检查通过'))
PY
