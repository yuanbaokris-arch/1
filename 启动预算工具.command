#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
export PATH="/Library/Frameworks/Python.framework/Versions/Current/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"
launcher_python=""
for candidate in python3.13 python3.12 python3.14 python3.15 python3; do
  if command -v "$candidate" >/dev/null 2>&1 && "$candidate" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)' >/dev/null 2>&1; then
    launcher_python="$(command -v "$candidate")"
    break
  fi
done
if [ -z "$launcher_python" ]; then
  echo '未找到 Python 3.12 或更新版本。'
  echo '请先从 https://www.python.org/downloads/macos/ 安装 Python，再打开本文件。'
  read -r -p '按回车键关闭……' || true
  exit 1
fi
exec "$launcher_python" launch.py
