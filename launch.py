"""Local desktop launcher. Installs pinned dependencies and opens the local app."""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import signal
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import webbrowser

ROOT = Path(__file__).resolve().parent
LOCAL_HOST = "127.0.0.1"


def read_health(port):
    try:
        # A local readiness check should stay local even when a machine uses a proxy.
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open(f"http://{LOCAL_HOST}:{port}/api/health", timeout=0.6) as response:
            result = json.load(response)
            return result if isinstance(result, dict) else None
    except (OSError, ValueError, urllib.error.URLError):
        return None


def choose_port(instance_id, first=8765, last=8785):
    for port in range(first, last + 1):
        existing = read_health(port)
        if existing and existing.get("status") == "ok" and existing.get("instance_id") == instance_id:
            return port, True
        try:
            with socket.socket() as probe:
                probe.bind((LOCAL_HOST, port))
            return port, False
        except OSError:
            continue
    raise RuntimeError("没有找到空闲端口，请关闭本工具之前打开的终端窗口后重试。")


def wait_ready(process, port, instance_id, timeout=30):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError("服务启动失败。请查看程序文件夹 data/mac-startup.log。")
        health = read_health(port)
        if health and health.get("status") == "ok" and health.get("instance_id") == instance_id:
            return
        time.sleep(0.3)
    raise RuntimeError("服务未能按时启动。请查看程序文件夹 data/mac-startup.log。")


def show_browser(port):
    address = f"http://{LOCAL_HOST}:{port}"
    if not webbrowser.open(address, new=2):
        print("请在这台电脑的浏览器地址栏输入：" + address, flush=True)


def stop_requested(_signal, _frame):
    raise KeyboardInterrupt


def main():
    if sys.version_info < (3, 12):
        raise RuntimeError("请先安装 Python 3.12 或更新版本。")
    signal.signal(signal.SIGTERM, stop_requested)
    if hasattr(signal, "SIGHUP"):
        signal.signal(signal.SIGHUP, stop_requested)
    os.chdir(ROOT)
    data_root = Path(os.environ.get("BUS_BUDGET_DATA_DIR", str(ROOT / "data"))).resolve()
    instance_id = hashlib.sha256(str(data_root).encode()).hexdigest()[:16]
    port, running = choose_port(instance_id)
    if running:
        print("预算工具已经在运行，正在打开浏览器。", flush=True)
        show_browser(port)
        return 0
    print("正在检查运行环境。首次启动需要联网安装 Python 依赖，请保持窗口打开。", flush=True)
    environment = os.environ.copy()
    environment["BUS_BUDGET_PYTHON"] = sys.executable
    subprocess.run(["bash", str(ROOT / "install.sh")], cwd=ROOT, env=environment, check=True)
    runtime_python = ROOT / ".venv" / "bin" / "python"
    if not runtime_python.is_file():
        raise RuntimeError("未找到运行环境，请重新启动；如仍失败，请提供终端中的错误文字。")
    (ROOT / "data").mkdir(exist_ok=True)
    process = None
    try:
        with (ROOT / "data" / "mac-startup.log").open("a", encoding="utf-8") as log:
            process = subprocess.Popen(
                [str(runtime_python), "-m", "uvicorn", "app:app", "--host", LOCAL_HOST, "--port", str(port)],
                cwd=ROOT, env=environment, stdout=log, stderr=subprocess.STDOUT,
            )
            wait_ready(process, port, instance_id)
            print("预算工具已启动，正在打开浏览器。", flush=True)
            print("请保留这个终端窗口。按 Control+C 停止服务，已保存的方案会保留在 data 文件夹。", flush=True)
            show_browser(port)
            return process.wait()
    except KeyboardInterrupt:
        print("\n正在停止预算工具，已保存的数据会保留。", flush=True)
        return 0
    finally:
        if process is not None and process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (RuntimeError, OSError, subprocess.CalledProcessError) as error:
        print("\n启动未完成：" + str(error), file=sys.stderr, flush=True)
        print("请保留上面的错误文字，按回车键关闭。", file=sys.stderr, flush=True)
        try:
            input()
        except (EOFError, KeyboardInterrupt):
            pass
        raise SystemExit(1)
