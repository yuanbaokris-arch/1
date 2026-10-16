"""Desktop readiness, port ownership, and Mac Calc discovery checks."""
import hashlib
import os
from pathlib import Path
import signal
import socket
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import Mock, patch

import budget_engine
import launch


class DesktopLauncherTests(unittest.TestCase):
    def test_reuses_only_our_healthy_instance(self):
        with patch.object(launch, 'read_health', return_value={'status': 'ok', 'instance_id': 'same'}):
            self.assertEqual(launch.choose_port('same', 8765, 8765), (8765, True))

    def test_occupied_other_service_is_not_reused(self):
        with socket.socket() as occupied:
            occupied.bind(('127.0.0.1', 0))
            port = occupied.getsockname()[1]
            with patch.object(launch, 'read_health', return_value={'status': 'ok', 'instance_id': 'other'}):
                with self.assertRaises(RuntimeError):
                    launch.choose_port('same', port, port)

    def test_failed_process_is_not_reported_ready(self):
        with self.assertRaisesRegex(RuntimeError, '启动失败'):
            launch.wait_ready(Mock(poll=Mock(return_value=1)), 8765, 'same')

    def test_configured_calc_path_must_be_executable(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'soffice'
            path.write_text('#!/bin/sh\nexit 0\n')
            path.chmod(0o600)
            with patch.dict(os.environ, {'BUS_BUDGET_SOFFICE': str(path)}):
                self.assertIsNone(budget_engine.spreadsheet_executable())
                path.chmod(0o700)
                self.assertEqual(budget_engine.spreadsheet_executable(), str(path))

    def test_finds_user_mac_application_outside_path(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'Applications/LibreOffice.app/Contents/MacOS/soffice'
            path.parent.mkdir(parents=True)
            path.write_text('#!/bin/sh\nexit 0\n')
            path.chmod(0o700)
            environment = {k: v for k, v in os.environ.items() if k != 'BUS_BUDGET_SOFFICE'}
            with patch.dict(os.environ, environment, clear=True), patch.object(budget_engine.shutil, 'which', return_value=None), patch.object(budget_engine.Path, 'home', return_value=Path(directory)):
                self.assertEqual(budget_engine.spreadsheet_executable(), str(path))

    def test_start_and_stop_real_local_service(self):
        with tempfile.TemporaryDirectory(prefix='bus-local-launch-') as directory:
            data = Path(directory) / 'data'
            expected = hashlib.sha256(str(data.resolve()).encode()).hexdigest()[:16]
            environment = os.environ.copy()
            environment['BUS_BUDGET_DATA_DIR'] = str(data)
            # Browser UI itself is Mac-only and is not claimed tested on Linux.
            environment['BROWSER'] = '/bin/true'
            process = subprocess.Popen([sys.executable, str(launch.ROOT / 'launch.py')], cwd=launch.ROOT,
                                       env=environment, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            port = None
            try:
                deadline = time.monotonic() + 40
                while time.monotonic() < deadline and process.poll() is None:
                    for candidate in range(8765, 8786):
                        health = launch.read_health(candidate)
                        if health and health.get('instance_id') == expected:
                            port = candidate
                            break
                    if port is not None:
                        break
                    time.sleep(0.3)
                self.assertIsNotNone(port, '本地服务未按时就绪')
                self.assertIsNone(process.poll())
                process.send_signal(signal.SIGINT)
                output, _ = process.communicate(timeout=15)
                self.assertEqual(process.returncode, 0, output)
                self.assertNotEqual((launch.read_health(port) or {}).get('instance_id'), expected)
            finally:
                if process.poll() is None:
                    process.send_signal(signal.SIGTERM)
                    try:
                        process.wait(timeout=15)
                    except subprocess.TimeoutExpired:
                        process.kill()
                        process.wait(timeout=5)


if __name__ == '__main__':
    unittest.main()
