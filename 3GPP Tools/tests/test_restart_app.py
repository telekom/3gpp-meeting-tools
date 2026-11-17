"""Unit tests for Restart Application functionality across shell surfaces."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
for _p in (str(ROOT_DIR), str(SRC_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from PyQt5.QtWidgets import QApplication, QMessageBox

from core.ui.command_palette import CommandPaletteDialog, build_static_command_items
from core.ui.system_status_bar import SystemStatusBar
from core.updater.app_updater import spawn_detached_relaunch


class TestRestartApplication(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_spawn_detached_relaunch_writes_helper_and_spawns_hidden(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            temp_root = Path(tmp) / "temp"
            app_root = Path(tmp) / "3GPP Tools"
            (app_root / "src").mkdir(parents=True)

            with (
                patch("core.updater.app_updater.get_temp_root", return_value=temp_root),
                patch("core.updater.app_updater.get_project_root", return_value=app_root / "src"),
                patch("core.updater.app_updater.subprocess.Popen") as popen_mock,
            ):
                helper_path = spawn_detached_relaunch(app_root=app_root, parent_pid=12345)

            self.assertTrue(helper_path.is_file())
            self.assertEqual(helper_path.name, "relaunch_app.py")
            script_text = helper_path.read_text(encoding="utf-8")
            self.assertIn("pid = 12345", script_text)
            self.assertIn("'-m', 'main_tools'", script_text)
            self.assertIn("PYTHONPATH", script_text)

            popen_mock.assert_called_once()
            args, kwargs = popen_mock.call_args
            self.assertEqual(args[0], [sys.executable, str(helper_path)])
            if os.name == "nt":
                self.assertEqual(kwargs.get("creationflags"), 0x08000000)

    def test_command_palette_includes_restart_action(self) -> None:
        triggered = []
        dummy_win = SimpleNamespace(restart_application=lambda: triggered.append(True))
        items = build_static_command_items(dummy_win)
        restart_items = [it for it in items if it.id == "sys:restart_app"]
        self.assertEqual(len(restart_items), 1)
        cmd = restart_items[0]
        self.assertEqual(cmd.title, "Restart Application")
        self.assertEqual(cmd.shortcut, "Ctrl+Shift+R")
        self.assertEqual(cmd.icon_name, "rotate-ccw")

        dlg = CommandPaletteDialog(items)
        try:
            dlg.populate_results("restart")
            matched_ids = []
            for row in range(dlg.results_list.count()):
                it = dlg.results_list.item(row)
                c = it.data(256 + 21)  # Qt.UserRole + 21
                if c is not None:
                    matched_ids.append(c.id)
            self.assertIn("sys:restart_app", matched_ids)
        finally:
            dlg.close()

        cmd.callback()
        self.assertEqual(triggered, [True])

    def test_system_status_bar_has_restart_menu_action(self) -> None:
        bar = SystemStatusBar()
        try:
            fired = []
            bar.restart_requested.connect(lambda: fired.append(True))
            actions = [act for act in bar.system_menu.actions() if "Restart Application" in act.text()]
            self.assertEqual(len(actions), 1)
            actions[0].trigger()
            self.assertEqual(fired, [True])
        finally:
            bar.stop_monitors()
            bar.mcp_timer.stop()

    def test_restart_application_idle_vs_busy(self) -> None:
        from main_window import DragDropUI

        class _FakeWindow:
            def __init__(self, is_processing: bool, queued: list) -> None:
                self._restarting = False
                self.queue_manager = SimpleNamespace(
                    is_processing=is_processing,
                    file_queue=list(queued),
                    clear_queue=MagicMock(),
                    abort_current_task=MagicMock(),
                )
                self.status_bar = MagicMock()
                self.close = MagicMock()

            restart_application = DragDropUI.restart_application

        # 1. Idle -> restarts immediately without confirmation
        idle_win = _FakeWindow(is_processing=False, queued=[])
        with (
            patch("main_window.QMessageBox.question") as q_mock,
            patch("core.updater.app_updater.spawn_detached_relaunch") as relaunch_mock,
            patch("main_window.QTimer.singleShot") as timer_mock,
        ):
            idle_win.restart_application()
            q_mock.assert_not_called()
            relaunch_mock.assert_called_once()
            self.assertTrue(idle_win._restarting)
            self.assertEqual(timer_mock.call_count, 2)

        # 2. Busy + user cancels -> does not restart
        busy_cancel = _FakeWindow(is_processing=True, queued=[])
        with (
            patch("main_window.QMessageBox.question", return_value=QMessageBox.No) as q_mock,
            patch("core.updater.app_updater.spawn_detached_relaunch") as relaunch_mock,
            patch("main_window.QTimer.singleShot") as timer_mock,
        ):
            busy_cancel.restart_application()
            q_mock.assert_called_once()
            relaunch_mock.assert_not_called()
            busy_cancel.queue_manager.clear_queue.assert_not_called()
            busy_cancel.queue_manager.abort_current_task.assert_not_called()
            self.assertFalse(busy_cancel._restarting)
            timer_mock.assert_not_called()

        # 3. Busy + user confirms -> clears/aborts queue and restarts
        busy_ok = _FakeWindow(is_processing=False, queued=["item1"])
        with (
            patch("main_window.QMessageBox.question", return_value=QMessageBox.Yes) as q_mock,
            patch("core.updater.app_updater.spawn_detached_relaunch") as relaunch_mock,
            patch("main_window.QTimer.singleShot") as timer_mock,
        ):
            busy_ok.restart_application()
            q_mock.assert_called_once()
            busy_ok.queue_manager.clear_queue.assert_called_once()
            busy_ok.queue_manager.abort_current_task.assert_called_once()
            relaunch_mock.assert_called_once()
            self.assertTrue(busy_ok._restarting)
            self.assertEqual(timer_mock.call_count, 2)


if __name__ == "__main__":
    unittest.main()
