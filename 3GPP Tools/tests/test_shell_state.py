"""Unit tests for Phase 1.6 QSettings layout and state persistence."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from PyQt5.QtCore import QSettings, Qt
from PyQt5.QtGui import QGuiApplication
from PyQt5.QtWidgets import QApplication, QMainWindow, QSplitter, QTabWidget, QWidget

from core.ui.bottom_drawer import BottomDrawer
from core.ui.nav_rail import NavRail
from core.ui.shell_state import (
    clamp_window_to_available_screens,
    restore_shell_state,
    save_shell_state,
)
from core.ui.system_status_bar import SystemStatusBar
from core.ui.ui_panels import ConsolePanel, QueuePanel


class _StubShellWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(1120, 760)
        self.nav_rail = NavRail(self)
        self.tabs = QTabWidget(self)
        for i in range(9):
            self.tabs.addTab(QWidget(), f"Tab {i}")
        self.splitter = QSplitter(Qt.Vertical, self)
        self.splitter.addWidget(self.tabs)
        self.bottom_drawer = BottomDrawer(ConsolePanel(), QueuePanel(), self)
        self.splitter.addWidget(self.bottom_drawer)
        self.status_bar = SystemStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.setCentralWidget(self.splitter)

    def cleanup(self):
        self.status_bar.stop_monitors()
        self.status_bar.mcp_timer.stop()


class TestShellStatePersistence(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_save_and_restore_roundtrip_via_qsettings(self):
        with tempfile.TemporaryDirectory() as tmp:
            ini_path = Path(tmp) / "shell_state.ini"
            cfg_save = QSettings(str(ini_path), QSettings.IniFormat)

            win1 = _StubShellWindow()
            try:
                win1.resize(1280, 820)
                win1.nav_rail.set_collapsed(True, manual=True)
                win1.bottom_drawer.preferred_height = 240
                win1.bottom_drawer.set_inner_sizes([720, 260])
                win1.bottom_drawer.set_open(True, emit=False)
                win1.bottom_drawer.apply_splitter_sizes(win1.splitter)
                win1.tabs.setCurrentIndex(6)
                expected_h = win1.bottom_drawer.height() if win1.bottom_drawer.height() >= 120 else 240
                save_shell_state(win1, settings=cfg_save)
            finally:
                win1.cleanup()

            cfg_load = QSettings(str(ini_path), QSettings.IniFormat)
            win2 = _StubShellWindow()
            try:
                restore_shell_state(win2, settings=cfg_load)
                self.assertEqual(win2.tabs.currentIndex(), 6)
                self.assertEqual(win2.nav_rail.active_index(), 6)
                self.assertTrue(win2.nav_rail.is_collapsed)
                self.assertTrue(win2.nav_rail.manual_collapsed)
                self.assertTrue(win2.bottom_drawer.is_open)
                self.assertEqual(win2.bottom_drawer.preferred_height, expected_h)
                self.assertTrue(win2.status_bar.drawer_toggle_btn.isChecked())
            finally:
                win2.cleanup()

    def test_disconnected_monitor_offscreen_window_clamped_to_primary_screen(self):
        win = QWidget()
        # Simulate saved coordinates on an unplugged external monitor far off-screen
        win.setGeometry(50000, 50000, 1100, 760)
        clamp_window_to_available_screens(win)

        primary = QGuiApplication.primaryScreen()
        if primary is not None:
            avail = primary.availableGeometry()
            self.assertTrue(
                avail.intersects(win.geometry()),
                msg=f"Clamped geometry {win.geometry()} should intersect primary {avail}",
            )


if __name__ == "__main__":
    unittest.main()
