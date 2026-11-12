"""Unit tests for Phase 1.7 shell keyboard shortcuts and focus order."""

from __future__ import annotations

import sys
import unittest

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QShortcut, QSplitter, QTabWidget, QWidget

from core.ui.bottom_drawer import BottomDrawer
from core.ui.nav_rail import DEFAULT_NAV_ITEMS, NavRail
from core.ui.system_status_bar import SystemStatusBar
from core.ui.ui_panels import ConsolePanel, QueuePanel
from main_window import DragDropUI


class _ShortcutStubWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.nav_rail = NavRail(self)
        self.help_btn = QPushButton("Help && Docs", self)
        self.settings_btn = QPushButton("Settings", self)
        self.nav_rail.register_footer_button(self.help_btn, "Help && Docs")
        self.nav_rail.register_footer_button(self.settings_btn, "Settings")

        self.tabs = QTabWidget(self)
        for spec in DEFAULT_NAV_ITEMS:
            self.tabs.addTab(QWidget(), spec.label)

        self.splitter = QSplitter(Qt.Vertical, self)
        self.splitter.addWidget(self.tabs)
        self.bottom_drawer = BottomDrawer(ConsolePanel(), QueuePanel(), self)
        self.splitter.addWidget(self.bottom_drawer)
        self.status_bar = SystemStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.setCentralWidget(self.splitter)

        self.toggle_bottom_drawer = lambda: DragDropUI.toggle_bottom_drawer(self)
        self._activate_tab_by_shortcut = lambda idx: DragDropUI._activate_tab_by_shortcut(self, idx)
        DragDropUI._setup_shortcuts_and_focus(self)

    def cleanup(self):
        self.status_bar.stop_monitors()
        self.status_bar.mcp_timer.stop()


class TestShellShortcutsAndFocus(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_ctrl_1_through_9_and_shell_shortcuts_registered(self):
        win = _ShortcutStubWindow()
        try:
            self.assertEqual(len(win.nav_shortcuts), 9)
            for idx, sc in enumerate(win.nav_shortcuts):
                self.assertEqual(sc.key(), QKeySequence(f"Ctrl+{idx + 1}"))
                sc.activated.emit()
                self.assertEqual(win.tabs.currentIndex(), idx)
                self.assertEqual(win.nav_rail.active_index(), idx)

            # Respect disabled tabs state
            win.tabs.setEnabled(False)
            win.nav_shortcuts[0].activated.emit()
            self.assertEqual(win.tabs.currentIndex(), 8)
            win.tabs.setEnabled(True)

            self.assertEqual(win.drawer_shortcut.key(), QKeySequence("Ctrl+J"))
            self.assertFalse(win.bottom_drawer.is_open)
            win.drawer_shortcut.activated.emit()
            self.assertTrue(win.bottom_drawer.is_open)

            self.assertEqual(win.rail_shortcut.key(), QKeySequence("Ctrl+\\"))
            self.assertFalse(win.nav_rail.is_collapsed)
            win.rail_shortcut.activated.emit()
            self.assertTrue(win.nav_rail.is_collapsed)
        finally:
            win.cleanup()

    def test_focus_policies_and_tab_order_chain(self):
        win = _ShortcutStubWindow()
        try:
            self.assertEqual(win.nav_rail.brand_badge.focusPolicy(), Qt.NoFocus)
            self.assertEqual(win.nav_rail.brand_title.focusPolicy(), Qt.NoFocus)

            ordered = win.nav_rail.ordered_focus_widgets()
            # Toggle btn + 9 nav buttons + 2 footer buttons
            self.assertEqual(len(ordered), 12)
            self.assertIs(ordered[0], win.nav_rail.toggle_btn)
            self.assertIs(ordered[-2], win.help_btn)
            self.assertIs(ordered[-1], win.settings_btn)

            # Next in focus chain from last rail button reaches tabs / drawer / status bar
            self.assertIsNotNone(win.settings_btn.nextInFocusChain())
        finally:
            win.cleanup()


if __name__ == "__main__":
    unittest.main()
