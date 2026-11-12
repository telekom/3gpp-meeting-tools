"""Unit tests for Phase 1.3 BottomDrawer and SystemStatusBar drawer badges."""

from __future__ import annotations

import logging
import sys
import unittest

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QSplitter, QWidget

from core.ui.bottom_drawer import BottomDrawer
from core.ui.system_status_bar import SystemStatusBar
from core.ui.ui_panels import ConsolePanel, QueuePanel


class TestBottomDrawerAndStatusBadges(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_drawer_starts_collapsed_and_toggles_with_splitter(self):
        splitter = QSplitter(Qt.Vertical)
        top = QWidget()
        splitter.addWidget(top)
        splitter.resize(900, 700)

        console = ConsolePanel()
        queue = QueuePanel()
        drawer = BottomDrawer(console, queue)
        splitter.addWidget(drawer)

        self.assertFalse(drawer.is_open)
        self.assertFalse(drawer.isVisibleTo(splitter))

        toggled_events = []
        drawer.toggled.connect(toggled_events.append)

        drawer.toggle()
        drawer.apply_splitter_sizes(splitter)
        self.assertTrue(drawer.is_open)
        self.assertEqual(toggled_events, [True])
        self.assertGreater(splitter.sizes()[1], 0)

        drawer.toggle()
        drawer.apply_splitter_sizes(splitter)
        self.assertFalse(drawer.is_open)
        self.assertEqual(toggled_events, [True, False])
        self.assertEqual(splitter.sizes()[1], 0)

    def test_unread_errors_accumulate_when_closed_and_clear_on_open(self):
        drawer = BottomDrawer(ConsolePanel(), QueuePanel())
        drawer.set_queue_count(2)
        self.assertEqual(drawer.queue_count, 2)
        self.assertEqual(drawer.queue_badge.text(), "2 items")

        # INFO does not increment unread error count
        drawer.record_log_event("Everything nominal", logging.INFO)
        self.assertEqual(drawer.unread_errors, 0)
        self.assertFalse(drawer.error_badge.isVisibleTo(drawer))

        # ERROR increments unread error badge while collapsed
        drawer.record_log_event("❌ Conversion failed", logging.ERROR)
        self.assertEqual(drawer.unread_errors, 1)
        self.assertEqual(drawer.error_badge.text(), "1 error")

        # Opening drawer clears unread errors automatically
        drawer.set_open(True)
        self.assertEqual(drawer.unread_errors, 0)
        self.assertFalse(drawer.error_badge.isVisibleTo(drawer))

    def test_status_bar_drawer_toggle_badge_and_signal(self):
        bar = SystemStatusBar()
        try:
            clicks = []
            bar.drawer_toggle_requested.connect(lambda: clicks.append(True))

            self.assertEqual(bar.drawer_toggle_btn.text(), "Console && Queue")
            bar.update_drawer_badge(False, 3, 2)
            self.assertEqual(bar.drawer_toggle_btn.text(), "Console && Queue · 3 queued · 2 errs")
            self.assertFalse(bar.drawer_toggle_btn.isChecked())

            bar.update_drawer_badge(True, 1, 0)
            self.assertEqual(bar.drawer_toggle_btn.text(), "Console && Queue · 1 queued")
            self.assertTrue(bar.drawer_toggle_btn.isChecked())

            bar.drawer_toggle_btn.click()
            self.assertEqual(clicks, [True])
        finally:
            bar.stop_monitors()
            bar.mcp_timer.stop()


if __name__ == "__main__":
    unittest.main()
