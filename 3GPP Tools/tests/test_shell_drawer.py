"""Unit tests for Phase 1.3 BottomDrawer and SystemStatusBar drawer badges."""

from __future__ import annotations

import logging
import sys
import unittest

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QSplitter, QWidget

from core.ui.bottom_drawer import BottomDrawer
from core.ui.system_status_bar import SystemStatusBar
from core.ui.theme import B
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

    def test_system_menu_in_status_bar_and_no_fixed_panel_buttons(self):
        console = ConsolePanel()
        queue = QueuePanel()

        # Buttons must not lock pixel size via setFixedSize for high-DPI adaptability
        for btn in (console.copy_log_btn, console.clear_btn, queue.abort_btn, queue.remove_btn, queue.clear_btn):
            self.assertGreater(btn.maximumWidth(), 10000)
            self.assertGreater(btn.maximumHeight(), 10000)

        # Compatibility no-op
        console.set_proxy_status(True, "Proxy active")

        bar = SystemStatusBar()
        try:
            actions = [act.text() for act in bar.system_menu.actions()]
            self.assertEqual(len(actions), 6)
            fired = []
            bar.task_manager_requested.connect(lambda: fired.append("task"))
            bar.db_maintenance_requested.connect(lambda: fired.append("db"))
            bar.network_config_requested.connect(lambda: fired.append("net"))
            bar.java_info_requested.connect(lambda: fired.append("java"))

            bar.system_menu.actions()[0].trigger()
            bar.system_menu.actions()[1].trigger()
            bar.system_menu.actions()[3].trigger()
            bar.system_menu.actions()[4].trigger()
            self.assertEqual(fired, ["task", "db", "net", "java"])
        finally:
            bar.stop_monitors()
            bar.mcp_timer.stop()

    def test_status_bar_minimalist_mode_and_hysteresis(self):
        bar = SystemStatusBar()
        try:
            # Seed states before collapsing
            bar._update_ollama_indicator(True, "qwen3:8b", ["qwen3:8b"], "")
            bar._update_network_indicator("Ramada", False, False)
            bar.update_drawer_badge(False, 3, 2)

            self.assertFalse(bar.is_minimal_mode)
            self.assertEqual(bar.system_btn.text(), "System")
            self.assertEqual(bar.assistant_dock_btn.text(), "Assistant")
            self.assertEqual(bar.search_dock_btn.text(), "Quick Search")
            self.assertEqual(bar.ollama_indicator.text(), "🦙 qwen3:8b")
            self.assertEqual(bar.proxy_indicator.text(), "🌐 Direct")
            self.assertEqual(bar.network_indicator.text(), "🌐 Ramada")
            self.assertFalse(bar.network_indicator.isHidden())

            # Shrink below STATUS_MINIMAL_COLLAPSE_PX -> enters minimalist mode
            bar.apply_responsive_width(B.STATUS_MINIMAL_COLLAPSE_PX - 20)
            self.assertTrue(bar.is_minimal_mode)
            self.assertEqual(bar.system_btn.text(), "")
            self.assertEqual(bar.assistant_dock_btn.text(), "")
            self.assertEqual(bar.search_dock_btn.text(), "")
            self.assertEqual(bar.ollama_indicator.text(), "🦙")
            self.assertEqual(bar.proxy_indicator.text(), "🌐")
            self.assertEqual(bar.drawer_toggle_btn.text(), "2!")
            self.assertTrue(bar.network_indicator.isHidden())
            self.assertIn("Local", bar.clock_indicator.text())

            # Idle drawer shows no text in minimal mode, active queue shows count
            bar.update_drawer_badge(False, 4, 0)
            self.assertEqual(bar.drawer_toggle_btn.text(), "4")
            bar.update_drawer_badge(False, 0, 0)
            self.assertEqual(bar.drawer_toggle_btn.text(), "")

            # 3GPP network IS shown compactly even in minimalist mode
            bar._update_network_indicator("3GPP-Meeting", True, True)
            self.assertFalse(bar.network_indicator.isHidden())
            self.assertEqual(bar.network_indicator.text(), "🟢 3GPP")

            # Hysteresis zone keeps minimal mode active
            bar.apply_responsive_width(B.STATUS_MINIMAL_COLLAPSE_PX + 10)
            self.assertTrue(bar.is_minimal_mode)

            # Expand >= STATUS_MINIMAL_EXPAND_PX -> restores full labels
            bar.apply_responsive_width(B.STATUS_MINIMAL_EXPAND_PX)
            self.assertFalse(bar.is_minimal_mode)
            self.assertEqual(bar.system_btn.text(), "System")
            self.assertEqual(bar.assistant_dock_btn.text(), "Assistant")
            self.assertEqual(bar.search_dock_btn.text(), "Quick Search")
            self.assertEqual(bar.drawer_toggle_btn.text(), "Console && Queue")
            self.assertEqual(bar.ollama_indicator.text(), "🦙 qwen3:8b")
            self.assertEqual(bar.network_indicator.text(), "🟢 3GPP-Meeting (Local Server Active)")
        finally:
            bar.stop_monitors()
            bar.mcp_timer.stop()


if __name__ == "__main__":
    unittest.main()
