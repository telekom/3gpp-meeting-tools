import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from PyQt5.QtCore import QPoint, QRect, QSettings, Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QSplitter, QTabWidget, QVBoxLayout, QWidget

from core.ui.nav_rail import NavRail
from core.ui.right_side_panel import (
    DEFAULT_PANEL_WIDTH,
    MAX_PANEL_WIDTH,
    MIN_PANEL_WIDTH,
    MODE_ASSISTANT,
    MODE_SPEC_SEARCH,
    RightSidePanel,
)
from core.ui.shell_state import restore_shell_state, save_shell_state
from core.ui.system_status_bar import SystemStatusBar
from core.ui.toast import ToastManager


class _MockRightPanelWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        central = QWidget(self)
        layout = QVBoxLayout(central)
        self.nav_rail = NavRail(self)
        self.workspace_splitter = QSplitter(Qt.Horizontal, self)
        self.workspace_splitter.setChildrenCollapsible(False)
        self.workspace_splitter.setHandleWidth(4)
        self.left_pane = QWidget(self)
        self.right_panel = RightSidePanel(self)
        self.workspace_splitter.addWidget(self.left_pane)
        self.workspace_splitter.addWidget(self.right_panel)
        self.workspace_splitter.setStretchFactor(0, 1)
        self.workspace_splitter.setStretchFactor(1, 0)
        self.workspace_splitter.setSizes([900, 0])
        layout.addWidget(self.workspace_splitter)
        self.setCentralWidget(central)

        self.tabs = QTabWidget(self)
        self.assistant_tab = QWidget(self)
        self.spec_search_tab = QWidget(self)
        self.tabs.addTab(self.spec_search_tab, "Spec Search")
        self.tabs.addTab(self.assistant_tab, "Assistant")
        self.status_bar = SystemStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.toast_manager = ToastManager(self)

        self.right_panel.toggled.connect(self._on_panel_toggled)
        self.right_panel.mode_changed.connect(
            lambda _m: self._on_panel_toggled(self.right_panel.is_open)
        )
        self.status_bar.assistant_dock_requested.connect(
            lambda: self.right_panel.toggle_mode(MODE_ASSISTANT)
        )
        self.status_bar.search_dock_requested.connect(
            lambda: self.right_panel.toggle_mode(MODE_SPEC_SEARCH)
        )

    def _on_panel_toggled(self, opened: bool):
        self.right_panel.apply_splitter_sizes(self.workspace_splitter)
        self.status_bar.update_right_panel_badges(opened, self.right_panel.mode)
        self.toast_manager.reposition()


class TestPhase3RightSidePanel(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_modes_toggle_and_width_clamping(self):
        panel = RightSidePanel()
        self.assertFalse(panel.is_open)
        self.assertEqual(panel.mode, MODE_ASSISTANT)
        self.assertEqual(panel.preferred_width, DEFAULT_PANEL_WIDTH)

        toggled_events = []
        mode_events = []
        promoted = []
        panel.toggled.connect(toggled_events.append)
        panel.mode_changed.connect(mode_events.append)
        panel.promote_requested.connect(promoted.append)

        # Toggle opens on requested mode
        panel.toggle_mode(MODE_SPEC_SEARCH)
        self.assertTrue(panel.is_open)
        self.assertEqual(panel.mode, MODE_SPEC_SEARCH)
        self.assertEqual(toggled_events, [True])
        self.assertEqual(mode_events, [MODE_SPEC_SEARCH])

        # Switching mode while open keeps panel open
        panel.toggle_mode(MODE_ASSISTANT)
        self.assertTrue(panel.is_open)
        self.assertEqual(panel.mode, MODE_ASSISTANT)
        self.assertEqual(toggled_events, [True])

        # Clicking active mode closes panel
        panel.toggle_mode(MODE_ASSISTANT)
        self.assertFalse(panel.is_open)
        self.assertEqual(toggled_events, [True, False])

        # Width clamping
        panel.preferred_width = 120
        self.assertEqual(panel.preferred_width, MIN_PANEL_WIDTH)
        panel.preferred_width = 1200
        self.assertEqual(panel.preferred_width, MAX_PANEL_WIDTH)

        # Promote emits current mode
        panel.open_mode(MODE_SPEC_SEARCH)
        panel.promote_btn.click()
        self.assertEqual(promoted, [MODE_SPEC_SEARCH])

    def test_splitter_sizes_and_capture(self):
        win = _MockRightPanelWindow()
        win.resize(1366, 768)
        win.show()
        self.app.processEvents()

        # Closed initially -> right width 0
        win.right_panel.apply_splitter_sizes(win.workspace_splitter)
        sizes_closed = win.workspace_splitter.sizes()
        self.assertEqual(sizes_closed[1], 0)

        # Open -> right width matches preferred_width (within handle tolerance)
        win.right_panel.preferred_width = 410
        win.right_panel.open_mode(MODE_ASSISTANT)
        self.app.processEvents()
        sizes_open = win.workspace_splitter.sizes()
        self.assertLessEqual(abs(sizes_open[1] - 410), 4)
        self.assertGreater(sizes_open[0], 500)

        # Manual splitter drag capture updates preferred width
        total = sum(sizes_open)
        win.workspace_splitter.setSizes([total - 460, 460])
        self.app.processEvents()
        win.right_panel.capture_width_from_splitter(win.workspace_splitter)
        self.assertLessEqual(abs(win.right_panel.preferred_width - 460), 4)
        win.close()

    def test_status_bar_dock_toggles(self):
        win = _MockRightPanelWindow()
        win.resize(1280, 768)
        win.show()
        self.app.processEvents()

        self.assertFalse(win.status_bar.assistant_dock_btn.isChecked())
        self.assertFalse(win.status_bar.search_dock_btn.isChecked())

        win.status_bar.assistant_dock_btn.click()
        self.app.processEvents()
        self.assertTrue(win.right_panel.is_open)
        self.assertEqual(win.right_panel.mode, MODE_ASSISTANT)
        self.assertTrue(win.status_bar.assistant_dock_btn.isChecked())
        self.assertFalse(win.status_bar.search_dock_btn.isChecked())

        win.status_bar.search_dock_btn.click()
        self.app.processEvents()
        self.assertTrue(win.right_panel.is_open)
        self.assertEqual(win.right_panel.mode, MODE_SPEC_SEARCH)
        self.assertFalse(win.status_bar.assistant_dock_btn.isChecked())
        self.assertTrue(win.status_bar.search_dock_btn.isChecked())

        win.status_bar.search_dock_btn.click()
        self.app.processEvents()
        self.assertFalse(win.right_panel.is_open)
        self.assertFalse(win.status_bar.search_dock_btn.isChecked())
        win.close()

    def test_toast_avoids_open_right_panel(self):
        win = _MockRightPanelWindow()
        win.resize(1400, 800)
        win.show()
        self.app.processEvents()

        toast = win.toast_manager.show_toast("Test notification", variant="info", duration_ms=0)
        self.app.processEvents()
        closed_right = toast.geometry().right()

        win.right_panel.preferred_width = 390
        win.right_panel.open_mode(MODE_ASSISTANT)
        self.app.processEvents()
        open_right = toast.geometry().right()

        panel_left = win.right_panel.mapTo(win, QPoint(0, 0)).x()
        self.assertLess(open_right, closed_right - 300)
        self.assertLessEqual(open_right, panel_left)
        toast.close()
        win.close()

    def test_nav_rail_dock_breakpoints(self):
        rail = NavRail()
        self.assertFalse(rail.is_collapsed)

        # Without dock open, 1200 keeps rail expanded
        rail.apply_responsive_width(1200, dock_open=False)
        self.assertFalse(rail.is_collapsed)

        # With dock open, < 1280 collapses rail automatically
        rail.apply_responsive_width(1240, dock_open=True)
        self.assertTrue(rail.is_collapsed)

        # Hysteresis zone (1280..1309) keeps current state
        rail.apply_responsive_width(1290, dock_open=True)
        self.assertTrue(rail.is_collapsed)

        # >= 1310 expands again
        rail.apply_responsive_width(1320, dock_open=True)
        self.assertFalse(rail.is_collapsed)

    def test_shell_state_round_trip(self):
        settings = QSettings("3GPPToolsTest", "Phase3RightPanelRoundTrip")
        settings.clear()

        win1 = _MockRightPanelWindow()
        win1.setGeometry(100, 100, 1440, 860)
        win1.show()
        self.app.processEvents()
        win1.right_panel.preferred_width = 435
        win1.right_panel.open_mode(MODE_SPEC_SEARCH)
        self.app.processEvents()

        save_shell_state(win1, settings)
        win1.close()

        win2 = _MockRightPanelWindow()
        with patch(
            "core.ui.shell_state.clamp_window_to_available_screens",
            side_effect=lambda rect: rect,
        ):
            restore_shell_state(win2, settings)
        win2.right_panel.apply_splitter_sizes(win2.workspace_splitter)

        self.assertTrue(win2.right_panel.is_open)
        self.assertEqual(win2.right_panel.mode, MODE_SPEC_SEARCH)
        self.assertEqual(win2.right_panel.preferred_width, 435)
        self.assertTrue(win2.status_bar.search_dock_btn.isChecked())
        settings.clear()
        win2.close()

    def test_main_window_shortcuts_and_promote(self):
        from main_window import DragDropUI

        host = _MockRightPanelWindow()
        host.toggle_right_panel = lambda mode: DragDropUI.toggle_right_panel(host, mode)
        host._on_right_panel_promote = lambda mode: DragDropUI._on_right_panel_promote(host, mode)
        host.right_panel.promote_requested.connect(host._on_right_panel_promote)

        # Toggle via helper bound from DragDropUI
        host.toggle_right_panel(MODE_ASSISTANT)
        self.assertTrue(host.right_panel.is_open)
        self.assertEqual(host.right_panel.mode, MODE_ASSISTANT)

        # Promote switches tab to Assistant and closes dock
        host.right_panel.promote_btn.click()
        self.assertFalse(host.right_panel.is_open)
        self.assertIs(host.tabs.currentWidget(), host.assistant_tab)

        # Promote Spec Search switches to spec_search_tab
        host.toggle_right_panel(MODE_SPEC_SEARCH)
        host.right_panel.promote_btn.click()
        self.assertFalse(host.right_panel.is_open)
        self.assertIs(host.tabs.currentWidget(), host.spec_search_tab)

        host.close()


if __name__ == "__main__":
    unittest.main()
