"""Unit tests for Phase 1.5 non-blocking Toast notification component."""

from __future__ import annotations

import sys
import unittest
from unittest.mock import MagicMock

from PyQt5.QtWidgets import QApplication, QMainWindow

from core.ui.toast import ToastAction, ToastManager


class TestToastManagerAndWidget(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_show_toast_stacks_max_visible_and_positions_bottom_right(self):
        host = QMainWindow()
        host.resize(1000, 700)
        mgr = ToastManager(host, max_visible=3)

        t1 = mgr.show_toast("First", variant="info", duration_ms=0)
        mgr.show_toast("Second", variant="success", duration_ms=0)
        mgr.show_toast("Third", variant="warning", duration_ms=0)
        t4 = mgr.show_toast("Fourth", variant="danger", duration_ms=0)

        # Max 3 active: oldest evicted automatically
        self.assertTrue(t1._closed)
        self.assertEqual(len(mgr.active_toasts), 3)
        self.assertEqual(mgr.active_toasts[-1], t4)

        # Bottom-right placement inside host bounds
        for toast in mgr.active_toasts:
            self.assertGreater(toast.x(), 500)
            self.assertLessEqual(toast.x() + toast.width(), host.width())
            self.assertLessEqual(toast.y() + toast.height(), host.height())

        mgr.dismiss_all()
        self.assertEqual(mgr.active_toasts, [])

    def test_toast_action_invokes_callback_and_dismisses(self):
        host = QMainWindow()
        host.resize(900, 600)
        mgr = ToastManager(host)

        clicked = []
        toast = mgr.show_toast(
            "Saved diagram.vsdx",
            message="C:/exports",
            variant="success",
            actions=[
                ToastAction("Open", lambda: clicked.append("open")),
                ToastAction("Save & Exit", lambda: clicked.append("save")),
            ],
            duration_ms=0,
        )

        # Button mnemonics escaped safely
        self.assertEqual(toast._action_buttons[0].text(), "Open")
        self.assertEqual(toast._action_buttons[1].text(), "Save && Exit")

        toast._action_buttons[0].click()
        self.assertEqual(clicked, ["open"])
        self.assertTrue(toast._closed)
        self.assertEqual(mgr.active_toasts, [])

    def test_main_window_conversion_success_emits_toast_with_actions(self):
        from main_window import DragDropUI

        win = MagicMock(spec=DragDropUI)
        win.code_tab = MagicMock()
        win.word_tab = MagicMock()
        win.batch_tab = MagicMock()

        DragDropUI.on_conversion_success(win, "C:/output/diagram_test.vsdx")

        win.show_toast.assert_called_once()
        args, kwargs = win.show_toast.call_args
        self.assertIn("diagram_test.vsdx", args[0])
        self.assertEqual(kwargs["variant"], "success")
        labels = [a.label for a in kwargs["actions"]]
        self.assertEqual(labels, ["Open", "Show in Folder"])


if __name__ == "__main__":
    unittest.main()
