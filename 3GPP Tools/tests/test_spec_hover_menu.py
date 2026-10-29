import sys
import unittest
from PyQt5.QtCore import QPoint, QRect, Qt
from PyQt5.QtGui import QCursor, QMouseEvent
from PyQt5.QtWidgets import QApplication, QMenu, QWidget

from modules.specifications.ui.components import HoverMenuButton


class TestHoverMenuButton(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        self.parent = QWidget()
        self.parent.setGeometry(100, 100, 400, 300)
        self.parent.show()
        self.btn = HoverMenuButton("⋮", self.parent)
        self.btn.setGeometry(10, 10, 22, 22)
        self.btn.show()

        self.menu = QMenu(self.btn)
        self.copy_sub = self.menu.addMenu("📋 Quick Copy")
        self.opt1 = self.copy_sub.addAction("Option 1")
        self.nested_sub = self.copy_sub.addMenu("🌐 Web Links")
        self.nested_opt = self.nested_sub.addAction("Nested Option")
        self.btn.setMenu(self.menu)
        QApplication.processEvents()

    def tearDown(self):
        if self.menu.isVisible():
            self.menu.hide()
        self.parent.close()
        QApplication.processEvents()

    def test_initial_state_and_timer(self):
        """Test initial timer configuration and attributes."""
        self.assertEqual(self.btn._hover_timer.interval(), 80)
        self.assertFalse(self.btn._pinned)
        self.assertEqual(self.btn._outside_count, 0)
        self.assertFalse(self.btn._hover_timer.isActive())

    def test_enter_event_opens_menu_and_starts_timer(self):
        """Test that hovering over the button opens the menu and starts the hover timer."""
        self.btn.enterEvent(None)
        QApplication.processEvents()
        self.assertTrue(self.menu.isVisible())
        self.assertTrue(self.btn._hover_timer.isActive())
        self.assertFalse(self.btn._pinned)

    def test_cursor_inside_button_and_main_menu(self):
        """Test that cursor over button or main menu is recognized as inside."""
        self.btn.enterEvent(None)
        QApplication.processEvents()

        # Over the button
        btn_center = self.btn.mapToGlobal(self.btn.rect().center())
        self.assertTrue(self.btn.is_pos_inside(btn_center))

        # Inside the main menu
        menu_center = self.menu.geometry().center()
        self.assertTrue(self.btn.is_pos_inside(menu_center))

    def test_cursor_inside_submenu_options(self):
        """
        Regression Test for User Request:
        When navigating to 'Quick copy' options or nested submenus, the mouse position
        inside the submenu must be detected as inside the menu hierarchy so the menu
        does NOT disappear.
        """
        self.btn.enterEvent(None)
        QApplication.processEvents()

        # Simulate Quick Copy submenu being popped up
        sub_pos = QPoint(self.menu.geometry().right() + 2, self.menu.geometry().top() + 10)
        self.copy_sub.popup(sub_pos)
        QApplication.processEvents()
        self.assertTrue(self.copy_sub.isVisible())

        # Mouse is inside the Quick Copy submenu options
        sub_center = self.copy_sub.geometry().center()
        self.assertTrue(
            self.btn.is_pos_inside(sub_center),
            "Cursor inside Quick Copy submenu must be recognized as inside!",
        )

        # Check transit corridor between main menu and submenu
        mid_x = (self.menu.geometry().right() + self.copy_sub.geometry().left()) // 2
        mid_y = self.copy_sub.geometry().center().y()
        corridor_point = QPoint(mid_x, mid_y)
        self.assertTrue(
            self.btn.is_pos_inside(corridor_point),
            "Cursor in transition corridor between parent and submenu must be recognized as inside!",
        )

        # Simulate nested sub-submenu popup (e.g. Web Links)
        nested_pos = QPoint(self.copy_sub.geometry().right() + 2, self.copy_sub.geometry().top() + 10)
        self.nested_sub.popup(nested_pos)
        QApplication.processEvents()
        self.assertTrue(self.nested_sub.isVisible())

        nested_center = self.nested_sub.geometry().center()
        self.assertTrue(
            self.btn.is_pos_inside(nested_center),
            "Cursor inside nested sub-submenu must be recognized as inside!",
        )

    def test_cursor_far_away_triggers_grace_period(self):
        """Test that moving the cursor far away requires 3 ticks before closing."""
        self.btn.enterEvent(None)
        QApplication.processEvents()
        self.assertTrue(self.menu.isVisible())

        far_point = QPoint(self.menu.geometry().right() + 500, self.menu.geometry().bottom() + 500)
        self.assertFalse(self.btn.is_pos_inside(far_point))

        # Manually trigger check with outside position by mocking QCursor.pos
        orig_pos = QCursor.pos
        try:
            QCursor.pos = lambda: far_point

            # Tick 1: count becomes 1, menu still visible
            self.btn._check_mouse_position()
            self.assertEqual(self.btn._outside_count, 1)
            self.assertTrue(self.menu.isVisible())

            # Tick 2: count becomes 2, menu still visible
            self.btn._check_mouse_position()
            self.assertEqual(self.btn._outside_count, 2)
            self.assertTrue(self.menu.isVisible())

            # Mouse briefly comes back inside -> count resets to 0
            QCursor.pos = lambda: self.menu.geometry().center()
            self.btn._check_mouse_position()
            self.assertEqual(self.btn._outside_count, 0)
            self.assertTrue(self.menu.isVisible())

            # Mouse goes outside again: 3 ticks close the menu
            QCursor.pos = lambda: far_point
            self.btn._check_mouse_position()
            self.btn._check_mouse_position()
            self.btn._check_mouse_position()
            self.assertFalse(self.menu.isVisible())
            self.assertFalse(self.btn._hover_timer.isActive())
        finally:
            QCursor.pos = orig_pos

    def test_left_click_pins_menu(self):
        """Test that clicking the button pins the menu open and stops hover timer."""
        press_ev = QMouseEvent(
            QMouseEvent.MouseButtonPress,
            QPoint(5, 5),
            Qt.LeftButton,
            Qt.LeftButton,
            Qt.NoModifier,
        )
        self.btn.mousePressEvent(press_ev)
        QApplication.processEvents()

        self.assertTrue(self.menu.isVisible())
        self.assertTrue(self.btn._pinned)
        self.assertFalse(self.btn._hover_timer.isActive())

        # While pinned, _check_mouse_position does nothing
        far_point = QPoint(self.menu.geometry().right() + 500, self.menu.geometry().bottom() + 500)
        orig_pos = QCursor.pos
        try:
            QCursor.pos = lambda: far_point
            self.btn._check_mouse_position()
            self.assertTrue(self.menu.isVisible())
        finally:
            QCursor.pos = orig_pos

        # Second click while pinned closes menu
        self.btn.mousePressEvent(press_ev)
        QApplication.processEvents()
        self.assertFalse(self.menu.isVisible())
        self.assertFalse(self.btn._pinned)


if __name__ == "__main__":
    unittest.main()
