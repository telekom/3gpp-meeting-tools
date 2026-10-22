import unittest
import sys
from pathlib import Path
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer

from core.utils.paths import get_project_root
from modules.specifications.core.database import SpecsDatabase
from modules.specifications.ui.ui_tabs import SpecificationsTab

app = QApplication.instance()
if app is None:
    app = QApplication(sys.argv)


class TestSpecFilteringPerformance(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.db_path = get_project_root() / "db" / "3gpp_data.db"
        cls.db = SpecsDatabase(cls.db_path)

    def setUp(self):
        self.tab = SpecificationsTab(self.db_path)
        self.tab._is_loaded = True

    def test_search_files_conditional_joins(self):
        """Verify search_files returns valid specifications both with and without faceted filters."""
        # Unfiltered search
        unfiltered = self.db.search_files()
        self.assertGreater(len(unfiltered), 0)

        # Scoped search
        scoped = self.db.search_files(series="23", spec_type="TS")
        self.assertGreater(len(scoped), 0)
        self.assertLessEqual(len(scoped), len(unfiltered))

    def test_filter_debounce_timer_initialization(self):
        """Verify both search and filter debounce timers are initialized with proper singleShot intervals."""
        self.assertTrue(self.tab.search_timer.isSingleShot())
        self.assertEqual(self.tab.search_timer.interval(), 300)

        self.assertTrue(self.tab.filter_timer.isSingleShot())
        self.assertEqual(self.tab.filter_timer.interval(), 150)

    def test_filter_interactions_trigger_debounce_timer(self):
        """Verify filter changes start the debounce timer rather than blocking synchronously."""
        # TSG button click
        self.tab._on_tsg_button_clicked("SA")
        self.assertTrue(self.tab.filter_timer.isActive())
        self.tab.filter_timer.stop()

        # Series change
        self.tab._on_series_filter_changed(["23"])
        self.assertTrue(self.tab.filter_timer.isActive())
        self.tab.filter_timer.stop()

        # Tech change
        self.tab._on_tech_combo_changed()
        self.assertTrue(self.tab.filter_timer.isActive())
        self.tab.filter_timer.stop()

        # Type change
        self.tab._on_type_button_clicked("TS")
        self.assertTrue(self.tab.filter_timer.isActive())
        self.tab.filter_timer.stop()

        # Search text change
        self.tab._on_search_changed()
        self.assertTrue(self.tab.search_timer.isActive())
        self.tab.search_timer.stop()

    def test_lazy_quick_copy_menu_in_column_2(self):
        """Verify Quick Copy submenu in Column 2 is deferred until aboutToShow is triggered."""
        self.tab.spec_search_input.setText("23.501")
        self.tab.refresh_table()

        self.assertGreater(self.tab.table.rowCount(), 0)
        cell_widget = self.tab.table.cellWidget(0, 2)
        self.assertIsNotNone(cell_widget)

        # Find doc_action_btn
        buttons = [child for child in cell_widget.children() if hasattr(child, "menu") and child.menu()]
        self.assertTrue(len(buttons) > 0)
        doc_btn = buttons[0]
        doc_menu = doc_btn.menu()

        copy_sub = None
        for act in doc_menu.actions():
            if act.menu() and "Quick Copy" in act.text():
                copy_sub = act.menu()
                break

        self.assertIsNotNone(copy_sub, "Quick Copy submenu should exist in Column 2 menu")
        # Submenu should be empty before user hovers/opens it
        self.assertEqual(len(copy_sub.actions()), 0)

        # Trigger aboutToShow
        copy_sub.aboutToShow.emit()
        self.assertGreater(len(copy_sub.actions()), 0)


if __name__ == "__main__":
    unittest.main()
