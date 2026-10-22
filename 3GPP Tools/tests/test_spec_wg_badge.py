import sys
import unittest
from pathlib import Path
from PyQt5.QtWidgets import QApplication

from modules.specifications.core.database import SpecsDatabase
from modules.specifications.ui.ui_tabs import get_wg_badge_style_and_desc, SpecificationsTab


class TestSpecWGBadge(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()
        cls.db_path = Path("src/db/3gpp_data.db")
        cls.db = SpecsDatabase(cls.db_path)

    def test_search_files_returns_working_group(self):
        # Query known spec 23.501
        rows = self.db.search_files(spec_number="23.501")
        self.assertGreater(len(rows), 0)
        first_row = rows[0]
        # Element 10 must be the primary working group
        self.assertGreaterEqual(len(first_row), 11)
        wg = first_row[10]
        self.assertEqual(wg, "SA2")

    def test_search_files_ran_and_ct_groups(self):
        # Test 38.413 (RAN3)
        rows_ran = self.db.search_files(spec_number="38.413")
        if rows_ran:
            self.assertEqual(rows_ran[0][10], "RAN3")

        # Test 24.501 (CT1)
        rows_ct = self.db.search_files(spec_number="24.501")
        if rows_ct:
            self.assertEqual(rows_ct[0][10], "CT1")

    def test_wg_badge_styles_and_descriptions(self):
        # SA group -> indigo style
        sa_style, sa_desc = get_wg_badge_style_and_desc("SA2")
        self.assertIn("#EEF2FF", sa_style)
        self.assertIn("#3730A3", sa_style)
        self.assertIn("Architecture", sa_desc)

        # RAN group -> emerald style
        ran_style, ran_desc = get_wg_badge_style_and_desc("RAN3")
        self.assertIn("#ECFDF5", ran_style)
        self.assertIn("#065F46", ran_style)
        self.assertIn("Architecture & Network Interfaces", ran_desc)

        # CT group -> amber style
        ct_style, ct_desc = get_wg_badge_style_and_desc("CT1")
        self.assertIn("#FFF7ED", ct_style)
        self.assertIn("#9A3412", ct_style)
        self.assertIn("Protocols", ct_desc)

        # Fallback / Other group -> slate style
        def_style, def_desc = get_wg_badge_style_and_desc("CUSTOM_WG")
        self.assertIn("#F1F5F9", def_style)
        self.assertIn("#475569", def_style)
        self.assertEqual(def_desc, "Working Group")

    def test_spec_tab_column_width_and_wg_badge_click(self):
        tab = SpecificationsTab(self.db_path)
        # Column 0 width should be 195 to fit action button, type badge, WG badge, and spec number
        self.assertEqual(tab.table.columnWidth(0), 195)

        # Populate a single spec and verify badges have exact uniform widths
        tab.spec_search_input.setText("23.501")
        tab.refresh_table()
        self.assertGreater(tab.table.rowCount(), 0)
        cell = tab.table.cellWidget(0, 0)
        layout = cell.layout()
        wg_badge = layout.itemAt(1).widget()
        spec_label = layout.itemAt(2).widget()
        self.assertEqual(wg_badge.text(), "SA2")
        self.assertEqual(wg_badge.width(), 42)
        self.assertIn("TS 23.501", spec_label.text())

        # Test 1-click filter by WG
        tab._on_wg_badge_clicked("SA2")
        self.assertEqual(tab.table_filters.get("group"), "SA2")
        self.assertEqual(tab._current_tsg, "SA")


if __name__ == "__main__":
    unittest.main()
