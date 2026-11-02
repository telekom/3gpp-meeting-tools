import sys
import unittest
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.specifications.core.database import SpecsDatabase
from modules.specifications.ui.ui_tabs import SpecificationsTab, USEFUL_WGS


class TestSpecWgTechFilters(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()
        cls.db_path = Path("src/db/3gpp_data.db")
        cls.db = SpecsDatabase(cls.db_path)

    def setUp(self):
        self.tab = SpecificationsTab(self.db_path)
        self.tab._is_loaded = True
        self.tab._set_wg_scope(False)

    def test_comboboxes_are_checkable(self):
        """Verify WG, Series, and Tech all use CheckableComboBox with uniform multi-selection."""
        self.assertIsInstance(self.tab.wg_combo, CheckableComboBox)
        self.assertIsInstance(self.tab.series_combo, CheckableComboBox)
        self.assertIsInstance(self.tab.tech_combo, CheckableComboBox)

        self.assertEqual(self.tab.wg_combo.title, "WG")
        self.assertEqual(self.tab.series_combo.title, "Series")
        self.assertEqual(self.tab.tech_combo.title, "Tech")

    def test_default_useful_wgs_scope(self):
        """Verify WG combobox is limited to useful/active WGs by default."""
        self.assertFalse(self.tab._show_all_wgs)
        self.assertTrue(self.tab.wg_scope_btn.text().startswith("Active WGs"))
        self.assertEqual(self.tab.wg_scope_btn.text(), "Active WGs (18)")
        self.assertFalse(self.tab.wg_scope_btn.isChecked())

        items = [self.tab.wg_combo.model().item(i).data(Qt.UserRole) for i in range(1, self.tab.wg_combo.model().rowCount())]
        self.assertEqual(len(items), 18)
        # First item is SA plenary, followed by SA1
        self.assertEqual(items[0], "SA")
        self.assertEqual(items[1], "SA1")

        # Useful WGs like SA2, RAN1, CT1 should be in the list
        self.assertIn("SA2", items)
        self.assertIn("RAN1", items)
        self.assertIn("CT1", items)

        # Legacy and disbanded groups like CN1, GERAN, SMG01, RAN6 should NOT be in the active list
        self.assertNotIn("CN1", items)
        self.assertNotIn("GERAN", items)
        self.assertNotIn("SMG01", items)
        self.assertNotIn("AHG", items)
        self.assertNotIn("RAN6", items)

    def test_toggle_to_full_wg_scope(self):
        """Verify toggling to full list reveals legacy WGs sorted alphabetically with (Legacy) display tags."""
        self.tab.wg_scope_btn.setChecked(True)
        self.tab._toggle_wg_scope()

        self.assertTrue(self.tab._show_all_wgs)
        self.assertEqual(self.tab.wg_scope_btn.text(), "Full List (33)")

        model = self.tab.wg_combo.model()
        items = [model.item(i).data(Qt.UserRole) for i in range(1, model.rowCount())]
        display_texts = [model.item(i).text() for i in range(1, model.rowCount())]

        self.assertEqual(len(items), 33)
        # In Full mode, items are sorted alphabetically, starting with AHG and CN1
        self.assertEqual(items[0], "AHG")
        self.assertEqual(items[1], "CN1")

        # Legacy groups should now be present with (Legacy) labels
        self.assertIn("CN1", items)
        self.assertIn("GERAN", items)
        self.assertIn("SMG01", items)
        self.assertIn("RAN6", items)

        self.assertIn("AHG (Legacy)", display_texts)
        self.assertIn("CN1 (Legacy)", display_texts)
        self.assertIn("GERAN (Legacy)", display_texts)
        self.assertIn("RAN6 (Legacy)", display_texts)
        self.assertIn("SA2", display_texts)  # Active groups remain clean

        # Toggle back to active
        self.tab.wg_scope_btn.setChecked(False)
        self.tab._toggle_wg_scope()
        self.assertFalse(self.tab._show_all_wgs)
        self.assertEqual(self.tab.wg_scope_btn.text(), "Active WGs (18)")
        items_reverted = [self.tab.wg_combo.model().item(i).data(Qt.UserRole) for i in range(1, self.tab.wg_combo.model().rowCount())]
        self.assertEqual(len(items_reverted), 18)
        self.assertNotIn("CN1", items_reverted)
        self.assertNotIn("RAN6", items_reverted)

    def test_multi_tech_selection(self):
        """Verify multi-select in Tech combobox updates filter and tags."""
        self.tab.tech_combo.setCheckedItems(["5G", "6G"])
        self.tab._on_tech_filter_changed(self.tab.tech_combo.getCheckedItems())

        self.assertEqual(self.tab.table_filters.get("tech"), "5G,6G")
        self.assertEqual(self.tab.tech_combo._refresh_text_and_tooltip(), "Tech: 2 sel")

        # Database search with multiple technologies
        specs = self.db.search_files(tech="5G,6G")
        self.assertGreater(len(specs), 0)

        # Check tag is rendered
        self.tab._update_active_tags()
        rendered = getattr(self.tab, "_rendered_tag_labels", [])
        self.assertTrue(any("Tech: 5G, 6G" in t or "Tech: 5G,6G" in t for t in rendered))

    def test_multi_wg_selection(self):
        """Verify multi-select in WG combobox updates filter and queries correctly."""
        self.tab.wg_combo.setCheckedItems(["SA2", "SA3"])
        self.tab._on_wg_filter_changed(self.tab.wg_combo.getCheckedItems())

        self.assertEqual(self.tab.table_filters.get("group"), "SA2,SA3")
        self.assertEqual(self.tab.wg_combo._refresh_text_and_tooltip(), "WG: 2 sel")

        # Search in DB
        specs = self.db.search_files(group="SA2,SA3")
        self.assertGreater(len(specs), 0)
        # All returned specs must have primary or secondary group SA2 or SA3
        for row in specs[:20]:
            primary_wg = row[10]
            self.assertIn(primary_wg, ("SA2", "SA3", "SA", ""))

    def test_tsg_button_scoping_and_reset(self):
        """Verify clicking TSG button scopes WG combobox to that TSG."""
        # Click SA TSG button
        self.tab._on_tsg_button_clicked("SA")
        self.assertEqual(self.tab._current_tsg, "SA")
        sa_items = [self.tab.wg_combo.model().item(i).data(Qt.UserRole) for i in range(1, self.tab.wg_combo.model().rowCount())]
        self.assertTrue(all(g and g.startswith("SA") for g in sa_items))
        self.assertIn("SA2", sa_items)
        self.assertNotIn("RAN1", sa_items)

        # Click All TSG button
        self.tab._on_tsg_button_clicked("")
        self.assertEqual(self.tab._current_tsg, "")
        all_items = [self.tab.wg_combo.model().item(i).data(Qt.UserRole) for i in range(1, self.tab.wg_combo.model().rowCount())]
        self.assertIn("SA2", all_items)
        self.assertIn("RAN1", all_items)
        self.assertIn("CT1", all_items)

    def test_legacy_badge_click_auto_expands_scope(self):
        """Verify clicking a legacy group badge in the table auto-expands scope to full list."""
        self.assertFalse(self.tab._show_all_wgs)
        self.tab._on_wg_badge_clicked("CN1")

        # Scope should have expanded to show_all_wgs=True to accommodate legacy group
        self.assertTrue(self.tab._show_all_wgs)
        self.assertEqual(self.tab.table_filters.get("group"), "CN1")
        self.assertEqual(self.tab.wg_combo.getCheckedItems(), ["CN1"])


if __name__ == "__main__":
    unittest.main()
