import sys
import unittest
from pathlib import Path
from PyQt5.QtWidgets import QApplication

from modules.meetings.ui.tdocs_components import CheckableComboBox


class TestTDocWindowFilterBar(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def test_checkable_combobox_filtering_and_elision(self):
        cb = CheckableComboBox("Company")
        cb.addItems(["Deutsche Telekom AG, SK Telecom", "Nokia", "Ericsson"])
        
        # Initially, all items are checked -> not actively filtering
        self.assertFalse(cb.isFiltered())
        self.assertTrue(cb.isAllChecked())

        # Select a single long company
        cb.setCheckedItems(["Deutsche Telekom AG, SK Telecom"])
        self.assertTrue(cb.isFiltered())
        self.assertFalse(cb.isAllChecked())
        self.assertEqual(cb.getCheckedItems(), ["Deutsche Telekom AG, SK Telecom"])
        self.assertIn("Deutsche Telekom AG", cb.toolTip())

    def test_tdocs_window_filter_bar_setup(self):
        import tempfile
        from modules.meetings.ui.tdocs_window import TDocsWindow

        with tempfile.TemporaryDirectory() as tmp_dir:
            agenda_path = Path(tmp_dir) / "Agenda" / "test.xlsx"
            agenda_path.parent.mkdir(parents=True, exist_ok=True)
            mtg_info = {"wg_name": "SA2", "meeting_number": "176", "is_electronic": 0}
            window = TDocsWindow(mtg_info, [], str(agenda_path))

            # Verify search input properties
            self.assertGreaterEqual(window.search_input.minimumWidth(), 160)
            self.assertTrue(window.search_input.isClearButtonEnabled())

            # Verify icon-only tracked WI buttons
            self.assertEqual(window.btn_followed_wis.text(), "🎯")
            self.assertEqual(window.btn_starred_wis.text(), "⭐")
            self.assertEqual(window.btn_followed_wis.width(), 30)
            self.assertEqual(window.btn_starred_wis.width(), 30)

            # Verify reset button exists and clears active filters
            self.assertTrue(hasattr(window, "btn_reset_filters"))
            window.search_input.setText("test query")
            window.btn_followed_wis.setChecked(True)
            window.btn_reset_filters.click()
            self.assertEqual(window.search_input.text(), "")
            self.assertFalse(window.btn_followed_wis.isChecked())


if __name__ == "__main__":
    unittest.main()
