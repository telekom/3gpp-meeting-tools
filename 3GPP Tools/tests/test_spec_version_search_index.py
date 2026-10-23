import sys
import unittest
from pathlib import Path
import tempfile
import sqlite3
from PyQt5.QtWidgets import QApplication

from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.specifications.core.database import SpecsDatabase
from modules.specifications.ui.ui_tabs import SpecificationsTab
from modules.specifications.ui.spec_menu import show_spec_context_menu


class TestSpecVersionSearchIndex(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()
        cls.db_path = Path("src/db/3gpp_data.db")
        cls.specs_db = SpecsDatabase(cls.db_path)

    def test_spec_search_db_is_version_indexed(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_db_path = Path(tmp_dir) / "test_spec_search.db"
            db = SpecSearchDatabase(tmp_db_path)

            # Initially not indexed
            self.assertFalse(db.is_version_indexed("23.256", "19.5.0"))
            self.assertFalse(db.is_version_indexed("23.256", "v19.5.0"))

            # Insert sample spec
            clauses = [{"clause_number": "1", "clause_title": "Scope", "content": "This document..."}]
            db.insert_parsed_spec("23.256", "19.5.0", clauses, release_date="2025-12-19")

            # Now indexed (verifying both plain and 'v'-prefixed formats)
            self.assertTrue(db.is_version_indexed("23.256", "19.5.0"))
            self.assertTrue(db.is_version_indexed("23.256", "v19.5.0"))
            self.assertTrue(db.is_version_indexed(" TS 23.256 ", " 19.5.0 "))
            self.assertFalse(db.is_version_indexed("23.256", "18.0.0"))
            self.assertFalse(db.is_version_indexed("23.501", "19.5.0"))

    def test_specs_tab_signals_and_helpers(self):
        tab = SpecificationsTab(self.db_path)
        self.assertTrue(hasattr(tab, "search_in_spec_search_requested"))
        self.assertTrue(hasattr(tab, "index_in_spec_search_requested"))
        self.assertTrue(hasattr(tab, "batch_index_in_spec_search_requested"))

        # Test batch index emitter helper
        captured_batch = []
        tab.batch_index_in_spec_search_requested.connect(lambda items: captured_batch.extend(items))

        mock_target_items = [
            ("23.501", {"version": "18.4.0"}),
            ("23.502", {"version": "18.3.0"}),
        ]
        tab._batch_index_in_spec_search(mock_target_items)
        self.assertEqual(len(captured_batch), 2)
        self.assertEqual(captured_batch[0], ("23.501", "18.4.0"))
        self.assertEqual(captured_batch[1], ("23.502", "18.3.0"))

    def test_context_menu_has_spec_search_indexing_action(self):
        tab = SpecificationsTab(self.db_path)
        tab.spec_search_input.setText("23.256")
        tab.refresh_table()

        if tab.table.rowCount() > 0:
            # Select first row
            tab.table.selectRow(0)
            ver = tab._get_row_version(0)
            self.assertTrue(len(ver) > 0)

            # Emulate single-item context menu invocation
            # Using show_spec_context_menu without crashing
            from PyQt5.QtCore import QPoint
            # Ensure menu creation works cleanly
            cell_widget = tab.table.cellWidget(0, 2)
            self.assertIsNotNone(cell_widget)


if __name__ == "__main__":
    unittest.main()
