import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import MagicMock
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

from modules.spec_search.core.spec_search_db import (
    SpecSearchDatabase,
    extract_major_release,
)
from modules.spec_search.ui.spec_search_components import (
    SpecCardWidget,
    SpecSearchVersionCardsWidget,
)
from modules.spec_search.ui.spec_search_dialogs import SpecSmartImportDialog
from modules.specifications.core.database import SpecsDatabase

app = QApplication.instance() or QApplication(sys.argv)


class TestSpecSmartImport(unittest.TestCase):
    def setUp(self):
        self.temp_dir = TemporaryDirectory()
        self.search_db_path = Path(self.temp_dir.name) / "test_spec_search.db"
        self.search_db = SpecSearchDatabase(self.search_db_path)
        self.cache_dir = Path(self.temp_dir.name) / "cache"
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        # Mock SpecsDatabase
        self.mock_specs_db = MagicMock(spec=SpecsDatabase)
        self.mock_specs_db.get_spec_details.return_value = {
            "title": "System architecture for the 5G System (5GS)",
            "primary_group": "SA2",
            "type": "TS",
        }
        # Available files in specs database for 23.501
        self.mock_specs_db.get_spec_files.return_value = [
            {"version": "20.3.0", "filename": "23501-k30.zip", "url": "https://3gpp.org/k30.zip", "upload_date": "2026-09-24"},
            {"version": "20.2.0", "filename": "23501-k20.zip", "url": "https://3gpp.org/k20.zip", "upload_date": "2026-06-22"},
            {"version": "19.9.0", "filename": "23501-j90.zip", "url": "https://3gpp.org/j90.zip", "upload_date": "2026-09-24"},
            {"version": "19.8.0", "filename": "23501-j80.zip", "url": "https://3gpp.org/j80.zip", "upload_date": "2026-06-22"},
            {"version": "18.12.0", "filename": "23501-i20.zip", "url": "https://3gpp.org/i20.zip", "upload_date": "2025-12-19"},
            {"version": "17.15.0", "filename": "23501-h50.zip", "url": "https://3gpp.org/h50.zip", "upload_date": "2025-12-19"},
            {"version": "16.20.0", "filename": "23501-g00.zip", "url": "https://3gpp.org/g00.zip", "upload_date": "2024-06-26"},
        ]

        # Seed Search DB:
        # Rel-20 has older 20.2.0 (Update available to 20.3.0)
        # Rel-18 has 18.12.0 (Up to date)
        # Rel-17 has 17.15.0 (Up to date)
        # Rel-19 is not indexed at all (New Release)
        # Rel-16 is not indexed at all (New Release)
        clauses = [{"clause_number": "1", "clause_title": "Scope", "content": "Text"}]
        self.search_db.insert_parsed_spec("23.501", "20.2.0", clauses, release_date="2026-06-22")
        self.search_db.insert_parsed_spec("23.501", "18.12.0", clauses, release_date="2025-12-19")
        self.search_db.insert_parsed_spec("23.501", "17.15.0", clauses, release_date="2025-12-19")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_smart_import_dialog_initialization_and_delta(self):
        dialog = SpecSmartImportDialog(
            specs_db=self.mock_specs_db,
            search_db=self.search_db,
            cache_dir=self.cache_dir,
            initial_spec="23.501",
        )

        self.assertEqual(dialog.current_spec, "23.501")
        self.assertEqual(dialog.table.rowCount(), 5)  # Rel-20, Rel-19, Rel-18, Rel-17, Rel-16

        # Check calculated statuses in releases_data
        status_map = {item["release"]: item["status"] for item in dialog.releases_data}
        self.assertEqual(status_map["Rel-20"], "update")   # 20.2.0 -> 20.3.0
        self.assertEqual(status_map["Rel-19"], "new")      # not indexed
        self.assertEqual(status_map["Rel-18"], "current")  # 18.12.0 matches
        self.assertEqual(status_map["Rel-17"], "current")  # 17.15.0 matches
        self.assertEqual(status_map["Rel-16"], "new")      # not indexed

        # By default, updates and new should be checked (Rel-20, Rel-19, Rel-16 = 3 rows)
        checked_releases = []
        for r in range(dialog.table.rowCount()):
            if dialog.table.item(r, 0).checkState() == Qt.Checked:
                checked_releases.append(dialog.table.item(r, 1).text())

        self.assertEqual(sorted(checked_releases), ["Rel-16", "Rel-19", "Rel-20"])
        self.assertTrue(dialog.btn_ingest.isEnabled())
        self.assertIn("3", dialog.btn_ingest.text())

    def test_smart_import_dialog_filter_chips(self):
        dialog = SpecSmartImportDialog(
            specs_db=self.mock_specs_db,
            search_db=self.search_db,
            cache_dir=self.cache_dir,
            initial_spec="23.501",
        )

        # 1. Clear All
        dialog._clear_selection()
        checked_count = sum(1 for r in range(dialog.table.rowCount()) if dialog.table.item(r, 0).checkState() == Qt.Checked)
        self.assertEqual(checked_count, 0)
        self.assertFalse(dialog.btn_ingest.isEnabled())

        # 2. Select All
        dialog._select_all()
        checked_count = sum(1 for r in range(dialog.table.rowCount()) if dialog.table.item(r, 0).checkState() == Qt.Checked)
        self.assertEqual(checked_count, 5)

        # 3. Select Rel-18+ (Rel-20, Rel-19, Rel-18)
        dialog._select_min_release(18)
        checked_releases = [
            dialog.table.item(r, 1).text()
            for r in range(dialog.table.rowCount())
            if dialog.table.item(r, 0).checkState() == Qt.Checked
        ]
        self.assertEqual(sorted(checked_releases), ["Rel-18", "Rel-19", "Rel-20"])

        # 4. Updates & New Only
        dialog._select_updates_and_new()
        checked_releases = [
            dialog.table.item(r, 1).text()
            for r in range(dialog.table.rowCount())
            if dialog.table.item(r, 0).checkState() == Qt.Checked
        ]
        self.assertEqual(sorted(checked_releases), ["Rel-16", "Rel-19", "Rel-20"])

    def test_smart_import_task_generation(self):
        dialog = SpecSmartImportDialog(
            specs_db=self.mock_specs_db,
            search_db=self.search_db,
            cache_dir=self.cache_dir,
            initial_spec="TS 23.501",
        )

        # Select Rel-18+ (Rel-20, Rel-19, Rel-18)
        dialog._select_min_release(18)
        dialog.chk_auto_supersede.setChecked(True)
        dialog._on_ingest_confirmed()

        self.assertEqual(len(dialog.selected_tasks), 3)
        versions_in_tasks = [t["version"] for t in dialog.selected_tasks]
        self.assertIn("20.3.0", versions_in_tasks)
        self.assertIn("19.9.0", versions_in_tasks)
        self.assertIn("18.12.0", versions_in_tasks)
        self.assertEqual(dialog.selected_tasks[0]["spec_number"], "23.501")
        self.assertTrue(dialog.auto_replace_older)

    def test_spec_card_smart_import_signal(self):
        versions = [
            {"id": 1, "spec_number": "23.501", "version": "18.12.0", "release_date": "2025-12-19"},
        ]
        card = SpecCardWidget("23.501", "5GS Architecture", versions)
        emitted_specs = []
        card.smart_import_requested.connect(lambda s: emitted_specs.append(s))

        card.smart_import_requested.emit("23.501")
        self.assertEqual(emitted_specs, ["23.501"])

        # Test forwarding through SpecSearchVersionCardsWidget
        cards_widget = SpecSearchVersionCardsWidget()
        forwarded_specs = []
        cards_widget.smart_import_requested.connect(lambda s: forwarded_specs.append(s))
        cards_widget.populate(versions)

        # Trigger on card
        card_in_widget = cards_widget._cards.get("23.501")
        self.assertIsNotNone(card_in_widget)
        card_in_widget.smart_import_requested.emit("23.501")
        self.assertEqual(forwarded_specs, ["23.501"])


if __name__ == "__main__":
    unittest.main()
