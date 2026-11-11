import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

from modules.spec_search.core.spec_search_db import (
    SpecSearchDatabase,
    extract_major_release,
    parse_version_tuple,
)
from modules.spec_search.ui.spec_search_components import (
    ReleasePillWidget,
    SpecCardWidget,
    SpecSearchVersionCardsWidget,
)
from modules.spec_search.ui.spec_search_dialogs import SpecSearchPruneDialog

app = QApplication.instance() or QApplication(sys.argv)


class TestSpecSearchRetentionAndPills(unittest.TestCase):
    def setUp(self):
        self.temp_dir = TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_spec_search.db"
        self.db = SpecSearchDatabase(self.db_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_extract_major_release(self):
        # Dot notation
        self.assertEqual(extract_major_release("15.16.0"), "Rel-15")
        self.assertEqual(extract_major_release("v15.17.0"), "Rel-15")
        self.assertEqual(extract_major_release("18.4.0"), "Rel-18")
        self.assertEqual(extract_major_release("20.2.0"), "Rel-20")
        self.assertEqual(extract_major_release("0.10.0"), "Rel-0")

        # 3-digit lettered codes
        self.assertEqual(extract_major_release("g30"), "Rel-16")
        self.assertEqual(extract_major_release("h20"), "Rel-17")
        self.assertEqual(extract_major_release("i40"), "Rel-18")

        # Edge cases
        self.assertEqual(extract_major_release(""), "Unknown")
        self.assertEqual(extract_major_release("draft"), "draft")

    def test_database_superseded_detection_and_pruning(self):
        clauses = [
            {"clause_number": "1", "clause_title": "Scope", "content": "Scope text"},
            {"clause_number": "2", "clause_title": "References", "content": "Ref text"},
        ]

        # Insert multiple versions for TS 23.501
        # Two versions in Rel-15: 15.16.0 and 15.17.0
        self.db.insert_parsed_spec("23.501", "15.16.0", clauses, release_date="2023-06-01")
        self.db.insert_parsed_spec("23.501", "15.17.0", clauses, release_date="2023-09-19")
        # One version in Rel-16
        self.db.insert_parsed_spec("23.501", "16.19.0", clauses, release_date="2024-03-20")
        # Two versions in Rel-18: 18.2.0 and 18.14.0
        self.db.insert_parsed_spec("23.501", "18.2.0", clauses, release_date="2024-01-10")
        self.db.insert_parsed_spec("23.501", "18.14.0", clauses, release_date="2026-06-22")

        # Insert for another spec TS 38.300
        self.db.insert_parsed_spec("38.300", "17.0.0", clauses, release_date="2022-03-15")

        # Check superseded versions
        superseded = self.db.get_superseded_versions()
        self.assertEqual(len(superseded), 2)  # 15.16.0 and 18.2.0 are superseded

        sup_versions = {s["version"]: s["superseded_by"] for s in superseded}
        self.assertIn("15.16.0", sup_versions)
        self.assertEqual(sup_versions["15.16.0"], "15.17.0")
        self.assertIn("18.2.0", sup_versions)
        self.assertEqual(sup_versions["18.2.0"], "18.14.0")

        # Prune only 15.16.0
        target_id = [s["id"] for s in superseded if s["version"] == "15.16.0"]
        pruned_count = self.db.prune_superseded_versions(target_id)
        self.assertEqual(pruned_count, 1)

        remaining_superseded = self.db.get_superseded_versions()
        self.assertEqual(len(remaining_superseded), 1)
        self.assertEqual(remaining_superseded[0]["version"], "18.2.0")

        # Prune remaining
        pruned_all = self.db.prune_superseded_versions()
        self.assertEqual(pruned_all, 1)
        self.assertEqual(len(self.db.get_superseded_versions()), 0)

    def test_auto_replace_older_on_ingest(self):
        clauses = [{"clause_number": "1", "clause_title": "Scope", "content": "Scope text"}]

        # Ingest 15.16.0 without auto-replace
        self.db.insert_parsed_spec("23.501", "15.16.0", clauses, release_date="2023-06-01")
        self.assertTrue(self.db.is_version_indexed("23.501", "15.16.0"))

        # Ingest 15.17.0 WITH auto_replace_older=True
        self.db.insert_parsed_spec(
            "23.501", "15.17.0", clauses, release_date="2023-09-19", auto_replace_older=True
        )

        # 15.16.0 should be automatically purged!
        self.assertFalse(self.db.is_version_indexed("23.501", "15.16.0"))
        # 15.17.0 should be present
        self.assertTrue(self.db.is_version_indexed("23.501", "15.17.0"))

    def test_release_pills_card_widget_interactions(self):
        sample_versions = [
            {"id": 10, "spec_number": "23.501", "version": "15.16.0", "release_date": "2023-06-01", "total_clauses": 100},
            {"id": 11, "spec_number": "23.501", "version": "15.17.0", "release_date": "2023-09-19", "total_clauses": 110},
            {"id": 12, "spec_number": "23.501", "version": "18.14.0", "release_date": "2026-06-22", "total_clauses": 200},
            {"id": 20, "spec_number": "38.300", "version": "17.0.0", "release_date": "2022-03-15", "total_clauses": 80},
        ]
        titles = {"23.501": "5G System Architecture", "38.300": "NR Overall Description"}

        widget = SpecSearchVersionCardsWidget()
        widget.populate(sample_versions, spec_titles=titles)

        # Verify cards
        self.assertEqual(len(widget._cards), 2)
        card_23501 = widget._cards["23.501"]
        self.assertEqual(len(card_23501.pills), 3)
        self.assertIn("5G System Architecture", card_23501.lbl_title.toolTip())

        # Test pills text formatting
        pill_texts = [p.text() for p in card_23501.pills]
        self.assertIn("R18: 18.14.0", pill_texts)
        self.assertIn("R15: 15.17.0", pill_texts)
        self.assertIn("R15: 15.16.0", pill_texts)

        # Initially all checked
        self.assertEqual(len(widget.get_selected_version_ids()), 4)

        # Test select_latest_per_release:
        # TS 23.501 has 18.14.0 (R18), 15.17.0 (R15), 15.16.0 (R15)
        # Should keep 18.14.0, 15.17.0, and 17.0.0 (38.300), unchecking 15.16.0
        widget.select_latest_per_release()
        selected_ids = widget.get_selected_version_ids()
        self.assertEqual(len(selected_ids), 3)
        self.assertNotIn(10, selected_ids)  # 15.16.0 unchecked
        self.assertIn(11, selected_ids)     # 15.17.0 checked
        self.assertIn(12, selected_ids)     # 18.14.0 checked
        self.assertIn(20, selected_ids)     # 17.0.0 checked

        # Test select_releases_min(18):
        # Only Rel-18 (id 12) should remain checked
        widget.select_releases_min(18)
        min18_ids = widget.get_selected_version_ids()
        self.assertEqual(min18_ids, [12])

        # Test deselect all and select all
        widget.deselect_all()
        self.assertEqual(len(widget.get_selected_version_ids()), 0)
        widget.select_all()
        self.assertEqual(len(widget.get_selected_version_ids()), 4)

        # Test filtering
        widget.filter_tree("NR")
        self.assertTrue(widget._cards["23.501"].isHidden())
        self.assertFalse(widget._cards["38.300"].isHidden())

        widget.filter_tree("")
        self.assertFalse(widget._cards["23.501"].isHidden())
        self.assertFalse(widget._cards["38.300"].isHidden())

    def test_prune_dialog_population(self):
        superseded = [
            {"id": 1, "spec_number": "23.501", "version": "15.16.0", "release": "Rel-15", "release_date": "2023-06-01", "total_clauses": 100, "superseded_by": "15.17.0"},
            {"id": 2, "spec_number": "23.501", "version": "18.2.0", "release": "Rel-18", "release_date": "2024-01-10", "total_clauses": 150, "superseded_by": "18.14.0"},
        ]
        dialog = SpecSearchPruneDialog(self.db, superseded)
        self.assertEqual(dialog.table.rowCount(), 2)
        self.assertEqual(dialog.table.item(0, 1).text(), "TS 23.501")
        self.assertEqual(dialog.table.item(0, 2).text(), "v15.16.0")
        self.assertEqual(dialog.table.item(0, 3).text(), "Rel-15")
        self.assertIn("15.17.0", dialog.table.item(0, 4).text())


if __name__ == "__main__":
    unittest.main()
