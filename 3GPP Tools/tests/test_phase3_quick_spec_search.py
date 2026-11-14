import sys
import tempfile
import time
import unittest
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from PyQt5.QtWidgets import QApplication

from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.spec_search.ui.quick_search_panel import QuickSpecSearchPanel


class TestPhase3QuickSpecSearch(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def _populate_db(self, db_path: Path) -> SpecSearchDatabase:
        db = SpecSearchDatabase(db_path)
        db.insert_parsed_spec(
            "23.501",
            "17.11.0",
            [
                {
                    "clause_number": "5.3.2",
                    "clause_title": "PEI and SUPI",
                    "content": "Legacy Rel-17 description of Permanent Equipment Identifier registration.",
                    "level": 3,
                    "order_index": 1,
                }
            ],
            release_date="2023-09-15",
            auto_replace_older=False,
        )
        db.insert_parsed_spec(
            "23.501",
            "18.6.0",
            [
                {
                    "clause_number": "5.3.2",
                    "clause_title": "PEI and SUPI",
                    "content": "Rel-18 registration procedure carrying Permanent Equipment Identifier over NAS.",
                    "level": 3,
                    "order_index": 1,
                },
                {
                    "clause_number": "5.6.1",
                    "clause_title": "DNN Selection",
                    "content": "Data Network Name validation during PDU session establishment.",
                    "level": 3,
                    "order_index": 2,
                },
            ],
            release_date="2024-12-20",
            auto_replace_older=False,
        )
        db.insert_parsed_spec(
            "24.501",
            "18.5.0",
            [
                {
                    "clause_number": "5.5.1.2",
                    "clause_title": "Initial registration",
                    "content": "UE initiates 5GMM registration request carrying PEI when configured.",
                    "level": 4,
                    "order_index": 1,
                }
            ],
            release_date="2024-12-10",
            auto_replace_older=False,
        )
        return db

    def test_latest_per_spec_scoping_async_search_preview_and_actions(self):
        with tempfile.TemporaryDirectory() as tmp:
            db_path = Path(tmp) / "spec_search.db"
            self._populate_db(db_path)

            panel = QuickSpecSearchPanel(db_path)
            panel.resize(390, 720)
            panel.show()
            self.app.processEvents()

            # Latest per spec is enabled by default -> only 2 version IDs (23.501 v18.6.0, 24.501 v18.5.0)
            self.assertTrue(panel.latest_only_btn.isChecked())
            self.assertEqual(len(panel.resolve_version_ids()), 2)

            # Spec filter with prefix normalization
            panel.spec_filter_input.setText("TS 23.501")
            self.assertEqual(len(panel.resolve_version_ids()), 1)

            # Unchecking Latest per Spec includes both Rel-17 and Rel-18 for 23.501
            panel.latest_only_btn.setChecked(False)
            self.assertEqual(len(panel.resolve_version_ids()), 2)

            # Reset filters and run async FTS5 search
            panel.spec_filter_input.clear()
            panel.latest_only_btn.setChecked(True)
            panel.search_input.setText("Permanent Equipment Identifier")
            panel._execute_search()

            deadline = time.monotonic() + 5.0
            while panel.results_list.count() == 0 and time.monotonic() < deadline:
                self.app.processEvents()
                time.sleep(0.02)
            self.app.processEvents()

            self.assertEqual(panel.results_list.count(), 1)
            self.assertIn("1 matching clause", panel.status_label.text())

            # Discard stale response check
            stale_req_id = panel._req_id - 1
            panel._on_results_ready(None, stale_req_id)
            self.assertEqual(panel.results_list.count(), 1)

            # Selecting result shows Clause Preview and emits deep-link signals
            search_calls = []
            specs_calls = []
            panel.open_in_spec_search_requested.connect(
                lambda s, v, c: search_calls.append((s, v, c))
            )
            panel.open_in_specs_requested.connect(lambda s, v: specs_calls.append((s, v)))

            first_item = panel.results_list.item(0)
            panel.results_list.setCurrentItem(first_item)
            panel._on_result_selected(first_item)
            self.app.processEvents()

            self.assertTrue(panel.preview_frame.isVisible())
            self.assertIn("23.501", panel.preview_title_label.text())
            self.assertIn("18.6.0", panel.preview_title_label.text())
            self.assertIn("Rel-18 registration procedure", panel.preview_browser.toPlainText())

            panel.open_search_btn.click()
            panel.open_specs_btn.click()
            panel.copy_citation_btn.click()
            self.app.processEvents()

            self.assertEqual(search_calls, [("23.501", "18.6.0", "5.3.2")])
            self.assertEqual(specs_calls, [("23.501", "18.6.0")])
            clip_text = QApplication.clipboard().text()
            self.assertIn("TS 23.501 v18.6.0 Clause 5.3.2", clip_text)
            self.assertIn("Rel-18 registration procedure", clip_text)

            panel.shutdown()
            panel.close()


if __name__ == "__main__":
    unittest.main()
