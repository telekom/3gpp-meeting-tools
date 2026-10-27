import sys
import unittest
from unittest.mock import MagicMock

from PyQt5.QtWidgets import QApplication

from modules.spec_search.ui.spec_clause_diff_dialog import SpecClauseDiffDialog


class TestSpecClauseDiffDialog(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        self.mock_db = MagicMock()
        self.mock_db.get_versions_for_spec.return_value = [
            {"version": "17.15.0", "release_date": "2025-12-19"},
            {"version": "16.20.0", "release_date": "2024-06-26"},
            {"version": "15.13.0", "release_date": "2022-03-23"},
        ]
        self.mock_db.get_clause_content_by_spec_ver.side_effect = lambda spec, ver, clause: {
            "16.20.0": {"content": "Line 1: Nnssf interface.\nLine 2: Common text.\n"},
            "17.15.0": {"content": "Line 1: Nnssf interface.\nLine 2: Common text.\nLine 3: Nnsacf interface added.\n"},
        }.get(ver, {"content": ""})

    def test_diff_dialog_initialization_and_preceding_version(self):
        dialog = SpecClauseDiffDialog(
            db=self.mock_db,
            spec_number="23.501",
            current_version="17.15.0",
            clause_number="4.2.6",
            clause_title="Service-based interfaces",
            base_version="16.20.0",
        )

        # Verify title & tabs
        self.assertIn("Compare Clause 4.2.6", dialog.windowTitle())
        self.assertEqual(dialog.tab_widget.count(), 3)
        self.assertEqual(dialog.tab_widget.tabText(0), "📊 Visual Diff")
        self.assertEqual(dialog.tab_widget.tabText(1), "🤖 AI Analysis Prompt")
        self.assertEqual(dialog.tab_widget.tabText(2), "🎯 Modifying TDoc")

        # Verify Attribution Button in Diff Bar
        self.assertTrue(hasattr(dialog, "btn_find_tdoc"))
        self.assertIn("Identify Modifying TDoc", dialog.btn_find_tdoc.text())

        # Verify combo selections
        self.assertEqual(dialog.target_ver_combo.currentData()["version"], "17.15.0")
        self.assertEqual(dialog.base_ver_combo.currentData()["version"], "16.20.0")

        # Verify visual diff content & stats
        diff_html = dialog.diff_browser.toHtml()
        self.assertIn("Nnsacf interface added", diff_html)
        self.assertIn("+1 addition(s)", dialog.diff_stats_lbl.text())

        # Verify Word Track Changes button
        self.assertIn("Track Changes", dialog.btn_word_compare.text())

    def test_diff_formatter_html(self):
        dialog = SpecClauseDiffDialog.__new__(SpecClauseDiffDialog)
        dialog.clause_number = "4.2.6"
        dialog.spec_number = "23.501"

        raw_diff = (
            "--- TS 23.501 v16.20.0\n"
            "+++ TS 23.501 v17.15.0\n"
            "@@ -1,2 +1,3 @@\n"
            " Context line\n"
            "-Old line\n"
            "+New line 1\n"
            "+New line 2\n"
        )

        html_out, additions, deletions = dialog._format_diff_to_html(raw_diff, "16.20.0", "17.15.0")
        self.assertEqual(additions, 2)
        self.assertEqual(deletions, 1)
        self.assertIn("New line 1", html_out)
        self.assertIn("Old line", html_out)

    def test_identical_diff_formatter(self):
        dialog = SpecClauseDiffDialog.__new__(SpecClauseDiffDialog)
        dialog.clause_number = "4.2.6"
        dialog.spec_number = "23.501"

        html_out, additions, deletions = dialog._format_diff_to_html("No text differences detected between these versions.", "16.20.0", "17.15.0")
        self.assertEqual(additions, 0)
        self.assertEqual(deletions, 0)
        self.assertIn("No Text Differences Detected", html_out)


if __name__ == "__main__":
    unittest.main()
