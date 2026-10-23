import sys
import unittest
from pathlib import Path
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import QApplication

from modules.spec_search.ui.spec_search_components import (
    SpecClauseInspector,
    SpecEvolutionStatusDelegate,
    SpecSearchVersionTreeWidget,
)
from modules.spec_search.ui.spec_search_models import SpecEvolutionMatrixModel
from modules.spec_search.ui.spec_search_tabs import SpecSearchTab

app = QApplication.instance() or QApplication(sys.argv)


class TestSpecSearchUIRedesign(unittest.TestCase):
    def test_version_tree_filter_and_selection(self):
        tree = SpecSearchVersionTreeWidget()
        sample_versions = [
            {"id": 1, "spec_number": "23.501", "version": "18.2.0", "release_date": "2024-03-20"},
            {"id": 2, "spec_number": "23.501", "version": "17.1.0", "release_date": "2023-01-15"},
            {"id": 3, "spec_number": "38.300", "version": "17.0.0", "release_date": "2023-06-10"},
        ]
        tree.populate(sample_versions)

        # Verify initial population
        self.assertEqual(tree.topLevelItemCount(), 3)  # All, 23.501, 38.300
        spec_item = tree.topLevelItem(1)
        self.assertIn("TS 23.501", spec_item.text(0))
        self.assertIn("(2)", spec_item.text(0))  # 2 versions under 23.501

        # Test filter: search for "38"
        tree.filter_tree("38")
        self.assertTrue(tree.topLevelItem(1).isHidden())   # TS 23.501 should be hidden
        self.assertFalse(tree.topLevelItem(2).isHidden())  # TS 38.300 should be visible

        # Clear filter
        tree.filter_tree("")
        self.assertFalse(tree.topLevelItem(1).isHidden())
        self.assertFalse(tree.topLevelItem(2).isHidden())

        # Test select all / deselect all
        tree.deselect_all()
        self.assertEqual(len(tree.get_selected_version_ids()), 0)
        tree.select_all()
        self.assertEqual(len(tree.get_selected_version_ids()), 3)

    def test_clause_inspector_two_tier_and_deduplication(self):
        inspector = SpecClauseInspector()

        # Track signals
        layout_toggled = []
        maximize_toggled = []
        inspector.layout_orientation_toggle_requested.connect(layout_toggled.append)
        inspector.maximize_toggle_requested.connect(maximize_toggled.append)

        # Test layout orientation toggle
        inspector.btn_layout_toggle.click()
        self.assertTrue(layout_toggled and layout_toggled[-1] is True)
        self.assertIn("Bottom", inspector.btn_layout_toggle.text())

        inspector.btn_layout_toggle.click()
        self.assertTrue(layout_toggled and layout_toggled[-1] is False)
        self.assertIn("Right Pane", inspector.btn_layout_toggle.text())

        # Test maximize toggle
        inspector.btn_maximize.click()
        self.assertTrue(maximize_toggled and maximize_toggled[-1] is True)
        self.assertEqual(inspector.btn_maximize.text(), "⤡")

        inspector.btn_maximize.click()
        self.assertTrue(maximize_toggled and maximize_toggled[-1] is False)
        self.assertEqual(inspector.btn_maximize.text(), "⤢")

        # Test display clause with repetitive heading
        raw_content = "5.15 Support of No-Transmit Zones\nThe following principles apply to NTZ.\nNTZ rules are defined here."
        inspector.display_clause(
            clause_number="5.15",
            clause_title="Support of No-Transmit Zones",
            spec_number="23.256",
            version="19.5.0",
            content=raw_content,
            release_date="2025-12-19",
            search_query="No-Transmit Zones",
            first_added_ver="v19.5.0",
        )

        html_content = inspector.browser.toHtml()
        self.assertIn("Clause 5.15", html_content)
        self.assertIn("Support of No-Transmit Zones", html_content)
        # Check that buttons are enabled
        self.assertTrue(inspector.btn_copy.isEnabled())
        self.assertTrue(inspector.btn_extract.isEnabled())
        self.assertTrue(inspector.btn_open_doc.isEnabled())
        self.assertTrue(inspector.btn_show_in_specs.isEnabled())
        self.assertTrue(inspector.btn_llm_diff.isEnabled())

    def test_evolution_matrix_model_clean_roles(self):
        import pandas as pd
        df = pd.DataFrame([
            {
                "spec_number": "23.256",
                "version": "19.5.0",
                "release_date": "2025-12-19",
                "clause_number": "5.15",
                "clause_title": "Support of No-Transmit Zones",
                "snippet_text": "Support of <mark>No-Transmit Zones</mark>",
                "clause_pk": 101,
                "order_index": 1,
            }
        ])
        model = SpecEvolutionMatrixModel(raw_df=df)
        self.assertEqual(model.rowCount(), 1)
        self.assertEqual(model.columnCount(), 3)  # Clause, Title, v19.5.0

        # Clause and Title
        self.assertEqual(model.data(model.index(0, 0), Qt.DisplayRole), "5.15")
        self.assertEqual(model.data(model.index(0, 1), Qt.DisplayRole), "Support of No-Transmit Zones")

        # Version status display role
        status_val = model.data(model.index(0, 2), Qt.DisplayRole)
        self.assertEqual(status_val, "🟢 Added")

        # Background role is None (handled by custom pill delegate)
        self.assertIsNone(model.data(model.index(0, 2), Qt.BackgroundRole))


if __name__ == "__main__":
    unittest.main()
