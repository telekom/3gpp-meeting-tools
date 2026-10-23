import sys
import unittest
from pathlib import Path
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

from modules.nas.ui.nas_dialogs import (
    PROTOCOL_SPEC_CATALOG,
    get_spec_metadata,
    get_spec_title,
    get_spec_tooltip,
)
from modules.nas.ui.nas_models import NASEvolutionMatrixModel
from modules.nas.ui.nas_components import (
    NASInspectorWidget,
    NASMessageItemDelegate,
    NASVersionTreeWidget,
)

# Ensure QApplication exists for Qt widgets
app = QApplication.instance() or QApplication(sys.argv)


class TestNASUiVisuals(unittest.TestCase):
    """Verifies visual enhancements, specification clue resolution, and model rendering in Protocols UI."""

    def test_protocol_spec_catalog_and_metadata(self):
        """Verifies that key 3GPP specifications resolve to rich descriptive metadata."""
        specs_to_test = ["24.501", "24.301", "29.244", "29.281", "38.331", "38.413", "38.415"]

        for s_num in specs_to_test:
            meta = get_spec_metadata(s_num)
            self.assertIn("acronym", meta)
            self.assertIn("wg", meta)
            self.assertIn("title", meta)
            self.assertIn("full_title", meta)
            self.assertIn("description", meta)
            self.assertTrue(len(meta["description"]) > 10)

            # Verify tooltip contains Working Group and description
            tooltip = get_spec_tooltip(s_num)
            self.assertIn(meta["wg"], tooltip)
            self.assertIn(meta["description"], tooltip)

    def test_heuristic_spec_metadata_fallback(self):
        """Verifies that uncatalogued specs gracefully infer series and WG."""
        meta = get_spec_metadata("24.999")
        self.assertEqual(meta["wg"], "CT1")
        self.assertEqual(meta["tech"], "Core")

        meta_ran = get_spec_metadata("38.399")
        self.assertEqual(meta_ran["wg"], "RAN2")

    def test_version_tree_population_and_clues(self):
        """Verifies that version tree items display Working Group badges and Release tags."""
        tree = NASVersionTreeWidget()
        versions = [
            {"id": 1, "spec_number": "24.501", "version": "20.0.0"},
            {"id": 2, "spec_number": "24.501", "version": "19.7.0"},
            {"id": 3, "spec_number": "29.244", "version": "20.0.0"},
        ]
        tree.populate(versions)

        # Top level has master toggle + 2 specs
        self.assertEqual(tree.topLevelItemCount(), 3)

        # Spec 24.501 should show [CT1] and 5GS NAS
        spec_item_24501 = tree.topLevelItem(1)
        self.assertIn("CT1", spec_item_24501.text(0))
        self.assertIn("24.501", spec_item_24501.text(0))

        # Child version should display Rel-20
        child_v20 = spec_item_24501.child(0)
        self.assertIn("Rel-20", child_v20.text(0))
        self.assertIn("v20.0.0", child_v20.text(0))

        # Test quick selection helpers
        tree.select_none()
        self.assertEqual(len(tree.get_selected_version_ids()), 0)

        tree.select_all()
        self.assertEqual(len(tree.get_selected_version_ids()), 3)

        tree.select_latest_only()
        selected_ids = tree.get_selected_version_ids()
        # Should select latest v20.0.0 for 24.501 (id 1) and latest v20.0.0 for 29.244 (id 3)
        self.assertEqual(len(selected_ids), 2)
        self.assertIn(1, selected_ids)
        self.assertIn(3, selected_ids)

    def test_evolution_matrix_headers_with_release(self):
        """Verifies that table headers format version columns with clear Release badges."""
        df = pd.DataFrame([
            {
                "iei": "5E",
                "ie_name": "PDU session identity 2",
                "type_reference": "9.11.4.10",
                "presence": "M",
                "format": "V",
                "length": "1",
                "spec_number": "24.501",
                "version": "19.7.0",
                "order_index": 1,
            },
            {
                "iei": "5E",
                "ie_name": "PDU session identity 2",
                "type_reference": "9.11.4.10",
                "presence": "M",
                "format": "V",
                "length": "1",
                "spec_number": "24.501",
                "version": "20.0.0",
                "order_index": 1,
            },
        ])
        model = NASEvolutionMatrixModel(df)
        self.assertGreaterEqual(model.columnCount(), 4)

        # Columns 0, 1, 2 are IEI, Field, Type
        self.assertEqual(model.headerData(0, Qt.Horizontal, Qt.DisplayRole), "IEI / ID")
        self.assertEqual(model.headerData(1, Qt.Horizontal, Qt.DisplayRole), "Field / Information Element")
        self.assertEqual(model.headerData(2, Qt.Horizontal, Qt.DisplayRole), "Type / Reference")

        # Column 3+ should contain Release tag
        v_header = model.headerData(3, Qt.Horizontal, Qt.DisplayRole)
        self.assertTrue("Rel-19" in v_header or "Rel-20" in v_header)

    def test_inspector_widget_stack_and_empty_state(self):
        """Verifies that the inspector card manages its empty state and populated state."""
        inspector = NASInspectorWidget()
        self.assertEqual(inspector.stack.currentIndex(), 0)  # Starts on empty state
        self.assertFalse(inspector.copy_btn.isEnabled())

        inspector.display_definitions(
            clause="8.2.1",
            ie_name="Test IE",
            spec_number="24.501",
            defs=[{
                "ie_name": "Test IE",
                "spec_number": "24.501",
                "version": "20.0.0",
                "raw_description": "<p>Sample description</p>",
            }],
            containing_msgs=[],
        )
        self.assertEqual(inspector.stack.currentIndex(), 1)  # Switches to content view
        self.assertTrue(inspector.copy_btn.isEnabled())

        inspector.clear_display()
        self.assertEqual(inspector.stack.currentIndex(), 0)  # Reverts to empty state
        self.assertFalse(inspector.copy_btn.isEnabled())


if __name__ == "__main__":
    unittest.main()
