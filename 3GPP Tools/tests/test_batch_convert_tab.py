"""Unit tests for BatchConvertTab state management and card compatibility."""

import sys
import unittest
from pathlib import Path
from PyQt5.QtWidgets import QApplication, QLineEdit, QTableWidget

from core.ui.ui_components import InteractiveDropLabel
from modules.puml2visio.ui.batch_convert_tab import BatchConvertTab


class TestBatchConvertTab(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def setUp(self):
        self.tab = BatchConvertTab()

    def test_initialization_and_backward_compatibility_aliases(self):
        """Verify that BatchConvertTab initializes properly with all backward-compatible aliases."""
        self.assertIsNotNone(self.tab)
        self.assertTrue(hasattr(self.tab, "card_puml"))
        self.assertTrue(hasattr(self.tab, "card_bridge"))
        self.assertTrue(hasattr(self.tab, "card_extractor"))
        self.assertTrue(hasattr(self.tab, "card_stencils"))

        # Verify backward-compatibility aliases
        self.assertTrue(hasattr(self.tab, "drop_label_puml"))
        self.assertIsInstance(self.tab.drop_label_puml, InteractiveDropLabel)
        self.assertTrue(hasattr(self.tab, "stencils_search_edit"))
        self.assertIsInstance(self.tab.stencils_search_edit, QLineEdit)
        self.assertTrue(hasattr(self.tab, "puml_table"))
        self.assertIsInstance(self.tab.puml_table, QTableWidget)
        self.assertTrue(hasattr(self.tab, "bridge_table"))
        self.assertIsInstance(self.tab.bridge_table, QTableWidget)
        self.assertTrue(hasattr(self.tab, "stencils_table"))
        self.assertIsInstance(self.tab.stencils_table, QTableWidget)

    def test_set_state_with_ready_and_legacy_message(self):
        """Verify set_state handles the exact invocation from main_window.py without TypeError."""
        legacy_msg = (
            "📥 Drag & Drop your .puml, .txt, .pptx, or .vsdx file(s) here\n\n"
            "(Batch converts between Visio & PowerPoint)"
        )
        # Must not raise TypeError
        self.tab.set_state("ready", legacy_msg)
        # Should set card-specific plantuml guidance
        self.assertIn("PlantUML", self.tab.drop_label_puml.text())

    def test_set_state_with_error(self):
        """Verify set_state handles error state with custom message."""
        err_msg = "❌ Core engine error. Java or PlantUML engine is unavailable."
        self.tab.set_state("error", err_msg)
        self.assertEqual(self.tab.drop_label_puml.text(), err_msg)

    def test_set_state_with_busy(self):
        """Verify set_state handles busy state."""
        busy_msg = "⚙️ Processing Queue...\n\nPlease wait until finished."
        self.tab.set_state("busy", busy_msg)
        self.assertEqual(self.tab.drop_label_puml.text(), busy_msg)

    def test_set_state_with_boolean(self):
        """Verify set_state handles boolean is_converting argument gracefully."""
        self.tab.set_state(True)
        self.tab.set_state(False)

    def test_on_conversion_success(self):
        """Verify on_conversion_success executes cleanly."""
        self.tab.on_conversion_success("/fake/path/test.vsdx")


if __name__ == "__main__":
    unittest.main()
