# --- File: tests/test_word_ui_phase2.py ---
import unittest
from pathlib import Path
from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for headless test runs
app = QApplication.instance() or QApplication([])

from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane
from modules.word_tools.ui.word_tabs import WordExtractorTab


class TestWordUiPhase2(unittest.TestCase):
    def test_document_selector_pane_controls(self):
        """Verify DocumentSelectorPane has browse, clear, and metadata controls."""
        pane = DocumentSelectorPane("Test Document")
        self.assertTrue(hasattr(pane, "btn_browse"))
        self.assertTrue(hasattr(pane, "btn_clear"))
        self.assertTrue(hasattr(pane, "lbl_meta"))
        self.assertTrue(hasattr(pane, "drop_zone"))
        self.assertFalse(pane.btn_clear.isEnabled())

        # Test set_file_path with an existing file (e.g. this test file)
        test_file = str(Path(__file__).resolve())
        pane.set_file_path(test_file)
        self.assertEqual(pane.get_inputs(), [test_file])
        self.assertTrue(pane.btn_clear.isEnabled())
        self.assertIn(Path(test_file).name, pane.lbl_meta.text())

        # Test clear_selection
        pane.clear_selection()
        self.assertEqual(pane.get_inputs(), [])
        self.assertFalse(pane.btn_clear.isEnabled())

    def test_compare_slot_swap(self):
        """Verify 1-click slot swap between Document A and Document B."""
        tab = WordExtractorTab()
        self.assertTrue(hasattr(tab, "btn_swap_slots"))
        self.assertTrue(hasattr(tab, "chk_cmp_formatting"))

        test_file_a = str(Path(__file__).resolve())
        test_file_b = str(Path(__file__).resolve().parent / "test_word_ui_phase1.py")

        tab.pane_a.set_file_path(test_file_a)
        tab.pane_b.set_file_path(test_file_b)

        self.assertEqual(tab.pane_a.get_inputs(), [test_file_a])
        self.assertEqual(tab.pane_b.get_inputs(), [test_file_b])

        # Click swap
        tab.btn_swap_slots.click()

        self.assertEqual(tab.pane_a.get_inputs(), [test_file_b])
        self.assertEqual(tab.pane_b.get_inputs(), [test_file_a])


if __name__ == "__main__":
    unittest.main()
