# --- File: tests/test_word_ui_phase4.py ---
import unittest
from pathlib import Path
from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for headless test runs
app = QApplication.instance() or QApplication([])

from modules.word_tools.ui.word_tabs import WordExtractorTab


class TestWordUiPhase4(unittest.TestCase):
    def test_visio_extractor_card_controls(self):
        """Verify Visio extractor card has modern DocumentSelectorPane and result frame."""
        tab = WordExtractorTab()
        self.assertTrue(hasattr(tab, "pane_visio"))
        self.assertTrue(hasattr(tab, "run_visio_btn"))
        self.assertTrue(hasattr(tab, "visio_result_frame"))
        self.assertTrue(hasattr(tab, "visio_open_folder_btn"))

        # Test staging a document in Visio pane
        sample_path = str(Path(__file__).resolve())
        tab.pane_visio.set_file_path(sample_path)
        self.assertEqual(tab.pane_visio.get_inputs(), [sample_path])

    def test_format_converter_engine_and_results(self):
        """Verify Format Converter has engine detection and result notification frame."""
        tab = WordExtractorTab()
        self.assertTrue(hasattr(tab, "pane_convert"))
        self.assertTrue(hasattr(tab, "conv_engine_lbl"))
        self.assertTrue(hasattr(tab, "conv_result_frame"))
        self.assertTrue(hasattr(tab, "conv_open_folder_btn"))

        # Test changing format combo updates engine indicator
        idx_md = tab.format_combo.findText("MD")
        tab.format_combo.setCurrentIndex(idx_md)
        self.assertIn("Markdown", tab.conv_engine_lbl.text())

        idx_pdf = tab.format_combo.findText("PDF")
        tab.format_combo.setCurrentIndex(idx_pdf)
        self.assertIn("Word COM", tab.conv_engine_lbl.text())

    def test_excerpt_selection_helpers(self):
        """Verify heading tree selection helpers and match count label."""
        tab = WordExtractorTab()
        self.assertTrue(hasattr(tab, "lbl_excerpt_match_count"))
        self.assertTrue(hasattr(tab, "excerpt_select_all_btn"))
        self.assertTrue(hasattr(tab, "excerpt_select_l1_btn"))


if __name__ == "__main__":
    unittest.main()
