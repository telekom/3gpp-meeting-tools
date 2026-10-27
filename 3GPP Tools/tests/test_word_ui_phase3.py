# --- File: tests/test_word_ui_phase3.py ---
import unittest
from pathlib import Path
from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for headless test runs
app = QApplication.instance() or QApplication([])

from modules.word_tools.ui.word_tabs import WordExtractorTab


class TestWordUiPhase3(unittest.TestCase):
    def test_cr_pcr_workbench_card(self):
        """Verify 3GPP Contribution & Templates Workbench controls and integration."""
        tab = WordExtractorTab()

        self.assertTrue(hasattr(tab, "card_cr_pcr"))
        self.assertTrue(hasattr(tab, "pane_cr"))
        self.assertTrue(hasattr(tab, "cr_mode_combo"))
        self.assertTrue(hasattr(tab, "btn_launch_cr"))
        self.assertTrue(hasattr(tab, "cr_card_dotx_lbl"))
        self.assertTrue(hasattr(tab, "btn_cr_install_dotx"))
        self.assertTrue(hasattr(tab, "btn_cr_open_tpl"))

        # Verify contribution modes
        modes = [tab.cr_mode_combo.itemData(i) for i in range(tab.cr_mode_combo.count())]
        self.assertIn("cr", modes)
        self.assertIn("pcr", modes)

        # Stage a baseline specification
        sample_path = str(Path(__file__).resolve())
        tab.pane_cr.set_file_path(sample_path)
        self.assertEqual(tab.pane_cr.get_inputs(), [sample_path])

        # Test switching to the CR/pCR workbench
        tab.set_active_operation(7)
        self.assertEqual(tab.stack.currentWidget(), tab.card_cr_pcr)
        self.assertTrue(tab.nav_btn_cr_pcr.isChecked())


if __name__ == "__main__":
    unittest.main()
