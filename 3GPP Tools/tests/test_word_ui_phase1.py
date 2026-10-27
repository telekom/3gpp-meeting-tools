# --- File: tests/test_word_ui_phase1.py ---
import unittest
from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for headless test runs
app = QApplication.instance() or QApplication([])

from modules.word_tools.ui.word_tabs import WordExtractorTab
from modules.word_tools.ui.word_settings_dialog import WordSettingsDialog
from modules.word_tools.core.word_config import WordConfig


class TestWordUiPhase1(unittest.TestCase):
    def test_word_tab_ribbon_and_nav_pills(self):
        """Verify 2-row ribbon, active word pill, and segmented navigation."""
        tab = WordExtractorTab()

        # Ribbon frame and Row 1 status pills exist
        self.assertTrue(hasattr(tab, "ribbon_frame"))
        self.assertTrue(hasattr(tab, "btn_active_word"))
        self.assertTrue(hasattr(tab, "dotx_badge"))
        self.assertTrue(hasattr(tab, "lo_badge"))
        self.assertTrue(hasattr(tab, "settings_btn"))

        # Verify navigation buttons exist and have correct labels (no unescaped &, no 'Redline')
        self.assertTrue(hasattr(tab, "nav_btn_compare"))
        self.assertEqual(tab.nav_btn_compare.text(), "⚖️ Compare Documents")
        self.assertTrue(hasattr(tab, "nav_btn_markdown"))
        self.assertEqual(tab.nav_btn_markdown.text(), "📝 Word → Markdown")
        self.assertTrue(hasattr(tab, "nav_btn_cr_pcr"))
        self.assertEqual(tab.nav_btn_cr_pcr.text(), "🏛️ 3GPP CR && Templates")
        self.assertTrue(hasattr(tab, "nav_btn_author"))
        self.assertEqual(tab.nav_btn_author.text(), "👥 Revision Authors")
        self.assertTrue(hasattr(tab, "nav_btn_slice"))
        self.assertEqual(tab.nav_btn_slice.text(), "✂️ Slice && Excerpt")
        self.assertTrue(hasattr(tab, "nav_btn_convert"))
        self.assertEqual(tab.nav_btn_convert.text(), "🔄 Format Converter")
        self.assertTrue(hasattr(tab, "nav_btn_visio"))
        self.assertEqual(tab.nav_btn_visio.text(), "🖼️ Visio && Media Extractor")

        # Verify cards exist in stack
        self.assertTrue(hasattr(tab, "card_cr_pcr"))
        self.assertTrue(hasattr(tab, "card_author_renamer"))
        self.assertEqual(tab.stack.count(), 9)

        # Test switching to Compare Documents
        tab.set_active_operation(4)
        self.assertEqual(tab.stack.currentIndex(), 4)
        self.assertTrue(tab.nav_btn_compare.isChecked())
        self.assertFalse(tab.nav_btn_markdown.isChecked())

        # Test switching to Markdown
        tab.set_active_operation(6)
        self.assertEqual(tab.stack.currentIndex(), 6)
        self.assertTrue(tab.nav_btn_markdown.isChecked())
        self.assertFalse(tab.nav_btn_compare.isChecked())

        # Test switching to 3GPP CR
        tab.set_active_operation(7)
        self.assertEqual(tab.stack.currentIndex(), 7)
        self.assertTrue(tab.nav_btn_cr_pcr.isChecked())

        # Test switching to Author Renamer
        tab.set_active_operation(8)
        self.assertEqual(tab.stack.currentIndex(), 8)
        self.assertTrue(tab.nav_btn_author.isChecked())

        # Test clicking nav pill button
        tab.nav_btn_convert.click()
        self.assertEqual(tab.stack.currentIndex(), 5)
        self.assertTrue(tab.nav_btn_convert.isChecked())

    def test_word_settings_dialog_loads_config(self):
        """Verify WordSettingsDialog loads configuration correctly."""
        dlg = WordSettingsDialog()
        self.assertTrue(hasattr(dlg, "chk_enable_sensitivity"))
        self.assertTrue(hasattr(dlg, "edit_label_id"))
        self.assertTrue(hasattr(dlg, "edit_label_name"))
        self.assertTrue(hasattr(dlg, "edit_company"))
        self.assertTrue(hasattr(dlg, "edit_lo_path"))

        cfg = WordConfig.load()
        self.assertEqual(dlg.edit_label_name.text(), cfg.get("sensitivity_level_label_name", "OFFEN"))
        dlg.close()


if __name__ == "__main__":
    unittest.main()
