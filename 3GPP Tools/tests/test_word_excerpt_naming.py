"""Unit and regression tests for Word and specification clause excerpt differentiated naming."""

import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.word_excerpt_extractor import HeadingInfo
from modules.word_tools.core.word_excerpt_naming import (
    NAMING_SCHEMES,
    SCHEME_CLAUSES_ONLY,
    SCHEME_CLAUSES_TIMESTAMP,
    SCHEME_GENERIC,
    SCHEME_SOURCE_CLAUSES,
    SCHEME_SOURCE_CLAUSES_TIMESTAMP,
    SCHEME_SOURCE_CLAUSES_TITLE,
    SCHEME_SOURCE_TIMESTAMP,
    SCHEME_TIMESTAMP_ONLY,
    build_clauses_token,
    build_source_token,
    extract_clause_identifier,
    generate_differentiated_export_name,
    normalize_selected_headings,
    sanitize_filename_token,
)


class TestWordExcerptNaming(unittest.TestCase):
    """Test suite for clause extraction differentiated naming functions."""

    def test_sanitize_filename_token(self):
        # Invalid characters replaced
        self.assertEqual(sanitize_filename_token('S2-2401234/CR:23501*draft?"<>|'), "S2-2401234_CR_23501_draft")
        # Collapses whitespace and underscores
        self.assertEqual(sanitize_filename_token("  Key   Derivation __ Function  "), "Key_Derivation_Function")
        # Strips dots and underscores from edges
        self.assertEqual(sanitize_filename_token("._test_file_."), "test_file")
        # Length bound
        long_name = "A" * 60
        self.assertEqual(len(sanitize_filename_token(long_name, max_len=30)), 30)

    def test_extract_clause_identifier(self):
        # Standard 3GPP clause with title
        cid, title = extract_clause_identifier("5.2 Key derivation for K_NF")
        self.assertEqual(cid, "5.2")
        self.assertEqual(title, "Key_derivation_for")

        # Clause prefix
        cid, title = extract_clause_identifier("Clause 6.3.1 Security Handling Procedures")
        self.assertEqual(cid, "6.3.1")
        self.assertEqual(title, "Security_Handling_Procedures")

        # Section prefix
        cid, title = extract_clause_identifier("Section 4 Architecture Reference Model")
        self.assertEqual(cid, "4")
        self.assertEqual(title, "Architecture_Reference_Model")

        # Annex with informative note
        cid, title = extract_clause_identifier("Annex A (informative): Change history")
        self.assertEqual(cid, "Annex_A")
        self.assertEqual(title, "Change_history")

        # Annex with subclause
        cid, title = extract_clause_identifier("Annex B.1 Verification procedures")
        self.assertEqual(cid, "Annex_B.1")
        self.assertEqual(title, "Verification_procedures")

        # Non-numbered heading
        cid, title = extract_clause_identifier("Introduction to Architecture")
        self.assertEqual(cid, "Introduction_to")
        self.assertEqual(title, "Architecture")

    def test_normalize_selected_headings_removes_children(self):
        # Heading 5.2 covers block 10 to 50
        h_parent = HeadingInfo(
            id="0:10", source_path="C:/docs/S2.docx", source_index=0,
            block_index=10, end_index=50, level=2, text="5.2 Key derivation",
        )
        # Child 5.2.1 is block 15 to 30
        h_child1 = HeadingInfo(
            id="0:15", source_path="C:/docs/S2.docx", source_index=0,
            block_index=15, end_index=30, level=3, text="5.2.1 General",
        )
        # Sibling 5.3 is block 50 to 90
        h_sibling = HeadingInfo(
            id="0:50", source_path="C:/docs/S2.docx", source_index=0,
            block_index=50, end_index=90, level=2, text="5.3 Security Handling",
        )

        all_headings = [h_parent, h_child1, h_sibling]

        # If both parent and child are selected, only parent is kept in normalized list
        normalized = normalize_selected_headings(all_headings, ["0:10", "0:15"])
        self.assertEqual(len(normalized), 1)
        self.assertEqual(normalized[0].id, "0:10")

        # If parent and sibling are selected, both are kept
        normalized2 = normalize_selected_headings(all_headings, ["0:10", "0:50"])
        self.assertEqual(len(normalized2), 2)
        self.assertEqual([h.id for h in normalized2], ["0:10", "0:50"])

    def test_build_clauses_token(self):
        h1 = HeadingInfo("0:10", "a.docx", 0, 10, 20, 2, "5.2 Key derivation")
        h2 = HeadingInfo("0:20", "a.docx", 0, 20, 30, 2, "5.3 Security handling")
        h3 = HeadingInfo("0:30", "a.docx", 0, 30, 40, 2, "6.1 Procedures")
        h4 = HeadingInfo("0:40", "a.docx", 0, 40, 50, 2, "6.2 Policy")

        # Single clause
        self.assertEqual(build_clauses_token([h1]), "Cl_5.2")
        # Single clause with title
        self.assertEqual(build_clauses_token([h1], include_title=True), "Cl_5.2_Key_derivation")
        # Two clauses
        self.assertEqual(build_clauses_token([h1, h2]), "Cl_5.2_5.3")
        # Three clauses
        self.assertEqual(build_clauses_token([h1, h2, h3]), "Cl_5.2_5.3_6.1")
        # Contiguous range
        h_c1 = HeadingInfo("0:1", "a.docx", 0, 1, 2, 2, "5.1 Scope")
        h_c2 = HeadingInfo("0:2", "a.docx", 0, 2, 3, 2, "5.2 General")
        h_c3 = HeadingInfo("0:3", "a.docx", 0, 3, 4, 2, "5.3 Details")
        h_c4 = HeadingInfo("0:4", "a.docx", 0, 4, 5, 2, "5.4 Summary")
        self.assertEqual(build_clauses_token([h_c1, h_c2, h_c3, h_c4]), "Cl_5.1-5.4")
        # More than 3 disparate clauses
        self.assertEqual(build_clauses_token([h1, h2, h3, h4]), "Cl_5.2_5.3_plus2")

    def test_build_source_token(self):
        self.assertEqual(build_source_token([]), "doc")
        self.assertEqual(build_source_token(["C:/path/S2-2401234.docx"]), "S2-2401234")
        self.assertEqual(
            build_source_token(["C:/path/S2-2401234.docx", "C:/path/S2-2401235.docx"]),
            "S2-2401234_S2-2401235",
        )
        self.assertEqual(
            build_source_token(["C:/path/A.docx", "C:/path/B.docx", "C:/path/C.docx"]),
            "A_plus2",
        )

    def test_generate_differentiated_export_name_all_schemes(self):
        h1 = HeadingInfo("0:10", "S2-2401234_CR_23501.docx", 0, 10, 20, 2, "5.2 Key derivation for K_NF")
        h2 = HeadingInfo("0:20", "S2-2401234_CR_23501.docx", 0, 20, 30, 2, "5.3 Security Handling")
        headings = [h1, h2]
        sources = ["C:/docs/S2-2401234_CR_23501.docx"]
        ts = "20261102_1425"

        # 1. Source + Clauses + Timestamp (Recommended)
        name1 = generate_differentiated_export_name(
            sources, headings, ["0:10"], scheme=SCHEME_SOURCE_CLAUSES_TIMESTAMP, timestamp=ts,
        )
        self.assertEqual(name1, "S2-2401234_CR_23501_Cl_5.2_20261102_1425.docx")

        # 2. Source + Clauses
        name2 = generate_differentiated_export_name(
            sources, headings, ["0:10", "0:20"], scheme=SCHEME_SOURCE_CLAUSES, timestamp=ts,
        )
        self.assertEqual(name2, "S2-2401234_CR_23501_Cl_5.2_5.3.docx")

        # 3. Source + Clauses & Title + Timestamp
        name3 = generate_differentiated_export_name(
            sources, headings, ["0:10"], scheme=SCHEME_SOURCE_CLAUSES_TITLE, timestamp=ts,
        )
        self.assertEqual(name3, "S2-2401234_CR_23501_Cl_5.2_Key_derivation_for_20261102_1425.docx")

        # 4. Source + Timestamp
        name4 = generate_differentiated_export_name(
            sources, headings, ["0:10"], scheme=SCHEME_SOURCE_TIMESTAMP, timestamp=ts,
        )
        self.assertEqual(name4, "S2-2401234_CR_23501_excerpt_20261102_1425.docx")

        # 5. Clauses + Timestamp
        name5 = generate_differentiated_export_name(
            sources, headings, ["0:10"], scheme=SCHEME_CLAUSES_TIMESTAMP, timestamp=ts,
        )
        self.assertEqual(name5, "Cl_5.2_20261102_1425.docx")

        # 6. Clauses Only
        name6 = generate_differentiated_export_name(
            sources, headings, ["0:10"], scheme=SCHEME_CLAUSES_ONLY, timestamp=ts,
        )
        self.assertEqual(name6, "Cl_5.2.docx")

        # 7. Timestamp Only
        name7 = generate_differentiated_export_name(
            sources, headings, ["0:10"], scheme=SCHEME_TIMESTAMP_ONLY, timestamp=ts,
        )
        self.assertEqual(name7, "excerpt_20261102_1425.docx")

        # 8. Generic
        name8 = generate_differentiated_export_name(
            sources, headings, ["0:10"], scheme=SCHEME_GENERIC, timestamp=ts,
        )
        self.assertEqual(name8, "extracted_sections.docx")

        # Custom prefix (e.g. for Specification Clause Dialog)
        name_spec = generate_differentiated_export_name(
            sources, headings, ["0:10"], scheme=SCHEME_SOURCE_CLAUSES_TIMESTAMP,
            custom_prefix="TS_23_501_v18_4_0", timestamp=ts,
        )
        self.assertEqual(name_spec, "TS_23_501_v18_4_0_Cl_5.2_20261102_1425.docx")

        # Markdown extension
        name_md = generate_differentiated_export_name(
            sources, headings, ["0:10"], scheme=SCHEME_SOURCE_CLAUSES, ext=".md",
        )
        self.assertTrue(name_md.endswith(".md"))

    def test_word_config_naming_scheme_persistence(self):
        current = WordConfig.get_excerpt_naming_scheme()
        self.assertTrue(isinstance(current, str))
        WordConfig.set_excerpt_naming_scheme("clauses_timestamp")
        self.assertEqual(WordConfig.get_excerpt_naming_scheme(), "clauses_timestamp")
        # Restore original
        WordConfig.set_excerpt_naming_scheme(current)


class TestWordExcerptCardUI(unittest.TestCase):
    """Test suite for WordExcerptCard UI differentiated export naming integration."""

    @classmethod
    def setUpClass(cls):
        from PyQt5.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        from modules.word_tools.ui.cards.excerpt_card import WordExcerptCard
        self.card = WordExcerptCard()

    def tearDown(self):
        self.card.deleteLater()

    def test_excerpt_card_naming_ui_components_exist(self):
        self.assertTrue(hasattr(self.card, "excerpt_naming_combo"))
        self.assertTrue(hasattr(self.card, "excerpt_filename_edit"))
        self.assertTrue(hasattr(self.card, "excerpt_reset_name_btn"))
        self.assertGreaterEqual(self.card.excerpt_naming_combo.count(), 7)

    def test_excerpt_card_preview_filename_updates_with_selection(self):
        h1 = HeadingInfo("0:10", "C:/docs/S2-2401234_CR_23501.docx", 0, 10, 20, 2, "5.2 Key derivation for K_NF")
        h2 = HeadingInfo("0:20", "C:/docs/S2-2401234_CR_23501.docx", 0, 20, 30, 2, "5.3 Security Handling")

        self.card._excerpt_paths = ["C:/docs/S2-2401234_CR_23501.docx"]
        self.card._excerpt_headings = [h1, h2]
        self.card._populate_excerpt_tree()

        # Initially no clauses selected
        initial_text = self.card.excerpt_filename_edit.text()
        self.assertIn("S2-2401234_CR_23501", initial_text)

        # Select all headings
        self.card._select_all_excerpt_headings()
        updated_text = self.card.excerpt_filename_edit.text()
        self.assertIn("Cl_5.2_5.3", updated_text)

        # Change scheme to Source + Clauses
        idx = self.card.excerpt_naming_combo.findData("source_clauses")
        self.card.excerpt_naming_combo.setCurrentIndex(idx)
        scheme_text = self.card.excerpt_filename_edit.text()
        self.assertEqual(scheme_text, "S2-2401234_CR_23501_Cl_5.2_5.3.docx")

        # Manually edit filename
        self.card.excerpt_filename_edit.setText("My_Custom_Export.docx")
        self.card._on_filename_text_edited("My_Custom_Export.docx")
        self.assertTrue(self.card._custom_filename_active)

        # Clear selection - text must remain custom
        self.card._clear_excerpt_selection()
        self.assertEqual(self.card.excerpt_filename_edit.text(), "My_Custom_Export.docx")

        # Reset button restores auto-generated name
        self.card.excerpt_reset_name_btn.click()
        self.assertFalse(self.card._custom_filename_active)
        self.assertIn("S2-2401234_CR_23501", self.card.excerpt_filename_edit.text())

    def test_word_settings_dialog_has_excerpt_naming_combo(self):
        from modules.word_tools.ui.word_settings_dialog import WordSettingsDialog
        dlg = WordSettingsDialog()
        self.assertTrue(hasattr(dlg, "combo_excerpt_naming"))
        self.assertGreaterEqual(dlg.combo_excerpt_naming.count(), 7)
        dlg.deleteLater()

    @patch("modules.word_tools.ui.cards.excerpt_card.WordExcerptThread")
    @patch("PyQt5.QtWidgets.QFileDialog.getSaveFileName")
    def test_extract_selected_sections_uses_differentiated_name(self, mock_get_save, mock_thread_cls):
        h1 = HeadingInfo("0:10", "C:/docs/S2-2401234.docx", 0, 10, 20, 2, "5.2 Key derivation")
        self.card._excerpt_paths = ["C:/docs/S2-2401234.docx"]
        self.card._excerpt_headings = [h1]
        self.card._populate_excerpt_tree()
        self.card._select_all_excerpt_headings()

        proposed_filename = self.card.excerpt_filename_edit.text()
        self.assertIn("Cl_5.2", proposed_filename)

        mock_get_save.return_value = (f"C:/docs/{proposed_filename}", "")
        mock_thread_instance = MagicMock()
        mock_thread_cls.return_value = mock_thread_instance

        self.card._extract_selected_sections()

        # Verify getSaveFileName was called with the differentiated filename
        self.assertTrue(mock_get_save.called)
        called_args = mock_get_save.call_args[0]
        self.assertIn(proposed_filename, called_args[2])


if __name__ == "__main__":
    unittest.main()
