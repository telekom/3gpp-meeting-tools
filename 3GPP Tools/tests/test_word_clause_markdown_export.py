"""Tests for Word clause excerpt export to Markdown and figures."""

import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for UI tests
app = QApplication.instance() or QApplication([])

from modules.word_tools.core.word_excerpt_extractor import (
    HeadingInfo,
    WordExcerptThread,
)
from modules.word_tools.ui.cards.excerpt_card import WordExcerptCard
from modules.word_tools.ui.word_tabs import WordExtractorTab
from modules.specifications.ui.spec_clause_dialog import SpecClauseDialog


class TestWordClauseMarkdownExport(unittest.TestCase):
    """Unit and integration tests for Word clause Markdown export."""

    def test_word_excerpt_thread_init_defaults(self):
        """Verify default parameters on WordExcerptThread."""
        thread = WordExcerptThread(
            source_paths=["dummy.docx"],
            headings=[],
            selected_ids=["0:0"],
            output_path="out.docx",
        )
        self.assertFalse(thread.export_markdown)
        self.assertEqual(thread.revisions, "accept")
        self.assertIsNone(thread.md_path)
        self.assertEqual(thread.images_count, 0)

    def test_word_excerpt_thread_custom_params(self):
        """Verify custom parameters on WordExcerptThread."""
        thread = WordExcerptThread(
            source_paths=["dummy.docx"],
            headings=[],
            selected_ids=["0:0"],
            output_path="out.docx",
            export_markdown=True,
            revisions="markup",
        )
        self.assertTrue(thread.export_markdown)
        self.assertEqual(thread.revisions, "markup")

    @patch("modules.word_tools.core.word_excerpt_extractor.WordExcerptEngine")
    @patch("modules.meetings.core.tdoc_markdown_converter.convert_document_to_markdown")
    def test_word_excerpt_thread_run_with_markdown(self, mock_convert_md, mock_engine_cls):
        """Verify WordExcerptThread invokes convert_document_to_markdown when export_markdown=True."""
        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = (Path(tmpdir) / "test_excerpt.docx").resolve()
            out_docx.touch()
            expected_md = (Path(tmpdir) / "test_excerpt.md").resolve()
            expected_md.touch()

            # Mock WordExcerptEngine.extract
            mock_engine_instance = MagicMock()
            mock_engine_instance.extract.return_value = out_docx
            mock_engine_cls.return_value = mock_engine_instance

            # Mock convert_document_to_markdown
            mock_export_res = MagicMock()
            mock_export_res.md_path = expected_md
            mock_export_res.images = 3
            mock_convert_md.return_value = mock_export_res

            thread = WordExcerptThread(
                source_paths=["dummy.docx"],
                headings=[],
                selected_ids=["0:0"],
                output_path=str(out_docx),
                export_markdown=True,
            )

            finished_paths = []
            progress_events = []
            thread.finished_path.connect(finished_paths.append)
            thread.progress.connect(lambda p, m: progress_events.append((p, m)))

            thread.run()

            # Verify WordExcerptEngine was called
            mock_engine_instance.extract.assert_called_once_with(["0:0"], str(out_docx))

            # Verify convert_document_to_markdown was called with correct parameters
            mock_convert_md.assert_called_once_with(
                out_docx.resolve(),
                out_dir=out_docx.parent,
                tdoc_id="test_excerpt",
                revisions="accept",
                media_dir=out_docx.parent / "test_excerpt_media",
                media_rel="test_excerpt_media",
            )

            # Verify thread state
            self.assertEqual(thread.md_path, expected_md)
            self.assertEqual(thread.images_count, 3)
            self.assertEqual(finished_paths, [str(out_docx)])
            self.assertTrue(any("Converting excerpt to Markdown" in msg for _, msg in progress_events))

    def test_word_excerpt_card_has_markdown_checkbox(self):
        """Verify WordExcerptCard UI contains the Markdown export checkbox."""
        card = WordExcerptCard()
        self.assertTrue(hasattr(card, "chk_also_markdown"))
        self.assertFalse(card.chk_also_markdown.isChecked())
        self.assertIn("Markdown", card.chk_also_markdown.text())
        self.assertIn("figures", card.chk_also_markdown.text().lower())

    def test_word_tabs_has_markdown_checkbox_alias(self):
        """Verify WordExtractorTab provides the chk_also_markdown alias for excerpt card."""
        tab = WordExtractorTab()
        self.assertTrue(hasattr(tab, "chk_also_markdown"))
        self.assertIs(tab.chk_also_markdown, tab.card_excerpt.chk_also_markdown)

    def test_spec_clause_dialog_has_markdown_checkbox(self):
        """Verify SpecClauseDialog UI contains the Markdown export checkbox and 2-tier footer layout."""
        from PyQt5.QtWidgets import QHeaderView
        with tempfile.TemporaryDirectory() as tmpdir:
            dlg = SpecClauseDialog(
                spec_num="23.501",
                spec_type="TS",
                version="17.4.0",
                fname="23501-h40.zip",
                url="https://www.3gpp.org/ftp/Specs/archive/23_series/23.501/23501-h40.zip",
                spec_dl_dir=Path(tmpdir),
            )
            self.assertTrue(hasattr(dlg, "chk_also_markdown"))
            self.assertFalse(dlg.chk_also_markdown.isChecked())
            self.assertIn("Markdown", dlg.chk_also_markdown.text())
            self.assertTrue(hasattr(dlg, "lbl_selection_count"))
            self.assertTrue(hasattr(dlg, "lbl_filter_matches"))
            self.assertTrue(hasattr(dlg, "open_after_cb"))
            self.assertTrue(hasattr(dlg, "btn_clear"))
            self.assertGreaterEqual(dlg.minimumWidth(), 740)
            header = dlg.clause_tree.header()
            self.assertEqual(header.sectionResizeMode(0), QHeaderView.Stretch)
            dlg.close()


if __name__ == "__main__":
    unittest.main()
