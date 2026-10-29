import os
import shutil
import tempfile
import unittest
from pathlib import Path

from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for headless test runs
app = QApplication.instance() or QApplication([])

from core.queue_manager import _TASK_REGISTRY
from modules.word_tools.core.document_opener import ensure_readable_word_document
from modules.word_tools.core.word_converter import WordConverterThread
from modules.word_tools.plugin_loader import register_word_plugin
from modules.word_tools.ui.word_tabs import WordExtractorTab


class TestWordToMarkdown(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        register_word_plugin()
        cls.test_dir = Path(tempfile.mkdtemp(prefix="test_word_to_md_"))

        # Find an existing test docx in the repo
        repo_root = Path(__file__).resolve().parent.parent
        cls.sample_docx = repo_root / "output" / "generated_test_cr.docx"
        if not cls.sample_docx.exists():
            # Search for any docx
            found = list(repo_root.glob("**/*.docx"))
            cls.sample_docx = found[0] if found else None

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.test_dir, ignore_errors=True)

    def test_task_registry_has_word_to_markdown(self):
        """Verify that word_to_markdown task is announced to QueueManager."""
        self.assertIn("word_to_markdown", _TASK_REGISTRY)
        entry = _TASK_REGISTRY["word_to_markdown"]
        self.assertEqual(entry["display_name"], "WORD → MARKDOWN")

    def test_ensure_readable_word_document_accessible(self):
        """Verify that accessible files return their path directly without snapshotting."""
        if not self.sample_docx or not self.sample_docx.exists():
            self.skipTest("No sample .docx found in repository.")

        usable, cleanup = ensure_readable_word_document(self.sample_docx)
        self.assertEqual(usable.resolve(), self.sample_docx.resolve())
        self.assertIsNone(cleanup)

    def test_word_converter_thread_markdown_export(self):
        """Verify WordConverterThread correctly converts a docx to markdown."""
        if not self.sample_docx or not self.sample_docx.exists():
            self.skipTest("No sample .docx found in repository.")

        out_dir = self.test_dir / "thread_test"
        out_dir.mkdir(parents=True, exist_ok=True)

        received_paths = []
        thread = WordConverterThread(
            doc_source=str(self.sample_docx),
            target_format="md",
            revisions="markup",
            extract_media=True,
            output_dir=out_dir,
            include_header=True,
        )
        thread.finished_path.connect(received_paths.append)
        thread.run()

        self.assertEqual(len(received_paths), 1)
        md_file = Path(received_paths[0])
        self.assertTrue(md_file.is_file())
        self.assertEqual(md_file.suffix.lower(), ".md")

        content = md_file.read_text(encoding="utf-8")
        self.assertGreater(len(content), 50)
        self.assertIn("generated_test_cr", content)

    def test_word_converter_thread_accept_revisions_mode(self):
        """Verify WordConverterThread honors accept revisions mode."""
        if not self.sample_docx or not self.sample_docx.exists():
            self.skipTest("No sample .docx found in repository.")

        out_dir = self.test_dir / "accept_test"
        out_dir.mkdir(parents=True, exist_ok=True)

        received_paths = []
        thread = WordConverterThread(
            doc_source=str(self.sample_docx),
            target_format="markdown",
            revisions="accept",
            output_dir=out_dir,
            include_header=False,
        )
        thread.finished_path.connect(received_paths.append)
        thread.run()

        self.assertEqual(len(received_paths), 1)
        md_file = Path(received_paths[0])
        self.assertTrue(md_file.is_file())

        content = md_file.read_text(encoding="utf-8")
        self.assertGreater(len(content), 20)

    def test_word_converter_thread_emits_failure_on_error(self):
        """Verify WordConverterThread emits conversion_failed signal on error."""
        nonexistent = self.test_dir / "does_not_exist.docx"
        errors = []
        thread = WordConverterThread(
            doc_source=str(nonexistent),
            target_format="md",
        )
        thread.conversion_failed.connect(errors.append)
        thread.run()

        self.assertEqual(len(errors), 1)
        self.assertTrue(len(errors[0]) > 0)

    def test_word_extractor_tab_ui_integration(self):
        """Verify WordExtractorTab includes the new Markdown operation and controls."""
        tab = WordExtractorTab()

        # Check op_combo has Convert Word to Markdown
        items = [tab.op_combo.itemText(i) for i in range(tab.op_combo.count())]
        self.assertIn("Convert Word to Markdown (.md)", items)

        # Check format_combo has MD
        formats = [tab.format_combo.itemText(i) for i in range(tab.format_combo.count())]
        self.assertIn("MD", formats)

        # Verify dedicated markdown card widgets exist
        self.assertTrue(hasattr(tab, "card_markdown"))
        self.assertTrue(hasattr(tab, "pane_markdown"))
        self.assertTrue(hasattr(tab, "md_revisions_combo"))
        self.assertTrue(hasattr(tab, "md_media_cb"))
        self.assertTrue(hasattr(tab, "md_header_cb"))
        self.assertTrue(hasattr(tab, "run_markdown_btn"))
        self.assertTrue(hasattr(tab, "md_result_frame"))

        # Switch to markdown card
        md_index = items.index("Convert Word to Markdown (.md)")
        tab.op_combo.setCurrentIndex(md_index)
        self.assertEqual(tab.stack.currentWidget(), tab.card_markdown)

    def test_word_extractor_tab_signals_and_success_flow(self):
        """Verify signal emission and on_conversion_success behavior."""
        tab = WordExtractorTab()

        # Test on_conversion_success updates UI
        test_md = self.test_dir / "sample_result.md"
        test_md.write_text("# Test Result\n\nSample markdown content.", encoding="utf-8")

        tab.on_conversion_success(str(test_md))
        self.assertFalse(tab.md_result_frame.isHidden())
        self.assertIn("sample_result.md", tab.md_result_lbl.text())
        self.assertEqual(tab.md_preview_edit.toPlainText().strip(), "# Test Result\n\nSample markdown content.")
        self.assertEqual(tab._last_converted_md, str(test_md.resolve()))

    def test_word_extractor_tab_error_recovery_never_stuck(self):
        """Verify UI re-enables controls and shows error banner on failure."""
        tab = WordExtractorTab()

        # Simulate conversion start
        tab.run_markdown_btn.setEnabled(False)
        tab.md_progress.setVisible(True)

        # Simulate failure signal arriving
        tab.on_conversion_failed("Permission denied")
        self.assertTrue(tab.run_markdown_btn.isEnabled())
        self.assertFalse(tab.md_progress.isVisible())
        self.assertFalse(tab.md_result_frame.isHidden())
        self.assertIn("Conversion failed", tab.md_result_lbl.text())

    def test_word_extractor_tab_idle_recovery(self):
        """Verify UI recovers button state when QueueManager becomes idle."""
        tab = WordExtractorTab()

        # Simulate conversion start
        tab.run_markdown_btn.setEnabled(False)
        tab.md_progress.setVisible(True)

        # Queue goes idle
        tab.on_processing_state_changed(False)
        self.assertTrue(tab.run_markdown_btn.isEnabled())
        self.assertFalse(tab.md_progress.isVisible())

    def test_word_markdown_card_ignores_non_markdown_extensions(self):
        """Verify WordMarkdownCard ignores non-markdown conversion events (e.g. PDF/VSDX)."""
        tab = WordExtractorTab()
        non_md = self.test_dir / "sample.pdf"
        non_md.write_text("dummy", encoding="utf-8")

        tab.on_conversion_success(str(non_md))
        self.assertTrue(tab.card_markdown.md_result_frame.isHidden())
        self.assertIsNone(tab.card_markdown._last_md_output_path)

    def test_main_window_markdown_handler(self):
        """Verify DragDropUI._on_word_convert_markdown_requested accepts dict and positional options."""
        from unittest.mock import MagicMock
        from main_window import DragDropUI

        mw = MagicMock()
        mw.queue_manager = MagicMock()
        DragDropUI._on_word_convert_markdown_requested(
            mw,
            "test.docx",
            {"revisions_mode": "accept", "extract_media": True, "output_dir": "out", "include_header": False},
        )

        mw.queue_manager.add_item.assert_called_once()
        args, kwargs = mw.queue_manager.add_item.call_args
        self.assertEqual(args[0], Path("test.docx"))
        self.assertEqual(args[1], "word_to_markdown")
        params = kwargs.get("params")
        self.assertEqual(params["revisions"], "accept")
        self.assertTrue(params["extract_media"])
        self.assertEqual(params["output_dir"], "out")
        self.assertFalse(params["include_header"])


if __name__ == "__main__":
    unittest.main()
