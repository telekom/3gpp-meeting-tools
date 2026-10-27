import os
import shutil
import tempfile
import unittest
from pathlib import Path

from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for headless test runs
app = QApplication.instance() or QApplication([])

from core.queue_manager import _TASK_REGISTRY
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


if __name__ == "__main__":
    unittest.main()
