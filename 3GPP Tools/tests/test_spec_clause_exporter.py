import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import MagicMock

from PyQt5.QtWidgets import QApplication

from modules.spec_search.core.spec_clause_exporter import (
    SpecClauseWordExtractorThread,
    copy_clause_to_clipboard,
    create_clause_docx_fallback,
    export_clause_to_markdown_file,
    generate_clause_html_email,
    generate_clause_markdown,
    generate_clause_markdown_quote,
    get_default_clause_export_path,
)
from modules.spec_search.ui.spec_search_components import SpecClauseInspector


class TestSpecClauseExporter(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def test_get_default_clause_export_path(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            path_docx = get_default_clause_export_path("23.256", "19.5.0", "5.15", ext=".docx", base_dir=base_dir)
            self.assertEqual(path_docx.name, "TS_23.256_v19.5.0_Clause_5_15.docx")
            self.assertTrue(path_docx.parent.exists())

            path_md = get_default_clause_export_path("TS 23.501", "v17.4.0", "4.2.6.1", ext=".md", base_dir=base_dir)
            self.assertEqual(path_md.name, "TS_23.501_v17.4.0_Clause_4_2_6_1.md")

    def test_generate_clause_markdown(self):
        md = generate_clause_markdown(
            spec_number="23.256",
            version="19.5.0",
            clause_number="5.15",
            clause_title="Disaster Roaming",
            content="When Disaster Roaming is activated, UE initiates registration.",
            release_date="2025-06-20",
        )
        self.assertIn("# 3GPP TS 23.256 v19.5.0 (2025-06-20)", md)
        self.assertIn("## Clause 5.15: Disaster Roaming", md)
        self.assertIn("When Disaster Roaming is activated, UE initiates registration.", md)

    def test_generate_clause_markdown_quote(self):
        quote = generate_clause_markdown_quote(
            spec_number="23.256",
            version="19.5.0",
            clause_number="5.15",
            clause_title="Disaster Roaming",
            content="Paragraph 1.\nParagraph 2.",
            release_date="2025-06-20",
        )
        self.assertIn("> [!NOTE]", quote)
        self.assertIn("> **3GPP TS 23.256 v19.5.0 (2025-06-20) — Clause 5.15: Disaster Roaming**", quote)
        self.assertIn("> Paragraph 1.", quote)
        self.assertIn("> Paragraph 2.", quote)

    def test_generate_clause_html_email(self):
        html_code = generate_clause_html_email(
            spec_number="23.256",
            version="19.5.0",
            clause_number="5.15",
            clause_title="Disaster Roaming",
            content="Line A.\n\nLine B with <brackets>.",
            release_date="2025-06-20",
        )
        self.assertIn("3GPP TS 23.256 v19.5.0 (2025-06-20) — Clause 5.15: Disaster Roaming", html_code)
        self.assertIn("Line A.", html_code)
        self.assertIn("Line B with &lt;brackets&gt;.", html_code)
        self.assertIn("font-family: Calibri", html_code)

    def test_copy_clause_to_clipboard(self):
        # Email mode (rich HTML + plain fallback)
        res = copy_clause_to_clipboard(
            "23.256", "19.5.0", "5.15", "Disaster Roaming", "Test clause text", mode="email"
        )
        self.assertEqual(res, "Copied for Email / Outlook!")
        cb = QApplication.clipboard()
        mime = cb.mimeData()
        self.assertTrue(mime.hasHtml())
        self.assertTrue(mime.hasText())
        self.assertIn("Disaster Roaming", mime.html())
        self.assertIn("Test clause text", mime.text())

        # Markdown quote mode
        res_md = copy_clause_to_clipboard(
            "23.256", "19.5.0", "5.15", "Disaster Roaming", "Test clause text", mode="markdown"
        )
        self.assertEqual(res_md, "Copied as Markdown!")
        self.assertIn("> [!NOTE]", cb.text())

        # Citation mode
        res_cite = copy_clause_to_clipboard(
            "23.256", "19.5.0", "5.15", "Disaster Roaming", "Test clause text", mode="citation"
        )
        self.assertEqual(res_cite, "Copied Citation!")
        self.assertEqual(cb.text(), "3GPP TS 23.256 v19.5.0, Clause 5.15 (Disaster Roaming)")

    def test_export_clause_to_markdown_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target_path = Path(tmpdir) / "test_clause.md"
            out = export_clause_to_markdown_file(
                "23.256",
                "19.5.0",
                "5.15",
                "Disaster Roaming",
                "Clause content goes here.",
                release_date="2025-06-20",
                output_path=target_path,
            )
            self.assertTrue(out.exists())
            text = out.read_text(encoding="utf-8")
            self.assertIn("# 3GPP TS 23.256 v19.5.0 (2025-06-20)", text)
            self.assertIn("Clause content goes here.", text)

    def test_create_clause_docx_fallback(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target_path = Path(tmpdir) / "test_clause.docx"
            out = create_clause_docx_fallback(
                "23.256",
                "19.5.0",
                "5.15",
                "Disaster Roaming",
                "First paragraph.\n\nSecond paragraph.",
                target_path,
                release_date="2025-06-20",
            )
            self.assertTrue(out.exists())
            self.assertGreater(out.stat().st_size, 1000)

    def test_spec_clause_inspector_extract_and_copy_integration(self):
        inspector = SpecClauseInspector()
        inspector.display_clause(
            clause_number="5.15",
            clause_title="Disaster Roaming",
            spec_number="23.256",
            version="19.5.0",
            content="Content for testing inspector buttons.",
            release_date="2025-06-20",
        )

        self.assertTrue(inspector.btn_copy.isEnabled())
        self.assertTrue(inspector.btn_extract.isEnabled())

        # Test extraction signal emission
        extract_signal_received = []
        inspector.extract_clause_requested.connect(
            lambda s, v, c, t, body, save_as: extract_signal_received.append((s, v, c, t, body, save_as))
        )
        inspector._on_extract_clicked(save_as=True)
        self.assertEqual(len(extract_signal_received), 1)
        self.assertEqual(extract_signal_received[0][0], "23.256")
        self.assertEqual(extract_signal_received[0][1], "19.5.0")
        self.assertEqual(extract_signal_received[0][2], "5.15")
        self.assertTrue(extract_signal_received[0][5])  # save_as is True

        # Test markdown export signal emission
        md_signal_received = []
        inspector.export_markdown_requested.connect(
            lambda s, v, c, t, body: md_signal_received.append((s, v, c, t, body))
        )
        inspector._on_export_markdown_clicked()
        self.assertEqual(len(md_signal_received), 1)
        self.assertEqual(md_signal_received[0][0], "23.256")
        self.assertEqual(md_signal_received[0][2], "5.15")

        # Test clear_display
        inspector.clear_display()
        self.assertFalse(inspector.btn_copy.isEnabled())
        self.assertFalse(inspector.btn_extract.isEnabled())

    def test_extractor_thread_markdown_fallback(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = Path(tmpdir) / "test_clause.md"
            worker = SpecClauseWordExtractorThread(
                spec_number="99.999",
                version="1.0.0",
                clause_number="1.1",
                clause_title="Test Title",
                clause_content="Test paragraph content line 1.\n\nTest paragraph line 2.",
                output_path=out_file,
                export_format="markdown",
                release_date="2025-01-01",
            )
            success_paths = []
            worker.finished_success.connect(lambda p: success_paths.append(p))
            worker.run()

            self.assertEqual(len(success_paths), 1)
            self.assertTrue(out_file.exists())
            text = out_file.read_text(encoding="utf-8")
            self.assertIn("# 3GPP TS 99.999 v1.0.0", text)
            self.assertIn("Clause 1.1: Test Title", text)
            self.assertIn("Test paragraph content line 1.", text)


if __name__ == "__main__":
    unittest.main()
