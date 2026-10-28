# --- File: tests/test_nas_exporter.py ---
import sys
import unittest
from pathlib import Path
from PyQt5.QtWidgets import QApplication

from modules.nas.core.nas_exporter import (
    build_matrix_html,
    build_matrix_markdown,
    build_matrix_tsv,
    copy_matrix_to_clipboard,
    build_definition_word_html,
    build_definition_markdown,
    build_definition_plain_text,
    copy_definition_to_clipboard,
    html_table_to_markdown,
)
from modules.nas.ui.nas_components import NASInspectorWidget
from modules.nas.ui.nas_tabs import NASTab

app = QApplication.instance() or QApplication(sys.argv)


class TestNASExporter(unittest.TestCase):
    """Verifies Word-compatible and Markdown clipboard export pipelines for Protocols tab."""

    def setUp(self):
        self.matrix_headers = ["IEI / ID", "Field / Information Element", "Type / Reference", "Rel-15\nv15.30.0", "Rel-16\nv16.22.0"]
        self.matrix_rows = [
            ["-", "BCCH-BCH-Message", "Clause 6.2.1", "SEQUENCE", "SEQUENCE"],
            ["-", "    └─ message", "BCCH-BCH-MessageType", "CHOICE", "CHOICE"],
            ["23", "Extended Protocol Discriminator", "Clause 9.2", "M | V | 1", "M | V | 1"],
            ["-", "New Feature IE", "Clause 9.9", "-", "O | TLV | 4"],
        ]
        self.diff_colors = [
            [None, None, None, None, None],
            [None, None, None, None, None],
            [None, None, None, None, None],
            [None, None, None, None, "#E8F5E9"],  # Soft green for added
        ]

    def test_build_matrix_html(self):
        """Verifies Word-compatible HTML table generation with inline styles and diff highlights."""
        html_out = build_matrix_html(
            title="TS 38.331 BCCH-BCH-Message Evolution",
            headers=self.matrix_headers,
            rows=self.matrix_rows,
            diff_colors=self.diff_colors,
        )

        self.assertIn("TS 38.331 BCCH-BCH-Message Evolution", html_out)
        self.assertIn("<table", html_out)
        self.assertIn("border-collapse: collapse;", html_out)
        self.assertIn("Calibri", html_out)
        self.assertIn("<th", html_out)
        self.assertIn("Rel-15 v15.30.0", html_out)
        # Verify diff color on the added cell
        self.assertIn("background-color: #E8F5E9;", html_out)
        # Verify indentation
        self.assertIn("&nbsp;&nbsp;&nbsp;&nbsp;└─ message", html_out)

    def test_build_matrix_markdown(self):
        """Verifies GitHub-Flavored Markdown table generation with pipe escaping."""
        md_out = build_matrix_markdown(
            title="Evolution: BCCH-BCH-Message",
            headers=self.matrix_headers,
            rows=self.matrix_rows,
        )

        self.assertTrue(md_out.startswith("### Evolution: BCCH-BCH-Message"))
        self.assertIn("| IEI / ID | Field / Information Element | Type / Reference | Rel-15 v15.30.0 | Rel-16 v16.22.0 |", md_out)
        self.assertIn("| --- | --- | --- | --- | --- |", md_out)
        self.assertIn("| 23 | Extended Protocol Discriminator | Clause 9.2 | M \\| V \\| 1 | M \\| V \\| 1 |", md_out)
        self.assertIn("&nbsp;&nbsp;&nbsp;&nbsp;└─ message", md_out)

    def test_build_matrix_tsv(self):
        """Verifies tab-separated output for spreadsheet paste."""
        tsv_out = build_matrix_tsv(self.matrix_headers, self.matrix_rows)
        lines = tsv_out.split("\n")
        self.assertEqual(len(lines), 5)
        self.assertIn("IEI / ID\tField / Information Element\tType / Reference\tRel-15 v15.30.0\tRel-16 v16.22.0", lines[0])
        self.assertIn("23\tExtended Protocol Discriminator\tClause 9.2\tM | V | 1\tM | V | 1", lines[3])

    def test_copy_matrix_to_clipboard(self):
        """Verifies clipboard population for Word (dual-format), Markdown, and TSV."""
        # 1. Word mode
        status_word = copy_matrix_to_clipboard("Test Matrix", self.matrix_headers, self.matrix_rows, self.diff_colors, mode="word")
        self.assertEqual(status_word, "Copied for Word / Outlook!")
        mime = QApplication.clipboard().mimeData()
        self.assertTrue(mime.hasHtml())
        self.assertTrue(mime.hasText())
        self.assertIn("<table", mime.html())
        self.assertIn("\t", mime.text())

        # 2. Markdown mode
        status_md = copy_matrix_to_clipboard("Test Matrix", self.matrix_headers, self.matrix_rows, mode="markdown")
        self.assertEqual(status_md, "Copied as Markdown Table!")
        self.assertIn("| --- |", QApplication.clipboard().text())

        # 3. TSV mode
        status_tsv = copy_matrix_to_clipboard("Test Matrix", self.matrix_headers, self.matrix_rows, mode="tsv")
        self.assertEqual(status_tsv, "Copied TSV (Excel)!")
        self.assertIn("\t", QApplication.clipboard().text())

    def test_build_definition_word_html(self):
        """Verifies Word-optimized HTML container with 3GPP citation header and cleaned links."""
        sample_html = (
            '<h2 style="color: #0369A1;">BCCH-BCH-Message (Clause 6.2.1)</h2>'
            '<div style="font-weight: bold;">ASN.1 Definition:</div>'
            '<pre>BCCH-BCH-Message ::= SEQUENCE { message BCCH-BCH-MessageType }</pre>'
            '<p>The BCCH-BCH-Message class is the set of RRC messages.</p>'
            '<a href="action:open_spec:38.331:19.3.0:6.2.1">Open Spec in Word</a>'
        )

        word_html = build_definition_word_html(
            html_content=sample_html,
            spec_number="38.331",
            version="19.3.0",
            ie_name="BCCH-BCH-Message",
            clause="6.2.1",
        )

        self.assertIn("3GPP TS 38.331 v19.3.0 — BCCH-BCH-Message (Clause 6.2.1)", word_html)
        self.assertIn("font-family: Calibri", word_html)
        self.assertNotIn('href="action:open_spec', word_html)
        self.assertIn("The BCCH-BCH-Message class is the set of RRC messages.", word_html)

    def test_build_definition_markdown(self):
        """Verifies HTML to GitHub-Flavored Markdown conversion."""
        sample_html = (
            '<h2 style="color: #0369A1;">BCCH-BCH-Message (Clause 6.2.1)</h2>'
            '<div style="font-weight: bold;">ASN.1 Definition:</div>'
            '<pre style="background: #FFF;">\nBCCH-BCH-Message ::= SEQUENCE {\n    message BCCH-BCH-MessageType\n}\n</pre>'
            '<div style="font-weight: bold;">Specification Details &amp; Tabular Definition:</div>'
            '<p>The BCCH-BCH-Message class is the set of RRC messages.</p>'
            '<div style="border-left: 4px solid #0284C7; padding: 6px;">'
            'Cross-Referenced from TS 24.301 v15.10.0 (Clause 9.9.3.x)'
            '</div>'
        )

        md = build_definition_markdown(
            html_content=sample_html,
            spec_number="38.331",
            version="19.3.0",
            ie_name="BCCH-BCH-Message",
            clause="6.2.1",
        )

        self.assertIn("### 3GPP TS 38.331 v19.3.0 — BCCH-BCH-Message (Clause 6.2.1)", md)
        self.assertIn("#### ASN.1 Definition", md)
        self.assertIn("```asn1\nBCCH-BCH-Message ::= SEQUENCE {\n    message BCCH-BCH-MessageType\n}\n```", md)
        self.assertIn("The BCCH-BCH-Message class is the set of RRC messages.", md)
        self.assertIn("> [!NOTE]", md)
        self.assertIn("Cross-Referenced from TS 24.301", md)

    def test_html_table_to_markdown_with_colspan(self):
        """Verifies table conversion with colspan expansion for bit-level layouts."""
        from bs4 import BeautifulSoup
        table_html = (
            '<table>'
            '<tr><th>Octet</th><th colspan="2">Field Description</th></tr>'
            '<tr><td>1</td><td>Type</td><td>Length</td></tr>'
            '</table>'
        )
        soup = BeautifulSoup(table_html, "html.parser")
        tbl_md = html_table_to_markdown(soup.find("table"))
        lines = tbl_md.split("\n")
        self.assertEqual(len(lines), 3)
        self.assertIn("| Octet | Field Description |  |", lines[0])
        self.assertIn("| --- | --- | --- |", lines[1])
        self.assertIn("| 1 | Type | Length |", lines[2])

    def test_copy_definition_to_clipboard(self):
        """Verifies clipboard population for IE definitions across all formats."""
        sample_html = "<pre>TAGS ::= BEGIN IMPORTS END</pre><p>Description</p>"

        # Word mode
        status_word = copy_definition_to_clipboard(sample_html, "38.331", "19.3.0", "TAGS", "TAGS", mode="word")
        self.assertEqual(status_word, "Copied for Word / Outlook!")
        mime = QApplication.clipboard().mimeData()
        self.assertTrue(mime.hasHtml())
        self.assertIn("3GPP TS 38.331 v19.3.0 — TAGS", mime.html())
        self.assertTrue(mime.hasText())

        # Markdown mode
        status_md = copy_definition_to_clipboard(sample_html, "38.331", "19.3.0", "TAGS", "TAGS", mode="markdown")
        self.assertEqual(status_md, "Copied as Markdown!")
        self.assertIn("```asn1", QApplication.clipboard().text())

        # Text mode
        status_txt = copy_definition_to_clipboard(sample_html, "38.331", "19.3.0", "TAGS", "TAGS", mode="text")
        self.assertEqual(status_txt, "Copied Plain Text!")
        self.assertIn("3GPP TS 38.331 v19.3.0 — TAGS", QApplication.clipboard().text())

        # Citation mode
        status_cite = copy_definition_to_clipboard(sample_html, "38.331", "19.3.0", "TAGS", "TAGS", mode="citation")
        self.assertEqual(status_cite, "Copied Citation!")
        self.assertEqual(QApplication.clipboard().text(), "3GPP TS 38.331 v19.3.0 — TAGS")

    def test_ui_components_menu_wiring(self):
        """Verifies that UI buttons in NASTab and NASInspectorWidget are equipped with formatted export menus."""
        # Inspector Widget
        inspector = NASInspectorWidget()
        self.assertEqual(inspector.copy_btn.text(), "📋 Copy Definition ▾")
        self.assertIsNotNone(inspector.copy_btn.menu())
        actions = [a.text() for a in inspector.copy_btn.menu().actions() if not a.isSeparator()]
        self.assertEqual(
            actions,
            [
                "📋 Copy for Word / Outlook (Rich Text)",
                "📝 Copy as Markdown",
                "📄 Copy Plain Text",
                "🔖 Copy Citation Only",
            ],
        )

        # Matrix Button in NASTab
        db_path = Path("src/db/3gpp_protocol_data.db")
        tab = NASTab(nas_db_path=db_path)
        self.assertEqual(tab.copy_matrix_btn.text(), "📋 Copy Table ▾")
        self.assertIsNotNone(tab.copy_matrix_btn.menu())
        matrix_actions = [a.text() for a in tab.copy_matrix_btn.menu().actions() if not a.isSeparator()]
        self.assertEqual(
            matrix_actions,
            [
                "📋 Copy for Word / Outlook (Rich Table)",
                "📝 Copy as Markdown Table",
                "📊 Copy as TSV (Excel / Plain Text)",
            ],
        )


if __name__ == "__main__":
    unittest.main()
