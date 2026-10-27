# --- File: tests/test_post_meeting_markdown_links.py ---
import re
import sys
import tempfile
import unittest
from pathlib import Path

# Ensure src is in sys.path
_src_dir = Path(__file__).resolve().parent.parent / "src"
if str(_src_dir) not in sys.path:
    sys.path.insert(0, str(_src_dir))

from PyQt5.QtWidgets import QApplication

from modules.meetings.core.post_meeting_exporter import (
    PostMeetingExporterThread,
    sanitize_folder_name,
    quote_markdown_path,
    safe_relpath,
    safe_markdown_rel_link,
    slugify_anchor,
)


class TestPostMeetingMarkdownLinks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def test_helpers(self):
        """Verify URL quoting, slugifying, and safe relative path helpers."""
        self.assertEqual(slugify_anchor("Ambient IoT"), "ambient-iot")
        self.assertEqual(slugify_anchor("5G_A-UP"), "5g_a-up")
        self.assertEqual(slugify_anchor("AI/ML in 5GS"), "ai-ml-in-5gs")

        self.assertEqual(
            quote_markdown_path("Rel-19/Ambient IoT/CR/S2-2601001.md"),
            "Rel-19/Ambient%20IoT/CR/S2-2601001.md",
        )
        self.assertEqual(
            quote_markdown_path("index.md#wi-ambient-iot"),
            "index.md#wi-ambient-iot",
        )

        p1 = Path("C:/Export/Rel-19/Ambient IoT/CR/doc.md")
        p2 = Path("C:/Export/index.md")
        rel = safe_relpath(p1, p2.parent)
        self.assertEqual(rel, "Rel-19/Ambient IoT/CR/doc.md")

        quoted_rel = safe_markdown_rel_link(p1, p2.parent)
        self.assertEqual(quoted_rel, "Rel-19/Ambient%20IoT/CR/doc.md")

    def test_post_meeting_export_links_validity(self):
        """
        Verify that all generated Markdown links in Master index.md, local WI index.md,
        and individual TDoc documents:
        1. Never point to bare directory paths (must target .md files).
        2. Never contain unencoded raw spaces in link destination URLs.
        3. Anchor links have matching HTML5 <a id="..."> targets without spaces.
        4. Breadcrumb navigation links exist to navigate between levels.
        """
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)

            rows = [
                # Incoming LS under 'Ambient IoT' with reply in S2-2602002
                {
                    "TDoc": "S2-2602001",
                    "Type": "LS in",
                    "TDoc Status": "replied to in S2-2602002",
                    "Related WIs": "Ambient IoT",
                    "Agenda Item": "5.1",
                    "Agenda item description": "Incoming LS (Ambient IoT)",
                    "Release": "Rel-19",
                    "Title": "LS on Ambient IoT Requirements",
                    "Reply in": "S2-2602002",
                },
                # Agreed reply LS out under 'Ambient IoT'
                {
                    "TDoc": "S2-2602002",
                    "Type": "LS out",
                    "TDoc Status": "agreed",
                    "Related WIs": "Ambient IoT",
                    "Agenda Item": "5.1",
                    "Agenda item description": "Outgoing LS (Ambient IoT)",
                    "Release": "Rel-19",
                    "Title": "Reply LS on Ambient IoT Requirements",
                    "Reply to": "S2-2602001",
                },
                # Agreed CR under 'Ambient IoT'
                {
                    "TDoc": "S2-2602003",
                    "Type": "CR",
                    "TDoc Status": "agreed",
                    "Related WIs": "Ambient IoT",
                    "Agenda Item": "5.2",
                    "Agenda item description": "Ambient IoT Architecture",
                    "Release": "Rel-19",
                    "Title": "CR on Ambient IoT Architecture",
                },
            ]

            export_dir = tmp_path / "Export" / "PostMeeting_Test"
            thread = PostMeetingExporterThread(
                meeting_dir=tmp_path,
                all_rows=rows,
                visible_rows=rows,
                agenda_map={},
                mtg_info={"wg_name": "SA2", "meeting_number": "160"},
                docs_ftp_url="https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_160_Athens/Docs",
                revisions_url="",
                scope_mode="whole",
                output_dir=export_dir,
                download_missing=False,
                generate_local_indices=True,
            )

            # Pre-populate dummy markdown in meeting_dir cache so converter is bypassed
            for r in rows:
                tid = r["TDoc"]
                tdoc_f = tmp_path / tid
                tdoc_f.mkdir(parents=True, exist_ok=True)
                (tdoc_f / f"{tid}.md").write_text(
                    f"# {tid}: {r['Title']}\n\nDummy body text.\n",
                    encoding="utf-8",
                )

            thread.run()

            # 1. Master index.md checks
            master_index = export_dir / "index.md"
            self.assertTrue(master_index.exists(), "Master index.md must be generated")
            master_content = master_index.read_text(encoding="utf-8")

            # Check for illegal bare directory links like `[Dossier Folder](Rel-19/Ambient IoT/)`
            bare_dir_match = re.search(r'\[([^\]]+)\]\(([^)]*\/)\)', master_content)
            self.assertIsNone(
                bare_dir_match,
                f"Found illegal bare directory link in master index: {bare_dir_match.group(0) if bare_dir_match else ''}",
            )

            # Check that all markdown links [text](url) in master index have no unencoded spaces
            # Regex extracts (url) in [text](url)
            link_urls = re.findall(r'\[(?:[^\]]*)\]\(([^)]+)\)', master_content)
            for url in link_urls:
                # Suffix title after quote allowed in markdown, but url itself shouldn't have raw unquoted spaces
                clean_url = url.split()[0]
                self.assertNotIn(
                    " ",
                    clean_url,
                    f"Link URL contains unencoded space in master index.md: {url}",
                )

            # Check that TOC links point to index.md
            self.assertIn("Dossier Index", master_content)
            self.assertIn("Rel-19/Ambient%20IoT/index.md", master_content)

            # Check that TOC anchor has no space and matches target <a id="...">
            toc_anchor_matches = re.findall(r'\[Ambient IoT\]\(#(wi-[^)]+)\)', master_content)
            self.assertTrue(toc_anchor_matches, "Should have TOC anchor link for Ambient IoT")
            for anc in toc_anchor_matches:
                self.assertNotIn(" ", anc, f"Anchor has spaces: {anc}")
                # Target must exist in document
                self.assertIn(f'<a id="{anc}"', master_content, f"Anchor target <a id='{anc}'> missing")

            # Check that table contains percent-encoded link to CR doc
            self.assertIn("Rel-19/Ambient%20IoT/CR/S2-2602003.md", master_content)

            # 2. Local WI index.md checks
            wi_index = export_dir / "Rel-19" / "Ambient IoT" / "index.md"
            self.assertTrue(wi_index.exists(), "Local WI index.md must exist")
            wi_content = wi_index.read_text(encoding="utf-8")

            # Must have breadcrumb back to master report
            self.assertIn("[⬅ Back to Master Report](../../index.md)", wi_content)

            # Table links in local WI index must be relative to wi_dir
            wi_links = re.findall(r'\[(?:[^\]]*)\]\(([^)]+)\)', wi_content)
            for url in wi_links:
                clean_url = url.split()[0]
                self.assertNotIn(
                    " ",
                    clean_url,
                    f"Link URL contains unencoded space in WI index.md: {url}",
                )

            # Check relative doc link in local WI index
            self.assertIn("CR/S2-2602003.md", wi_content)

            # 3. Individual TDoc document checks
            cr_doc = export_dir / "Rel-19" / "Ambient IoT" / "CR" / "S2-2602003.md"
            self.assertTrue(cr_doc.exists(), "CR doc must exist")
            cr_content = cr_doc.read_text(encoding="utf-8")

            # Must have breadcrumbs
            self.assertIn("[📁 Master Report](../../../index.md)", cr_content)
            self.assertIn("[🎯 Ambient IoT Dossier](../index.md)", cr_content)

            # Check reciprocal links in LS IN document
            ls_in_doc = export_dir / "Rel-19" / "Ambient IoT" / "LS_IN" / "S2-2602001.md"
            self.assertTrue(ls_in_doc.exists(), "LS IN doc must exist")
            ls_in_content = ls_in_doc.read_text(encoding="utf-8")

            # Reciprocal link to reply LS OUT
            self.assertIn("> [!TIP]", ls_in_content)
            self.assertIn("[S2-2602002](../LS_OUT/S2-2602002.md)", ls_in_content)


if __name__ == "__main__":
    unittest.main()
