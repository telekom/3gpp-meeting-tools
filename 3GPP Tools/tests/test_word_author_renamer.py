# --- File: tests/test_word_author_renamer.py ---
import io
import shutil
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from modules.word_tools.core.author_renamer import (
    AuthorStats,
    inspect_authors,
    replace_authors,
    W_NS,
    CP_NS,
    DC_NS,
)


def create_sample_docx(target_path: Path) -> Path:
    """Creates a minimal valid .docx package containing revisions and comments."""
    content_types = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        b'<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        b'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        b'</Types>'
    )

    doc_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}">'
        f'<w:body>'
        f'<w:p>'
        f'<w:r><w:t>Baseline text</w:t></w:r>'
        f'<w:ins w:id="1" w:author="Alice (EXT)" w:date="2026-10-27T10:00:00Z">'
        f'<w:r><w:t> Inserted by Alice</w:t></w:r>'
        f'</w:ins>'
        f'<w:del w:id="2" w:author="Alice (EXT)" w:date="2026-10-27T10:01:00Z">'
        f'<w:r><w:delText>Deleted by Alice</w:delText></w:r>'
        f'</w:del>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rPrChange w:id="3" w:author="Bob" w:date="2026-10-27T10:05:00Z"/>'
        f'</w:rPr>'
        f'<w:t>Formatted text</w:t>'
        f'</w:r>'
        f'</w:p>'
        f'</w:body>'
        f'</w:document>'
    ).encode("utf-8")

    comments_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:comments xmlns:w="{W_NS}">'
        f'<w:comment w:id="1" w:author="Alice (EXT)" w:initials="AE" w:date="2026-10-27T10:02:00Z">'
        f'<w:p><w:r><w:t>Comment by Alice</w:t></w:r></w:p>'
        f'</w:comment>'
        f'</w:comments>'
    ).encode("utf-8")

    core_props_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<cp:coreProperties xmlns:cp="{CP_NS}" xmlns:dc="{DC_NS}">'
        f'<dc:creator>Alice (EXT)</dc:creator>'
        f'<cp:lastModifiedBy>Bob</cp:lastModifiedBy>'
        f'</cp:coreProperties>'
    ).encode("utf-8")

    with zipfile.ZipFile(target_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/comments.xml", comments_xml)
        z.writestr("docProps/core.xml", core_props_xml)

    return target_path


class TestWordAuthorRenamer(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_author_renamer_"))
        self.sample_docx = self.temp_dir / "sample_contribution.docx"
        create_sample_docx(self.sample_docx)

    def tearDown(self):
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_inspect_authors(self):
        """Verify author stats extraction across document and comments."""
        stats = inspect_authors(self.sample_docx)
        self.assertEqual(len(stats), 2)

        # Alice should be first (3 changes)
        alice = stats[0]
        self.assertEqual(alice.name, "Alice (EXT)")
        self.assertEqual(alice.insertions, 1)
        self.assertEqual(alice.deletions, 1)
        self.assertEqual(alice.formatting, 0)
        self.assertEqual(alice.comments, 1)
        self.assertEqual(alice.initials, "AE")
        self.assertEqual(alice.total, 3)

        # Bob should be second (1 change)
        bob = stats[1]
        self.assertEqual(bob.name, "Bob")
        self.assertEqual(bob.insertions, 0)
        self.assertEqual(bob.deletions, 0)
        self.assertEqual(bob.formatting, 1)
        self.assertEqual(bob.comments, 0)
        self.assertEqual(bob.total, 1)

    def test_replace_authors_new_copy(self):
        """Verify replacing authors into a new copy."""
        out_docx = self.temp_dir / "anonymized_contribution.docx"
        mapping = {
            "Alice (EXT)": "Deutsche Telekom",
            "Bob": "Deutsche Telekom",
        }
        success, count, msg = replace_authors(
            docx_path=self.sample_docx,
            output_path=out_docx,
            mapping=mapping,
            new_initials="DT",
            update_metadata=True,
            backup=False,
        )
        self.assertTrue(success)
        self.assertEqual(count, 4)  # 2 in doc (Alice), 1 in doc (Bob), 1 in comments (Alice)

        # Inspect new document
        new_stats = inspect_authors(out_docx)
        self.assertEqual(len(new_stats), 1)
        dt = new_stats[0]
        self.assertEqual(dt.name, "Deutsche Telekom")
        self.assertEqual(dt.total, 4)
        self.assertEqual(dt.initials, "DT")

        # Verify docProps/core.xml
        with zipfile.ZipFile(out_docx, "r") as z:
            core_xml = z.read("docProps/core.xml")
            root = ET.fromstring(core_xml)
            creator = root.find(f".//{{{DC_NS}}}creator")
            last_mod = root.find(f".//{{{CP_NS}}}lastModifiedBy")
            self.assertIsNotNone(creator)
            self.assertEqual(creator.text, "Deutsche Telekom")
            self.assertIsNotNone(last_mod)
            self.assertEqual(last_mod.text, "Deutsche Telekom")

    def test_replace_authors_inplace_with_backup(self):
        """Verify in-place replacement and .bak creation."""
        mapping = {"Alice (EXT)": "Deutsche Telekom"}
        success, count, msg = replace_authors(
            docx_path=self.sample_docx,
            output_path=self.sample_docx,
            mapping=mapping,
            new_initials="DT",
            update_metadata=False,
            backup=True,
        )
        self.assertTrue(success)
        self.assertEqual(count, 3)

        # Verify .bak backup exists
        bak_file = self.sample_docx.with_suffix(".docx.bak")
        self.assertTrue(bak_file.is_file())

        # Inspect backup: should still have original authors
        bak_stats = inspect_authors(bak_file)
        author_names = {s.name for s in bak_stats}
        self.assertIn("Alice (EXT)", author_names)

        # Inspect modified file: should have Deutsche Telekom and Bob
        mod_stats = inspect_authors(self.sample_docx)
        mod_names = {s.name for s in mod_stats}
        self.assertIn("Deutsche Telekom", mod_names)
        self.assertIn("Bob", mod_names)
        self.assertNotIn("Alice (EXT)", mod_names)

    def test_xml_entity_escaping(self):
        """Verify that company names with special characters ('AT&T & Co.') are properly escaped."""
        out_docx = self.temp_dir / "att_contribution.docx"
        mapping = {"Alice (EXT)": "AT&T & Co."}
        success, count, msg = replace_authors(
            docx_path=self.sample_docx,
            output_path=out_docx,
            mapping=mapping,
            new_initials="ATT",
            update_metadata=True,
        )
        self.assertTrue(success)

        # Re-reading via ElementTree must succeed without ParseError
        with zipfile.ZipFile(out_docx, "r") as z:
            doc_data = z.read("word/document.xml")
            root = ET.fromstring(doc_data)
            # Find ins element and verify author attribute
            ins = root.find(f".//{{{W_NS}}}ins")
            self.assertEqual(ins.attrib.get(f"{{{W_NS}}}author"), "AT&T & Co.")

    def test_scrub_initials(self):
        """Verify that passing new_initials='' scrubs reviewer initials from comments."""
        out_docx = self.temp_dir / "scrubbed_contribution.docx"
        mapping = {"Alice (EXT)": "Editor"}
        success, count, msg = replace_authors(
            docx_path=self.sample_docx,
            output_path=out_docx,
            mapping=mapping,
            new_initials="",  # Empty string indicates scrubbing
        )
        self.assertTrue(success)

        with zipfile.ZipFile(out_docx, "r") as z:
            comm_data = z.read("word/comments.xml")
            root = ET.fromstring(comm_data)
            comment = root.find(f".//{{{W_NS}}}comment")
            self.assertEqual(comment.attrib.get(f"{{{W_NS}}}author"), "Editor")
            self.assertNotIn(f"{{{W_NS}}}initials", comment.attrib)


class TestWordAuthorRenamerUi(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from PyQt5.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_author_ui_"))
        self.sample_docx = self.temp_dir / "sample_ui_test.docx"
        create_sample_docx(self.sample_docx)

    def tearDown(self):
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_ui_card_inspection_and_presets(self):
        """Verify that WordExtractorTab author card inspects documents and manages presets."""
        from PyQt5.QtCore import Qt
        from modules.word_tools.ui.word_tabs import WordExtractorTab

        tab = WordExtractorTab()
        tab.set_active_operation(8)
        self.assertEqual(tab.stack.currentIndex(), 8)

        # Feed the sample document
        tab._inspect_selected_author_document(str(self.sample_docx))
        self.assertEqual(tab.author_table.rowCount(), 2)

        # Verify rows populated
        item_alice = tab.author_table.item(0, 1)
        self.assertEqual(item_alice.text(), "Alice (EXT)")
        chk_alice = tab.author_table.item(0, 0)
        self.assertEqual(chk_alice.checkState(), Qt.Checked)

        item_bob = tab.author_table.item(1, 1)
        self.assertEqual(item_bob.text(), "Bob")

        # Test Deselect All & Select All
        tab._select_all_authors(False)
        self.assertEqual(tab.author_table.item(0, 0).checkState(), Qt.Unchecked)
        self.assertEqual(tab.author_table.item(1, 0).checkState(), Qt.Unchecked)

        tab._select_all_authors(True)
        self.assertEqual(tab.author_table.item(0, 0).checkState(), Qt.Checked)

        # Test Presets
        tab._apply_author_preset("Editor", "Ed")
        self.assertEqual(tab.edit_target_author.text(), "Editor")
        self.assertEqual(tab.edit_target_initials.text(), "Ed")

        tab._apply_author_preset("Rapporteur", "Rapp")
        self.assertEqual(tab.edit_target_author.text(), "Rapporteur")
        self.assertEqual(tab.edit_target_initials.text(), "Rapp")

        # Test Initials derivation
        self.assertEqual(tab._derive_initials("Deutsche Telekom"), "DT")
        self.assertEqual(tab._derive_initials("Qualcomm Technologies"), "QT")
        self.assertEqual(tab._derive_initials("Editor"), "Edi")


if __name__ == "__main__":
    unittest.main()
