# --- File: tests/test_document_sanitizer.py ---
import io
import shutil
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from modules.word_tools.core.document_sanitizer import (
    LanguageStats,
    DocumentHygieneAudit,
    SanitizeConfig,
    inspect_document,
    sanitize_document,
    W_NS,
    A_NS,
    CP_NS,
    DC_NS,
)


def create_rich_sample_docx(target_path: Path) -> Path:
    """Creates a comprehensive sample docx with languages, changes, and styles."""
    content_types = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        b'<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        b'<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        b'<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        b'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        b'</Types>'
    )

    styles_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{W_NS}">'
        f'<w:docDefaults>'
        f'<w:rPrDefault>'
        f'<w:rPr><w:lang w:val="de-DE"/></w:rPr>'
        f'</w:rPrDefault>'
        f'</w:docDefaults>'
        f'<w:style w:type="paragraph" w:styleId="Normal">'
        f'<w:name w:val="Normal"/>'
        f'<w:rPr><w:lang w:val="de-DE"/></w:rPr>'
        f'</w:style>'
        f'</w:styles>'
    ).encode("utf-8")

    doc_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}" xmlns:a="{A_NS}">'
        f'<w:body>'
        # Paragraph 1: German text with an insertion
        f'<w:p>'
        f'<w:pPr><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:lang w:val="de-DE"/></w:rPr><w:t>Hallo Welt</w:t></w:r>'
        f'<w:ins w:id="1" w:author="Alice" w:date="2026-10-28T08:00:00Z">'
        f'<w:r><w:rPr><w:lang w:val="de-DE"/></w:rPr><w:t> Neu von Alice</w:t></w:r>'
        f'</w:ins>'
        f'</w:p>'
        # Paragraph 2: Pure language change revision mark (de-DE -> en-US)
        f'<w:p>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:b/>'
        f'<w:lang w:val="en-US"/>'
        f'<w:rPrChange w:id="2" w:author="Bob" w:date="2026-10-28T08:05:00Z">'
        f'<w:rPr><w:b/><w:lang w:val="de-DE"/></w:rPr>'
        f'</w:rPrChange>'
        f'</w:rPr>'
        f'<w:t>Pure language changed text</w:t>'
        f'</w:r>'
        f'</w:p>'
        # Paragraph 3: Compound revision mark (added italic AND changed language de-DE -> en-US)
        f'<w:p>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:b/>'
        f'<w:i/>'
        f'<w:lang w:val="en-US"/>'
        f'<w:rPrChange w:id="3" w:author="Bob" w:date="2026-10-28T08:06:00Z">'
        f'<w:rPr><w:b/><w:lang w:val="de-DE"/></w:rPr>'
        f'</w:rPrChange>'
        f'</w:rPr>'
        f'<w:t>Compound changed text</w:t>'
        f'</w:r>'
        f'</w:p>'
        # Paragraph 4: DrawingML text run with lang attribute
        f'<w:p>'
        f'<w:r>'
        f'<a:rPr lang="de-DE"><a:t>Drawing text</a:t></a:rPr>'
        f'</w:r>'
        f'</w:p>'
        f'</w:body>'
        f'</w:document>'
    ).encode("utf-8")

    header_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:hdr xmlns:w="{W_NS}">'
        f'<w:p><w:r><w:rPr><w:lang w:val="de-DE"/></w:rPr><w:t>Header Text</w:t></w:r></w:p>'
        f'</w:hdr>'
    ).encode("utf-8")

    comments_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:comments xmlns:w="{W_NS}">'
        f'<w:comment w:id="1" w:author="Alice" w:initials="AL" w:date="2026-10-28T08:02:00Z">'
        f'<w:p><w:r><w:t>Comment text</w:t></w:r></w:p>'
        f'</w:comment>'
        f'</w:comments>'
    ).encode("utf-8")

    core_props_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<cp:coreProperties xmlns:cp="{CP_NS}" xmlns:dc="{DC_NS}">'
        f'<dc:creator>Original Author</dc:creator>'
        f'<cp:lastModifiedBy>Editor</cp:lastModifiedBy>'
        f'</cp:coreProperties>'
    ).encode("utf-8")

    with zipfile.ZipFile(target_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("word/styles.xml", styles_xml)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/header1.xml", header_xml)
        z.writestr("word/comments.xml", comments_xml)
        z.writestr("docProps/core.xml", core_props_xml)

    return target_path


class TestDocumentSanitizer(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_sanitizer_"))
        self.sample_docx = self.temp_dir / "sample_doc.docx"
        create_rich_sample_docx(self.sample_docx)

    def tearDown(self):
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_inspect_document(self):
        """Verifies deep inspection of languages, authors, and change marks."""
        audit = inspect_document(self.sample_docx)

        self.assertEqual(audit.file_name, "sample_doc.docx")
        self.assertEqual(audit.language_stats.default_language, "de-DE")
        # Check language breakdown
        self.assertIn("de-DE", audit.language_stats.language_counts)
        self.assertIn("en-US", audit.language_stats.language_counts)
        self.assertGreaterEqual(audit.language_stats.language_change_marks, 2)

        # Check authors
        author_names = {a.name for a in audit.authors}
        self.assertIn("Alice", author_names)
        self.assertIn("Bob", author_names)
        self.assertEqual(audit.total_insertions, 1)
        self.assertEqual(audit.total_comments, 1)
        self.assertEqual(audit.doc_creator, "Original Author")

    def test_sanitize_language_unification_and_purge(self):
        """
        Verifies:
        1. All languages become en-GB in styles, docDefaults, runs, headers, and DrawingML.
        2. Pure language change rPrChange is purged completely.
        3. Compound change rPrChange is retained, but inner lang is synchronized to en-GB (bold/italic kept).
        """
        out_docx = self.temp_dir / "sanitized_doc.docx"
        cfg = SanitizeConfig(
            unify_language=True,
            target_language="en-GB",
            apply_to_all_stories=True,
            update_styles_and_defaults=True,
            purge_language_changes=True,
            update_authors=False,
            backup=False,
        )

        success, mod_count, msg = sanitize_document(self.sample_docx, out_docx, cfg)
        self.assertTrue(success)
        self.assertGreater(mod_count, 0)

        with zipfile.ZipFile(out_docx, "r") as z:
            # 1. Check styles.xml docDefaults
            styles_xml = z.read("word/styles.xml")
            s_root = ET.fromstring(styles_xml)
            doc_def_lang = s_root.find(f".//{{{W_NS}}}docDefaults//{{{W_NS}}}lang")
            self.assertIsNotNone(doc_def_lang)
            self.assertEqual(doc_def_lang.attrib.get(f"{{{W_NS}}}val"), "en-GB")

            # Check Normal style
            norm_lang = s_root.find(f".//{{{W_NS}}}style[@{{{W_NS}}}styleId='Normal']//{{{W_NS}}}lang")
            self.assertIsNotNone(norm_lang)
            self.assertEqual(norm_lang.attrib.get(f"{{{W_NS}}}val"), "en-GB")

            # 2. Check header1.xml
            hdr_xml = z.read("word/header1.xml")
            h_root = ET.fromstring(hdr_xml)
            h_lang = h_root.find(f".//{{{W_NS}}}lang")
            self.assertIsNotNone(h_lang)
            self.assertEqual(h_lang.attrib.get(f"{{{W_NS}}}val"), "en-GB")

            # 3. Check document.xml
            doc_xml = z.read("word/document.xml")
            d_root = ET.fromstring(doc_xml)

            # A. Paragraph 1 runs should be en-GB
            first_p_lang = d_root.find(f".//{{{W_NS}}}p//{{{W_NS}}}lang")
            self.assertEqual(first_p_lang.attrib.get(f"{{{W_NS}}}val"), "en-GB")

            # B. Pure language change (rPrChange w:id="2") should be completely removed!
            rpr_change_2 = d_root.find(f".//{{{W_NS}}}rPrChange[@{{{W_NS}}}id='2']")
            self.assertIsNone(rpr_change_2, "Pure language rPrChange was not purged!")

            # C. Compound change (rPrChange w:id="3") should be retained for its italic change,
            # but inner lang must now be en-GB!
            rpr_change_3 = d_root.find(f".//{{{W_NS}}}rPrChange[@{{{W_NS}}}id='3']")
            self.assertIsNotNone(rpr_change_3, "Compound rPrChange was incorrectly purged!")
            inner_lang = rpr_change_3.find(f".//{{{W_NS}}}rPr/{{{W_NS}}}lang")
            self.assertIsNotNone(inner_lang)
            self.assertEqual(inner_lang.attrib.get(f"{{{W_NS}}}val"), "en-GB")

            # D. DrawingML lang attribute
            drawing_rpr = d_root.find(f".//{{{A_NS}}}rPr")
            self.assertIsNotNone(drawing_rpr)
            self.assertEqual(drawing_rpr.attrib.get("lang"), "en-GB")

    def test_sanitize_author_replacement_and_metadata(self):
        """Verifies author consolidation and docProps synchronization."""
        out_docx = self.temp_dir / "anonymized_doc.docx"
        cfg = SanitizeConfig(
            unify_language=False,
            update_authors=True,
            author_mapping={"Alice": "Deutsche Telekom", "Bob": "Deutsche Telekom"},
            new_initials="DT",
            update_metadata=True,
            backup=False,
        )

        success, mod_count, msg = sanitize_document(self.sample_docx, out_docx, cfg)
        self.assertTrue(success)

        audit = inspect_document(out_docx)
        self.assertEqual(len(audit.authors), 1)
        self.assertEqual(audit.authors[0].name, "Deutsche Telekom")
        self.assertEqual(audit.authors[0].initials, "DT")

        # Verify metadata
        with zipfile.ZipFile(out_docx, "r") as z:
            core_xml = z.read("docProps/core.xml")
            c_root = ET.fromstring(core_xml)
            creator = c_root.find(f".//{{{DC_NS}}}creator")
            self.assertEqual(creator.text, "Deutsche Telekom")

    def test_purge_all_formatting_and_strip_comments(self):
        """Verifies optional stripping of all formatting revisions and comments."""
        out_docx = self.temp_dir / "cleaned_doc.docx"
        cfg = SanitizeConfig(
            unify_language=False,
            update_authors=False,
            purge_all_formatting_changes=True,
            remove_all_comments=True,
            backup=False,
        )

        success, mod_count, msg = sanitize_document(self.sample_docx, out_docx, cfg)
        self.assertTrue(success)

        with zipfile.ZipFile(out_docx, "r") as z:
            doc_xml = z.read("word/document.xml")
            d_root = ET.fromstring(doc_xml)

            # All rPrChange should be gone
            self.assertEqual(len(d_root.findall(f".//{{{W_NS}}}rPrChange")), 0)
            # Text insertion (<w:ins>) MUST still be preserved!
            self.assertEqual(len(d_root.findall(f".//{{{W_NS}}}ins")), 1)

            # Comments should be empty
            comm_xml = z.read("word/comments.xml")
            c_root = ET.fromstring(comm_xml)
            self.assertEqual(len(c_root.findall(f".//{{{W_NS}}}comment")), 0)


class TestDocumentSanitizerUi(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from PyQt5.QtWidgets import QApplication
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_sanitizer_ui_"))
        self.sample_docx = self.temp_dir / "sample_ui_doc.docx"
        create_rich_sample_docx(self.sample_docx)

    def tearDown(self):
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_ui_inspection_and_presets(self):
        """Verifies UI audit badges, table population, and preset switching."""
        from modules.word_tools.ui.word_tabs import WordExtractorTab

        tab = WordExtractorTab()
        tab.set_active_operation(8)
        self.assertEqual(tab.stack.currentIndex(), 8)

        # Inspect the rich sample
        tab._inspect_selected_author_document(str(self.sample_docx))

        # Check badges
        self.assertIn("de-DE", tab.lbl_kpi_lang.text())
        self.assertIn("Authors: 2", tab.lbl_kpi_authors.text())
        self.assertIn("Lang Changes: 2", tab.lbl_kpi_lang_rev.text())
        self.assertIn("Comments: 1", tab.lbl_kpi_comments.text())

        # Test Presets
        tab._apply_preset_lang_only()
        self.assertTrue(tab.chk_sec_lang.isChecked())
        self.assertEqual(tab.combo_target_lang.currentData(), "en-GB")
        self.assertFalse(tab.chk_sec_authors.isChecked())

        tab._apply_preset_3gpp()
        self.assertTrue(tab.chk_sec_lang.isChecked())
        self.assertTrue(tab.chk_sec_authors.isChecked())
        self.assertEqual(tab.edit_target_author.text(), "Deutsche Telekom")

        tab._apply_preset_anonymize()
        self.assertTrue(tab.chk_sec_cleanup.isChecked())
        self.assertTrue(tab.chk_remove_comments.isChecked())
        self.assertEqual(tab.edit_target_author.text(), "Contributor")


if __name__ == "__main__":
    unittest.main()
