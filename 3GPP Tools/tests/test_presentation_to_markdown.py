"""
Unit tests for PowerPoint Presentation to Markdown Converter (presentation_to_markdown.py).
"""

import os
import posixpath
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

from modules.meetings.core.exporters.tdoc2md import DOC_EXTENSIONS, PRESENTATION_EXTENSIONS
from modules.meetings.core.exporters.export_manifest import (
    is_export_complete,
    plan_incremental_session,
    ExportManifest,
    reconstruct_manifest_from_directory,
)
from modules.meetings.core.tdocs.tdoc_markdown_converter import is_conversion_stale, CONVERTER_VERSION
from modules.meetings.core.exporters.post_meeting_exporter import PostMeetingExporterThread
from modules.word_tools.core.presentation_to_markdown import (
    ExtractedSlide,
    _format_markdown_table,
    _extract_via_openxml,
    convert_presentation_to_markdown,
)
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane


def _create_minimal_test_pptx(dest_path: Path) -> Path:
    """Creates a minimal well-formed OpenXML .pptx archive for unit testing."""
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dest_path, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
</Types>""",
        )

        # _rels/.rels
        z.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>""",
        )

        # ppt/presentation.xml
        z.writestr(
            "ppt/presentation.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <p:sldIdLst>
        <p:sldId id="256" r:id="rId2"/>
    </p:sldIdLst>
</p:presentation>""",
        )

        # ppt/_rels/presentation.xml.rels
        z.writestr(
            "ppt/_rels/presentation.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="/ppt/slides/slide1.xml"/>
</Relationships>""",
        )

        # ppt/slides/slide1.xml
        z.writestr(
            "ppt/slides/slide1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            <!-- Title Shape -->
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1"/>
                    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
                    <p:nvPr><p:ph type="title"/></p:nvPr>
                </p:nvSpPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:p><a:r><a:t>Discussion Paper on 6G Architecture</a:t></a:r></a:p>
                </p:txBody>
            </p:sp>
            <!-- Content Shape -->
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Content 1"/>
                    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
                    <p:nvPr><p:ph type="body"/></p:nvPr>
                </p:nvSpPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:p><a:pPr lvl="0"/><a:r><a:t>Proposal for converged data framework</a:t></a:r></a:p>
                    <a:p><a:pPr lvl="1"/><a:r><a:t>Sub-item on SA2 KI#21</a:t></a:r></a:p>
                </p:txBody>
            </p:sp>
            <!-- Table Shape -->
            <p:graphicFrame>
                <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                            <a:tr>
                                <a:tc><a:txBody><a:p><a:r><a:t>Work Item</a:t></a:r></a:p></a:txBody></a:tc>
                                <a:tc><a:txBody><a:p><a:r><a:t>Release</a:t></a:r></a:p></a:txBody></a:tc>
                            </a:tr>
                            <a:tr>
                                <a:tc><a:txBody><a:p><a:r><a:t>FS_6G_ARC</a:t></a:r></a:p></a:txBody></a:tc>
                                <a:tc><a:txBody><a:p><a:r><a:t>Rel-20</a:t></a:r></a:p></a:txBody></a:tc>
                            </a:tr>
                        </a:tbl>
                    </a:graphicData>
                </a:graphic>
            </p:graphicFrame>
        </p:spTree>
    </p:cSld>
</p:sld>""",
        )

        # ppt/slides/_rels/slide1.xml.rels
        z.writestr(
            "ppt/slides/_rels/slide1.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="/ppt/notesSlides/notesSlide1.xml"/>
</Relationships>""",
        )

        # ppt/notesSlides/notesSlide1.xml
        z.writestr(
            "ppt/notesSlides/notesSlide1.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Notes Body"/>
                    <p:nvPr><p:ph type="body"/></p:nvPr>
                </p:nvSpPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:p><a:r><a:t>Note to delegates: Please review TR 23.801-01 before discussion.</a:t></a:r></a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
</p:notes>""",
        )

        # ppt/media/dummy_figure.png
        z.writestr("ppt/media/dummy_figure.png", b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR...")

    return dest_path


class TestPresentationToMarkdown(unittest.TestCase):

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_ppt2md_"))

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_extension_definitions(self):
        """Verify .pptx and .ppt are first-class extensions in exporter registry."""
        self.assertIn(".pptx", DOC_EXTENSIONS)
        self.assertIn(".ppt", DOC_EXTENSIONS)
        self.assertIn(".pptx", PRESENTATION_EXTENSIONS)
        self.assertIn(".ppt", PRESENTATION_EXTENSIONS)

    def test_markdown_table_formatter(self):
        """Verify 2D string matrix renders into clean GitHub Flavored Markdown table."""
        rows = [
            ["Topic", "Spec", "Status"],
            ["URSP", "TS 24.526", "Normative"],
            ["V2X Policy", "TS 24.588 | Annex", "Study"],
        ]
        md = _format_markdown_table(rows)
        self.assertIn("| Topic | Spec | Status |", md)
        self.assertIn("| :--- | :--- | :--- |", md)
        self.assertIn("| URSP | TS 24.526 | Normative |", md)
        # Pipe characters in cells should be escaped
        self.assertIn("TS 24.588 \\| Annex", md)

    def test_openxml_presentation_parser(self):
        """Verify pure Python OpenXML parsing extracts titles, bullets, tables, and notes."""
        pptx_path = self.temp_dir / "S2-2612345.pptx"
        _create_minimal_test_pptx(pptx_path)

        media_dir = self.temp_dir / "media"
        slides, warnings = _extract_via_openxml(pptx_path, media_dir)

        self.assertEqual(len(slides), 1)
        slide = slides[0]
        self.assertEqual(slide.title, "Discussion Paper on 6G Architecture")
        self.assertTrue(any("Proposal for converged data framework" in p for p in slide.body_paragraphs))
        self.assertTrue(any("Sub-item on SA2 KI#21" in p for p in slide.body_paragraphs))

        # Table extraction
        self.assertEqual(len(slide.tables), 1)
        tbl = slide.tables[0]
        self.assertEqual(tbl[0], ["Work Item", "Release"])
        self.assertEqual(tbl[1], ["FS_6G_ARC", "Rel-20"])

        # Speaker notes
        self.assertIn("Note to delegates: Please review TR 23.801-01 before discussion.", slide.speaker_notes)

        # Embedded media extraction
        extracted_fig = media_dir / "dummy_figure.png"
        self.assertTrue(extracted_fig.exists())

    def test_convert_presentation_to_markdown_full(self):
        """Verify convert_presentation_to_markdown produces compliant ExportedDoc."""
        pptx_path = self.temp_dir / "S2-2612345.pptx"
        _create_minimal_test_pptx(pptx_path)

        out_dir = self.temp_dir / "out"
        meta = {
            "Title": "Discussion Paper on 6G Architecture",
            "Source": "Deutsche Telekom",
            "Type": "Discussion",
            "Agenda Item": "8.1",
        }

        exported = convert_presentation_to_markdown(
            ppt_path=pptx_path,
            out_dir=out_dir,
            tdoc_id="S2-2612345",
            meta=meta,
        )

        self.assertEqual(exported.tdoc_id, "S2-2612345")
        self.assertTrue(exported.md_path.exists())
        self.assertEqual(exported.md_path.name, "S2-2612345.md")

        text = exported.text
        # Header verification
        self.assertIn("# S2-2612345: Discussion Paper on 6G Architecture", text)
        self.assertIn("**Source:** Deutsche Telekom", text)
        self.assertIn("**Agenda Item:** 8.1", text)

        # Slide content
        self.assertIn("## Slide 1: Discussion Paper on 6G Architecture", text)
        self.assertIn("Proposal for converged data framework", text)
        self.assertIn("| Work Item | Release |", text)

        # Speaker notes callout
        self.assertIn("> [!NOTE] Speaker Notes", text)
        self.assertIn("> Note to delegates: Please review TR 23.801-01 before discussion.", text)

    def test_document_selector_pane_accepts_presentations(self):
        """Verify DocumentSelectorPane recognizes presentation file extensions."""
        import sys
        from PyQt5.QtWidgets import QApplication

        app = QApplication.instance() or QApplication(sys.argv)
        pane = DocumentSelectorPane(
            "Test Selector",
            accepted_extensions=[".docx", ".doc", ".pptx", ".ppt", ".md"],
        )
        self.assertIn(".pptx", pane.accepted_extensions)
        self.assertIn(".ppt", pane.accepted_extensions)

    def test_is_export_complete_and_stale_detection(self):
        """Verify incomplete or empty presentation exports are detected as stale."""
        t_dir = self.temp_dir / "stale_tests"
        t_dir.mkdir(parents=True, exist_ok=True)

        # 1. Non-existent file
        non_existent = t_dir / "non_existent.md"
        self.assertFalse(is_export_complete(non_existent))
        self.assertTrue(is_conversion_stale(None, non_existent))

        # 2. Empty 0-byte file
        empty_md = t_dir / "empty.md"
        empty_md.write_text("", encoding="utf-8")
        self.assertFalse(is_export_complete(empty_md))
        self.assertTrue(is_conversion_stale(None, empty_md))

        # 3. Truncated tiny stub (< 120 bytes)
        stub_md = t_dir / "stub.md"
        stub_md.write_text("# S2-260001\nSmall stub.", encoding="utf-8")
        self.assertFalse(is_export_complete(stub_md))
        self.assertTrue(is_conversion_stale(None, stub_md))

        # 4. Presentation header without any slides
        no_slides_md = t_dir / "no_slides.md"
        no_slides_md.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; type=presentation; mode=presentation --->\n"
            "# S2-260002: Discussion Paper\n"
            "- **TDoc:** S2-260002\n"
            "- **Original file:** `S2-260002.pptx`\n"
            "- **Slides:** 0 slides\n"
            "---\n\n"
            "Some text but no slide sections at all.\n",
            encoding="utf-8",
        )
        self.assertFalse(is_export_complete(no_slides_md))
        self.assertTrue(is_conversion_stale(None, no_slides_md))

        # 5. Presentation with missing referenced slide image
        missing_img_md = t_dir / "missing_img.md"
        missing_img_md.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; type=presentation; mode=presentation --->\n"
            "# S2-260003: Architecture Concepts\n"
            "- **TDoc:** S2-260003\n"
            "- **Original file:** `S2-260003.pptx`\n"
            "- **Slides:** 1 slides\n"
            "---\n\n"
            "## Slide 1: Introduction\n\n"
            "![Slide 1: Introduction](media/S2-260003_slide_01.png)\n",
            encoding="utf-8",
        )
        self.assertFalse(is_export_complete(missing_img_md))
        self.assertTrue(is_conversion_stale(None, missing_img_md))

        # 6. Valid complete presentation with image on disk
        valid_img_dir = t_dir / "media"
        valid_img_dir.mkdir(parents=True, exist_ok=True)
        img_file = valid_img_dir / "S2-260003_slide_01.png"
        img_file.write_bytes(b"\x89PNG\r\n\x1a\n")

        self.assertTrue(is_export_complete(missing_img_md))
        self.assertFalse(is_conversion_stale(None, missing_img_md))

    def test_plan_incremental_session_recovers_empty_powerpoint_export(self):
        """Verify incremental planning converts empty/incomplete PowerPoint exports instead of skipping."""
        exp_dir = self.temp_dir / "incremental_test"
        disc_dir = exp_dir / "Rel-19" / "FS_6G" / "Discussion"
        disc_dir.mkdir(parents=True, exist_ok=True)

        # Existing incomplete (0-byte) export for S2-2600010
        bad_md = disc_dir / "S2-2600010.md"
        bad_md.write_text("", encoding="utf-8")

        # Existing complete export for S2-2600020
        good_md = disc_dir / "S2-2600020.md"
        good_md.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n"
            "# S2-2600020: Standard CR\n"
            "- **TDoc:** S2-2600020\n"
            "- **Source:** Operator\n"
            "---\n\n"
            "Detailed text content describing changes in normative specification.\n" * 5,
            encoding="utf-8",
        )

        manifest = ExportManifest(
            meeting="SA2 #166",
            export_type="pre_meeting",
            sessions=[{"session_id": 1, "added_tdocs": ["S2-2600010", "S2-2600020"]}],
            tdocs={
                "S2-2600010": {"tdoc_id": "S2-2600010", "rel_path": "Rel-19/FS_6G/Discussion/S2-2600010.md"},
                "S2-2600020": {"tdoc_id": "S2-2600020", "rel_path": "Rel-19/FS_6G/Discussion/S2-2600020.md"},
            },
        )

        incoming_entries = {
            "Rel-19": {
                "FS_6G": {
                    "categories": {
                        "DISC": [
                            {"tdoc_id": "S2-2600010", "type": "Discussion", "target_md_abs_path": bad_md},
                            {"tdoc_id": "S2-2600020", "type": "CR", "target_md_abs_path": good_md},
                            {"tdoc_id": "S2-2600030", "type": "Discussion"},
                        ]
                    }
                }
            }
        }

        plan = plan_incremental_session(manifest, incoming_entries, exp_dir, force_reconvert=False)
        self.assertTrue(plan.is_incremental)
        self.assertIn("S2-2600020", plan.tdocs_to_skip, "Valid complete file should be skipped")
        self.assertIn("S2-2600010", plan.tdocs_to_update, "Empty/incomplete file must be scheduled for update/re-conversion")
        self.assertNotIn("S2-2600010", plan.tdocs_to_skip, "Empty/incomplete file must NOT be skipped")
        self.assertIn("S2-2600030", plan.new_tdocs, "Newly introduced TDoc must be in new_tdocs")

    def test_reconstruct_manifest_ignores_incomplete_export(self):
        """Verify reconstruct_manifest_from_directory skips 0-byte or incomplete files."""
        exp_dir = self.temp_dir / "reconstruct_test"
        d_dir = exp_dir / "General" / "General" / "Discussion"
        d_dir.mkdir(parents=True, exist_ok=True)

        (exp_dir / "index.md").write_text("# Master Index\n**Generated:** 2026-11-09 12:00\n", encoding="utf-8")
        (d_dir / "S2-2699991.md").write_text("", encoding="utf-8")  # Empty broken export
        (d_dir / "S2-2699992.md").write_text(  # Valid complete export
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n"
            "# S2-2699992: Title\n- **Source:** Vendor\n---\nValid content.\n" * 5,
            encoding="utf-8",
        )

        manifest = reconstruct_manifest_from_directory(exp_dir, {}, [], export_type="pre_meeting")
        self.assertIsNotNone(manifest)
        self.assertNotIn("S2-2699991", manifest.tdocs, "Incomplete/empty file must be omitted from manifest")
        self.assertIn("S2-2699992", manifest.tdocs, "Complete file must be indexed in manifest")

    def test_find_word_doc_finds_pptx_and_ppt(self):
        """Verify _find_word_doc finds .pptx and .ppt files."""
        folder = self.temp_dir / "find_doc_test"
        folder.mkdir(parents=True, exist_ok=True)
        pptx_file = folder / "S2-260123_Discussion.pptx"
        pptx_file.write_bytes(b"dummy")

        thread = PostMeetingExporterThread(
            meeting_dir=self.temp_dir,
            all_rows=[],
            visible_rows=[],
            agenda_map={},
            mtg_info={},
            docs_ftp_url="",
            revisions_url="",
        )
        found = thread._find_word_doc(folder, "S2-260123")
        self.assertIsNotNone(found)
        self.assertEqual(found.name, "S2-260123_Discussion.pptx")


if __name__ == "__main__":
    unittest.main()
