import os
import shutil
import tempfile
import time
import unittest
from pathlib import Path

from core.utils.paths import to_extended_path
from modules.meetings.core.llm_exporter import (
    TDocMarkdownExporterThread,
    LLMExporterThread,
    TDOC_MD_EXTRACTOR_VERSION,
)
from modules.meetings.core.tdoc_markdown_converter import (
    convert_document_to_markdown,
    is_conversion_stale,
    write_index,
    ExportedDoc,
    CONVERTER_VERSION,
)


class TestTDocMarkdownExporter(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="3gpp_test_md_exp_"))
        self.meeting_dir = self.tmp_dir / "Meeting_101"
        self.meeting_dir.mkdir(parents=True, exist_ok=True)

    def tearDown(self):
        shutil.rmtree(to_extended_path(self.tmp_dir), ignore_errors=True)

    def test_exporter_backward_compatibility_alias(self):
        self.assertIs(LLMExporterThread, TDocMarkdownExporterThread)

    def test_stale_detection_logic(self):
        tdoc_id = "S2-2400100"
        folder = self.meeting_dir / tdoc_id
        folder.mkdir(parents=True, exist_ok=True)
        docx_file = folder / f"{tdoc_id}.docx"
        docx_file.write_bytes(b"dummy_docx_content")
        md_file = folder / f"{tdoc_id}.md"

        # 1. Stale when MD does not exist
        self.assertTrue(is_conversion_stale(docx_file, md_file))

        # 2. Write fresh MD file with header tag
        time.sleep(0.05)
        md_file.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n# {tdoc_id}\nContent",
            encoding="utf-8"
        )
        self.assertFalse(is_conversion_stale(docx_file, md_file))

        # 3. Stale when source docx is updated (newer mtime)
        time.sleep(0.05)
        docx_file.write_bytes(b"updated_docx_content")
        self.assertTrue(is_conversion_stale(docx_file, md_file))

        # Refresh MD
        time.sleep(0.05)
        md_file.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n# {tdoc_id}\nContent",
            encoding="utf-8"
        )
        self.assertFalse(is_conversion_stale(docx_file, md_file))

        # 4. Stale when converter version is bumped
        self.assertTrue(is_conversion_stale(docx_file, md_file, expected_version="9.9.9"))

        # 5. Stale when revision mode is mismatched
        self.assertTrue(is_conversion_stale(docx_file, md_file, expected_mode="accept"))

    def test_bulk_markdown_export_from_cache(self):
        tdoc1 = "S2-2400001"
        tdoc2 = "S2-2400002"

        # Create mock meeting folder with cached markdown and media
        tdoc1_folder = self.meeting_dir / tdoc1
        tdoc1_folder.mkdir(parents=True, exist_ok=True)
        tdoc1_media = tdoc1_folder / "media"
        tdoc1_media.mkdir(parents=True, exist_ok=True)
        (tdoc1_media / f"{tdoc1}_img1.png").write_bytes(b"dummy_png_bytes")

        # Canonical markdown dossier
        md1 = tdoc1_folder / f"{tdoc1}.md"
        md1.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n# {tdoc1}: Test Title 1\n\nContent for tdoc 1\n![img](media/{tdoc1}_img1.png)",
            encoding="utf-8"
        )

        tdoc2_folder = self.meeting_dir / tdoc2
        tdoc2_folder.mkdir(parents=True, exist_ok=True)
        md2 = tdoc2_folder / f"{tdoc2}.md"
        md2.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n# {tdoc2}: Test Title 2\n\nContent for tdoc 2",
            encoding="utf-8"
        )

        tdocs_list = [
            {"TDoc": tdoc1, "Title": "Test Title 1", "Type": "CR", "Agenda Item": "5.1", "TDoc Status": "Agreed", "Source": "Operator A"},
            {"TDoc": tdoc2, "Title": "Test Title 2", "Type": "Discussion", "Agenda Item": "5.2", "TDoc Status": "Noted", "Source": "Vendor B"},
        ]

        exporter = TDocMarkdownExporterThread(
            meeting_dir=self.meeting_dir,
            tdocs_list=tdocs_list,
            docs_ftp_url="",
            revisions_url="",
            is_bulk=True,
        )

        results = []
        exporter.finished.connect(lambda success, msg: results.append((success, msg)))
        exporter.run()

        self.assertTrue(results, "Exporter thread should have emitted finished signal")
        success, msg = results[0]
        self.assertTrue(success, f"Exporter failed: {msg}")

        # Check export directory: Export/Markdown_YYYY-MM-DD_HHMMSS/
        export_dir = exporter.export_dir
        self.assertTrue(export_dir.exists(), f"Export directory {export_dir} must exist")
        self.assertTrue(export_dir.name.startswith("Markdown_"), "Export folder should follow Markdown_YYYY-MM-DD_HHMMSS naming")

        # Each TDoc must be exported into its own subfolder
        dossier1 = export_dir / tdoc1
        dossier2 = export_dir / tdoc2
        self.assertTrue(dossier1.is_dir(), f"Dossier subfolder for {tdoc1} must exist")
        self.assertTrue(dossier2.is_dir(), f"Dossier subfolder for {tdoc2} must exist")

        md1_path = dossier1 / f"{tdoc1}.md"
        md2_path = dossier2 / f"{tdoc2}.md"
        index_path = export_dir / "index.md"
        media_img_path = dossier1 / "media" / f"{tdoc1}_img1.png"

        self.assertTrue(md1_path.exists(), f"{tdoc1}.md was not exported to its dossier")
        self.assertTrue(md2_path.exists(), f"{tdoc2}.md was not exported to its dossier")
        self.assertTrue(index_path.exists(), "index.md manifest was not created")
        self.assertTrue(media_img_path.exists(), "Media image was not copied to tdoc1/media")

        # Check index.md contents link to subfolder dossiers
        index_content = index_path.read_text(encoding="utf-8")
        self.assertIn(f"[{tdoc1}]({tdoc1}/{tdoc1}.md)", index_content)
        self.assertIn(f"[{tdoc2}]({tdoc2}/{tdoc2}.md)", index_content)
        self.assertIn("Test Title 1", index_content)
        self.assertIn("Test Title 2", index_content)

    def test_single_markdown_export(self):
        tdoc_id = "S2-2400999"
        tdoc_folder = self.meeting_dir / tdoc_id
        tdoc_folder.mkdir(parents=True, exist_ok=True)
        local_md = tdoc_folder / f"{tdoc_id}.md"
        local_md.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n# {tdoc_id}: Single Doc Title\n\nSingle doc body",
            encoding="utf-8"
        )

        tdocs_list = [
            {"TDoc": tdoc_id, "Title": "Single Doc Title", "Type": "CR", "Agenda Item": "9.1", "TDoc Status": "Agreed"}
        ]

        exporter = TDocMarkdownExporterThread(
            meeting_dir=self.meeting_dir,
            tdocs_list=tdocs_list,
            docs_ftp_url="",
            revisions_url="",
            is_bulk=False,
        )

        results = []
        exporter.finished.connect(lambda success, msg: results.append((success, msg)))
        exporter.run()

        self.assertTrue(results)
        success, msg = results[0]
        self.assertTrue(success)

        export_dir = exporter.export_dir
        single_md = export_dir / tdoc_id / f"{tdoc_id}.md"
        self.assertTrue(single_md.exists(), f"{single_md} was not created in dossier subfolder")
        self.assertIn("Single doc body", single_md.read_text(encoding="utf-8"))

    def test_docx_conversion_to_markdown_dossier(self):
        import docx

        tdoc_id = "S2-2405555"
        tdoc_folder = self.meeting_dir / tdoc_id
        tdoc_folder.mkdir(parents=True, exist_ok=True)

        # Create a sample .docx file
        doc = docx.Document()
        doc.add_heading("3GPP TSG-SA WG2 Meeting #160", level=1)
        doc.add_paragraph("This is an architectural proposal for 5G core network procedures.")
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Param"
        table.cell(0, 1).text = "Value"
        table.cell(1, 0).text = "SMF"
        table.cell(1, 1).text = "Active"
        docx_path = tdoc_folder / f"{tdoc_id}.docx"
        doc.save(str(docx_path))

        tdocs_list = [
            {"TDoc": tdoc_id, "Title": "5GC Architectural Proposal", "Type": "CR", "Agenda Item": "7.3", "TDoc Status": "Agreed", "Source": "Operator X"}
        ]

        exporter = TDocMarkdownExporterThread(
            meeting_dir=self.meeting_dir,
            tdocs_list=tdocs_list,
            docs_ftp_url="",
            revisions_url="",
            is_bulk=True,
        )

        results = []
        exporter.finished.connect(lambda success, msg: results.append((success, msg)))
        exporter.run()

        self.assertTrue(results)
        success, msg = results[0]
        self.assertTrue(success, f"Exporter failed: {msg}")

        # 1. Local canonical dossier in meeting_dir / tdoc_id must exist
        local_md = tdoc_folder / f"{tdoc_id}.md"
        self.assertTrue(local_md.exists(), "Local canonical Markdown file must be created")

        # 2. Export dossier in Export/Markdown_.../tdoc_id/ must exist
        export_dir = exporter.export_dir
        exported_md = export_dir / tdoc_id / f"{tdoc_id}.md"
        self.assertTrue(exported_md.exists(), f"{exported_md} was not generated")

        content = exported_md.read_text(encoding="utf-8")
        self.assertIn("5G core network procedures", content)
        self.assertIn("SMF", content)

        # 3. Check index.md links to subfolder
        index_file = export_dir / "index.md"
        self.assertTrue(index_file.exists())
        self.assertIn(f"[{tdoc_id}]({tdoc_id}/{tdoc_id}.md)", index_file.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
