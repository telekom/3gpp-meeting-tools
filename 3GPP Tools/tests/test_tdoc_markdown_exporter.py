import os
import shutil
import tempfile
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
    write_index,
    ExportedDoc,
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

    def test_bulk_markdown_export_from_cache(self):
        tdoc1 = "S2-2400001"
        tdoc2 = "S2-2400002"

        # Create mock meeting folder with cached markdown and media
        tdoc1_folder = self.meeting_dir / tdoc1
        tdoc1_folder.mkdir(parents=True, exist_ok=True)
        tdoc1_media = tdoc1_folder / "media"
        tdoc1_media.mkdir(parents=True, exist_ok=True)
        (tdoc1_media / f"{tdoc1}_img1.png").write_bytes(b"dummy_png_bytes")

        cache1 = tdoc1_folder / f"{tdoc1}_LLM_v{TDOC_MD_EXTRACTOR_VERSION}.md"
        cache1.write_text(f"# {tdoc1}: Test Title 1\n\nContent for tdoc 1\n![img](media/{tdoc1}_img1.png)", encoding="utf-8")

        tdoc2_folder = self.meeting_dir / tdoc2
        tdoc2_folder.mkdir(parents=True, exist_ok=True)
        cache2 = tdoc2_folder / f"{tdoc2}_LLM_v{TDOC_MD_EXTRACTOR_VERSION}.md"
        cache2.write_text(f"# {tdoc2}: Test Title 2\n\nContent for tdoc 2", encoding="utf-8")

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

        export_dir = self.meeting_dir / "Export" / "Markdown"
        self.assertTrue(export_dir.exists(), "Export/Markdown directory must exist")

        md1_path = export_dir / f"{tdoc1}.md"
        md2_path = export_dir / f"{tdoc2}.md"
        index_path = export_dir / "index.md"
        media_img_path = export_dir / "media" / f"{tdoc1}_img1.png"

        self.assertTrue(md1_path.exists(), f"{tdoc1}.md was not exported")
        self.assertTrue(md2_path.exists(), f"{tdoc2}.md was not exported")
        self.assertTrue(index_path.exists(), "index.md manifest was not created")
        self.assertTrue(media_img_path.exists(), "Media image was not copied to export/media")

        # Check index.md contents
        index_content = index_path.read_text(encoding="utf-8")
        self.assertIn(f"[{tdoc1}]({tdoc1}.md)", index_content)
        self.assertIn(f"[{tdoc2}]({tdoc2}.md)", index_content)
        self.assertIn("Test Title 1", index_content)
        self.assertIn("Test Title 2", index_content)

    def test_single_markdown_export(self):
        tdoc_id = "S2-2400999"
        tdoc_folder = self.meeting_dir / tdoc_id
        tdoc_folder.mkdir(parents=True, exist_ok=True)
        cache_file = tdoc_folder / f"{tdoc_id}_LLM_v{TDOC_MD_EXTRACTOR_VERSION}.md"
        cache_file.write_text(f"# {tdoc_id}: Single Doc Title\n\nSingle doc body", encoding="utf-8")

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

        export_dir = self.meeting_dir / "Export" / "Markdown"
        single_md = export_dir / f"{tdoc_id}.md"
        self.assertTrue(single_md.exists(), f"{single_md} was not created")
        self.assertIn("Single doc body", single_md.read_text(encoding="utf-8"))


    def test_docx_conversion_to_markdown(self):
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

        export_dir = self.meeting_dir / "Export" / "Markdown"
        md_file = export_dir / f"{tdoc_id}.md"
        self.assertTrue(md_file.exists(), f"{md_file} was not generated")

        content = md_file.read_text(encoding="utf-8")
        self.assertIn("5G core network procedures", content)
        self.assertIn("SMF", content)

        # Check index.md
        index_file = export_dir / "index.md"
        self.assertTrue(index_file.exists())
        self.assertIn(f"[{tdoc_id}]({tdoc_id}.md)", index_file.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
