import os
import shutil
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch, MagicMock

from core.utils.paths import to_extended_path
from modules.meetings.core.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.llm_exporter import (
    TDocMarkdownExporterThread,
    LLMExporterThread,
    TDOC_MD_EXTRACTOR_VERSION,
)
from modules.meetings.core.tdoc_markdown_converter import (
    convert_document_to_markdown,
    is_conversion_stale,
    is_plain_document_type,
    resolve_revision_mode_for_type,
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

    def test_plain_document_type_resolution(self):
        """Verify that LS documents are classified as plain documents, while CRs are not."""
        plain_types = [
            "LS in", "ls in", "LS IN", "LS (in)", "Incoming LS",
            "LS out", "ls out", "LS OUT", "LS (out)", "Outgoing LS",
            "LS", "ls", "Liaison", "Liaison Statement", "Reply LS",
            "LS_IN", "LS_OUT",
        ]
        for t in plain_types:
            self.assertTrue(is_plain_document_type(t), f"'{t}' must be identified as plain document")
            self.assertEqual(resolve_revision_mode_for_type(t), "accept",
                             f"'{t}' must resolve to 'accept' revision mode")

        markup_types = [
            "CR", "cr", "pCR", "p-cr", "Draft CR", "CR (draft)", "Discussion", "Report", "WID", "", None
        ]
        for t in markup_types:
            self.assertFalse(is_plain_document_type(t), f"'{t}' must NOT be identified as plain document")
            self.assertEqual(resolve_revision_mode_for_type(t), "markup",
                             f"'{t}' must default to 'markup' revision mode")

        # Explicit mode overrides should be preserved
        self.assertEqual(resolve_revision_mode_for_type("LS in", default_mode="reject"), "reject")
        self.assertEqual(resolve_revision_mode_for_type("CR", default_mode="accept"), "accept")

    def test_plain_document_export_omits_markup_and_legend(self):
        """
        Verify that LS documents with Word tracked changes are exported in 'accept' mode
        without change marks ([INSERTED]/[DELETED]) and without markup legend,
        while CR documents retain change marks and markup legend.
        """
        import docx
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        folder = self.meeting_dir / "Docs"
        folder.mkdir(parents=True, exist_ok=True)

        def create_doc_with_track_changes(filename: str) -> Path:
            doc = docx.Document()
            doc.add_heading("Technical Statement", level=1)
            p = doc.add_paragraph("Base text: ")

            # Add an inserted run: <w:ins><w:r><w:t>new addition</w:t></w:r></w:ins>
            ins = OxmlElement("w:ins")
            ins.set(qn("w:id"), "1")
            ins.set(qn("w:author"), "Author")
            r_ins = OxmlElement("w:r")
            t_ins = OxmlElement("w:t")
            t_ins.text = "new addition"
            r_ins.append(t_ins)
            ins.append(r_ins)
            p._p.append(ins)

            # Add a deleted run: <w:del><w:r><w:delText>old removed</w:delText></w:r></w:del>
            deleted = OxmlElement("w:del")
            deleted.set(qn("w:id"), "2")
            deleted.set(qn("w:author"), "Author")
            r_del = OxmlElement("w:r")
            t_del = OxmlElement("w:delText")
            t_del.text = "old removed"
            r_del.append(t_del)
            deleted.append(r_del)
            p._p.append(deleted)

            path = folder / filename
            doc.save(str(path))
            return path

        ls_docx = create_doc_with_track_changes("LS_Test.docx")
        cr_docx = create_doc_with_track_changes("CR_Test.docx")

        # 1. Convert with Type: "LS in"
        ls_out = folder / "ls_out"
        ls_res = convert_document_to_markdown(
            ls_docx, out_dir=ls_out, tdoc_id="S2-2400010",
            meta={"Type": "LS in", "Title": "Incoming Liaison on 5GS"}
        )
        ls_text = ls_res.text
        # Header version stamp must reflect mode=accept
        self.assertIn("mode=accept", ls_text)
        self.assertIn("(revisions: accept)", ls_text)
        # Must NOT contain markup legend
        self.assertNotIn("Markup legend:", ls_text)
        self.assertNotIn("[INSERTED:", ls_text)
        self.assertNotIn("[DELETED:", ls_text)
        # Plain text must contain the addition cleanly
        self.assertIn("new addition", ls_text)
        # Deleted text must be dropped
        self.assertNotIn("old removed", ls_text)

        # 2. Convert with Type: "CR"
        cr_out = folder / "cr_out"
        cr_res = convert_document_to_markdown(
            cr_docx, out_dir=cr_out, tdoc_id="S2-2400020",
            meta={"Type": "CR", "Title": "Normative Spec Revision"}
        )
        cr_text = cr_res.text
        # Header version stamp must reflect mode=markup
        self.assertIn("mode=markup", cr_text)
        self.assertIn("(revisions: markup)", cr_text)
        # Must contain markup legend
        self.assertIn("Markup legend:", cr_text)
        # Must contain inserted marker
        self.assertIn("[INSERTED: new addition]", cr_text)
        self.assertIn("[DELETED: old removed]", cr_text)

    def test_stale_detection_auto_upgrades_ls_markup_to_accept(self):
        """
        Verify that an LS document previously cached with mode=markup is detected
        as stale when checked with expected_mode=accept, and re-converting updates it to accept mode.
        """
        tdoc_id = "S2-2400030"
        folder = self.meeting_dir / tdoc_id
        folder.mkdir(parents=True, exist_ok=True)
        docx_file = folder / f"{tdoc_id}.docx"
        docx_file.write_bytes(b"dummy_docx_content")
        md_file = folder / f"{tdoc_id}.md"

        # Simulate older export that generated mode=markup for this LS
        md_file.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n"
            f"# {tdoc_id}: LS Title\n\n- **Type:** LS in\n- **Exporter:** tdoc2md v2.1.0 (revisions: markup)\n\n"
            f"> **Markup legend:** `[INSERTED: ...]`\n\nBody",
            encoding="utf-8"
        )

        expected_mode = resolve_revision_mode_for_type("LS in")
        self.assertEqual(expected_mode, "accept")

        # Must be flagged as stale due to mode mismatch!
        self.assertTrue(is_conversion_stale(docx_file, md_file, expected_mode=expected_mode))

        # Re-save with mode=accept
        md_file.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=accept --->\n"
            f"# {tdoc_id}: LS Title\n\n- **Type:** LS in\n- **Exporter:** tdoc2md v2.1.0 (revisions: accept)\n\nBody",
            encoding="utf-8"
        )
        # Now it is up-to-date!
        self.assertFalse(is_conversion_stale(docx_file, md_file, expected_mode=expected_mode))

    def test_file_handler_atomic_download_and_cache_validation(self):
        """Verify TDocFileHandler downloads atomically, verifies archive integrity, and recovers from corrupt files."""
        tdoc_dir = self.meeting_dir / "S2-2400999"
        tdoc_dir.mkdir(parents=True, exist_ok=True)
        zip_path = tdoc_dir / "S2-2400999.zip"

        # 1. Create a corrupt (0-byte) zip file simulating an aborted download
        zip_path.write_bytes(b"")
        self.assertTrue(zip_path.exists())

        # Build a valid in-memory zip payload for the mock response
        import io
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as z:
            z.writestr("S2-2400999.docx", b"dummy_docx_content")
        valid_zip_bytes = zip_buffer.getvalue()

        mock_resp = MagicMock()
        mock_resp.iter_content = MagicMock(return_value=[valid_zip_bytes])
        mock_resp.raise_for_status = MagicMock()

        with patch("core.network.session.NetworkSession.get_instance") as mock_get_inst:
            mock_session = MagicMock()
            mock_session.get.return_value = mock_resp
            mock_get_inst.return_value = mock_session

            # Call download_and_extract_tdoc; it must detect corrupt zip, delete it, and download fresh
            extracted = TDocFileHandler.download_and_extract_tdoc(
                "S2-2400999", "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/Docs", tdoc_dir
            )
            self.assertTrue(zip_path.exists())
            self.assertGreater(zip_path.stat().st_size, 0)
            self.assertTrue(zipfile.is_zipfile(to_extended_path(zip_path)))
            self.assertEqual(len(extracted), 1)
            self.assertTrue(extracted[0].exists())

    def test_file_handler_cancellation_check(self):
        """Verify TDocFileHandler aborts immediately when cancellation is requested."""
        tdoc_dir = self.meeting_dir / "S2-2400888"
        tdoc_dir.mkdir(parents=True, exist_ok=True)

        with self.assertRaises(InterruptedError):
            TDocFileHandler.download_and_extract_tdoc(
                "S2-2400888",
                "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/Docs",
                tdoc_dir,
                cancel_check=lambda: True,
            )

    def test_exporter_parallel_predownload_phase(self):
        """
        Verify that TDocMarkdownExporterThread pre-downloads only missing TDocs in parallel,
        skips already cached ones, and properly converts them.
        """
        # 1. Cached TDoc (already has docx and md)
        cached_tdoc = "S2-2400010"
        cached_folder = self.meeting_dir / cached_tdoc
        cached_folder.mkdir(parents=True, exist_ok=True)
        (cached_folder / f"{cached_tdoc}.docx").write_bytes(b"cached_content")
        (cached_folder / f"{cached_tdoc}.md").write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n# {cached_tdoc}\nCached text",
            encoding="utf-8"
        )

        # 2. Missing TDocs
        missing_tdoc1 = "S2-2400011"
        missing_tdoc2 = "S2-2400012"

        tdocs_list = [
            {"TDoc": cached_tdoc, "Title": "Cached TDoc", "Type": "CR"},
            {"TDoc": missing_tdoc1, "Title": "Missing TDoc 1", "Type": "Discussion"},
            {"TDoc": missing_tdoc2, "Title": "Missing TDoc 2", "Type": "Discussion"},
        ]

        exporter = TDocMarkdownExporterThread(
            meeting_dir=self.meeting_dir,
            tdocs_list=tdocs_list,
            docs_ftp_url="https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/Docs",
            revisions_url="https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/INBOX/Revisions",
            is_bulk=True,
        )

        downloaded_tdocs = []

        def fake_download(tdoc_id, folder):
            downloaded_tdocs.append(tdoc_id)
            folder.mkdir(parents=True, exist_ok=True)
            doc_file = folder / f"{tdoc_id}.docx"
            doc_file.write_bytes(b"downloaded_content")
            # Write a fresh md file so convert_document_to_markdown isn't doing full docx parsing on dummy bytes
            md_file = folder / f"{tdoc_id}.md"
            md_file.write_text(
                f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=accept --->\n# {tdoc_id}\nFresh text",
                encoding="utf-8"
            )
            return doc_file

        with patch.object(exporter, "_download_and_extract_tdoc", side_effect=fake_download):
            progress_messages = []
            exporter.progress.connect(progress_messages.append)
            exporter.run()

        # Cached TDoc must NOT have been scheduled for download
        self.assertNotIn(cached_tdoc, downloaded_tdocs)
        # Missing TDocs must have been downloaded
        self.assertIn(missing_tdoc1, downloaded_tdocs)
        self.assertIn(missing_tdoc2, downloaded_tdocs)

        # Progress messages should show parallel download and conversion
        joined_progress = " ".join(progress_messages)
        self.assertIn("in parallel", joined_progress)
        self.assertIn("Converting", joined_progress)


if __name__ == "__main__":
    unittest.main()
