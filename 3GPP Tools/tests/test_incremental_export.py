# --- File: tests/test_incremental_export.py ---
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

# Ensure src is in sys.path
_src_dir = Path(__file__).resolve().parent.parent / "src"
if str(_src_dir) not in sys.path:
    sys.path.insert(0, str(_src_dir))

from PyQt5.QtWidgets import QApplication

from modules.meetings.core.exporters.pre_meeting_exporter import PreMeetingExporterThread
from modules.meetings.core.exporters.post_meeting_exporter import PostMeetingExporterThread
from modules.meetings.core.exporters.export_manifest import (
    load_manifest,
    load_or_reconstruct_manifest,
    plan_incremental_session,
    merge_manifest_into_tree,
    save_manifest,
    ExportManifest,
    MANIFEST_FILENAME,
)


class TestIncrementalExport(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="3gpp_test_incremental_")
        self.meeting_dir = Path(self.temp_dir)
        self.export_dir = self.meeting_dir / "Export" / "PreMeeting_SA2_178_Test"
        self.mtg_info = {"wg_name": "SA2", "meeting_number": "178"}

        # Realistic TDoc rows
        self.all_rows = [
            {
                "TDoc": "S2-2611001",
                "Type": "pCR",
                "TDoc Status": "Available",
                "Related WIs": "FS_5G_AIML_Ph2",
                "Agenda Item": "18.2",
                "Agenda item description": "AI/ML Phase 2",
                "Release": "Rel-19",
                "Title": "KI#3 Solution for Model Distribution",
                "Source": "Qualcomm",
                "_Sanitized_Companies": ["Qualcomm"],
                "_All_Sanitized_Companies": ["Qualcomm"],
            },
            {
                "TDoc": "S2-2611002",
                "Type": "pCR",
                "TDoc Status": "Available",
                "Related WIs": "FS_5G_AIML_Ph2",
                "Agenda Item": "18.2",
                "Agenda item description": "AI/ML Phase 2",
                "Release": "Rel-19",
                "Title": "KI#3 Solution Evaluation",
                "Source": "Ericsson",
                "_Sanitized_Companies": ["Ericsson"],
                "_All_Sanitized_Companies": ["Ericsson"],
            },
            {
                "TDoc": "S2-2611010",
                "Type": "CR",
                "TDoc Status": "Available",
                "Related WIs": "TEI18",
                "Agenda Item": "5.1",
                "Agenda item description": "Architecture Solutions",
                "Release": "Rel-18",
                "Title": "Core Network Architecture Corrections",
                "Source": "Deutsche Telekom",
                "_Sanitized_Companies": ["Deutsche Telekom"],
                "_All_Sanitized_Companies": ["Deutsche Telekom"],
            },
            {
                "TDoc": "S2-2611020",
                "Type": "Discussion",
                "TDoc Status": "Available",
                "Related WIs": "FS_5G_AIML_Ph2",
                "Agenda Item": "18.2",
                "Agenda item description": "AI/ML Phase 2",
                "Release": "Rel-19",
                "Title": "Discussion on Training Inference Latencies",
                "Source": "Nokia",
                "_Sanitized_Companies": ["Nokia"],
                "_All_Sanitized_Companies": ["Nokia"],
            },
            {
                "TDoc": "S2-2611030",
                "Type": "pCR",
                "TDoc Status": "Available",
                "Related WIs": "5GSAT_Ph3",
                "Agenda Item": "19.1",
                "Agenda item description": "Satellite Access Phase 3",
                "Release": "Rel-19",
                "Title": "Satellite Routing Optimization",
                "Source": "Thales",
                "_Sanitized_Companies": ["Thales"],
                "_All_Sanitized_Companies": ["Thales"],
            },
        ]

        # Create dummy source word files in meeting cache so no network is needed
        for r in self.all_rows:
            tid = r["TDoc"]
            tdoc_folder = self.meeting_dir / tid
            tdoc_folder.mkdir(parents=True, exist_ok=True)
            dummy_md = tdoc_folder / f"{tid}.md"
            dummy_md.write_text(f"# {tid}: {r['Title']}\n\nDummy Markdown content for testing.", encoding="utf-8")

    def tearDown(self):
        try:
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        except Exception:
            pass

    def test_initial_and_incremental_export_flow(self):
        """Tests initial export followed by an incremental export adding new TDocs."""
        # 1. Initial Export: Export only TEI18 and S2-2611001 (custom selection)
        initial_wis = {"tei18"}
        exporter_init = PreMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=self.all_rows,
            visible_rows=self.all_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="custom",
            custom_wis=initial_wis,
            download_missing=False,
            output_dir=self.export_dir,
        )

        success_results = []
        exporter_init.finished.connect(lambda ok, msg: success_results.append((ok, msg)))
        exporter_init.run()

        self.assertTrue(success_results[0][0], f"Initial export failed: {success_results[0][1]}")
        self.assertTrue(self.export_dir.is_dir())

        # Check manifest created
        manifest_1 = load_manifest(self.export_dir)
        self.assertIsNotNone(manifest_1)
        self.assertEqual(len(manifest_1.sessions), 1)
        self.assertEqual(manifest_1.sessions[0]["session_id"], 1)
        self.assertIn("S2-2611010", manifest_1.tdocs)
        self.assertEqual(len(manifest_1.tdocs), 1)

        # Check master index.md
        master_idx = self.export_dir / "index.md"
        self.assertTrue(master_idx.is_file())
        master_content = master_idx.read_text(encoding="utf-8")
        self.assertIn("S2-2611010", master_content)
        self.assertIn("Export Changelog & LLM Incremental Update Guide", master_content)
        self.assertIn("Initial Export", master_content)

        # 2. Incremental Export: Re-run with Whole Meeting scope targeting the SAME directory
        exporter_inc = PreMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=self.all_rows,
            visible_rows=self.all_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="whole",
            download_missing=False,
            output_dir=self.export_dir,
        )

        inc_results = []
        exporter_inc.finished.connect(lambda ok, msg: inc_results.append((ok, msg)))
        exporter_inc.run()

        self.assertTrue(inc_results[0][0], f"Incremental export failed: {inc_results[0][1]}")

        # Check manifest updated with session #2
        manifest_2 = load_manifest(self.export_dir)
        self.assertIsNotNone(manifest_2)
        self.assertEqual(len(manifest_2.sessions), 2)
        session_2 = manifest_2.sessions[1]
        self.assertEqual(session_2["session_id"], 2)
        self.assertIn("Incremental Update #1", session_2["label"])
        # All 5 TDocs should now be registered
        self.assertEqual(len(manifest_2.tdocs), 5)
        self.assertIn("S2-2611001", session_2["added_tdocs"])
        self.assertIn("S2-2611002", session_2["added_tdocs"])
        self.assertIn("S2-2611020", session_2["added_tdocs"])
        self.assertIn("S2-2611030", session_2["added_tdocs"])
        self.assertNotIn("S2-2611010", session_2["added_tdocs"])  # was already in session 1

        # Check master index.md contains both old and new TDocs
        master_content_2 = master_idx.read_text(encoding="utf-8")
        self.assertIn("S2-2611010", master_content_2)
        self.assertIn("S2-2611001", master_content_2)
        self.assertIn("S2-2611002", master_content_2)
        self.assertIn("S2-2611020", master_content_2)
        self.assertIn("S2-2611030", master_content_2)

        # Check Changelog table in index.md
        self.assertIn("Newly Added TDocs in Latest Update (+4)", master_content_2)
        self.assertIn("KI#3 Solution for Model Distribution", master_content_2)
        self.assertIn("Suggested Instruction for LLM Analysis:", master_content_2)

        # Check incremental header notice inside newly added file
        new_md_file = self.export_dir / "Rel-19" / "FS_5G_AIML_Ph2" / "pCR" / "S2-2611001.md"
        self.assertTrue(new_md_file.is_file())
        new_content = new_md_file.read_text(encoding="utf-8")
        self.assertIn("Incremental Addition:", new_content)
        self.assertIn("Incremental Update #1", new_content)

    def test_corpus_preservation_with_narrow_scope(self):
        """
        Tests that running an incremental export with a narrow filter (e.g. only 1 new WI)
        does NOT erase previously exported WIs from master and local index files.
        """
        # Session 1: Export TEI18 (S2-2611010)
        exp1 = PreMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=self.all_rows,
            visible_rows=self.all_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="custom",
            custom_wis={"tei18"},
            download_missing=False,
            output_dir=self.export_dir,
        )
        exp1.run()

        # Session 2: Export ONLY 5GSAT_Ph3 (S2-2611030) targeting same folder
        exp2 = PreMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=self.all_rows,
            visible_rows=self.all_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="custom",
            custom_wis={"5gsat_ph3"},
            download_missing=False,
            output_dir=self.export_dir,
        )
        exp2.run()

        # Master index MUST contain BOTH TEI18 (S2-2611010) and 5GSAT_Ph3 (S2-2611030)
        master_content = (self.export_dir / "index.md").read_text(encoding="utf-8")
        self.assertIn("S2-2611010", master_content, "Previous TEI18 TDoc was accidentally lost from master index!")
        self.assertIn("S2-2611030", master_content, "New 5GSAT TDoc missing from master index!")

        manifest = load_manifest(self.export_dir)
        self.assertEqual(len(manifest.tdocs), 2)
        self.assertIn("S2-2611010", manifest.tdocs)
        self.assertIn("S2-2611030", manifest.tdocs)

    def test_reconstruction_from_legacy_export_directory(self):
        """
        Tests that an export folder created without .export_manifest.json is safely
        reconstructed on the fly and subsequent incremental exports succeed.
        """
        # Session 1 export
        exp1 = PreMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=self.all_rows,
            visible_rows=self.all_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="custom",
            custom_wis={"tei18"},
            download_missing=False,
            output_dir=self.export_dir,
        )
        exp1.run()

        # Simulate legacy export: delete .export_manifest.json
        manifest_file = self.export_dir / MANIFEST_FILENAME
        self.assertTrue(manifest_file.is_file())
        manifest_file.unlink()
        self.assertFalse(manifest_file.is_file())

        # Test reconstruction
        reconstructed = load_or_reconstruct_manifest(self.export_dir, self.mtg_info, self.all_rows)
        self.assertIsNotNone(reconstructed)
        self.assertIn("S2-2611010", reconstructed.tdocs)
        self.assertTrue(manifest_file.is_file(), "Manifest should have been recreated on disk.")

        # Now run an incremental export on this reconstructed folder
        exp2 = PreMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=self.all_rows,
            visible_rows=self.all_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="custom",
            custom_wis={"5gsat_ph3"},
            download_missing=False,
            output_dir=self.export_dir,
        )
        exp2.run()

        manifest_after = load_manifest(self.export_dir)
        self.assertEqual(len(manifest_after.sessions), 2)
        self.assertIn("S2-2611010", manifest_after.tdocs)
        self.assertIn("S2-2611030", manifest_after.tdocs)

    def test_post_meeting_incremental_export(self):
        """Tests incremental update on PostMeetingExporterThread."""
        post_dir = self.meeting_dir / "Export" / "PostMeeting_SA2_178_Test"

        # Set status to Agreed for post-meeting requirements
        for r in self.all_rows:
            r["TDoc Status"] = "Agreed"

        # Initial post-meeting run with TEI18
        exp1 = PostMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=self.all_rows,
            visible_rows=self.all_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="custom",
            custom_wis={"tei18"},
            include_disc=True,
            download_missing=False,
            output_dir=post_dir,
        )
        exp1.run()

        self.assertTrue((post_dir / "index.md").is_file())
        manifest_1 = load_manifest(post_dir)
        self.assertIsNotNone(manifest_1)
        self.assertEqual(manifest_1.export_type, "post_meeting")
        self.assertIn("S2-2611010", manifest_1.tdocs)

        # Incremental post-meeting run adding AIML
        exp2 = PostMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=self.all_rows,
            visible_rows=self.all_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="custom",
            custom_wis={"fs_5g_aiml_ph2"},
            include_disc=True,
            download_missing=False,
            output_dir=post_dir,
        )
        exp2.run()

        manifest_2 = load_manifest(post_dir)
        self.assertEqual(len(manifest_2.sessions), 2)
        master_content = (post_dir / "index.md").read_text(encoding="utf-8")
        self.assertIn("S2-2611010", master_content)
        self.assertIn("S2-2611001", master_content)
        self.assertIn("Export Changelog & LLM Incremental Update Guide", master_content)

    def test_reserved_stub_upgraded_when_docx_becomes_available(self):
        """
        Tests that an initially unuploaded/reserved TDoc exported as a metadata-only stub
        is properly detected as incomplete and upgraded to full Markdown content when
        its docx document becomes available in an incremental export run.
        """
        # S2-2611493 starts as Reserved and NOT present on disk
        reserved_row = {
            "TDoc": "S2-2611493",
            "Type": "pCR",
            "TDoc Status": "Reserved",
            "Related WIs": "FS_5G_AIML_Ph2",
            "Agenda Item": "18.2",
            "Agenda item description": "AI/ML Phase 2",
            "Release": "Rel-19",
            "Title": "Model Checkpointing and Verification",
            "Source": "Samsung",
            "_Sanitized_Companies": ["Samsung"],
            "_All_Sanitized_Companies": ["Samsung"],
        }
        test_rows = list(self.all_rows) + [reserved_row]

        # Ensure no docx exists for S2-2611493 initially
        s2_folder = self.meeting_dir / "S2-2611493"
        if s2_folder.exists():
            shutil.rmtree(s2_folder)

        # 1. Run Initial Export with S2-2611493 reserved
        exp1 = PreMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=test_rows,
            visible_rows=test_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="whole",
            download_missing=False,
            output_dir=self.export_dir,
        )
        exp1.run()

        stub_path = self.export_dir / "Rel-19" / "FS_5G_AIML_Ph2" / "pCR" / "S2-2611493.md"
        self.assertTrue(stub_path.is_file(), "Reserved stub should have been created on initial export.")
        stub_content = stub_path.read_text(encoding="utf-8")
        self.assertIn("Status: Reserved / Not Uploaded", stub_content)

        # Master index should show reserved status
        master_content = (self.export_dir / "index.md").read_text(encoding="utf-8")
        self.assertIn("⏳ Reserved", master_content)

        # 2. Later: The author uploads the docx and it is cached locally in the meeting folder!
        s2_folder.mkdir(parents=True, exist_ok=True)
        docx_file = s2_folder / "S2-2611493.docx"
        import zipfile
        with zipfile.ZipFile(docx_file, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types/>")
        # Also create dummy converted md in cache so _ensure_tdoc_markdown finds it
        cache_md = s2_folder / "S2-2611493.md"
        cache_md.write_text("# S2-2611493: Model Checkpointing\n\nFull technical specification text...", encoding="utf-8")

        # 3. Run Incremental Export targeting the SAME export directory
        exp2 = PreMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=test_rows,
            visible_rows=test_rows,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="",
            revisions_url="",
            scope_mode="whole",
            download_missing=False,
            output_dir=self.export_dir,
        )
        exp2.run()

        # The stub file MUST have been replaced with the full technical text!
        updated_content = stub_path.read_text(encoding="utf-8")
        self.assertNotIn("Status: Reserved / Not Uploaded", updated_content, "Reserved stub was not upgraded!")
        self.assertIn("Full technical specification text", updated_content)

        # Master index must no longer mark S2-2611493 as ⏳ Reserved!
        master_updated = (self.export_dir / "index.md").read_text(encoding="utf-8")
        self.assertIn("Model Checkpointing", master_updated)
        self.assertIn("Updated / Re-converted TDocs", master_updated)


if __name__ == "__main__":
    unittest.main()
