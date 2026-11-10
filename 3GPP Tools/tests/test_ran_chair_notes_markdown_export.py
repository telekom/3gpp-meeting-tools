# --- File: tests/test_ran_chair_notes_markdown_export.py ---
"""
Comprehensive tests for RAN meeting Chairman's Notes agreements extraction,
live scanning for running meetings, and inclusion in Markdown exports (Post-Meeting & Summary Reports).
"""
import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

# Ensure src is on sys.path
_root = Path(__file__).resolve().parent.parent
_src = _root / "src"
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import docx
from PyQt5.QtWidgets import QApplication

from modules.meetings.core.exporters.markdown_exporter import MarkdownExporterThread, natural_ai_sort_key
from modules.meetings.core.exporters.post_meeting_exporter import PostMeetingExporterThread
from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser
from modules.meetings.ui.dialogs.post_meeting_export_dialog import PostMeetingExportDialog


class TestRanChairNotesMarkdownExport(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.meeting_dir = Path(self.temp_dir.name)
        self.agenda_dir = self.meeting_dir / "Agenda"
        self.chair_notes_dir = self.agenda_dir / "ChairNotes"
        os.makedirs(self.chair_notes_dir, exist_ok=True)

        self.mtg_info = {
            "wg_name": "RAN2",
            "meeting_number": "135bis",
            "start_date": "2026-10-12",
        }

    def tearDown(self):
        self.temp_dir.cleanup()

    def _create_sample_ran2_notes_docx(self, path: Path) -> Path:
        """Creates a sample RAN2 block-based notes document with agreements."""
        doc = docx.Document()
        doc.add_heading("9.3.1.1\tUser Plane Radio Protocols", level=3)
        doc.add_paragraph("(FS_6G_Radio, leading WG: RAN2; REL-20; SID: RP-253876)")

        doc.add_paragraph("R2-2607326\t6G L2 Architecture\tDeutsche Telekom\tdiscussion\tagreement\tRel-20\tFS_6G_Radio")
        doc.add_paragraph("Noted")

        # Boxed table agreement
        tbl = doc.add_table(rows=1, cols=1)
        tbl.cell(0, 0).text = (
            "Agreements\n"
            "1. PDCP and RLC layers are kept separated in 6GR.\n"
            "2. Support selective out-of-order delivery at PDCP receiver."
        )

        doc.add_heading("9.3.1.2\tControl Plane Architecture", level=3)
        doc.add_paragraph("(FS_6G_Radio, leading WG: RAN2; REL-20; SID: RP-253876)")
        doc.add_paragraph("R2-2607674\t6G UL Data Burst Reporting\tHuawei\tdiscussion\tagreement\tRel-20\tFS_6G_Radio")
        doc.add_paragraph("Agreements")
        doc.add_paragraph("- The type 1 burst reporting is realized by BSR/DSR.")
        doc.add_paragraph("- Burst information of non-AS buffer data is FFS.")

        doc.save(str(path))
        return path

    def test_scan_and_cache_chair_notes(self):
        """Verifies scan_and_cache_chair_notes discovers files, parses agreements, and caches them."""
        notes_file = self.chair_notes_dir / "R2_135b_Chair_Notes_clean.docx"
        self._create_sample_ran2_notes_docx(notes_file)

        cache_file = self.agenda_dir / "chair_notes_cache.json"
        self.assertFalse(cache_file.exists())

        # First scan: should parse and create cache
        agenda_data, agreements, files_scanned = ChairNotesAgreementsParser.scan_and_cache_chair_notes(
            self.meeting_dir, mtg_info=self.mtg_info
        )

        self.assertEqual(files_scanned, 1)
        self.assertIn("R2-2607326", agenda_data)
        self.assertIn("R2-2607674", agenda_data)
        self.assertGreaterEqual(len(agreements), 2)
        self.assertTrue(cache_file.exists())

        # Verify cache content
        with open(cache_file, "r", encoding="utf-8") as f:
            cache_payload = json.load(f)
        self.assertIn("agenda_data", cache_payload)
        self.assertIn("agreements", cache_payload)
        self.assertEqual(len(cache_payload["agreements"]), len(agreements))

        # Second scan: cache is fresh, should return cached data without re-reading
        agenda_data2, agreements2, files_scanned2 = ChairNotesAgreementsParser.scan_and_cache_chair_notes(
            self.meeting_dir, mtg_info=self.mtg_info, force_rescan=False
        )
        self.assertEqual(len(agreements2), len(agreements))

    def test_markdown_exporter_includes_ran_agreements_in_ai_and_master(self):
        """
        Verifies that MarkdownExporterThread exports agreements into:
        1. Respective Agenda Item .md files (e.g. 9.3.1.1.md)
        2. Master Agreements.md with full statistics and callouts
        3. Company.md if company contributions resulted in agreements
        """
        notes_file = self.chair_notes_dir / "R2_135b_Chair_Notes_clean.docx"
        self._create_sample_ran2_notes_docx(notes_file)

        tdocs_data = [
            {
                "TDoc": "R2-2607326",
                "Title": "6G L2 Architecture",
                "Source": "Deutsche Telekom",
                "Type": "discussion",
                "For": "agreement",
                "Agenda Item": "9.3.1.1",
                "TDoc Status": "agreed",
            },
            {
                "TDoc": "R2-2607674",
                "Title": "6G UL Data Burst Reporting",
                "Source": "Huawei",
                "Type": "discussion",
                "For": "agreement",
                "Agenda Item": "9.3.1.2",
                "TDoc Status": "agreed",
            },
        ]

        exporter = MarkdownExporterThread(
            meeting_dir=self.meeting_dir,
            tdocs_data=tdocs_data,
            docs_url="https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_135bis/Docs/",
            mtg_info=self.mtg_info,
            include_chair_agreements=True,
        )

        success = False
        msg = ""

        def on_finished(s, m):
            nonlocal success, msg
            success = s
            msg = m

        exporter.finished.connect(on_finished)
        exporter.run()

        self.assertTrue(success, f"Exporter failed: {msg}")

        export_dir = self.meeting_dir / "Export"
        self.assertTrue(export_dir.is_dir())

        # 1. Verify AI 9.3.1.1.md has Chairman's Notes Agreements section
        ai_file = export_dir / "9.3.1.1.md"
        self.assertTrue(ai_file.exists(), "9.3.1.1.md was not created")
        with open(ai_file, "r", encoding="utf-8") as f:
            ai_content = f.read()
        self.assertIn("Following Agreements & Working Assumptions were captured in Chairman's Notes:", ai_content)
        self.assertIn("PDCP and RLC layers are kept separated in 6GR", ai_content)
        self.assertIn("R2-2607326", ai_content)

        # 2. Verify master Agreements.md is created
        master_ag_file = export_dir / "Agreements.md"
        self.assertTrue(master_ag_file.exists(), "Agreements.md was not created")
        with open(master_ag_file, "r", encoding="utf-8") as f:
            master_content = f.read()
        self.assertIn("RAN2#135bis — Chairman's Notes Agreements & Working Assumptions", master_content)
        self.assertIn("Summary:", master_content)
        self.assertIn("Agenda Item 9.3.1.1", master_content)
        self.assertIn("Agenda Item 9.3.1.2", master_content)
        self.assertIn("selective out-of-order delivery", master_content)

        # 3. Verify Company.md mentions company agreement
        company_file = export_dir / "Company.md"
        self.assertTrue(company_file.exists(), "Company.md was not created")
        with open(company_file, "r", encoding="utf-8") as f:
            company_content = f.read()
        self.assertIn("Following Company Contributions resulted in Agreements / Working Assumptions:", company_content)
        self.assertIn("R2-2607326", company_content)

    def test_post_meeting_dialog_ran_scanning_and_badges(self):
        """Verifies PostMeetingExportDialog auto-detects RAN, auto-scans notes, and supports in-dialog scanning."""
        notes_file = self.chair_notes_dir / "R2_135b_Chair_Notes_clean.docx"
        self._create_sample_ran2_notes_docx(notes_file)

        tdocs_data = [
            {
                "TDoc": "R2-2607326",
                "Title": "6G L2 Architecture",
                "Source": "Deutsche Telekom",
                "Type": "discussion",
                "For": "agreement",
                "Agenda Item": "9.3.1.1",
                "TDoc Status": "agreed",
                "Related WIs": "FS_6G_Radio",
                "Release": "Rel-20",
            }
        ]

        dialog = PostMeetingExportDialog(
            parent=None,
            meeting_dir=self.meeting_dir,
            all_rows=tdocs_data,
            visible_rows=tdocs_data,
            agenda_map={},
            mtg_info=self.mtg_info,
            followed_tokens=set(),
            starred_tokens=set(),
        )

        # Auto-scan should have loaded the agreements
        self.assertTrue(dialog.is_ran)
        self.assertTrue(dialog.chk_chair_agreements.isChecked())
        self.assertIn("RAN Text Agreements", dialog.chk_chair_agreements.text())
        self.assertGreaterEqual(len(dialog.chair_agreements), 2)
        self.assertIn("loaded", dialog.lbl_chair_notes_status.text())

        # Test re-scan button trigger with mocked QMessageBox
        from unittest.mock import patch
        with patch("PyQt5.QtWidgets.QMessageBox.information") as mock_info:
            dialog._on_scan_chair_notes_clicked()
            self.assertGreaterEqual(len(dialog.chair_agreements), 2)
            self.assertTrue(mock_info.called)

    def test_post_meeting_exporter_includes_agreements_md_and_ai_wi_inheritance(self):
        """
        Verifies PostMeetingExporterThread associates AI-level agreements with the
        AI's Work Item and generates Agreements.md in the export folder.
        """
        notes_file = self.chair_notes_dir / "R2_135b_Chair_Notes_clean.docx"
        self._create_sample_ran2_notes_docx(notes_file)

        tdocs_data = [
            {
                "TDoc": "R2-2607326",
                "Title": "6G L2 Architecture",
                "Source": "Deutsche Telekom",
                "Type": "CR",
                "For": "agreement",
                "Agenda Item": "9.3.1.1",
                "TDoc Status": "agreed",
                "Related WIs": "FS_6G_Radio",
                "Release": "Rel-20",
            }
        ]

        export_dir = self.meeting_dir / "Export_Post"
        exporter = PostMeetingExporterThread(
            meeting_dir=self.meeting_dir,
            all_rows=tdocs_data,
            visible_rows=tdocs_data,
            agenda_map={},
            mtg_info=self.mtg_info,
            docs_ftp_url="https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_135bis/Docs/",
            revisions_url="",
            scope_mode="whole",
            custom_wis=set(),
            custom_ais=set(),
            include_ls_in=True,
            include_pcr=True,
            include_cr=True,
            include_ls_out=True,
            include_disc=True,
            include_chair_agreements=True,
            include_general_ls=True,
            generate_local_indices=True,
            download_missing=False,
            force_reconvert=False,
            output_dir=export_dir,
        )

        exporter.run()

        # 1. Master index.md should exist and reference Chairman's Notes Agreements
        master_index = export_dir / "index.md"
        self.assertTrue(master_index.exists())
        with open(master_index, "r", encoding="utf-8") as f:
            idx_text = f.read()
        self.assertIn("Chairman's Notes Agreements", idx_text)

        # 2. Standalone Agreements.md should be generated in export_dir
        agreements_file = export_dir / "Agreements.md"
        self.assertTrue(agreements_file.exists(), "Agreements.md was not generated in post-meeting export")
        with open(agreements_file, "r", encoding="utf-8") as f:
            ag_text = f.read()
        self.assertIn("Chairman's Notes Agreements & Working Assumptions", ag_text)
        self.assertIn("PDCP and RLC layers are kept separated", ag_text)

    def test_scan_existing_chair_notes_controller_action(self):
        """Verifies scan_existing_chair_notes in tdocs_agenda_controller scans, merges into model, and reports."""
        from unittest.mock import MagicMock, patch
        from modules.meetings.ui.tdocs_agenda_controller import scan_existing_chair_notes

        notes_file = self.chair_notes_dir / "R2_135b_Chair_Notes_clean.docx"
        self._create_sample_ran2_notes_docx(notes_file)

        mock_window = MagicMock()
        mock_window.meeting_dir = self.meeting_dir
        mock_window.mtg_info = self.mtg_info
        mock_window.agenda_map = {}
        mock_window.model = MagicMock()

        with patch("PyQt5.QtWidgets.QMessageBox.information") as mock_info, \
             patch("PyQt5.QtCore.QTimer.singleShot"):
            scan_existing_chair_notes(mock_window)

            self.assertTrue(mock_window.model.merge_agenda_data.called)
            self.assertTrue(mock_window._refresh_comboboxes.called)
            self.assertTrue(mock_info.called)
            args, _ = mock_info.call_args
            self.assertIn("Agreements & Decisions Extracted", args[2])


if __name__ == "__main__":
    unittest.main()
