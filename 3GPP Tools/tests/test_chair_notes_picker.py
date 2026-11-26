# --- File: tests/test_chair_notes_picker.py ---
"""
Unit and UI test suite for multi-chair / multi-track Chairman's Notes classification,
the interactive ChairNotesPickerDialog, and Chairman's Notes download flows across SA and RAN.
"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtWidgets import QApplication

from modules.meetings.core.chair_notes_classifier import (
    analyze_candidate_files,
    classify_chair_track,
    compute_recency_score,
)
from modules.meetings.core.chair_notes_downloader import (
    ChairNotesDownloaderThread,
    ChairNotesListingThread,
)
from modules.meetings.ui.chair_notes_picker_dialog import ChairNotesPickerDialog
from modules.meetings.ui.tdocs_window import TDocsWindow


class TestChairNotesPicker(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_classify_chair_track(self):
        """Verifies chair/track classification across SA2 per-chair files, sessions, and agendas."""
        # SA2 per-chair track files
        lbl, tid = classify_chair_track("ChairNotes_Andy_v01.docx")
        self.assertEqual(lbl, "Chair: Andy")
        self.assertEqual(tid, "chair_andy")

        lbl, tid = classify_chair_track("ChairNotes_hector_v02.docx")
        self.assertEqual(lbl, "Chair: Hector")
        self.assertEqual(tid, "chair_hector")

        lbl, tid = classify_chair_track("ChairNotes_Ruicheng_v01.docx")
        self.assertEqual(lbl, "Chair: Ruicheng")
        self.assertEqual(tid, "chair_ruicheng")

        lbl, tid = classify_chair_track("ChairNotes_Jian_v02.docx")
        self.assertEqual(lbl, "Chair: Jian")
        self.assertEqual(tid, "chair_jian")

        # Session / Breakout track
        lbl, tid = classify_chair_track("R2_135b_Chair_Notes_sessionA_v01.docx")
        self.assertEqual(lbl, "Session A")
        self.assertEqual(tid, "session_a")

        # Working Group Agenda Prefix (RAN3)
        lbl, tid = classify_chair_track("RAN3_133bis_agenda_20261016_0900.docx")
        self.assertEqual(lbl, "Agenda: RAN3_133BIS_AGENDA")
        self.assertEqual(tid, "ag_ran3_133bis_agenda")

        # TDoc ID
        lbl, tid = classify_chair_track("S2-2612345_ChairNotes.docx")
        self.assertEqual(lbl, "TDoc S2-2612345")
        self.assertEqual(tid, "tdoc_s2-2612345")

        # Consolidated / Final / Versioned
        lbl, tid = classify_chair_track("ChairNotes_final.docx")
        self.assertEqual(lbl, "Consolidated / Main")
        self.assertEqual(tid, "main")

        lbl, tid = classify_chair_track("ChairNotes_v02.docx")
        self.assertEqual(lbl, "Consolidated / Main")
        self.assertEqual(tid, "main")

    def test_compute_recency_score(self):
        """Verifies recency scoring properly orders versions, EOM/final markers, and dates."""
        score_v1 = compute_recency_score("ChairNotes_Andy_v01.docx")
        score_v2 = compute_recency_score("ChairNotes_Andy_v02.docx")
        score_v3 = compute_recency_score("ChairNotes_Andy_v03.docx")
        self.assertGreater(score_v3, score_v2)
        self.assertGreater(score_v2, score_v1)

        # Final / EOM bonus
        score_final = compute_recency_score("ChairNotes_Andy_final.docx")
        self.assertGreater(score_final, score_v3)

        # ISO Timestamps (e.g. 14:00 vs 09:00)
        score_morning = compute_recency_score("RAN3_133bis_agenda_20261016_0900.docx")
        score_afternoon = compute_recency_score("RAN3_133bis_agenda_20261016_1400.docx")
        score_eom = compute_recency_score("RAN3_133bis_agenda_20261016_EOM.docx")
        self.assertGreater(score_afternoon, score_morning)
        self.assertGreater(score_eom, score_afternoon)

    def test_analyze_candidate_files_pre_picks_latest_per_chair(self):
        """Verifies candidate file analysis pre-picks the latest file per chair and marks downloaded status."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            local_dir = Path(tmp_dir)
            # Create a file that is already downloaded
            (local_dir / "ChairNotes_Andy_v02.docx").touch()

            candidates = [
                ("ChairNotes_Andy_v01.docx", "http://server/Andy_v1"),
                ("ChairNotes_Andy_v02.docx", "http://server/Andy_v2"),
                ("ChairNotes_Andy_v03.docx", "http://server/Andy_v3"),
                ("ChairNotes_hector_v01.docx", "http://server/hector_v1"),
                ("ChairNotes_hector_v02.docx", "http://server/hector_v2"),
                ("ChairNotes_Ruicheng_v01.docx", "http://server/Ruicheng_v1"),
            ]

            analyzed = analyze_candidate_files(candidates, local_dir)
            self.assertEqual(len(analyzed), 6)

            by_name = {rec["filename"]: rec for rec in analyzed}

            # Andy: v03 should be pre-picked; v01 and v02 should NOT be pre-picked
            self.assertTrue(by_name["ChairNotes_Andy_v03.docx"]["is_latest_for_chair"])
            self.assertFalse(by_name["ChairNotes_Andy_v02.docx"]["is_latest_for_chair"])
            self.assertFalse(by_name["ChairNotes_Andy_v01.docx"]["is_latest_for_chair"])

            # Hector: v02 should be pre-picked; v01 should NOT
            self.assertTrue(by_name["ChairNotes_hector_v02.docx"]["is_latest_for_chair"])
            self.assertFalse(by_name["ChairNotes_hector_v01.docx"]["is_latest_for_chair"])

            # Ruicheng: single file should be pre-picked
            self.assertTrue(by_name["ChairNotes_Ruicheng_v01.docx"]["is_latest_for_chair"])

            # Check download status detection
            self.assertTrue(by_name["ChairNotes_Andy_v02.docx"]["is_already_downloaded"])
            self.assertFalse(by_name["ChairNotes_Andy_v03.docx"]["is_already_downloaded"])
            self.assertFalse(by_name["ChairNotes_hector_v02.docx"]["is_already_downloaded"])

    def test_chair_notes_picker_dialog_ui(self):
        """Verifies ChairNotesPickerDialog renders table rows, pre-picks latest, and handles filter actions."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            local_dir = Path(tmp_dir)
            (local_dir / "ChairNotes_Andy_v02.docx").touch()

            candidates = [
                ("ChairNotes_Andy_v01.docx", "http://server/Andy_v1"),
                ("ChairNotes_Andy_v02.docx", "http://server/Andy_v2"),
                ("ChairNotes_Andy_v03.docx", "http://server/Andy_v3"),
                ("ChairNotes_hector_v01.docx", "http://server/hector_v1"),
                ("ChairNotes_hector_v02.docx", "http://server/hector_v2"),
            ]

            mtg_info = {"wg_name": "SA2", "meeting_number": "177"}
            dlg = ChairNotesPickerDialog(
                parent=None,
                file_tuples=candidates,
                local_dir=local_dir,
                mtg_info=mtg_info,
            )

            try:
                self.assertEqual(dlg.table.rowCount(), 5)

                # Pre-picked items: Andy_v03 and hector_v02 should be checked!
                selected = dlg.get_selected_files()
                selected_names = [f[0] for f in selected]
                self.assertIn("ChairNotes_Andy_v03.docx", selected_names)
                self.assertIn("ChairNotes_hector_v02.docx", selected_names)
                self.assertNotIn("ChairNotes_Andy_v01.docx", selected_names)
                self.assertNotIn("ChairNotes_Andy_v02.docx", selected_names)
                self.assertNotIn("ChairNotes_hector_v01.docx", selected_names)
                self.assertEqual(len(selected), 2)

                # Test 'Select All'
                dlg._select_all()
                self.assertEqual(len(dlg.get_selected_files()), 5)

                # Test 'Clear All'
                dlg._clear_all()
                self.assertEqual(len(dlg.get_selected_files()), 0)

                # Test 'Select Pre-picked (Recommended)'
                dlg._select_recommended()
                selected_rec = [f[0] for f in dlg.get_selected_files()]
                self.assertEqual(sorted(selected_rec), ["ChairNotes_Andy_v03.docx", "ChairNotes_hector_v02.docx"])

                # Test 'Select Un-downloaded Only' (Andy_v02 is downloaded, so other 4 are un-downloaded)
                dlg._select_undownloaded_only()
                selected_new = [f[0] for f in dlg.get_selected_files()]
                self.assertEqual(len(selected_new), 4)
                self.assertNotIn("ChairNotes_Andy_v02.docx", selected_new)

                # Test search filtering
                dlg._filter_table("Hector")
                visible_rows = [r for r in range(dlg.table.rowCount()) if not dlg.table.isRowHidden(r)]
                self.assertEqual(len(visible_rows), 2)

                dlg._filter_table("")
                all_visible = [r for r in range(dlg.table.rowCount()) if not dlg.table.isRowHidden(r)]
                self.assertEqual(len(all_visible), 5)

                # Overwrite flag
                self.assertTrue(dlg.should_overwrite())
                dlg.overwrite_cb.setChecked(False)
                self.assertFalse(dlg.should_overwrite())
            finally:
                dlg.close()

    def test_downloader_thread_with_selected_files(self):
        """Verifies ChairNotesDownloaderThread accepts selected_files and runs download_files_sync."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            local_dir = Path(tmp_dir)
            selected = [
                ("ChairNotes_Andy_v03.docx", "http://server/Andy_v3"),
                ("ChairNotes_hector_v02.docx", "http://server/hector_v2"),
            ]

            thread = ChairNotesDownloaderThread(
                candidate_base_urls=["http://server"],
                target_dir=local_dir,
                selected_files=selected,
            )

            self.assertEqual(thread.selected_files, selected)
            self.assertTrue(thread.overwrite)
            self.assertIsNot(ChairNotesDownloaderThread.download_finished, ChairNotesDownloaderThread.finished)
            self.assertIsNot(ChairNotesListingThread.listing_finished, ChairNotesListingThread.finished)

    def test_download_files_sync_uses_part_file_and_streaming_progress(self):
        """Verifies atomic .part staging, streaming progress updates, and .part cleanup on error."""

        with tempfile.TemporaryDirectory() as tmp_dir:
            target_dir = Path(tmp_dir)
            payload_chunks = [b"A" * 65536, b"B" * 65536]
            total_len = sum(len(c) for c in payload_chunks)

            mock_resp = MagicMock()
            mock_resp.__enter__ = MagicMock(return_value=mock_resp)
            mock_resp.__exit__ = MagicMock(return_value=False)
            mock_resp.headers = {"Content-Length": str(total_len)}
            mock_resp.iter_content.return_value = payload_chunks

            mock_session = MagicMock()
            mock_session.get.return_value = mock_resp

            progress_msgs = []
            ticks = iter([10.0, 10.3, 10.6])
            with patch("modules.meetings.core.chair_notes_downloader.time.monotonic", side_effect=lambda: next(ticks, 11.0)):
                downloaded = ChairNotesDownloaderThread.download_files_sync(
                    mock_session,
                    [("Combined_ChairNotes.doc", "https://example.org/Combined_ChairNotes.doc")],
                    target_dir,
                    progress_callback=progress_msgs.append,
                    overwrite=True,
                )

            self.assertEqual(downloaded, ["Combined_ChairNotes.doc"])
            self.assertTrue((target_dir / "Combined_ChairNotes.doc").exists())
            self.assertFalse((target_dir / "Combined_ChairNotes.doc.part").exists())
            self.assertTrue(any("%" in m and "KB" in m for m in progress_msgs), f"Expected byte progress in {progress_msgs}")

            # Verify .part is cleaned up if stream fails mid-download
            failing_resp = MagicMock()
            failing_resp.__enter__ = MagicMock(return_value=failing_resp)
            failing_resp.__exit__ = MagicMock(return_value=False)
            failing_resp.headers = {"Content-Length": "2048"}

            def _boom(chunk_size=65536):
                yield b"partial"
                raise RuntimeError("connection reset")

            failing_resp.iter_content.side_effect = _boom
            mock_session.get.return_value = failing_resp

            failed_dl = ChairNotesDownloaderThread.download_files_sync(
                mock_session,
                [("Broken_Notes.doc", "https://example.org/Broken_Notes.doc")],
                target_dir,
                overwrite=True,
            )
            self.assertEqual(failed_dl, [])
            self.assertFalse((target_dir / "Broken_Notes.doc").exists())
            self.assertFalse((target_dir / "Broken_Notes.doc.part").exists())

    def test_controller_refresh_btn_unlock_and_seq_guard(self):
        """Verifies on_chair_notes_finished & safety fallback re-enable refresh_btn and guard stale timers."""
        from PyQt5.QtWidgets import QMainWindow, QPushButton
        from modules.meetings.ui.tdocs_agenda_controller import (
            _ensure_refresh_btn_unlocked,
            _restore_refresh_btn,
            _set_refresh_btn_busy,
            on_chair_notes_finished,
        )

        win = QMainWindow()
        win.meeting_dir = Path(".")
        win.refresh_btn = QPushButton("🔄 Refresh", win)
        win.refresh_btn.setToolTip("Default Tooltip")

        try:
            seq1 = _set_refresh_btn_busy(win, "Downloading Notes...")
            self.assertFalse(win.refresh_btn.isEnabled())

            # Simulate abnormal thread exit without on_chair_notes_finished: safety fallback unlocks
            _ensure_refresh_btn_unlocked(win, seq1)
            self.assertTrue(win.refresh_btn.isEnabled())
            self.assertEqual(win.refresh_btn.text(), "🔄 Refresh")
            self.assertEqual(win.refresh_btn.toolTip(), "Default Tooltip")

            # Simulate normal finish followed immediately by a new busy operation before 4s timer fires
            scheduled_callbacks = []
            with patch("modules.meetings.ui.tdocs_agenda_controller.QMessageBox") as mock_mb, \
                 patch("modules.meetings.ui.tdocs_agenda_controller.QTimer.singleShot", side_effect=lambda ms, fn: scheduled_callbacks.append(fn)):
                mock_mb.return_value.clickedButton.return_value = None
                _set_refresh_btn_busy(win, "Downloading Notes...")
                on_chair_notes_finished(win, True, 1, ["Combined.doc"], "ok")

            self.assertTrue(win.refresh_btn.isEnabled())
            self.assertEqual(win.refresh_btn.text(), "✅ 1 Notes")
            self.assertEqual(len(scheduled_callbacks), 1)

            # Start another operation before the 4s timer fires
            _set_refresh_btn_busy(win, "Getting Agenda...")
            scheduled_callbacks[0]()  # Stale 4s callback must NOT clobber the new busy label!
            self.assertFalse(win.refresh_btn.isEnabled())
            self.assertEqual(win.refresh_btn.text(), "⏳ Getting Agenda...")

            _restore_refresh_btn(win)
        finally:
            win.close()

    def test_sa2_window_chairnotes_resources_menu(self):
        """Verifies SA2 TDocsWindow populates Local ChairNotes folder and Open Notes action when notes exist."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            meeting_dir = Path(tmp_dir)
            notes_dir = meeting_dir / "Agenda" / "ChairNotes"
            notes_dir.mkdir(parents=True, exist_ok=True)
            note_file = notes_dir / "ChairNotes_Andy_v03.docx"
            note_file.touch()

            agenda_list = meeting_dir / "Agenda" / "TDocs_List.xlsx"
            agenda_list.touch()

            sa_info = {"wg_name": "SA2", "meeting_number": "177", "url_key": ""}
            win_sa = TDocsWindow(sa_info, tdocs_data=[], filepath=str(agenda_list))
            try:
                # Trigger resources menu population
                win_sa._populate_resources_menu()
                action_texts = [a.text() for a in win_sa.folder_menu.actions()]

                # Verify ChairNotes Folder and Open Notes are present in SA2!
                self.assertTrue(any("Local: ChairNotes Folder" in t for t in action_texts))
                self.assertTrue(any("Open Notes: ChairNotes_Andy_v03.docx" in t for t in action_texts))
            finally:
                win_sa.close()


if __name__ == "__main__":
    unittest.main()
