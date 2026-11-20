"""Unit tests for workspace window session persistence and restoration."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
for _p in (str(ROOT_DIR), str(SRC_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from PyQt5.QtCore import QDate, QSettings, Qt
from PyQt5.QtGui import QGuiApplication
from PyQt5.QtWidgets import QApplication, QWidget

from core.ui.window_session import (
    RESTORE_MODE_ALWAYS,
    RESTORE_MODE_OFF,
    RESTORE_MODE_RESTART_ONLY,
    apply_window_geometry,
    collect_open_windows_state,
    consume_pending_restart_restore,
    encode_window_geometry,
    get_window_restore_mode,
    load_open_windows_state,
    mark_pending_restart_restore,
    restore_open_windows,
    save_open_windows_state,
    set_window_restore_mode,
    should_restore_windows,
)
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.ui.agreement_tracker_dialog import AgreementTrackerDialog
from modules.meetings.ui.contributions.contribution_dialog import ContributionReportDialog
from modules.meetings.ui.dialogs.topic_dossier_dialog import TopicDossierDialog
from modules.meetings.ui.session_plan.session_plan_window import SessionPlanWindow
from modules.meetings.ui.tdocs_window import TDocsWindow
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.ui.email_window import EmailManagerWindow
from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.spec_search.ui.spec_clause_diff_dialog import SpecClauseDiffDialog
from modules.specifications.core.spec_history_db import SpecHistoryDatabase
from modules.specifications.ui.spec_clause_dialog import SpecClauseExcerptDialog
from modules.specifications.ui.spec_history_dialog import SpecVersionHistoryDialog
from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog


class TestWindowSession(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def _new_settings(self, tmp_dir: str, name: str = "session.ini") -> QSettings:
        return QSettings(str(Path(tmp_dir) / name), QSettings.IniFormat)

    def test_restore_modes_and_escape_hatches(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            cfg = self._new_settings(tmp)
            self.assertEqual(get_window_restore_mode(cfg), RESTORE_MODE_ALWAYS)

            # Invalid value normalizes to always
            set_window_restore_mode("unknown_mode", settings=cfg)
            self.assertEqual(get_window_restore_mode(cfg), RESTORE_MODE_ALWAYS)

            # Always mode restores on normal launch unless --no-restore or Shift held
            self.assertTrue(should_restore_windows(settings=cfg, argv=["main_tools.py"], modifiers=Qt.NoModifier))
            self.assertFalse(
                should_restore_windows(settings=cfg, argv=["main_tools.py", "--no-restore"], modifiers=Qt.NoModifier)
            )
            self.assertFalse(
                should_restore_windows(settings=cfg, argv=["main_tools.py"], modifiers=Qt.ShiftModifier)
            )

            # Restart-only mode requires pending restart flag, and consumes it once
            set_window_restore_mode(RESTORE_MODE_RESTART_ONLY, settings=cfg)
            self.assertEqual(get_window_restore_mode(cfg), RESTORE_MODE_RESTART_ONLY)
            self.assertFalse(should_restore_windows(settings=cfg, argv=["main_tools.py"], modifiers=Qt.NoModifier))

            mark_pending_restart_restore(settings=cfg)
            self.assertTrue(should_restore_windows(settings=cfg, argv=["main_tools.py"], modifiers=Qt.NoModifier))
            # Second cold launch no longer has pending flag
            self.assertFalse(should_restore_windows(settings=cfg, argv=["main_tools.py"], modifiers=Qt.NoModifier))

            # --no-restore still consumes pending restart flag
            mark_pending_restart_restore(settings=cfg)
            self.assertFalse(
                should_restore_windows(settings=cfg, argv=["main_tools.py", "--no-restore"], modifiers=Qt.NoModifier)
            )
            self.assertFalse(consume_pending_restart_restore(settings=cfg))

            # Off mode never restores even after restart
            set_window_restore_mode(RESTORE_MODE_OFF, settings=cfg)
            mark_pending_restart_restore(settings=cfg)
            self.assertFalse(should_restore_windows(settings=cfg, argv=["main_tools.py"], modifiers=Qt.NoModifier))

    def test_geometry_encode_restore_clamps_offscreen(self) -> None:
        w1 = QWidget()
        w1.setGeometry(50000, 50000, 960, 680)
        b64 = encode_window_geometry(w1)
        self.assertTrue(bool(b64))
        w1.close()

        w2 = QWidget()
        apply_window_geometry(w2, b64)
        primary = QGuiApplication.primaryScreen()
        if primary is not None:
            self.assertTrue(primary.availableGeometry().intersects(w2.geometry()))
        w2.close()

    def test_studio_and_spec_dialogs_session_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            cfg = self._new_settings(tmp)

            mdb = MeetingsDatabase(tmp_path / "meetings.db")
            hdb = SpecHistoryDatabase(tmp_path / "history.db")
            sdb = SpecSearchDatabase(tmp_path / "search.db")

            # 1. Open and customize several primary workspace windows
            dossier = TopicDossierDialog(mdb)
            dossier.resize(980, 640)
            dossier.combo_cycle_window.setCurrentIndex(6)
            dossier.edit_date_from.setDate(QDate(2025, 2, 1))
            dossier.edit_date_to.setDate(QDate(2026, 6, 15))
            dossier.chk_discussions.setChecked(False)
            dossier.show()

            tracker = AgreementTrackerDialog(mdb)
            tracker.keyword_edit.setText("paging, LTM")
            tracker.wi_token_widget.set_tokens(["NR_AIML"])
            tracker.ai_token_widget.set_tokens(["8.7"])
            tracker.chk_wi_ai_or.setChecked(True)
            tracker.chk_ffs.setChecked(False)
            tracker.show()

            contrib = ContributionReportDialog(mdb)
            contrib.wi_token_widget.set_tokens(["FS_6G_Radio"])
            contrib.ai_token_widget.set_tokens(["9.3.1"])
            contrib.chk_wi_ai_or.setChecked(True)
            contrib.company_widget.set_selection(["Ericsson"], primary_only=True, include_tentative=True)
            contrib.show()

            history = SpecVersionHistoryDialog(
                spec_number="23.501",
                parent=None,
                history_db=hdb,
                specs_db=None,
            )
            history.search_input.setText("S2-2601234")
            history.clause_input.setText("5.15")
            history.chk_clause_filter.setChecked(True)
            history.chk_origin_ranking.setChecked(True)
            history.show()

            diff = SpecClauseDiffDialog(
                db=sdb,
                spec_number="23.501",
                current_version="18.5.0",
                clause_number="4.2.2",
                clause_title="Registration",
            )
            diff.tab_widget.setCurrentIndex(1)
            diff.focus_combo.setCurrentIndex(2)
            diff.show()

            cr_dlg = CrPcrDialog(
                spec_num="",
                meeting_info={"wg_name": "SA2", "meeting_number": "177", "location": "Prague"},
                initial_mode="cr",
                auto_start=False,
            )
            cr_dlg.input_title.setText("Test Contribution")
            cr_dlg.show()

            fake_meetings_tab = SimpleNamespace(
                db=mdb,
                tdoc_windows={},
                topic_dossier_dialog=dossier,
                agreement_tracker_dialog=tracker,
                contribution_dialog=contrib,
                enable_dates_cb=SimpleNamespace(isChecked=lambda: False),
                wg_filter=SimpleNamespace(
                    getCheckedItems=lambda: [],
                    model=lambda: SimpleNamespace(rowCount=lambda: 1),
                ),
            )
            fake_specs_tab = SimpleNamespace(db=None, history_db=hdb)
            fake_spec_search_tab = SimpleNamespace(db=sdb, cache_dir=tmp_path, specs_db=None)
            fake_main = SimpleNamespace(
                meetings_tab=fake_meetings_tab,
                specs_tab=fake_specs_tab,
                spec_search_tab=fake_spec_search_tab,
            )

            descriptors = collect_open_windows_state(fake_main)
            types = {d["type"] for d in descriptors}
            self.assertEqual(
                types,
                {
                    "topic_dossier",
                    "agreement_tracker",
                    "contribution_report",
                    "spec_history",
                    "spec_clause_diff",
                    "cr_pcr",
                },
            )
            save_open_windows_state(descriptors, settings=cfg)

            for w in (dossier, tracker, contrib, history, diff, cr_dlg):
                w.close()

            # 2. Restore onto fresh fake main window
            restored_meetings_tab = SimpleNamespace(
                db=mdb,
                tdoc_windows={},
                topic_dossier_dialog=None,
                agreement_tracker_dialog=None,
                contribution_dialog=None,
                enable_dates_cb=SimpleNamespace(isChecked=lambda: False),
                wg_filter=SimpleNamespace(
                    getCheckedItems=lambda: [],
                    model=lambda: SimpleNamespace(rowCount=lambda: 1),
                ),
            )
            restored_specs_tab = SimpleNamespace(db=None, history_db=hdb)
            restored_spec_search_tab = SimpleNamespace(db=sdb, cache_dir=tmp_path, specs_db=None)
            restored_main = SimpleNamespace(
                meetings_tab=restored_meetings_tab,
                specs_tab=restored_specs_tab,
                spec_search_tab=restored_spec_search_tab,
            )

            count = restore_open_windows(
                restored_main,
                settings=cfg,
                argv=["main_tools.py"],
                modifiers=Qt.NoModifier,
            )
            self.assertEqual(count, 6)

            r_dossier = restored_meetings_tab.topic_dossier_dialog
            self.assertIsNotNone(r_dossier)
            self.assertEqual(r_dossier.combo_cycle_window.currentIndex(), 6)
            self.assertEqual(r_dossier.edit_date_from.date(), QDate(2025, 2, 1))
            self.assertFalse(r_dossier.chk_discussions.isChecked())

            r_tracker = restored_meetings_tab.agreement_tracker_dialog
            self.assertIsNotNone(r_tracker)
            self.assertEqual(r_tracker.keyword_edit.text(), "paging, LTM")
            self.assertEqual(r_tracker.wi_token_widget.get_tokens(), ["NR_AIML"])
            self.assertEqual(r_tracker.ai_token_widget.get_tokens(), ["8.7"])
            self.assertTrue(r_tracker.chk_wi_ai_or.isChecked())
            self.assertFalse(r_tracker.chk_ffs.isChecked())

            r_contrib = restored_meetings_tab.contribution_dialog
            self.assertIsNotNone(r_contrib)
            self.assertEqual(r_contrib.wi_token_widget.get_tokens(), ["FS_6G_Radio"])
            self.assertEqual(r_contrib.ai_token_widget.get_tokens(), ["9.3.1"])
            self.assertEqual(r_contrib.company_widget.get_selected_companies(), {"Ericsson"})
            self.assertTrue(r_contrib.company_widget.is_primary_only())
            self.assertTrue(r_contrib.company_widget.is_include_tentative())

            r_hist = getattr(restored_specs_tab, "_history_dialog", None)
            self.assertIsNotNone(r_hist)
            self.assertEqual(r_hist.search_input.text(), "S2-2601234")
            self.assertEqual(r_hist.clause_input.text(), "5.15")
            self.assertTrue(r_hist.chk_clause_filter.isChecked())
            self.assertTrue(r_hist.chk_origin_ranking.isChecked())

            r_diff = getattr(restored_spec_search_tab, "_diff_dialog", None)
            self.assertIsNotNone(r_diff)
            self.assertEqual(r_diff.tab_widget.currentIndex(), 1)
            self.assertEqual(r_diff.focus_combo.currentIndex(), 2)

            r_cr = restored_specs_tab._active_cr_dialogs[0]
            self.assertEqual(r_cr.current_mode, "cr")
            self.assertEqual(r_cr.input_title.text(), "Test Contribution")

            for w in (r_dossier, r_tracker, r_contrib, r_hist, r_diff, r_cr):
                w.close()

    def test_missing_tdoc_file_skips_gracefully_without_modal(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            cfg = self._new_settings(tmp)
            save_open_windows_state(
                [
                    {
                        "type": "tdocs_window",
                        "geometry": "",
                        "maximized": False,
                        "state": {
                            "filepath": str(Path(tmp) / "nonexistent_tdocs.xlsx"),
                            "mtg_info": {"mtg_id": "99999", "wg_name": "SA2", "meeting_number": "999"},
                        },
                    }
                ],
                settings=cfg,
            )
            fake_main = SimpleNamespace(
                meetings_tab=SimpleNamespace(tdoc_windows={}),
            )
            with patch("PyQt5.QtWidgets.QMessageBox.warning") as warn_mock:
                count = restore_open_windows(
                    fake_main,
                    settings=cfg,
                    argv=["main_tools.py"],
                    modifiers=Qt.NoModifier,
                )
            self.assertEqual(count, 0)
            warn_mock.assert_not_called()

    def test_corrupt_saved_payload_handled_safely(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            cfg = self._new_settings(tmp)
            cfg.setValue("session/open_windows", "{not-valid-json")
            self.assertEqual(load_open_windows_state(cfg), [])

    def test_tdocs_session_plan_email_and_excerpt_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            cfg = self._new_settings(tmp)
            meeting_dir = tmp_path / "SA2#177"
            agenda_dir = meeting_dir / "Agenda"
            agenda_dir.mkdir(parents=True, exist_ok=True)
            tdocs_xlsx = agenda_dir / "TDoc_List_Meeting_177.xlsx"
            tdocs_xlsx.write_bytes(b"placeholder")
            spec_dl_dir = tmp_path / "specs" / "23.501"
            spec_dl_dir.mkdir(parents=True, exist_ok=True)

            mtg_info = {
                "mtg_id": "17700",
                "wg_name": "SA2",
                "meeting_number": "177",
                "location": "Prague",
                "start_date": "2026-10-12",
                "end_date": "2026-10-16",
            }
            tdocs_data = [
                {"TDoc": "S2-2600001", "Title": "Test TDoc 1", "Agenda Item": "6.1", "Source": "Ericsson"},
                {"TDoc": "S2-2600002", "Title": "Test TDoc 2", "Agenda Item": "6.2", "Source": "Nokia"},
            ]

            tdoc_win = TDocsWindow(mtg_info, tdocs_data, str(tdocs_xlsx), filter_settings=cfg)
            tdoc_win.show()
            idx_row1 = tdoc_win.proxy.index(1, 1)
            tdoc_win.table.selectRow(1)
            tdoc_win.table.setCurrentIndex(idx_row1)
            selected_before = tdoc_win.get_session_state()["selected_tdoc"]

            sp_win = SessionPlanWindow(meeting_dir=meeting_dir, mtg_info=mtg_info, agenda_map={}, tracked_tokens=set())
            sp_win.btn_agenda.setChecked(True)
            sp_win._on_view_mode_changed(2)
            sp_win.search_input.setText("6.2")
            sp_win.show()
            tdoc_win._session_plan_window = sp_win

            em_win = EmailManagerWindow(meeting_dir, {"S2-2600001": "6.1"})
            em_win.tdoc_search_input.setText("S2-2600001")
            em_win.search_input.setText("agreement")
            em_win.show()
            tdoc_win.email_window = em_win

            inbox_dlg = MeetingEmailInboxDialog(
                tracked_db=TrackedThreadsDatabase(agenda_dir / "tracked_threads.db"),
                general_email_db=GeneralEmailDatabase(agenda_dir / "emails.db"),
                meeting_id="SA2#177",
                settings=cfg,
            )
            inbox_dlg.show()
            tdoc_win._inbox_dialog = inbox_dlg

            excerpt_dlg = SpecClauseExcerptDialog(
                spec_num="23.501",
                spec_type="TS",
                version="18.5.0",
                fname="23501-i50.zip",
                url="",
                spec_dl_dir=spec_dl_dir,
                auto_start=False,
            )
            excerpt_dlg.filter_input.setText("5.15")
            excerpt_dlg.show()

            fake_meetings_tab = SimpleNamespace(tdoc_windows={"17700": tdoc_win})
            fake_specs_tab = SimpleNamespace(_active_excerpt_dialogs=[excerpt_dlg])
            fake_main = SimpleNamespace(meetings_tab=fake_meetings_tab, specs_tab=fake_specs_tab)

            descriptors = collect_open_windows_state(fake_main)
            self.assertEqual(
                [d["type"] for d in descriptors],
                [
                    "tdocs_window",
                    "session_plan",
                    "email_manager",
                    "meeting_email_inbox",
                    "spec_clause_excerpt",
                ],
            )
            save_open_windows_state(descriptors, settings=cfg)

            for w in (excerpt_dlg, inbox_dlg, em_win, sp_win, tdoc_win):
                w.close()

            # Restore using lightweight openers that instantiate fresh windows
            restored_tab = SimpleNamespace(tdoc_windows={})

            def _fake_open_tdocs(info, fp, quiet=False):
                w = TDocsWindow(info, tdocs_data, fp, filter_settings=cfg)
                w._open_session_plan_window = lambda: setattr(
                    w,
                    "_session_plan_window",
                    SessionPlanWindow(meeting_dir=meeting_dir, mtg_info=info, agenda_map={}, tracked_tokens=set()),
                )
                w._open_email_manager = lambda: setattr(
                    w,
                    "email_window",
                    EmailManagerWindow(meeting_dir, {"S2-2600001": "6.1"}),
                )
                w._open_email_inbox = lambda: setattr(
                    w,
                    "_inbox_dialog",
                    MeetingEmailInboxDialog(
                        tracked_db=TrackedThreadsDatabase(agenda_dir / "tracked_threads.db"),
                        general_email_db=GeneralEmailDatabase(agenda_dir / "emails.db"),
                        meeting_id="SA2#177",
                        settings=cfg,
                    ),
                )
                restored_tab.tdoc_windows[info["mtg_id"]] = w

            restored_tab._open_tdocs_window = _fake_open_tdocs
            restored_specs_tab = SimpleNamespace(_active_excerpt_dialogs=[])
            restored_main = SimpleNamespace(meetings_tab=restored_tab, specs_tab=restored_specs_tab)

            with patch.object(SpecClauseExcerptDialog, "_start_preparation", lambda self: None):
                count = restore_open_windows(
                    restored_main,
                    settings=cfg,
                    argv=["main_tools.py"],
                    modifiers=Qt.NoModifier,
                )
            self.assertEqual(count, 5)

            r_tdoc = restored_tab.tdoc_windows["17700"]
            self.assertEqual(r_tdoc.get_session_state()["selected_tdoc"], selected_before)

            r_sp = r_tdoc._session_plan_window
            self.assertEqual(r_sp.view_stack.currentIndex(), 2)
            self.assertEqual(r_sp.search_input.text(), "6.2")

            r_em = r_tdoc.email_window
            self.assertEqual(r_em.tdoc_search_input.text(), "S2-2600001")
            self.assertEqual(r_em.search_input.text(), "agreement")

            self.assertIsNotNone(r_tdoc._inbox_dialog)
            r_exc = restored_specs_tab._active_excerpt_dialogs[0]
            self.assertEqual(r_exc.filter_input.text(), "5.15")

            for w in (r_exc, r_tdoc._inbox_dialog, r_em, r_sp, r_tdoc):
                w.close()


if __name__ == "__main__":
    unittest.main()
