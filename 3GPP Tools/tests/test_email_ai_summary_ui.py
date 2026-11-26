# --- File: tests/test_email_ai_summary_ui.py ---
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt5.QtCore import QSettings, Qt, QUrl
from PyQt5.QtWidgets import QApplication

project_root = Path(__file__).resolve().parent.parent
src_dir = project_root / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from modules.emails.core.ai_summary import EmailSummaryRecord, EmailSummaryStore, SummaryScope
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.ui.ai_summary import InlineEmailSummaryCard
from modules.emails.ui.briefings import EmailBriefingsDialog
from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog
from modules.emails.ui.tdoc_emails_dialog import TDocEmailsDialog
from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog

_APP = QApplication.instance() or QApplication([])


class TestEmailAiSummaryUi(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self._tmp.name)
        self.agenda_dir = self.tmp_path / "Agenda"
        self.agenda_dir.mkdir(parents=True, exist_ok=True)
        self.settings = QSettings(str(self.tmp_path / "ui.ini"), QSettings.IniFormat)

        self.tdb = TrackedThreadsDatabase(self.agenda_dir / "tracked_threads.db")
        self.gdb = GeneralEmailDatabase(self.agenda_dir / "emails.db")
        self.store = EmailSummaryStore(self.agenda_dir / "email_summaries.db")

        self.tid = self.tdb.create_or_get_thread(
            "URSP rule update",
            "ursp rule update",
            meeting_id="S2-174",
            wg="SA2",
            linked_ais=["6.1"],
            linked_tdocs=["S2-2602001"],
        )
        self.tdb.upsert_messages(self.tid, [
            {"id": "M1", "subject": "URSP rule update", "sender_name": "Alice", "company": "Ericsson", "date_sent": "2026-04-14 10:00:00", "body_text": "Proposal text", "is_read": 0},
            {"id": "M2", "subject": "Re: URSP rule update", "sender_name": "Bob", "company": "Nokia", "date_sent": "2026-04-14 11:00:00", "body_text": "Support", "is_read": 0},
        ])
        self.gdb.save_emails_batch(
            [
                {"id": "G1", "folder_tag": "SA2", "subject": "S2-2602001 comment", "sender_name": "Carol", "company": "Huawei", "date_received": "2026-04-14 12:00:00", "body_text": "Comment on S2-2602001", "is_read": 0},
            ],
            [{"email_id": "G1", "tdoc_id": "S2-2602001", "rev_matched": "", "match_location": "Subject"}],
        )

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_inline_summary_card_visibility_and_signals(self):
        card = InlineEmailSummaryCard()
        msgs = self.tdb.get_thread_messages(self.tid)
        card.bind_scope(self.store, SummaryScope("thread", self.tid, "URSP rule update"), msgs)
        self.assertFalse(card.isVisible())  # 0 px footprint when no summary exists yet

        self.store.upsert_summary(
            EmailSummaryRecord(
                scope_type="thread",
                scope_key=self.tid,
                title="URSP rule update",
                summary_md="### 1. Topic\nDetails [Msg #1](jump://msg/1)",
                covered_message_ids=["M1"],
                cutoff_utc="2026-04-14 10:00:00",
                provider="codex",
                model="gpt-6.2-astra",
            )
        )
        fresh = card.bind_scope(self.store, SummaryScope("thread", self.tid, "URSP rule update"), msgs)
        self.assertTrue(card.isVisible())
        self.assertEqual(fresh.new_count, 1)
        self.assertIn("+1", card.btn_update.text())

        clicked_urls = []
        card.anchor_clicked.connect(clicked_urls.append)
        card.anchor_clicked.emit(QUrl("jump://msg/1"))
        self.assertEqual(len(clicked_urls), 1)
        card.shutdown_worker()
        card.close()

    def test_inbox_contextual_ai_strip_and_multiselect_summarize(self):
        dlg = MeetingEmailInboxDialog(
            tracked_db=self.tdb,
            general_email_db=self.gdb,
            meeting_id="S2-174",
            tdoc_ai_lookup={"S2-2602001": "6.1"},
            wg="SA2",
            ai_suggestions=["6.1 — 5GS Policy"],
            settings=self.settings,
        )
        dlg._set_mode("ALL")
        dlg.show()
        self.assertFalse(dlg.ai_briefing_strip.isVisible())

        # Filter down to 1 AI -> contextual strip appears
        dlg.ai_combo.setCheckedItems(["6.1"])
        dlg._apply_filters()
        self.assertTrue(dlg.ai_briefing_strip.isVisible())
        self.assertIn("6.1", dlg.ai_briefing_strip_lbl.text())

        # Multi-selection keeps Summarize button enabled
        dlg.ai_combo.setCheckedItems([])
        dlg._apply_filters()
        self.assertGreaterEqual(dlg.table.rowCount(), 2)
        dlg.table.selectAll()
        dlg._on_selection_changed()
        self.assertTrue(dlg.btn_llm_export.isEnabled())
        self.assertTrue(dlg.btn_llm_export.isVisible())
        self.assertIn("Summarize", dlg.btn_llm_export.text())
        dlg.close()

    def test_tracked_thread_tdoc_emails_and_briefings_modeless(self):
        tt_dlg = TrackedThreadDialog(self.tdb, self.tid)
        self.assertTrue(hasattr(tt_dlg, "summary_card"))
        tt_dlg.close()

        td_dlg = TDocEmailsDialog("S2-2602001", ["S2-2602001"], self.gdb.db_path, wg="SA2")
        self.assertTrue(hasattr(td_dlg, "summary_card"))
        td_dlg.close()

        br_dlg = EmailBriefingsDialog(
            tracked_db=self.tdb,
            general_email_db=self.gdb,
            summary_db_path=self.store.db_path,
            meeting_id="S2-174",
            wg="SA2",
            ai_label_map={"6.1": "6.1 — 5GS Policy"},
            followed_ais={"6.1"},
            tdoc_ai_lookup={"S2-2602001": "6.1"},
            initial_ai_code="6.1",
        )
        self.assertIsNone(br_dlg.parent())
        self.assertEqual(br_dlg.windowModality(), Qt.NonModal)
        self.assertTrue(bool(br_dlg.windowFlags() & Qt.Window))
        self.assertIn("6.1", br_dlg.scope_title_lbl.text())
        self.assertTrue(hasattr(br_dlg, "btn_m365_copilot"))
        br_dlg.close()

    def test_m365_copilot_handoff_dialog_modeless_and_roundtrip_save(self):
        card = InlineEmailSummaryCard()
        msgs = self.tdb.get_thread_messages(self.tid)
        card.bind_scope(self.store, SummaryScope("thread", self.tid, "URSP rule update"), msgs)
        self.assertFalse(card.isVisible())

        with patch("modules.emails.core.ai_summary.m365_handoff.webbrowser.open", return_value=True) as mock_web, patch(
            "modules.emails.ui.ai_summary.m365_copilot_dialog.get_m365_copilot_target", return_value="web"
        ):
            handoff_dlg = card.open_m365_copilot_handoff(mode="auto", auto_launch=True)

        self.assertIsNotNone(handoff_dlg)
        self.assertIsNone(handoff_dlg.parent())
        self.assertEqual(handoff_dlg.windowModality(), Qt.NonModal)
        self.assertTrue(bool(handoff_dlg.windowFlags() & Qt.Window))
        self.assertTrue(mock_web.called)
        self.assertIn("URSP rule update", QApplication.clipboard().text())

        handoff_dlg.response_edit.setPlainText("### 1. Core Contention\n- M365 Copilot summary [Msg #1](jump://msg/1)")
        handoff_dlg._on_save_clicked()

        saved = self.store.get_summary("thread", self.tid)
        self.assertIsNotNone(saved)
        self.assertEqual(saved.provider, "m365-copilot")
        self.assertEqual(set(saved.covered_message_ids), {"M1", "M2"})
        self.assertTrue(card.isVisible())
        self.assertIn("M365 Copilot summary", card.body_browser.toPlainText())
        card.shutdown_worker()
        card.close()

    def test_general_email_sync_dialog_default_buffer_5_and_live_window(self):
        from PyQt5.QtCore import QDate
        from modules.emails.ui.general_email_dialog import GeneralEmailSyncDialog

        sample_folders = [
            {"folder_path": "Mailbox/SA2", "tag": "WG", "color": "#1E6FD9", "include_subfolders": True},
        ]
        with patch("modules.emails.ui.general_email_dialog.load_wg_email_config", return_value=sample_folders):
            dlg = GeneralEmailSyncDialog("SA2", "2026-11-16", "2026-11-20")
            self.assertEqual(dlg.get_params()["buffer"], 5)
            self.assertEqual(dlg.dt_start.displayFormat(), "yyyy-MM-dd")
            self.assertGreaterEqual(dlg.dt_start.minimumWidth(), 130)
            self.assertTrue(dlg.btn_sync.isEnabled())
            self.assertIn("2026-11-11", dlg.lbl_effective_window.text())
            self.assertIn("2026-11-25", dlg.lbl_effective_window.text())

            # Clicking preset changes buffer & recomputes effective window live
            for dval, btn in dlg._preset_buttons:
                if dval == 10:
                    btn.click()
                    break
            self.assertEqual(dlg.get_params()["buffer"], 10)
            self.assertIn("2026-11-06", dlg.lbl_effective_window.text())
            self.assertIn("2026-11-30", dlg.lbl_effective_window.text())

            # Inverted date range disables Start Sync
            dlg.dt_end.setDate(QDate(2026, 11, 10))
            self.assertFalse(dlg.btn_sync.isEnabled())

            dlg.btn_reset_dates.click()
            self.assertEqual(dlg.get_params()["buffer"], 5)
            self.assertTrue(dlg.btn_sync.isEnabled())
            dlg.close()

        # Zero folders disables Start Sync with actionable prompt
        with patch("modules.emails.ui.general_email_dialog.load_wg_email_config", return_value=[]):
            empty_dlg = GeneralEmailSyncDialog("SA2", "2026-11-16", "2026-11-20")
            self.assertFalse(empty_dlg.btn_sync.isEnabled())
            self.assertIn("No Outlook folders", empty_dlg.lbl_folders.text())
            empty_dlg.close()

    def test_inline_summary_card_error_state_and_briefings_no_clipping(self):
        from modules.emails.ui.briefings.mixins.actions import _compact_ai_nav_label, _natural_ai_sort_key

        card = InlineEmailSummaryCard()
        msgs = self.tdb.get_thread_messages(self.tid)
        card.bind_scope(self.store, SummaryScope("thread", self.tid, "URSP rule update"), msgs)
        card.trigger_summarize = lambda **_kw: None  # Prevent real worker spawn
        card.setVisible(True)

        # Simulate worker failure + finished -> never displays 'Ready' on error!
        card._on_worker_failed("codex CLI could not reach its backend (workspace routing discovery failed).")
        card._on_worker_finished()
        self.assertTrue(card.error_banner.isVisible())
        self.assertFalse(card.body_browser.isVisible())
        self.assertIn("failed", card.freshness_pill.text().lower())
        self.assertNotEqual(card.freshness_pill.text().strip().lower(), "ready")
        self.assertIn("Retry", card.btn_rebuild.text())

        # Dismissing error when no summary exists collapses card back to 0 px
        card.error_banner.btn_dismiss.click()
        self.assertFalse(card.error_banner.isVisible())
        self.assertFalse(card.isVisible())
        card.close()

        # Verify natural sorting (20.6.2 before 20.6.17) and compact nav label
        codes = ["20.6.17", "20.6.2", "20.6.0"]
        self.assertEqual(sorted(codes, key=_natural_ai_sort_key), ["20.6.0", "20.6.2", "20.6.17"])
        self.assertEqual(
            _compact_ai_nav_label("20.6.0", "20.6.0 — Study on Architecture for 6G System (FS_6G_ARC) — Miscellaneous"),
            "20.6.0 — Miscellaneous",
        )

        br_dlg = EmailBriefingsDialog(
            tracked_db=self.tdb,
            general_email_db=self.gdb,
            summary_db_path=self.store.db_path,
            meeting_id="S2-174",
            wg="SA2",
            ai_label_map={"6.1": "6.1 — Very Long Study Title That Should Wrap — Subtopic"},
            followed_ais={"6.1"},
            tdoc_ai_lookup={"S2-2602001": "6.1"},
            initial_ai_code="6.1",
        )
        br_dlg.show()
        self.assertLessEqual(br_dlg.leaves_frame.maximumHeight(), 45)
        br_dlg._on_worker_failed("codex CLI could not reach its backend (timeout).")
        self.assertTrue(br_dlg.error_banner.isVisible())
        br_dlg.error_banner.btn_dismiss.click()
        self.assertFalse(br_dlg.error_banner.isVisible())
        br_dlg.close()

    def test_summary_reader_layout_modes_popout_and_jump_restore(self):
        from modules.emails.ui.ai_summary import PopoutSummaryWindow

        # Save a summary for the tracked thread and TDoc first
        self.store.upsert_summary(
            EmailSummaryRecord(
                scope_type="thread",
                scope_key=self.tid,
                title="URSP rule update",
                summary_md="### 1. Topic\nSummary body [Msg #1](jump://msg/1)",
                covered_message_ids=["M1", "M2"],
                cutoff_utc="2026-04-14 11:00:00",
                provider="m365-copilot",
                model="M365 Copilot",
            )
        )
        self.store.upsert_summary(
            EmailSummaryRecord(
                scope_type="tdoc",
                scope_key="S2-2602001",
                title="S2-2602001",
                summary_md="### 1. Topic\nTDoc summary body [Msg #1](jump://msg/1)",
                covered_message_ids=["G1"],
                cutoff_utc="2026-04-14 12:00:00",
                provider="m365-copilot",
                model="M365 Copilot",
            )
        )

        with patch(
            "modules.emails.ui.ai_summary.layout_controller.QSettings",
            return_value=self.settings,
        ):
            td_dlg = TDocEmailsDialog("S2-2602001", ["S2-2602001"], self.gdb.db_path, wg="SA2")
            ctrl = td_dlg.summary_layout_ctrl
            td_dlg.show()

            # Side-by-Side (Summary Left | Emails Right)
            ctrl.set_mode("split_h")
            self.assertEqual(td_dlg.reader_splitter.orientation(), Qt.Horizontal)
            self.assertEqual(td_dlg.reader_splitter.indexOf(td_dlg.summary_card), 0)
            self.assertEqual(td_dlg.reader_splitter.indexOf(td_dlg.browser), 1)
            self.assertTrue(td_dlg.browser.isVisible())

            # Summary Only hides browser; clicking jump://msg/1 restores Side-by-Side
            ctrl.set_mode("summary_only")
            self.assertTrue(td_dlg.summary_card.isVisible())
            self.assertFalse(td_dlg.browser.isVisible())
            td_dlg.summary_card.anchor_clicked.emit(QUrl("jump://msg/1"))
            self.assertEqual(ctrl.current_mode(), "split_h")
            self.assertTrue(td_dlg.browser.isVisible())

            # Slim collapse -> thread_only
            td_dlg.summary_card.btn_toggle.click()
            self.assertEqual(ctrl.current_mode(), "thread_only")
            self.assertLessEqual(td_dlg.summary_card.maximumHeight(), 80)
            self.assertFalse(td_dlg.summary_card.body_browser.isVisible())
            self.assertTrue(td_dlg.browser.isVisible())

            # Pop Out -> modeless PopoutSummaryWindow (mirrors live content, keeps inline card in splitter)
            ctrl.set_mode("popout")
            pop = ctrl._popout_win
            self.assertIsInstance(pop, PopoutSummaryWindow)
            self.assertIsNone(pop.parent())
            self.assertEqual(pop.windowModality(), Qt.NonModal)
            self.assertTrue(bool(pop.windowFlags() & Qt.Window))
            self.assertTrue(pop.isVisible())
            self.assertFalse(td_dlg.summary_card.isVisible())
            self.assertTrue(td_dlg.browser.isVisible())
            self.assertIn("TDoc summary body", pop.browser.toPlainText())

            # Anchor clicked inside PopoutSummaryWindow forwards to summary_card -> scrolls parent reader
            pop_jumps = []
            td_dlg.summary_card.anchor_clicked.connect(pop_jumps.append)
            pop.anchor_clicked.emit(QUrl("jump://msg/1"))
            self.assertEqual(len(pop_jumps), 1)
            self.assertEqual(pop_jumps[0].toString(), "jump://msg/1")

            pop.close()
            QApplication.processEvents()
            self.assertIsNone(ctrl._popout_win)
            self.assertTrue(td_dlg.summary_card.isVisible())
            self.assertEqual(td_dlg.reader_splitter.indexOf(td_dlg.summary_card), 0)
            self.assertNotEqual(ctrl.current_mode(), "popout")
            td_dlg.close()

            # TrackedThreadDialog and MeetingEmailInboxDialog also expose summary_layout_ctrl
            tt_dlg = TrackedThreadDialog(self.tdb, self.tid)
            self.assertTrue(hasattr(tt_dlg, "summary_layout_ctrl"))
            tt_dlg.close()

            inbox_dlg = MeetingEmailInboxDialog(
                tracked_db=self.tdb,
                general_email_db=self.gdb,
                meeting_id="S2-174",
                tdoc_ai_lookup={"S2-2602001": "6.1"},
                wg="SA2",
                ai_suggestions=["6.1 — 5GS Policy"],
                settings=self.settings,
            )
            self.assertTrue(hasattr(inbox_dlg, "summary_layout_ctrl"))
            inbox_dlg.close()
            QApplication.processEvents()

    def test_m365_save_from_collapsed_thread_only_restores_split_and_self_heals_sizes(self):
        # Seed a saved summary and corrupted tiny split height + thread_only mode in settings
        rec = EmailSummaryRecord(
            scope_type="tdoc",
            scope_key="S2-2602001",
            title="S2-2602001",
            summary_md="### 1. Topic\nInitial summary [Msg #1](jump://msg/1)",
            covered_message_ids=["G1"],
            cutoff_utc="2026-04-14 12:00:00",
            provider="m365-copilot",
            model="m365_copilot",
        )
        self.store.upsert_summary(rec)

        bad_settings = QSettings(str(self.tmp_path / "corrupt_layout.ini"), QSettings.IniFormat)
        bad_settings.setValue("tdoc_emails/mode", "thread_only")
        bad_settings.setValue("tdoc_emails/last_split_mode", "split_v")
        bad_settings.setValue("tdoc_emails/sizes_v", "38,600")
        bad_settings.sync()

        with patch(
            "modules.emails.ui.ai_summary.layout_controller.QSettings",
            return_value=bad_settings,
        ):
            td_dlg = TDocEmailsDialog("S2-2602001", ["S2-2602001"], self.gdb.db_path, wg="SA2")
            td_dlg.resize(680, 760)
            ctrl = td_dlg.summary_layout_ctrl
            td_dlg.show()
            QApplication.processEvents()
            narrow_max_h = td_dlg.summary_card.maximumHeight()
            self.assertGreater(narrow_max_h, 55)

            # Widening window after show() automatically recomputes collapsed header back down to 1 row
            td_dlg.resize(1380, 760)
            QApplication.processEvents()
            self.assertEqual(ctrl.current_mode(), "thread_only")
            self.assertFalse(td_dlg.summary_card.body_browser.isVisible())
            self.assertLessEqual(td_dlg.summary_card.maximumHeight(), 48)
            self.assertEqual(td_dlg.summary_card.model_pill.text(), "💠 M365 Copilot")

            # Simulate pasting/saving an updated M365 Copilot summary while collapsed in thread_only
            updated_rec = EmailSummaryRecord(
                scope_type="tdoc",
                scope_key="S2-2602001",
                title="S2-2602001",
                summary_md="### 1. Topic\nUpdated Copilot output [Msg #1](jump://msg/1)",
                covered_message_ids=["G1"],
                cutoff_utc="2026-04-14 12:00:00",
                provider="m365_copilot",
                model="m365_copilot",
            )
            self.store.upsert_summary(updated_rec)
            td_dlg.summary_card._on_external_summary_saved(updated_rec)
            QApplication.processEvents()

            # Must cleanly exit thread_only into split_v, unlock maximumHeight, show body, and self-heal sizes_v
            self.assertEqual(ctrl.current_mode(), "split_v")
            self.assertGreater(td_dlg.summary_card.maximumHeight(), 10000)
            self.assertTrue(td_dlg.summary_card.body_browser.isVisible())
            self.assertTrue(td_dlg.browser.isVisible())
            restored = ctrl._restore_split_sizes("split_v")
            self.assertGreaterEqual(restored[0], 120)
            self.assertEqual(td_dlg.summary_card.model_pill.text(), "💠 M365 Copilot")

            td_dlg.close()
            QApplication.processEvents()


if __name__ == "__main__":
    unittest.main()
