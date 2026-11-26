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

from core.ai.llm_runner import RunnerResult
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


if __name__ == "__main__":
    unittest.main()
