# --- File: tests/test_emeeting_unified_inbox.py ---
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt5.QtCore import QSettings, QUrl
from PyQt5.QtWidgets import QApplication

from modules.emails.core.emeeting_signals import (
    analyze_message_signals,
    build_message_jump_strip_html,
    compute_group_signals,
    detect_objection_in_text,
    extract_revisions_from_text,
    extract_start_end_blocks,
    format_signal_badges,
    linkify_shorthand_revisions,
)
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.stats.dataset_builder import build_unified_email_stats_dataset
from modules.emails.core.thread_markdown_exporter import format_emeeting_body_with_callouts
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.ui.meeting_inbox.dialog import MeetingEmailInboxDialog
from modules.emails.ui.meeting_inbox.tdoc_grouping import (
    build_grouped_tdoc_items,
    is_item_effectively_parked,
    is_priority_item,
)
from modules.meetings.core.tdocs.tdocs_db import TDocsDatabase


class TestEMeetingUnifiedInbox(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def test_start_end_extraction_and_callout_rendering(self):
        body = (
            "Hello colleagues,\n"
            "<<START>>\n"
            "We support S2-2601234 with the changes proposed in r02.\n"
            "<<END>>\n"
            "Best regards,\nAlice"
        )
        blocks, free_txt = extract_start_end_blocks(body)
        self.assertEqual(len(blocks), 1)
        self.assertIn("changes proposed in r02", blocks[0])
        self.assertIn("Hello colleagues,", free_txt)
        self.assertNotIn("<<START>>", free_txt)

        max_rev, rev_tokens = extract_revisions_from_text(body, primary_tdoc="S2-2601234")
        self.assertEqual(max_rev, 2)
        self.assertIn("S2-2601234r02", rev_tokens)

        linked = linkify_shorthand_revisions(blocks[0], primary_tdoc="S2-2601234")
        self.assertIn("[r02](tdoc:S2-2601234R02)", linked)

        html_out = format_emeeting_body_with_callouts(
            body,
            primary_tdoc="S2-2601234",
            interactive_links=True,
            enable_tdoc_links=True,
        )
        self.assertIn("e-Meeting Formal Comment", html_out)
        self.assertIn("tdoc:S2-2601234R02", html_out)

    def test_objection_detection_and_smart_unparking(self):
        has_obj, snip = detect_objection_in_text("We do not support this proposal as written.")
        self.assertTrue(has_obj)
        self.assertIn("We do not support", snip)

        # Quoted objection should be ignored
        quoted_only, _ = detect_objection_in_text("> We do not support this proposal.\nWe agree with the rapporteur.")
        self.assertFalse(quoted_only)

        msgs_r01 = [
            {
                "id": "m1",
                "subject": "[SA2#162] S2-2601234",
                "body_text": "<<START>> Please see r01 <<END>>",
                "company": "Huawei",
                "is_read": False,
            },
            {
                "id": "m2",
                "subject": "Re: [SA2#162] S2-2601234",
                "body_text": "We cannot agree to r01 due to roaming impacts.",
                "company": "Ericsson",
                "is_read": False,
            },
        ]
        sig_parked = compute_group_signals(
            msgs_r01,
            primary_tdoc="S2-2601234",
            is_objected=True,
            objected_at_rev="r01",
        )
        self.assertTrue(sig_parked["effectively_parked"])
        self.assertFalse(sig_parked["unparked_by_new_rev"])
        self.assertIn("🛑", format_signal_badges(sig_parked))

        # Now author posts r02 after the objection at r01 -> automatically unparks!
        msgs_r02 = list(msgs_r01) + [
            {
                "id": "m3",
                "subject": "Re: [SA2#162] S2-2601234",
                "body_text": "<<START>> Uploaded r02 addressing Ericsson's concern <<END>>",
                "company": "Huawei",
                "is_read": False,
            }
        ]
        sig_unparked = compute_group_signals(
            msgs_r02,
            primary_tdoc="S2-2601234",
            is_objected=True,
            objected_at_rev="r01",
        )
        self.assertFalse(sig_unparked["effectively_parked"])
        self.assertTrue(sig_unparked["unparked_by_new_rev"])
        self.assertIn("🛑→🆕r02", format_signal_badges(sig_unparked))

    def test_legacy_emeeting_migration_and_zero_data_loss(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = Path(tmpdir) / "Agenda" / "emails.db"
            db_path.parent.mkdir(parents=True, exist_ok=True)
            conn = sqlite3.connect(db_path)
            conn.execute(
                """
                CREATE TABLE emails (
                    id TEXT PRIMARY KEY,
                    tdoc_id TEXT,
                    agenda_item TEXT,
                    sender_name TEXT,
                    company TEXT,
                    date_received TEXT,
                    subject TEXT,
                    revisions_mentioned TEXT,
                    short_text TEXT,
                    free_text TEXT,
                    msg_path TEXT,
                    outlook_location TEXT,
                    sender_email TEXT
                )
                """
            )
            conn.execute(
                """
                INSERT INTO emails VALUES (
                    'LEG-001', 'S2-2605555', '6.1', 'Bob', 'Nokia', '2026-11-20 10:00:00',
                    '[SA2#162] S2-2605555 QoS', 'S2-2605555r01', 'We support r01.', 'Additional remarks.',
                    '', 'Source', 'bob@nokia.com'
                )
                """
            )
            conn.commit()
            conn.close()

            # Initializing GeneralEmailDatabase auto-migrates legacy `emails` rows
            gdb = GeneralEmailDatabase(db_path)
            all_rows = gdb.get_all_emails_with_tdocs()
            self.assertEqual(len(all_rows), 1)
            self.assertEqual(all_rows[0]["id"], "LEG-001")
            self.assertIn("<<START>>", all_rows[0]["body_text"])
            self.assertIn("We support r01.", all_rows[0]["body_text"])
            self.assertIn("S2-2605555", all_rows[0]["matched_tdocs"])

            # Zero-data-loss check: upserting with empty body_text/clean_body never wipes cached body
            gdb.save_emails_batch(
                [
                    {
                        "id": "LEG-001",
                        "folder_tag": "SA2",
                        "subject": "[SA2#162] S2-2605555 QoS",
                        "sender_name": "Bob",
                        "sender_email": "bob@nokia.com",
                        "company": "Nokia",
                        "date_received": "2026-11-20 10:00:00",
                        "body_text": "",
                        "clean_body": "",
                        "attachments": [],
                        "is_read": 0,
                        "flag_status": 2,
                    }
                ],
                [{"email_id": "LEG-001", "tdoc_id": "S2-2605555", "rev_matched": "S2-2605555r01", "match_location": "Subject"}],
                sync_unread=True,
            )
            reloaded = gdb.get_all_emails_with_tdocs()[0]
            self.assertIn("We support r01.", reloaded["body_text"])
            self.assertIn("We support r01.", reloaded["clean_body"])

    def test_inbox_by_tdoc_view_priority_and_two_way_tdoc_sync(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            agenda_dir = Path(tmpdir) / "Agenda"
            gdb = GeneralEmailDatabase(agenda_dir / "emails.db")
            tdb = TrackedThreadsDatabase(agenda_dir / "tracked_threads.db")
            udb = TDocsDatabase(agenda_dir / "user_tdocs.db")
            settings = QSettings(str(agenda_dir / "inbox_test.ini"), QSettings.IniFormat)

            gdb.save_emails_batch(
                [
                    {
                        "id": "E-1",
                        "folder_tag": "SA2",
                        "subject": "[SA2#162, AI 6.1] S2-2601111 KI#1 solution",
                        "sender_name": "Alice",
                        "sender_email": "alice@huawei.com",
                        "company": "Huawei",
                        "date_received": "2026-11-21 09:00:00",
                        "body_text": "<<START>>\nPlease review r01.\n<<END>>",
                        "clean_body": "<<START>>\nPlease review r01.\n<<END>>",
                        "attachments": [],
                        "is_read": 0,
                        "flag_status": 0,
                    },
                    {
                        "id": "E-2",
                        "folder_tag": "SA2",
                        "subject": "Re: [SA2#162, AI 6.1] S2-2601111 KI#1 solution",
                        "sender_name": "Erik",
                        "sender_email": "erik@ericsson.com",
                        "company": "Ericsson",
                        "date_received": "2026-11-21 10:00:00",
                        "body_text": "We do not support r01 yet.",
                        "clean_body": "We do not support r01 yet.",
                        "attachments": [],
                        "is_read": 0,
                        "flag_status": 0,
                    },
                ],
                [
                    {"email_id": "E-1", "tdoc_id": "S2-2601111", "rev_matched": "S2-2601111r01", "match_location": "Subject"},
                    {"email_id": "E-2", "tdoc_id": "S2-2601111", "rev_matched": "S2-2601111r01", "match_location": "Subject"},
                ],
                sync_unread=True,
            )

            tdoc_meta_map = {
                "S2-2601111": {
                    "ai": "6.1",
                    "sub_item": "KI#1",
                    "wis": ["FS_XRM"],
                    "title": "New QoS Monitoring Procedure",
                    "source": "Huawei, China Mobile",
                    "type": "pCR",
                    "tdoc_status": "Available",
                    "my_status": "🟢 Support",
                    "my_notes": "Initial note",
                    "is_objected": False,
                    "objected_at_rev": "",
                }
            }
            synced_calls = []

            dlg = MeetingEmailInboxDialog(
                tracked_db=tdb,
                general_email_db=gdb,
                meeting_id="SA2#162E",
                tdoc_ai_lookup={"S2-2601111": "6.1"},
                tdoc_meta_map=tdoc_meta_map,
                ai_suggestions=["6.1 — XRM Study"],
                wi_suggestions=["FS_XRM"],
                settings=settings,
                tdocs_db=udb,
                user_data_callback=lambda tid, st, notes: synced_calls.append((tid, st, notes)),
            )
            try:
                # Switch to By TDoc mode
                dlg._set_view_mode("TDOCS")
                self.assertEqual(dlg._view_mode, "TDOCS")
                self.assertEqual(dlg.table.rowCount(), 1)
                self.assertIn("S2-2601111", dlg.table.item(0, 1).text())
                self.assertIn("Huawei", dlg.table.item(0, 1).text())
                self.assertIn("2 cos", dlg.table.item(0, 4).text())
                self.assertIn("🆕r01", dlg.table.item(0, 0).text())
                self.assertIn("⚠️", dlg.table.item(0, 0).text())
                self.assertIn("📋", dlg.table.item(0, 0).text())

                # Jump strip should be visible (not hidden) for the 2-message TDoc group
                self.assertFalse(dlg.detail_jump_strip_lbl.isHidden())
                self.assertIn("jump://msg/1", dlg.detail_jump_strip_lbl.text())
                dlg._on_anchor_clicked(QUrl("jump://msg/1"))

                # Priority mode should include this item (since My Status is '🟢 Support' and unread)
                dlg._set_mode("PRIORITY")
                self.assertEqual(dlg.table.rowCount(), 1)

                # Mark as Objected -> parks out of Priority mode!
                dlg._on_toggle_objected_clicked()
                self.assertTrue(udb.get_all_user_data()["S2-2601111"]["is_objected"])
                self.assertEqual(udb.get_all_user_data()["S2-2601111"]["objected_at_rev"], "r01")
                self.assertEqual(dlg.table.rowCount(), 0)

                # Switch to All mode and change My Status -> two-way callback fires
                dlg._set_mode("ALL")
                self.assertEqual(dlg.table.rowCount(), 1)
                idx_obj = dlg.detail_status_combo.findText("🔴 Object")
                dlg.detail_status_combo.setCurrentIndex(idx_obj)
                self.assertTrue(any(c[0] == "S2-2601111" and c[1] == "🔴 Object" for c in synced_calls))

                # Build generic statistics dataset
                stats_data = build_unified_email_stats_dataset(
                    all_inbox_items=dlg._all_items,
                    tdoc_ai_lookup=dlg.tdoc_ai_lookup,
                    tdoc_meta_map=dlg._tdoc_meta_map,
                )
                self.assertEqual(len(stats_data), 2)
                self.assertEqual(stats_data[0]["agenda_item"], "6.1")
            finally:
                dlg.close()

    def test_subfolder_recursion_deduplication_and_folders_dialog(self):
        from unittest.mock import patch
        from modules.emails.core.outlook_client import OutlookClient
        from modules.emails.ui.general_email_dialog import GeneralEmailFoldersDialog

        class _FakeFolder:
            def __init__(self, name: str, path: str, children=None):
                self.Name = name
                self.FolderPath = path
                self.Folders = list(children or [])

        sub_ai1 = _FakeFolder("AI_6_1", "user@domain/Inbox/SA2/AI_6_1")
        sub_ai2 = _FakeFolder("AI_6_2", "user@domain/Inbox/SA2/AI_6_2")
        sa2_folder = _FakeFolder("SA2", "user@domain/Inbox/SA2", [sub_ai1, sub_ai2])
        other_sub = _FakeFolder("Other", "user@domain/Inbox/Other")
        inbox_folder = _FakeFolder("Inbox", "user@domain/Inbox", [sa2_folder, other_sub])

        visited = set()
        # 1. Expand specific SA2 folder with include_subfolders=True first
        sa2_tree = OutlookClient.iter_folder_tree(
            sa2_folder,
            include_subfolders=True,
            visited_keys=visited,
        )
        self.assertEqual([f.Name for f in sa2_tree], ["SA2", "AI_6_1", "AI_6_2"])

        # 2. Expand parent Inbox with include_subfolders=True -> SA2 and its subfolders are skipped (not scanned twice)
        inbox_tree = OutlookClient.iter_folder_tree(
            inbox_folder,
            include_subfolders=True,
            visited_keys=visited,
        )
        self.assertEqual([f.Name for f in inbox_tree], ["Inbox", "Other"])

        # 3. Expand with include_subfolders=False -> only top folder returned
        shallow = OutlookClient.iter_folder_tree(sa2_folder, include_subfolders=False)
        self.assertEqual([f.Name for f in shallow], ["SA2"])

        # 4. Verify GeneralEmailFoldersDialog 4-column UI and include_subfolders persistence
        saved_payload = []
        sample_cfg = [
            {"folder_path": "user@domain/Inbox/Standardisierung/SA2", "tag": "WG", "color": "#0078D7"},
            {"folder_path": "user@domain/Inbox", "tag": "Inbox", "color": "#E20074", "include_subfolders": False},
        ]
        with patch("modules.emails.ui.general_email_dialog.load_wg_email_config", return_value=sample_cfg), patch(
            "modules.emails.ui.general_email_dialog.save_wg_email_config",
            side_effect=lambda wg, flist: saved_payload.extend(flist),
        ):
            dlg = GeneralEmailFoldersDialog("SA2")
            try:
                self.assertEqual(dlg.table.columnCount(), 4)
                self.assertEqual(dlg.table.rowCount(), 2)
                self.assertTrue(dlg._get_row_include_subfolders(0))
                self.assertFalse(dlg._get_row_include_subfolders(1))
                dlg._save_and_close()
                self.assertEqual(len(saved_payload), 2)
                self.assertTrue(saved_payload[0]["include_subfolders"])
                self.assertFalse(saved_payload[1]["include_subfolders"])
            finally:
                dlg.close()


if __name__ == "__main__":
    unittest.main()
