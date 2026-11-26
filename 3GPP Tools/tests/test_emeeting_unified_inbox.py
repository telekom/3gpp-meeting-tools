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

    def test_underscore_tdoc_ai_wi_extraction_and_ai_to_wi_inference(self):
        from types import SimpleNamespace
        from modules.emails.core.general_email_sync import GeneralEmailSyncThread
        from modules.emails.core.tracked_threads.parsing import auto_suggest_links, infer_wis_for_ais

        subj = (
            "Re: [SA2#178, FS_enIMS_ARC, AI#20.10.1, S2-2612919_was_S2-2612519] "
            "23.801-04: [KI#3, New Solution] Association and correlation"
        )
        atts = ["S2-2613005_S2-2612919_was_S2-2612519_New Solution for KI3.docx"]
        agenda_map = {
            "20.10.1": SimpleNamespace(description="KI#3, FS_enIMS_ARC New Solution"),
            "6.1": SimpleNamespace(description="FS_XRM Monitoring"),
        }
        tdoc_meta_map = {
            "S2-2612919": {"ai": "20.10.1", "sub_item": "KI#3", "wis": ["FS_enIMS_ARC"]},
            "S2-2612519": {"ai": "20.10.1", "sub_item": "KI#3", "wis": ["FS_enIMS_ARC"]},
            "S2-2601111": {"ai": "6.1", "sub_item": "KI#1", "wis": ["FS_XRM"]},
        }

        matches = GeneralEmailSyncThread.extract_tdocs_from_parts(subj, "Please see attached.", attachments=atts)
        matched_ids = {m["tdoc_id"]: m["location"] for m in matches}
        self.assertEqual(matched_ids.get("S2-2612919"), "Subject")
        self.assertEqual(matched_ids.get("S2-2612519"), "Subject")
        self.assertEqual(matched_ids.get("S2-2613005"), "Body")

        sugg = auto_suggest_links(
            subj,
            "",
            agenda_map=agenda_map,
            valid_ais=["20.10.1", "6.1"],
            valid_wis=["FS_enIMS_ARC", "FS_XRM"],
            tdoc_meta_map=tdoc_meta_map,
            attachments=atts,
        )
        self.assertEqual(sugg["ais"], ["20.10.1"])
        self.assertEqual(sugg["wis"], ["FS_enIMS_ARC"])
        self.assertEqual(sugg["tdocs"], ["S2-2612519", "S2-2612919", "S2-2613005"])

        # AI -> WI inference even when subject has NO WI and NO TDocs!
        inferred = infer_wis_for_ais(
            ["6.1"],
            agenda_map=agenda_map,
            tdoc_meta_map=tdoc_meta_map,
            valid_wis=["FS_enIMS_ARC", "FS_XRM"],
        )
        self.assertEqual(inferred, ["FS_XRM"])

    def test_lossless_deduplication_across_general_and_tracked_dbs(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            agenda_dir = Path(tmpdir) / "Agenda"
            gdb = GeneralEmailDatabase(agenda_dir / "emails.db")
            tdb = TrackedThreadsDatabase(agenda_dir / "tracked_threads.db")

            # Simulate colleague's duplicate scenario: same email in 2 folders with different EntryIDs
            # One copy is Read with attachment, the other is Unread without attachment
            gdb.save_emails_batch(
                [
                    {
                        "id": "EID-INBOX-12",
                        "folder_tag": "Inbox",
                        "subject": "Re: [SA2#178] S2-2612919",
                        "sender_name": "Pallab Gupta (Nokia)",
                        "sender_email": "pallab@nokia.com",
                        "company": "Nokia",
                        "date_received": "2026-11-20 14:57:05",
                        "body_text": "Hi Nishant, I have tried to provide some clarification text.",
                        "clean_body": "Hi Nishant, I have tried to provide some clarification text.",
                        "attachments": ["S2-2613005_S2-2612919_was_S2-2612519.docx"],
                        "is_read": 1,
                        "flag_status": 0,
                    },
                    {
                        "id": "EID-WG-13",
                        "folder_tag": "WG",
                        "subject": "Re: [SA2#178] S2-2612919",
                        "sender_name": "Pallab Gupta",
                        "sender_email": "pallab@nokia.com",
                        "company": "Nokia",
                        "date_received": "2026-11-20 14:57:05",
                        "body_text": "Hi Nishant, I have tried to provide some clarification text.",
                        "clean_body": "Hi Nishant, I have tried to provide some clarification text.",
                        "attachments": [],
                        "is_read": 0,
                        "flag_status": 2,
                    },
                ],
                [
                    {"email_id": "EID-INBOX-12", "tdoc_id": "S2-2612919", "rev_matched": "", "match_location": "Subject"},
                    {"email_id": "EID-WG-13", "tdoc_id": "S2-2612519", "rev_matched": "", "match_location": "Quoted"},
                ],
                sync_unread=True,
            )

            rows = gdb.get_all_emails_with_tdocs()
            self.assertEqual(len(rows), 1)
            merged = rows[0]
            self.assertEqual(merged["is_read"], 1)
            self.assertEqual(merged["flag_status"], 2)
            self.assertEqual(merged["attachments"], ["S2-2613005_S2-2612919_was_S2-2612519.docx"])
            matched_set = set(merged["matched_tdocs"].split(","))
            self.assertEqual(matched_set, {"S2-2612919", "S2-2612519"})

            # TrackedThreadMessages deduplication when merging or upserting overlapping messages
            tid = tdb.create_or_get_thread("Thread X", "thread x", meeting_id="SA2#178", wg="SA2")
            tdb.upsert_messages(
                tid,
                [
                    {
                        "id": "MSG-LOCAL-1",
                        "subject": "Thread X",
                        "sender_name": "Nishant Gupta (QSIO)",
                        "sender_email": "nishant@qti.qualcomm.com",
                        "company": "Qualcomm",
                        "date_sent": "2026-11-20 06:19:12",
                        "body_text": "Please add the following ENs.",
                        "clean_body": "Please add the following ENs.",
                        "is_read": 0,
                        "attachments": [],
                    },
                    {
                        "id": "EID-OUTLOOK-11",
                        "subject": "Thread X",
                        "sender_name": "Nishant Gupta",
                        "sender_email": "nishant@qti.qualcomm.com",
                        "company": "Qualcomm",
                        "date_sent": "2026-11-20 06:19:12",
                        "body_text": "Please add the following ENs.",
                        "clean_body": "Please add the following ENs.",
                        "is_read": 1,
                        "attachments": ["EN.docx"],
                    },
                ],
                sync_unread=True,
            )
            tmsgs = tdb.get_thread_messages(tid)
            self.assertEqual(len(tmsgs), 1)
            self.assertEqual(tmsgs[0]["id"], "EID-OUTLOOK-11")
            self.assertEqual(tmsgs[0]["is_read"], 1)
            self.assertEqual(tmsgs[0]["attachments"], ["EN.docx"])

    def test_batch_multi_thread_drop_and_dual_db_indexing(self):
        from modules.emails.core.tracked_threads.collection_and_pinning import pin_threads_from_seeds_batch

        with tempfile.TemporaryDirectory() as tmpdir:
            agenda_dir = Path(tmpdir) / "Agenda"
            gdb = GeneralEmailDatabase(agenda_dir / "emails.db")
            tdb = TrackedThreadsDatabase(agenda_dir / "tracked_threads.db")

            tdoc_meta_map = {
                "S2-2612919": {"ai": "20.10.1", "sub_item": "KI#3", "wis": ["FS_enIMS_ARC"]},
                "S2-2609999": {"ai": "6.1", "sub_item": "KI#1", "wis": ["FS_XRM"]},
            }

            seeds = [
                (
                    {
                        "id": "M1",
                        "subject": "[SA2#178, AI#20.10.1, S2-2612919_was_S2-2612519] Solution",
                        "sender_name": "Alice",
                        "sender_email": "alice@nokia.com",
                        "company": "Nokia",
                        "date_sent": "2026-11-21 10:00:00",
                        "body_text": "Initial draft.",
                        "clean_body": "Initial draft.",
                        "conversation_id": "CID-A",
                        "is_read": 0,
                        "flag_status": 0,
                        "attachments": [],
                    },
                    None,
                ),
                (
                    {
                        "id": "M2",
                        "subject": "Re: [SA2#178, AI#20.10.1, S2-2612919_was_S2-2612519] Solution",
                        "sender_name": "Bob",
                        "sender_email": "bob@qualcomm.com",
                        "company": "Qualcomm",
                        "date_sent": "2026-11-21 10:05:00",
                        "body_text": "We support.",
                        "clean_body": "We support.",
                        "conversation_id": "CID-A",
                        "is_read": 0,
                        "flag_status": 0,
                        "attachments": [],
                    },
                    None,
                ),
                (
                    {
                        "id": "M3",
                        "subject": "[SA2#178, AI 6.1] S2-2609999 XRM Discussion",
                        "sender_name": "Carol",
                        "sender_email": "carol@huawei.com",
                        "company": "Huawei",
                        "date_sent": "2026-11-21 11:00:00",
                        "body_text": "Please review.",
                        "clean_body": "Please review.",
                        "conversation_id": "CID-B",
                        "is_read": 0,
                        "flag_status": 0,
                        "attachments": [],
                    },
                    None,
                ),
            ]

            res = pin_threads_from_seeds_batch(
                tdb,
                seeds,
                meeting_id="SA2#178",
                wg="SA2",
                valid_wis=["FS_enIMS_ARC", "FS_XRM"],
                tdoc_meta_map=tdoc_meta_map,
                general_email_db=gdb,
                refresh_outlook=False,
            )
            self.assertEqual(len(res["thread_ids"]), 2)
            self.assertEqual(res["message_count"], 3)

            threads = tdb.get_threads_for_meeting("SA2#178")
            self.assertEqual(len(threads), 2)
            th_a = next(t for t in threads if "CID-A" in t["conversation_ids"])
            th_b = next(t for t in threads if "CID-B" in t["conversation_ids"])
            self.assertEqual(th_a["linked_ais"], ["20.10.1"])
            self.assertEqual(th_a["linked_wis"], ["FS_enIMS_ARC"])
            self.assertEqual(th_b["linked_ais"], ["6.1"])
            self.assertEqual(th_b["linked_wis"], ["FS_XRM"])
            self.assertEqual(len(tdb.get_thread_messages(th_a["thread_id"])), 2)

            # Verify GeneralEmailDatabase was also indexed cleanly for Related Emails dialogs!
            gen_rows = gdb.get_all_emails_with_tdocs()
            self.assertEqual(len(gen_rows), 3)
            self.assertTrue(any("S2-2612919" in r["matched_tdocs"] for r in gen_rows))
            self.assertTrue(any("S2-2609999" in r["matched_tdocs"] for r in gen_rows))

    def test_startup_cleanup_on_dirty_dbs_tdoc_links_and_live_ai_to_wi(self):
        from modules.emails.core.thread_markdown_exporter import linkify_tdoc_tokens

        # 1. Verify underscore-delimited TDoc linkification in Markdown viewer
        linked = linkify_tdoc_tokens("Re: S2-2612919_was_S2-2612519")
        self.assertIn("[S2-2612919](tdoc:S2-2612919)", linked)
        self.assertIn("[S2-2612519](tdoc:S2-2612519)", linked)

        # 2. Verify opening pre-existing dirty DBs cleans up duplicates automatically on startup
        with tempfile.TemporaryDirectory() as tmpdir:
            agenda_dir = Path(tmpdir) / "Agenda"
            agenda_dir.mkdir(parents=True, exist_ok=True)
            edb_path = agenda_dir / "emails.db"
            tdb_path = agenda_dir / "tracked_threads.db"

            # Seed dirty general_emails + duplicate tracked_threads directly via raw SQLite
            conn_e = sqlite3.connect(edb_path)
            conn_e.execute("""
                CREATE TABLE general_emails (
                    id TEXT PRIMARY KEY, folder_tag TEXT, subject TEXT, sender_name TEXT,
                    sender_email TEXT, company TEXT, date_received TEXT, body_text TEXT,
                    clean_body TEXT DEFAULT '', attachments TEXT DEFAULT '[]',
                    is_read INTEGER DEFAULT 0, is_ignored INTEGER DEFAULT 0, flag_status INTEGER DEFAULT 0
                )
            """)
            conn_e.execute("""
                CREATE TABLE general_email_tdoc_matches (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, email_id TEXT, tdoc_id TEXT,
                    rev_matched TEXT, match_location TEXT
                )
            """)
            conn_e.execute(
                "INSERT INTO general_emails VALUES ('OLD-1','Inbox','Subj','Pallab Gupta','p@n.com','Nokia','2026-11-20 14:57:00','Body text','Body text','[\"a.docx\"]',1,0,0)"
            )
            conn_e.execute(
                "INSERT INTO general_emails VALUES ('OLD-2','WG','Subj','Pallab Gupta','p@n.com','Nokia','2026-11-20 14:57:59','Body text','Body text','[]',0,0,2)"
            )
            conn_e.execute("INSERT INTO general_email_tdoc_matches (email_id, tdoc_id, rev_matched, match_location) VALUES ('OLD-1','S2-2612919','','Subject')")
            conn_e.execute("INSERT INTO general_email_tdoc_matches (email_id, tdoc_id, rev_matched, match_location) VALUES ('OLD-2','S2-2612519','','Quoted')")
            conn_e.commit()
            conn_e.close()

            gdb = GeneralEmailDatabase(edb_path)
            cleaned_gen = gdb.get_all_emails_with_tdocs()
            self.assertEqual(len(cleaned_gen), 1)
            self.assertEqual(cleaned_gen[0]["is_read"], 1)
            self.assertEqual(cleaned_gen[0]["flag_status"], 2)
            self.assertEqual(cleaned_gen[0]["attachments"], ["a.docx"])

            # Seed dirty tracked_threads with duplicate thread subjects and duplicate messages
            conn_t = sqlite3.connect(tdb_path)
            conn_t.execute("""
                CREATE TABLE tracked_threads (
                    thread_id TEXT PRIMARY KEY, title TEXT NOT NULL, normalized_subject TEXT NOT NULL,
                    conversation_ids TEXT DEFAULT '[]', subject_patterns TEXT DEFAULT '[]',
                    outlook_folders TEXT DEFAULT '[]', meeting_id TEXT DEFAULT '', wg TEXT DEFAULT '',
                    linked_ais TEXT DEFAULT '[]', linked_wis TEXT DEFAULT '[]', linked_tdocs TEXT DEFAULT '[]',
                    linked_sub_items TEXT DEFAULT '[]', unread_count INTEGER DEFAULT 0, message_count INTEGER DEFAULT 0,
                    last_message_date TEXT DEFAULT '', last_synced_at TEXT DEFAULT '', created_at TEXT DEFAULT '',
                    user_status TEXT DEFAULT '⚪ Neutral', user_notes TEXT DEFAULT '', user_tags TEXT DEFAULT '[]',
                    flagged_count INTEGER DEFAULT 0, is_objected INTEGER DEFAULT 0, objected_at_rev TEXT DEFAULT ''
                )
            """)
            conn_t.execute("""
                CREATE TABLE tracked_thread_messages (
                    id TEXT PRIMARY KEY, thread_id TEXT NOT NULL, subject TEXT DEFAULT '',
                    sender_name TEXT DEFAULT '', sender_email TEXT DEFAULT '', company TEXT DEFAULT '',
                    date_sent TEXT DEFAULT '', body_text TEXT DEFAULT '', clean_body TEXT DEFAULT '',
                    message_id_header TEXT DEFAULT '', in_reply_to TEXT DEFAULT '', conversation_id TEXT DEFAULT '',
                    is_read INTEGER DEFAULT 0, flag_status INTEGER DEFAULT 0, attachments TEXT DEFAULT '[]'
                )
            """)
            conn_t.execute("INSERT INTO tracked_threads (thread_id, title, normalized_subject, meeting_id, message_count, last_message_date) VALUES ('THR-1','Same Topic','same topic','SA2#178',1,'2026-11-20 14:57:00')")
            conn_t.execute("INSERT INTO tracked_threads (thread_id, title, normalized_subject, meeting_id, message_count, last_message_date) VALUES ('THR-2','Same Topic','same topic','SA2#178',1,'2026-11-20 14:57:00')")
            conn_t.execute("INSERT INTO tracked_thread_messages (id, thread_id, subject, sender_name, sender_email, date_sent, body_text, clean_body, is_read) VALUES ('TM-1','THR-1','Same Topic','Pallab','p@n.com','2026-11-20 14:57:00','Hello','Hello',1)")
            conn_t.execute("INSERT INTO tracked_thread_messages (id, thread_id, subject, sender_name, sender_email, date_sent, body_text, clean_body, is_read) VALUES ('TM-2','THR-2','Same Topic','Pallab','p@n.com','2026-11-20 14:57:00','Hello','Hello',0)")
            conn_t.commit()
            conn_t.close()

            tdb = TrackedThreadsDatabase(tdb_path)
            threads_after = tdb.get_threads_for_meeting("SA2#178")
            self.assertEqual(len(threads_after), 1)
            msgs_after = tdb.get_thread_messages(threads_after[0]["thread_id"])
            self.assertEqual(len(msgs_after), 1)
            self.assertEqual(msgs_after[0]["is_read"], 1)

            # 3. Verify live AI -> WI inference in MeetingEmailInboxDialog inline assignment bar
            tdb.upsert_messages(
                threads_after[0]["thread_id"],
                [{
                    "id": "TM-UNLINKED",
                    "subject": "General topic",
                    "sender_name": "Alice",
                    "sender_email": "alice@test.com",
                    "company": "Nokia",
                    "date_sent": "2026-11-21 10:00:00",
                    "body_text": "No WI in body.",
                    "clean_body": "No WI in body.",
                    "is_read": 0,
                }],
            )
            settings = QSettings(str(agenda_dir / "ai_wi_test.ini"), QSettings.IniFormat)
            dlg = MeetingEmailInboxDialog(
                tracked_db=tdb,
                general_email_db=gdb,
                meeting_id="SA2#178",
                tdoc_meta_map={"S2-2601111": {"ai": "6.1", "wis": ["FS_XRM"]}},
                ai_suggestions=["6.1 — XRM Study"],
                wi_suggestions=["FS_XRM"],
                settings=settings,
            )
            try:
                dlg._set_mode("ALL")
                dlg.assign_ais_pills.setText("6.1")
                dlg._on_inline_links_changed()
                self.assertIn("FS_XRM", dlg.assign_wis_pills.tags())
            finally:
                dlg.close()

    def test_sync_watermark_persistence_and_restart_gap(self):
        from datetime import datetime, timedelta, timezone
        from modules.emails.core.sync_watermark import (
            LAST_SYNC_META_KEY,
            format_utc_iso,
            resolve_sync_window,
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = Path(tmpdir) / "Agenda" / "emails.db"
            gdb = GeneralEmailDatabase(db_path)
            last_dt = datetime(2026, 11, 17, 10, 0, 0, tzinfo=timezone.utc)
            gdb.set_sync_meta(LAST_SYNC_META_KEY, format_utc_iso(last_dt))

            # Simulate closing the app / restarting laptop and reopening 3 days later
            gdb_reopened = GeneralEmailDatabase(db_path)
            now_dt = datetime(2026, 11, 20, 12, 30, 0, tzinfo=timezone.utc)
            win = resolve_sync_window(
                gdb_reopened,
                meeting_start="2026-11-16",
                meeting_end="2026-11-20",
                sync_mode="incremental",
                now_utc=now_dt,
            )
            self.assertFalse(win["is_first_sync"])
            self.assertEqual(win["since_dt_utc"], last_dt - timedelta(minutes=5))
            self.assertEqual(win["start_date"], "")
            self.assertEqual(win["end_date"], "")
            self.assertIn("catch-up 3d", win["gap_description"])

            # Wiping DB resets watermark cleanly
            gdb_reopened.wipe_generic_emails()
            self.assertEqual(gdb_reopened.get_sync_meta(LAST_SYNC_META_KEY, ""), "")

    def test_first_time_sync_uses_meeting_window_not_all_history(self):
        from datetime import datetime, timezone
        from modules.emails.core.sync_watermark import resolve_sync_window

        with tempfile.TemporaryDirectory() as tmpdir:
            gdb = GeneralEmailDatabase(Path(tmpdir) / "Agenda" / "emails.db")
            now_dt = datetime(2026, 11, 20, 12, 0, 0, tzinfo=timezone.utc)

            # First incremental sync with no watermark falls back to meeting window - 3 days
            win_inc = resolve_sync_window(
                gdb,
                meeting_start="2026-11-16",
                meeting_end="2026-11-20",
                sync_mode="incremental",
                days_buffer=3,
                now_utc=now_dt,
            )
            self.assertTrue(win_inc["is_first_sync"])
            self.assertEqual(win_inc["start_date"], "2026-11-16")
            self.assertEqual(win_inc["end_date"], "2026-11-20")
            self.assertEqual(win_inc["days_buffer"], 3)
            self.assertIsNotNone(win_inc["since_dt_utc"])
            self.assertLessEqual(win_inc["since_dt_utc"].date().isoformat(), "2026-11-13")

            # Explicit Full Sync keeps since_dt_utc=None and honors explicit dates/buffer
            win_full = resolve_sync_window(
                gdb,
                meeting_start="2026-11-10",
                meeting_end="2026-11-14",
                sync_mode="full",
                days_buffer=5,
                now_utc=now_dt,
            )
            self.assertEqual(win_full["mode"], "full")
            self.assertIsNone(win_full["since_dt_utc"])
            self.assertEqual(win_full["start_date"], "2026-11-10")
            self.assertEqual(win_full["end_date"], "2026-11-14")
            self.assertEqual(win_full["days_buffer"], 5)

            # Default buffer when omitted is ±5 days
            win_default = resolve_sync_window(
                gdb,
                meeting_start="2026-11-16",
                meeting_end="2026-11-20",
                sync_mode="incremental",
                now_utc=now_dt,
            )
            self.assertEqual(win_default["days_buffer"], 5)
            self.assertLessEqual(win_default["since_dt_utc"].date().isoformat(), "2026-11-11")

    def test_apply_mapi_since_restriction_dasl_and_fallback(self):
        from datetime import datetime, timedelta, timezone
        from modules.emails.core.sync_watermark import (
            apply_mapi_since_restriction,
            format_dasl_since_filter,
        )

        tz_plus1 = timezone(timedelta(hours=1))
        local_dt = datetime(2026, 11, 20, 14, 55, 0, tzinfo=tz_plus1)
        dasl_inbox = format_dasl_since_filter(local_dt, is_sent_folder=False)
        self.assertEqual(dasl_inbox, '@SQL="urn:schemas:httpmail:datereceived" >= \'2026-11-20 13:55:00\'')
        dasl_sent = format_dasl_since_filter(local_dt, is_sent_folder=True)
        self.assertEqual(dasl_sent, '@SQL="urn:schemas:httpmail:date" >= \'2026-11-20 13:55:00\'')

        class _FilteredTable:
            pass

        filtered_tbl = _FilteredTable()
        seen_filters = []

        class _RawTable:
            def Restrict(self, filt):
                seen_filters.append(filt)
                return filtered_tbl

        out_tbl, applied = apply_mapi_since_restriction(_RawTable(), local_dt)
        self.assertTrue(applied)
        self.assertIs(out_tbl, filtered_tbl)
        self.assertEqual(seen_filters, [dasl_inbox])

        class _FailingTable:
            def Restrict(self, _filt):
                raise RuntimeError("store does not support DASL Restrict")

        raw_fail = _FailingTable()
        out_fail, applied_fail = apply_mapi_since_restriction(raw_fail, local_dt)
        self.assertFalse(applied_fail)
        self.assertIs(out_fail, raw_fail)

    def test_inbox_sync_worker_incremental_writes_watermark_and_scoped_preserves(self):
        from datetime import datetime
        from unittest.mock import patch
        from modules.emails.core.outlook_client import OutlookClient
        from modules.emails.core.sync_watermark import LAST_SYNC_META_KEY
        from modules.emails.core.tracked_threads_workers import InboxOutlookSyncWorker

        with tempfile.TemporaryDirectory() as tmpdir:
            agenda_dir = Path(tmpdir) / "Agenda"
            gdb = GeneralEmailDatabase(agenda_dir / "emails.db")
            tdb = TrackedThreadsDatabase(agenda_dir / "tracked_threads.db")

            restrict_calls = []
            recent_row = {
                "EntryID": "EID-RECENT-1",
                "MessageClass": "IPM.Note",
                "ReceivedTime": datetime(2026, 11, 22, 12, 0, 0),
                "UnRead": True,
                "FlagStatus": 0,
            }

            class _FakeMail:
                Class = 43
                EntryID = "EID-RECENT-1"
                Subject = "Re: [SA2#178] S2-2612999 Update"
                Body = "Looks good for S2-2612999."
                HTMLBody = ""
                SenderName = "Alice"
                SenderEmailAddress = "alice@nokia.com"
                SenderEmailType = "SMTP"
                ReceivedTime = datetime(2026, 11, 22, 12, 0, 0)
                UnRead = True
                FlagStatus = 0
                Attachments = []

            class _FakeRow:
                def __init__(self, data):
                    self._data = data

                def __call__(self, key):
                    if key not in self._data:
                        raise KeyError(key)
                    return self._data[key]

            class _FakeTable:
                def __init__(self, rows):
                    self._rows = list(rows)
                    self._idx = 0

                def Restrict(self, filt):
                    restrict_calls.append(filt)
                    # Server-side restriction leaves only the 1 recent row instead of 500 historical ones
                    return _FakeTable([recent_row])

                def Sort(self, *_args):
                    pass

                @property
                def Columns(self):
                    return type("Cols", (), {"Add": lambda *_a: None})()

                def GetRowCount(self):
                    return len(self._rows)

                @property
                def EndOfTable(self):
                    return self._idx >= len(self._rows)

                def GetNextRow(self):
                    r = self._rows[self._idx]
                    self._idx += 1
                    return _FakeRow(r)

            fake_folder = type(
                "FakeFolder",
                (),
                {
                    "Name": "WG",
                    "FolderPath": "\\\\me@example.com\\Inbox\\WG",
                    "Folders": [],
                    "GetTable": lambda self: _FakeTable([recent_row] * 50),
                },
            )()

            class _FakeNS:
                def GetDefaultFolder(self, _idx):
                    return type("F", (), {"FolderPath": ""})()

                def GetItemFromID(self, eid):
                    if eid == "EID-RECENT-1":
                        return _FakeMail()
                    raise KeyError(eid)

            with patch.object(OutlookClient, "get_namespace", return_value=_FakeNS()), patch.object(
                OutlookClient, "get_folder_by_path", return_value=fake_folder
            ), patch(
                "modules.emails.core.tracked_threads_workers.load_wg_email_config",
                return_value=[{"folder_path": "\\\\me@example.com\\Inbox\\WG", "tag": "WG"}],
            ):
                w = InboxOutlookSyncWorker(
                    tdb.db_path,
                    gdb.db_path,
                    meeting_id="SA2#178",
                    wg="SA2",
                    start_date="2026-11-16",
                    end_date="2026-11-22",
                    sync_mode="incremental",
                )
                results = []
                w.finished_sync.connect(lambda ok, nt, ng, err: results.append((ok, nt, ng, err)))
                w.run()

            self.assertTrue(restrict_calls)
            self.assertEqual(len(results), 1)
            self.assertTrue(results[0][0])
            self.assertGreaterEqual(results[0][2], 1)
            saved_wm = gdb.get_sync_meta(LAST_SYNC_META_KEY, "")
            self.assertTrue(saved_wm.endswith("Z"))

            # Scoped sync must NOT overwrite global last_sync_utc watermark
            sentinel = "2026-11-20T08:00:00Z"
            gdb.set_sync_meta(LAST_SYNC_META_KEY, sentinel)
            with patch.object(OutlookClient, "get_namespace", return_value=_FakeNS()), patch(
                "modules.emails.core.tracked_threads_workers.sync_scoped_inbox_items",
                return_value=(0, 1),
            ):
                w_scoped = InboxOutlookSyncWorker(
                    tdb.db_path,
                    gdb.db_path,
                    meeting_id="SA2#178",
                    wg="SA2",
                    scoped_thread_ids=[],
                    scoped_email_ids=["EID-RECENT-1"],
                )
                w_scoped.run()
            self.assertEqual(gdb.get_sync_meta(LAST_SYNC_META_KEY, ""), sentinel)

    def test_live_sync_toggle_persists_per_meeting(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            agenda_dir = Path(tmpdir) / "Agenda"
            gdb = GeneralEmailDatabase(agenda_dir / "emails.db")
            tdb = TrackedThreadsDatabase(agenda_dir / "tracked_threads.db")
            ini_settings = QSettings(str(agenda_dir / "live_sync.ini"), QSettings.IniFormat)

            dlg = MeetingEmailInboxDialog(
                tracked_db=tdb,
                general_email_db=gdb,
                meeting_id="SA2#178",
                settings=ini_settings,
            )
            emitted = []
            dlg.live_sync_toggled.connect(lambda val: emitted.append(bool(val)))
            try:
                self.assertEqual(dlg.sync_outlook_btn.text(), "⚡ Sync New")
                self.assertFalse(dlg.is_live_sync_enabled())
                self.assertFalse(dlg.btn_live_sync.isChecked())
                self.assertFalse(dlg._live_sync_timer.isActive())
                self.assertIsNotNone(dlg.findChild(object, "inboxTopBar"))
                self.assertIsNotNone(dlg.findChild(object, "inboxViewSegment"))
                self.assertIsNotNone(dlg.findChild(object, "inboxFilterBar"))
                self.assertEqual(
                    dlg.sync_outlook_btn.minimumHeight(),
                    dlg.btn_live_sync.minimumHeight(),
                )
                self.assertEqual(
                    dlg.btn_live_sync.minimumHeight(),
                    dlg.btn_stats.minimumHeight(),
                )

                dlg.search_edit.setText("test")
                dlg.btn_hide_objected.setChecked(True)
                dlg.btn_clear_filters.click()
                self.assertEqual(dlg.search_edit.text(), "")
                self.assertFalse(dlg.btn_hide_objected.isChecked())

                dlg.btn_live_sync.setChecked(True)
                self.assertTrue(dlg.is_live_sync_enabled())
                self.assertTrue(dlg._live_sync_timer.isActive())
                self.assertEqual(emitted, [True])
            finally:
                dlg.close()

            self.assertFalse(dlg._live_sync_timer.isActive())

            # Reopening Inbox for the same meeting restores Live Auto-Sync state automatically
            dlg2 = MeetingEmailInboxDialog(
                tracked_db=tdb,
                general_email_db=gdb,
                meeting_id="SA2#178",
                settings=ini_settings,
            )
            try:
                self.assertTrue(dlg2.is_live_sync_enabled())
                self.assertTrue(dlg2.btn_live_sync.isChecked())
                self.assertTrue(dlg2._live_sync_timer.isActive())
            finally:
                dlg2.close()


if __name__ == "__main__":
    unittest.main()
