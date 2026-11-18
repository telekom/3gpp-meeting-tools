import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from modules.emails.core.thread_markdown_exporter import export_thread_to_markdown
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import normalize_subject


def test_normalize_subject_strips_prefixes_and_tags():
    assert normalize_subject("Re: [SA2#178] FS_6G_ARC: KI#1_1: Interim conclusion") == "FS_6G_ARC: KI#1_1: Interim conclusion"
    assert normalize_subject("AW: Fwd: [SA2#178]   KI#1_1 NAS Routing  ") == "KI#1_1 NAS Routing"


def test_tracked_threads_db_crud_merge_and_read(tmp_path):
    db = TrackedThreadsDatabase(tmp_path / "test_threads.db")
    t1 = db.create_or_get_thread(
        title="KI#1_1 NAS Routing",
        normalized_subject="FS_6G_ARC: KI#1_1: Interim conclusion",
        conversation_id="CID-1",
        meeting_id="SA2#178",
        wg="SA2",
        linked_ais=["KI#1_1"],
        linked_wis=["FS_6G_ARC"],
    )
    t2 = db.create_or_get_thread(
        title="KI#1_1 NAS Routing (split)",
        normalized_subject="KI#1_1: Interim conclusion",
        conversation_id="CID-2",
        meeting_id="SA2#178",
        wg="SA2",
    )

    added1 = db.upsert_messages(
        t1,
        [
            {
                "id": "M1",
                "subject": "Re: [SA2#178] FS_6G_ARC: KI#1_1: Interim conclusion",
                "sender_name": "Alice (Nokia)",
                "sender_email": "alice@nokia.com",
                "company": "Nokia",
                "date_sent": "2026-11-18 10:00:00",
                "body_text": "We support Option 1 in S2-2601234.",
                "clean_body": "We support Option 1 in S2-2601234.",
                "is_read": 0,
            }
        ],
    )
    assert added1 == 1
    db.upsert_messages(
        t2,
        [
            {
                "id": "M2",
                "subject": "AW: KI#1_1: Interim conclusion",
                "sender_name": "Bob (Ericsson)",
                "sender_email": "bob@ericsson.com",
                "company": "Ericsson",
                "date_sent": "2026-11-18 11:00:00",
                "body_text": "We prefer Option 2.",
                "clean_body": "We prefer Option 2.",
                "is_read": 0,
            }
        ],
    )

    db.merge_threads(t1, [t2])
    assert db.get_thread(t2) is None
    merged = db.get_thread(t1)
    assert merged["message_count"] == 2
    assert merged["unread_count"] == 2
    assert set(merged["conversation_ids"]) == {"CID-1", "CID-2"}

    db.mark_thread_read(t1)
    assert db.get_thread(t1)["unread_count"] == 0

    msgs = db.get_thread_messages(t1)
    md = export_thread_to_markdown(merged, msgs)
    assert "# Email Thread: KI#1_1 NAS Routing" in md
    assert "S2-2601234" in md
    assert "Alice (Nokia)" in md
    assert "Bob (Ericsson)" in md
