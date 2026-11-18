import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    linkify_markdown_urls,
    unwrap_outlook_safelinks,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import (
    _extract_folder_paths,
    auto_suggest_links,
    normalize_subject,
)


def test_normalize_subject_strips_prefixes_and_tags():
    assert normalize_subject("Re: [SA2#178] FS_6G_ARC: KI#1_1: Interim conclusion") == "FS_6G_ARC: KI#1_1: Interim conclusion"
    assert normalize_subject("AW: Fwd: [SA2#178]   KI#1_1 NAS Routing  ") == "KI#1_1 NAS Routing"
    assert normalize_subject("Re: [SA2#178, FS_6G_ARC, KI#1.1] Interim conclusion for KI#1.1 NAS Routing") == "Interim conclusion for KI#1.1 NAS Routing"
    assert _extract_folder_paths([{"folder_path": "a/Inbox", "tag": "Inbox"}, "b/SA2"]) == ["a/Inbox", "b/SA2"]
    links = auto_suggest_links("Re: [SA2#178] FS_6G_ARC: KI#1_1: Interim conclusion", "See S2-2601234 for details.")
    assert "KI#1_1" in links["ais"]
    assert "FS_6G_ARC" in links["wis"]
    assert "S2-2601234" in links["tdocs"]


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


def test_pill_autocomplete_and_tdoc_family_expansion(tmp_path):
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])

    from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog

    db = TrackedThreadsDatabase(tmp_path / "test_pills.db")
    tid = db.create_or_get_thread(
        title="NAS Routing",
        normalized_subject="Interim conclusion for KI#1.1 NAS Routing",
        meeting_id="SA2#178",
        wg="SA2",
    )

    def _expand(tdocs):
        out = []
        for t in tdocs:
            tu = t.upper()
            out.append(tu)
            if tu == "S2-2612673":
                out.append("S2-2612999")
        return list(dict.fromkeys(out))

    dlg = TrackedThreadDialog(
        db,
        tid,
        ai_suggestions=["20.6.1 — FS_6G_ARC: KI#1_1 NAS Routing"],
        wi_suggestions=["FS_6G_ARC"],
        tdoc_suggestions=["S2-2612673 — Proposal for interim agreements"],
        tdoc_family_expander=_expand,
    )
    dlg.ais_pills._on_completer_activated("20.6.1 — FS_6G_ARC: KI#1_1 NAS Routing")
    dlg.wis_pills._on_completer_activated("FS_6G_ARC")
    dlg.tdocs_pills._on_completer_activated("S2-2612673 — Proposal for interim agreements")

    updated = db.get_thread(tid)
    assert updated["linked_ais"] == ["20.6.1"]
    assert updated["linked_wis"] == ["FS_6G_ARC"]
    assert updated["linked_tdocs"] == ["S2-2612673", "S2-2612999"]

    # Verify removing a pill and processing DeferredDelete does NOT delete input_edit
    from PyQt5.QtCore import QEvent
    from PyQt5.QtWidgets import QPushButton
    from modules.meetings.ui.topic_pills import TagPillItem
    pills = dlg.ais_pills.findChildren(TagPillItem)
    assert len(pills) == 1
    pills[0].findChild(QPushButton).click()
    app.sendPostedEvents(None, QEvent.DeferredDelete)
    app.processEvents()
    assert dlg.ais_pills.tags() == []
    assert dlg.ais_pills.input_edit.text() == ""
    dlg.ais_pills._on_completer_activated("20.6.1 — FS_6G_ARC: KI#1_1 NAS Routing")
    assert db.get_thread(tid)["linked_ais"] == ["20.6.1"]
    dlg.close()


def test_safelinks_unwrapping_and_clickable_links():
    raw = (
        "Check https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178/Docs/S2-2612673.zip"
        "<https://eur03.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.3gpp.org%2Fftp%2Ftsg_sa%2FWG2_Arch%2FTSGS2_178%2FDocs%2FS2-2612673.zip&data=05%7C02%7C> "
        "and also https://eur01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fportal.3gpp.org%2FHome.aspx&data=123."
    )
    unwrapped = unwrap_outlook_safelinks(raw)
    assert "safelinks.protection.outlook.com" not in unwrapped
    assert unwrapped.count("https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178/Docs/S2-2612673.zip") == 1
    assert "https://portal.3gpp.org/Home.aspx." in unwrapped

    linked = linkify_markdown_urls(raw)
    assert "[https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178/Docs/S2-2612673.zip](https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178/Docs/S2-2612673.zip)" in linked
    assert "[https://portal.3gpp.org/Home.aspx](https://portal.3gpp.org/Home.aspx)." in linked
