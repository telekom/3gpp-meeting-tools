import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    extract_thread_document_urls,
    is_cacheable_document_url,
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
    assert normalize_subject("回复: [SA2#178, FS_6G_ARC, KI#2] [S2-2612666, KI#2, interim conclusion] Interim conclusion for KI#2.") == "Interim conclusion for KI#2."
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
        sub_item_suggestions=["NAS Routing/transport — AI 20.6.1"],
    )
    dlg.ais_pills._on_completer_activated("20.6.1 — FS_6G_ARC: KI#1_1 NAS Routing")
    dlg.sub_items_pills._on_completer_activated("NAS Routing/transport — AI 20.6.1")
    dlg.wis_pills._on_completer_activated("FS_6G_ARC")
    dlg.tdocs_pills._on_completer_activated("S2-2612673 — Proposal for interim agreements")

    updated = db.get_thread(tid)
    assert updated["linked_ais"] == ["20.6.1"]
    assert updated["linked_sub_items"] == ["NAS Routing/transport"]
    assert updated["linked_wis"] == ["FS_6G_ARC"]
    assert updated["linked_tdocs"] == ["S2-2612673", "S2-2612999"]

    from modules.meetings.ui.tdocs_models import TDocsTableModel
    m = TDocsTableModel(tmp_path, [], {}, {})
    m.sync_tracked_thread_rows([updated])
    assert m._data[0]["Sub-Item"] == "NAS Routing/transport"
    assert "NAS Routing/transport" in m._data[0]["_linked_sub_items"]

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

    thread = {"thread_id": "THR-1", "title": "Test", "normalized_subject": "Test"}
    msgs = [
        {
            "id": "00000000ABCD1234",
            "subject": "Test",
            "sender_name": "Alice",
            "company": "Nokia",
            "date_sent": "2026-11-18 12:00:00",
            "body_text": "See attached.",
            "clean_body": "See attached.",
            "attachments": ["S2-2612673_draft.docx"],
        }
    ]
    md_ui = export_thread_to_markdown(thread, msgs, interactive_links=True)
    assert "[📧 Open in Outlook](outlook://msg/00000000ABCD1234)" in md_ui
    assert "[📎 S2-2612673_draft.docx](outlook://att/00000000ABCD1234/S2-2612673_draft.docx)" in md_ui

    md_export = export_thread_to_markdown(thread, msgs, interactive_links=False)
    assert "outlook://" not in md_export
    assert "**Attachments:** S2-2612673_draft.docx" in md_export

    # Test 10.10.10.10 space-containing document links & Proofpoint urldefense unwrapping
    raw_local_doc = (
        "Draft placed at http://10.10.10.10/ftp/SA/SA2/Inbox/Drafts/FS_6G_ARC/KI#1/S2-2612888-KK#1.1 Conclusion-6GNAS-Routing-Baseline-v2.docx "
        "and https://urldefense.com/v3/__https:/www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178_Calgary_2026-11/Docs/S2-2611551.zip__;!!CTRNKA9wMg0ARbw!nCj6zB "
    )
    unwrapped_doc = unwrap_outlook_safelinks(raw_local_doc)
    assert "http://10.10.10.10/ftp/SA/SA2/Inbox/Drafts/FS_6G_ARC/KI%231/S2-2612888-KK%231.1%20Conclusion-6GNAS-Routing-Baseline-v2.docx" in unwrapped_doc
    assert "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178_Calgary_2026-11/Docs/S2-2611551.zip" in unwrapped_doc

    doc_urls = extract_thread_document_urls([{"body_text": raw_local_doc}])
    assert len(doc_urls) == 2
    assert all(is_cacheable_document_url(u) for u in doc_urls)

    # Test Safelink with doubly percent-encoded %23/%20/%5B/%5D and already-unquoted DB text with brackets
    bracket_safelink = (
        "Uploaded a draft <https://eur05.safelinks.protection.outlook.com/?url="
        "http%3A%2F%2F10.10.10.10%2Fftp%2FSA%2FSA2%2FInbox%2FDrafts%2FFS_6G_ARC%2FKI%25234%2F"
        "Draft%2520S2-2612711%2520%255BKI%25234%252C%2520bullet%25233%255D%2520Conclusion.docx"
        "&data=05%7C02%7Cx&sdata=abc&reserved=0> for review"
    )
    unwrapped_bracket = linkify_markdown_urls(bracket_safelink)
    expected_url = (
        "http://10.10.10.10/ftp/SA/SA2/Inbox/Drafts/FS_6G_ARC/KI%234/"
        "Draft%20S2-2612711%20%5BKI%234%2C%20bullet%233%5D%20Conclusion.docx"
    )
    assert f"[{expected_url}]({expected_url})" in unwrapped_bracket
    assert extract_thread_document_urls([{"body_text": bracket_safelink}]) == [expected_url]

    legacy_unquoted = (
        "Uploaded http://10.10.10.10/ftp/SA/SA2/Inbox/Drafts/FS_6G_ARC/KI#4/"
        "Draft S2-2612711 [KI#4, bullet#3] Conclusion.docx for review"
    )
    repaired = extract_thread_document_urls([{"body_text": legacy_unquoted}])
    assert len(repaired) == 1
    assert repaired[0].endswith("Conclusion.docx")
    assert "%5BKI%234,%20bullet%233%5D" in repaired[0]


def test_newest_first_ordering_filters_and_llm_workspace(tmp_path):
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])

    from modules.emails.core.tracked_threads_service import prepare_thread_llm_workspace
    from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog

    db = TrackedThreadsDatabase(tmp_path / "test_order.db")
    tid = db.create_or_get_thread("Order Thread", "Order Thread", meeting_id="SA2#178", wg="SA2")
    db.upsert_messages(
        tid,
        [
            {
                "id": "000000001111",
                "subject": "Order Thread",
                "sender_name": "Alice",
                "company": "Nokia",
                "date_sent": "2026-11-18 09:00:00",
                "body_text": "First view: Option A",
                "clean_body": "First view: Option A",
                "is_read": 1,
            },
            {
                "id": "000000002222",
                "subject": "Order Thread",
                "sender_name": "Bob",
                "company": "Ericsson",
                "date_sent": "2026-11-18 10:00:00",
                "body_text": "Second view: Option B",
                "clean_body": "Second view: Option B",
                "is_read": 0,
            },
        ],
    )
    thread = db.get_thread(tid)
    msgs = db.get_thread_messages(tid)

    md_newest = export_thread_to_markdown(thread, msgs, interactive_links=True, newest_first=True)
    assert "🔵 **NEW**" in md_newest
    assert "· ·" not in md_newest
    assert "outlook://reply/000000002222" in md_newest
    assert "outlook://copy/000000002222" in md_newest

    # Inline <Ashok> tag must be escaped in interactive UI Markdown so QTextBrowser does not swallow following lines
    msgs_with_tag = [
        {
            "id": "000000003333",
            "subject": "Order Thread",
            "sender_name": "Ashok",
            "company": "Samsung",
            "date_sent": "2026-11-18 11:00:00",
            "body_text": "1. Item one\n<Ashok> What does this mean?\n2. Item two",
            "clean_body": "1. Item one\n<Ashok> What does this mean?\n2. Item two",
            "is_read": 0,
        }
    ]
    md_tag = export_thread_to_markdown(thread, msgs_with_tag, interactive_links=True)
    assert "&lt;Ashok&gt; What does this mean?" in md_tag
    assert "2. Item two" in md_tag
    disc_newest = md_newest.split("## Discussion (Newest First)", 1)[1]
    assert disc_newest.index("Bob (Ericsson)") < disc_newest.index("Alice (Nokia)")

    md_oldest = export_thread_to_markdown(thread, msgs, interactive_links=False, newest_first=False)
    disc_oldest = md_oldest.split("## Chronological Discussion", 1)[1]
    assert disc_oldest.index("Alice (Nokia)") < disc_oldest.index("Bob (Ericsson)")

    dlg = TrackedThreadDialog(db, tid)
    dlg.show()
    dlg.msg_search_edit.setText("Option B")
    assert "1 of 2" in dlg.msg_count_lbl.text()
    dlg.msg_search_edit.clear()
    idx_nokia = dlg.company_filter_combo.findData("Nokia")
    assert idx_nokia >= 0
    dlg.company_filter_combo.setCurrentIndex(idx_nokia)
    assert "1 of 2" in dlg.msg_count_lbl.text()
    dlg.close()

    cache_dir = tmp_path / "cache"
    prep = prepare_thread_llm_workspace(db, tid, cache_dir, newest_first=False, cache_linked_docs=False)
    assert prep["thread_md"].exists()
    ws_txt = prep["thread_md"].read_text(encoding="utf-8")
    assert ws_txt.index("Alice (Nokia)") < ws_txt.index("Bob (Ericsson)")


def test_ctrl_wheel_reader_zoom_on_thread_browser(tmp_path):
    from PyQt5.QtCore import QEvent, QPoint, QPointF, Qt
    from PyQt5.QtGui import QKeyEvent, QWheelEvent
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])

    from core.ui.font_scale import install_reader_zoom
    from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog

    install_reader_zoom()
    db = TrackedThreadsDatabase(tmp_path / "test_zoom.db")
    tid = db.create_or_get_thread("Zoom Thread", "Zoom Thread", meeting_id="SA2#178", wg="SA2")
    dlg = TrackedThreadDialog(db, tid)
    dlg.show()

    vp = dlg.browser.viewport()
    base_steps = int(dlg.browser.property("_reader_zoom_steps") or 0)
    wheel_up = QWheelEvent(
        QPointF(10, 10),
        QPointF(10, 10),
        QPoint(0, 0),
        QPoint(0, 120),
        Qt.NoButton,
        Qt.ControlModifier,
        Qt.ScrollUpdate,
        False,
    )
    QApplication.sendEvent(vp, wheel_up)
    assert int(dlg.browser.property("_reader_zoom_steps") or 0) == base_steps + 1

    wheel_down = QWheelEvent(
        QPointF(10, 10),
        QPointF(10, 10),
        QPoint(0, 0),
        QPoint(0, -120),
        Qt.NoButton,
        Qt.ControlModifier,
        Qt.ScrollUpdate,
        False,
    )
    QApplication.sendEvent(vp, wheel_down)
    assert int(dlg.browser.property("_reader_zoom_steps") or 0) == base_steps

    QApplication.sendEvent(vp, wheel_up)
    assert int(dlg.browser.property("_reader_zoom_steps") or 0) == base_steps + 1

    ctrl_zero = QKeyEvent(QEvent.KeyPress, Qt.Key_0, Qt.ControlModifier)
    QApplication.sendEvent(dlg.browser, ctrl_zero)
    assert int(dlg.browser.property("_reader_zoom_steps") or 0) == 0
    dlg.close()


def test_unified_notes_status_and_tags_for_tdocs_and_threads(tmp_path):
    import sqlite3
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])

    from modules.meetings.core.tdocs.tdocs_db import TDocsDatabase
    from modules.meetings.ui.tdocs_dialogs import InteractiveNotesDialog
    from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel
    from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog

    # 1. Test TDocsDatabase schema migration from legacy DB without 'tags' column
    legacy_db_path = tmp_path / "legacy_user_tdocs.db"
    conn = sqlite3.connect(legacy_db_path)
    conn.execute("CREATE TABLE user_tdocs (tdoc_id TEXT PRIMARY KEY, status TEXT, notes TEXT)")
    conn.execute("INSERT INTO user_tdocs VALUES ('S2-2612673', '🟢 Support', 'Baseline note')")
    conn.commit()
    conn.close()

    tdb = TDocsDatabase(legacy_db_path)
    tdb.upsert("S2-2612673", "🟢 Support", "Updated note", "Key Issue, Follow-up, key issue")
    assert tdb.get_all()["S2-2612673"]["tags"] == "Key Issue, Follow-up"
    # 3-arg backward-compatible upsert must preserve existing tags
    tdb.upsert("S2-2612673", "🟢 Support", "Updated note 2")
    assert tdb.get_all()["S2-2612673"]["tags"] == "Key Issue, Follow-up"
    assert tdb.get_all_unique_tags() == ["Follow-up", "Key Issue"]

    # 2. Test InteractiveNotesDialog with tag suggestions + TagCloudWidget
    saved_calls = []
    dlg_notes = InteractiveNotesDialog(
        None,
        "S2-2612673",
        {"TDoc": "S2-2612673", "Secretary Remarks": "Discussed", "My Status": "⚪ Neutral", "My Notes": "", "My Tags": ""},
        lambda tid, st, nt, tg: saved_calls.append((tid, st, nt, tg)),
        all_existing_tags=["Key Issue", "Urgent"],
    )
    dlg_notes.status_combo.setCurrentText("🟡 Monitor")
    dlg_notes.my_notes.setPlainText("Check revision tomorrow")
    dlg_notes.tags_cloud.add_tag("Key Issue")
    dlg_notes.tags_cloud.add_tag("Urgent")
    dlg_notes._on_save_clicked()
    assert saved_calls == [("S2-2612673", "🟡 Monitor", "Check revision tomorrow", "Key Issue, Urgent")]

    # 3. Test TrackedThreadsDatabase annotations + Markdown export + Table display & filtering
    ttdb = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid = ttdb.create_or_get_thread("NAS Routing", "NAS Routing", meeting_id="SA2#178", wg="SA2", linked_ais=["20.6.1"])
    ttdb.update_thread_annotations(tid, user_status="🟢 Support", user_notes="Align with ZTE option", user_tags=["Key Issue"])
    t_row = ttdb.get_thread(tid)
    md = export_thread_to_markdown(t_row, [])
    assert "- **My Status:** 🟢 Support" in md
    assert "- **My Tags:** Key Issue" in md
    assert "> [!NOTE] Delegate Notes" in md
    assert "> Align with ZTE option" in md

    model = TDocsTableModel(tmp_path, [{"TDoc": "S2-2612673", "Title": "Doc", "Type": "pCR", "Source": "ZTE", "Agenda Item": "20.6.1", "TDoc Status": "available"}], tdb.get_all(), {})
    model.sync_tracked_thread_rows([t_row])
    notes_col = model._headers.index("My Notes")
    # Row 0 (S2-2612673) and Row 1 (THR-...) both show tag preview
    from PyQt5.QtCore import Qt
    assert "🏷️ Key Issue" in str(model.data(model.index(0, notes_col), Qt.DisplayRole))
    assert "🏷️ Key Issue" in str(model.data(model.index(1, notes_col), Qt.DisplayRole))

    proxy = TDocsFilterProxyModel()
    proxy.setSourceModel(model)
    proxy.setTypeFilters(["pCR", "📧 Email Thread"])
    proxy.setAIFilters(["20.6.1"])
    proxy.setWIFilters(model.get_related_wis(0) + model.get_related_wis(1))
    proxy.setStatusFilters(["available", "📧 Active Thread"])
    proxy.setCompanyFilters(["ZTE", "Other"])
    proxy.setMyStatusFilters(["🟢 Support"])
    proxy.setTagFilters(["Key Issue"], active=True)
    assert proxy.rowCount() == 2
    proxy.setTagFilters(["NonExistent"], active=True)
    assert proxy.rowCount() == 0

    # 4. Test TrackedThreadDialog tags pill auto-save
    tdlg = TrackedThreadDialog(ttdb, tid, all_existing_tags=["Key Issue", "Priority"])
    assert "🟢 Support" in tdlg.notes_btn.text()
    tdlg.tags_pills.add_tag("Priority")
    assert set(ttdb.get_thread(tid)["user_tags"]) == {"Key Issue", "Priority"}
    tdlg.close()


def test_mark_read_syncs_with_outlook(tmp_path):
    from unittest.mock import patch
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.general_email_db import GeneralEmailDatabase
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.core.tracked_threads_service import mark_thread_read_with_outlook

    class FakeMailItem:
        def __init__(self, unread: bool):
            self.UnRead = unread
            self.save_calls = 0

        def Save(self):
            self.save_calls += 1

    fake_items = {
        "EID-UNREAD-1": FakeMailItem(unread=True),
        "EID-READ-2": FakeMailItem(unread=False),
        "EID-GEN-1": FakeMailItem(unread=True),
    }

    class FakeNamespace:
        def GetItemFromID(self, eid):
            if eid not in fake_items:
                raise KeyError(eid)
            return fake_items[eid]

    ttdb = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid = ttdb.create_or_get_thread("KI#2 read sync", "KI#2 read sync", meeting_id="SA2#178")
    ttdb.upsert_messages(
        tid,
        [
            {"id": "EID-UNREAD-1", "received_iso": "2026-11-18 10:00:00", "is_read": 0},
            {"id": "EID-READ-2", "received_iso": "2026-11-18 11:00:00", "is_read": 1},
            {"id": "MSG-SYNTHETIC-3", "received_iso": "2026-11-18 12:00:00", "is_read": 0},
        ],
    )

    with patch.object(OutlookClient, "get_namespace", return_value=FakeNamespace()):
        updated = mark_thread_read_with_outlook(ttdb, tid, sync_outlook=True)
        assert updated == 1
        assert fake_items["EID-UNREAD-1"].UnRead is False
        assert fake_items["EID-UNREAD-1"].save_calls == 1
        assert fake_items["EID-READ-2"].save_calls == 0

    assert all(int(m["is_read"]) == 1 for m in ttdb.get_thread_messages(tid))

    # Refreshing from Outlook syncs is_read=1 while preserving local read marks
    tid2 = ttdb.create_or_get_thread("KI#3 refresh sync", "KI#3 refresh sync", meeting_id="SA2#178")
    ttdb.upsert_messages(
        tid2,
        [{"id": "EID-SYNC-BACK", "received_iso": "2026-11-18 09:00:00", "is_read": 0}],
    )
    assert ttdb.get_thread_messages(tid2)[0]["is_read"] == 0
    ttdb.upsert_messages(
        tid2,
        [{"id": "EID-SYNC-BACK", "received_iso": "2026-11-18 09:00:00", "is_read": 1}],
    )
    assert ttdb.get_thread_messages(tid2)[0]["is_read"] == 1

    # GeneralEmailDatabase read/unread syncs with Outlook
    gdb = GeneralEmailDatabase(db_path=str(tmp_path / "general.db"))
    gdb.save_emails_batch(
        [
            {
                "id": "EID-GEN-1",
                "folder_tag": "Inbox",
                "sender_name": "Alice",
                "subject": "S2-2612673 comments",
                "date_received": "2026-11-18 10:00:00",
                "is_read": 0,
            }
        ],
        [{"email_id": "EID-GEN-1", "tdoc_id": "S2-2612673", "rev_matched": "r00", "match_location": "Subject"}],
    )
    with patch.object(OutlookClient, "get_namespace", return_value=FakeNamespace()):
        assert gdb.set_tdocs_read_status({"S2-2612673"}, True) == 1
        assert fake_items["EID-GEN-1"].UnRead is False
        assert gdb.set_tdocs_read_status({"S2-2612673"}, False) == 1
        assert fake_items["EID-GEN-1"].UnRead is True
        assert gdb.mark_all_read() == 1
        assert fake_items["EID-GEN-1"].UnRead is False


def test_per_message_read_unread_and_flag_sync(tmp_path):
    from unittest.mock import patch
    from PyQt5.QtCore import Qt, QUrl
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.general_email_db import GeneralEmailDatabase
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.core.thread_markdown_exporter import export_thread_to_markdown
    from modules.emails.core.tracked_threads_service import (
        set_thread_message_flag_with_outlook,
        set_thread_message_read_with_outlook,
    )
    from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog
    from modules.emails.ui.tdoc_emails_dialog import TDocEmailsDialog
    from modules.meetings.ui.tdocs_models import TDocsTableModel

    class FakeMailItem:
        def __init__(self, unread: bool, flag_status: int = 0):
            self.UnRead = unread
            self.FlagStatus = flag_status
            self.save_calls = 0

        def Save(self):
            self.save_calls += 1

    fake_items = {
        "EID-M1": FakeMailItem(unread=True, flag_status=0),
        "EID-M2": FakeMailItem(unread=False, flag_status=2),
        "EID-GEN-1": FakeMailItem(unread=False, flag_status=0),
    }

    class FakeNamespace:
        def GetItemFromID(self, eid):
            if eid not in fake_items:
                raise KeyError(eid)
            return fake_items[eid]

    ttdb = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid = ttdb.create_or_get_thread("KI#5 Per-Message", "KI#5 Per-Message", meeting_id="SA2#178")
    ttdb.upsert_messages(
        tid,
        [
            {"id": "EID-M1", "sender_name": "Alice", "company": "Samsung", "date_sent": "2026-11-18 10:00:00", "is_read": 0, "flag_status": 0},
            {"id": "EID-M2", "sender_name": "Bob", "company": "Ericsson", "date_sent": "2026-11-18 11:00:00", "is_read": 1, "flag_status": 2},
            {"id": "MSG-LOCAL-1", "sender_name": "Carol", "company": "Nokia", "date_sent": "2026-11-18 12:00:00", "is_read": 1, "flag_status": 0},
        ],
    )
    t_row = ttdb.get_thread(tid)
    assert t_row["unread_count"] == 1
    assert t_row["flagged_count"] == 1

    # Interactive markdown exposes per-message read/unread and flag links + badges
    md = export_thread_to_markdown(t_row, ttdb.get_thread_messages(tid), interactive_links=True)
    assert "outlook://read/EID-M1" in md
    assert "outlook://flag/EID-M1" in md
    assert "outlook://unread/EID-M2" in md
    assert "outlook://unflag/EID-M2" in md
    assert "🚩 **FLAGGED**" in md

    with patch.object(OutlookClient, "get_namespace", return_value=FakeNamespace()):
        # Toggle individual read/unread and flag via service & anchor clicks
        dlg = TrackedThreadDialog(ttdb, tid)
        dlg._on_anchor_clicked(QUrl("outlook://read/EID-M1"))
        dlg._on_anchor_clicked(QUrl("outlook://flag/EID-M1"))
        assert fake_items["EID-M1"].UnRead is False
        assert fake_items["EID-M1"].FlagStatus == 2

        dlg._on_anchor_clicked(QUrl("outlook://unread/EID-M2"))
        dlg._on_anchor_clicked(QUrl("outlook://unflag/EID-M2"))
        assert fake_items["EID-M2"].UnRead is True
        assert fake_items["EID-M2"].FlagStatus == 0

        # Synthetic MSG- IDs toggle locally without failing Outlook lookup
        assert set_thread_message_read_with_outlook(ttdb, "MSG-LOCAL-1", False, sync_outlook=True) == 0
        assert set_thread_message_flag_with_outlook(ttdb, "MSG-LOCAL-1", True, sync_outlook=True) == 0
        dlg._load_data()

        t_after = ttdb.get_thread(tid)
        assert t_after["unread_count"] == 2
        assert t_after["flagged_count"] == 2
        assert "Flagged: 2" in dlg.meta_label.text()

        # In-thread state filter works for All / Unread Only / Flagged Only
        dlg.state_filter_combo.setCurrentIndex(dlg.state_filter_combo.findData("unread"))
        assert "Showing 2 of 3" in dlg.msg_count_lbl.text()
        dlg.state_filter_combo.setCurrentIndex(dlg.state_filter_combo.findData("flagged"))
        assert "Showing 2 of 3" in dlg.msg_count_lbl.text()
        dlg.close()

    # TDocEmailsDialog & TDocsTableModel show and sync 2-state flags
    gdb = GeneralEmailDatabase(db_path=str(tmp_path / "general.db"))
    gdb.save_emails_batch(
        [
            {
                "id": "EID-GEN-1",
                "folder_tag": "SA2",
                "sender_name": "Alice",
                "company": "Samsung",
                "subject": "S2-2612673 check",
                "date_received": "2026-11-18 10:00:00",
                "is_read": 0,
                "flag_status": 0,
            }
        ],
        [{"email_id": "EID-GEN-1", "tdoc_id": "S2-2612673", "rev_matched": "r00", "match_location": "Subject"}],
    )
    with patch.object(OutlookClient, "get_namespace", return_value=FakeNamespace()):
        edlg = TDocEmailsDialog("S2-2612673", ["S2-2612673"], tmp_path / "general.db", wg="SA2")
        assert edlg._get_selected_indices() == [0]
        edlg._toggle_selected_flag()
        assert fake_items["EID-GEN-1"].FlagStatus == 2
        assert "🚩" in edlg.model.item(0, 0).text()
        assert "1 Flagged" in edlg.lbl_counts.text()

        counts = gdb.get_email_counts_per_tdoc()
        assert counts["S2-2612673"]["flagged"] == 1
        model = TDocsTableModel(
            tmp_path,
            [{"TDoc": "S2-2612673", "Title": "Doc", "Type": "pCR", "Source": "Samsung", "Agenda Item": "20.6.1"}],
            {},
            {},
        )
        model.email_counts = counts
        model._recompute_cached_email_stats()
        model.sync_tracked_thread_rows([ttdb.get_thread(tid)])
        email_col = model._headers.index("Emails")
        assert "🚩 1" in str(model.data(model.index(0, email_col), Qt.DisplayRole))
        assert "🚩 2" in str(model.data(model.index(1, email_col), Qt.DisplayRole))

        edlg._toggle_selected_flag()
        assert fake_items["EID-GEN-1"].FlagStatus == 0
        edlg.close()


def test_lean_thread_dialog_and_idle_cli_launch(tmp_path):
    from unittest.mock import patch
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.core.thread_markdown_exporter import export_thread_to_markdown
    from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog

    db = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid = db.create_or_get_thread("Lean Dialog Test", "Lean Dialog Test", meeting_id="SA2#178", linked_ais=["20.6.1"])
    db.upsert_messages(
        tid,
        [{"id": "EID-L1", "sender_name": "Alice", "company": "Samsung", "date_sent": "2026-11-18 10:00:00", "body_text": "Proposal body text", "is_read": 0}],
    )
    t_row = db.get_thread(tid)
    msgs = db.get_thread_messages(tid)

    md_compact = export_thread_to_markdown(t_row, msgs, interactive_links=True, include_metadata=False)
    assert "## Thread Metadata" not in md_compact
    assert "Proposal body text" in md_compact

    md_full = export_thread_to_markdown(t_row, msgs, interactive_links=False, include_metadata=True)
    assert "## Thread Metadata" in md_full

    dlg = TrackedThreadDialog(db, tid)
    dlg._settings.setValue("links_panel_expanded", False)
    dlg._settings.setValue("show_metadata_header", False)
    dlg.links_toggle_btn.setChecked(False)
    dlg.show_meta_action.setChecked(False)
    assert not dlg.links_panel.isVisible()
    assert "Links & Tags (1)" in dlg.links_toggle_btn.text()
    assert "## Thread Metadata" not in dlg.browser.toPlainText()

    dlg.links_toggle_btn.setChecked(True)
    assert not dlg.links_panel.isHidden()

    with patch("modules.emails.ui.tracked_thread_dialog.launch_cli_in_terminal", return_value=True) as mock_launch:
        dlg._on_launch_cli("codex")
        mock_launch.assert_called_once()
        assert mock_launch.call_args.kwargs.get("auto_execute") is False
        assert mock_launch.call_args.args[0] == "codex"
        assert (mock_launch.call_args.args[1] / "_CLI_CONTEXT.md").is_file()
    dlg.close()


def test_strict_dropdown_cascading_subitems_and_ki_resolution(tmp_path):
    from types import SimpleNamespace
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.core.tracked_threads_service import auto_suggest_links
    from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog

    # 1. KI/AI/WI resolution against meeting agenda_map and valid_wis
    agenda_map = {
        "20.6.1": SimpleNamespace(description="FS_6G_ARC: KI#1_1 NAS Routing"),
        "20.6.2": SimpleNamespace(description="FS_6G_ARC: KI#4 Other aspects"),
        "20.7.1": SimpleNamespace(description="FS_OTHER: KI#4 Something else"),
    }
    res_disamb = auto_suggest_links(
        "Re: [SA2#178, FS_6G_ARC, KI#4] Interim conclusion",
        agenda_map=agenda_map,
        valid_wis=["FS_6G_ARC", "FS_OTHER"],
    )
    assert res_disamb["ais"] == ["20.6.2"]
    assert res_disamb["wis"] == ["FS_6G_ARC"]

    res_unmapped = auto_suggest_links(
        "Re: [KI#99, UNKNOWN_WI] Topic",
        agenda_map=agenda_map,
        valid_wis=["FS_6G_ARC"],
    )
    assert res_unmapped["ais"] == []
    assert res_unmapped["wis"] == []

    # 2. Strict dropdowns, cascading Sub-Topics, parent-removal pruning, and preservation of legacy links
    db = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid = db.create_or_get_thread(
        "Cascading Links Thread",
        "Cascading Links Thread",
        meeting_id="SA2#178",
        linked_ais=["LEGACY_KI#4", "20.6.1"],
        linked_wis=["LEGACY_WI"],
    )
    ai_to_sub = {
        "20.6.1": ["NAS Routing/transport", "Shared Sub"],
        "20.6.2": ["KI#4 Bullet 3", "Shared Sub"],
    }
    dlg = TrackedThreadDialog(
        db,
        tid,
        ai_suggestions=["20.6.1 — FS_6G_ARC: KI#1_1", "20.6.2 — FS_6G_ARC: KI#4"],
        wi_suggestions=["FS_6G_ARC"],
        ai_to_sub_items=ai_to_sub,
    )
    # Strict dropdowns hide/lock free-text input
    assert dlg.ais_pills.strict_dropdown and dlg.ais_pills.input_edit.isReadOnly()
    assert dlg.sub_items_pills.strict_dropdown and dlg.wis_pills.strict_dropdown
    assert not dlg.tdocs_pills.strict_dropdown and not dlg.tags_pills.strict_dropdown
    # Legacy AIs/WIs are NOT auto-cleaned on load
    assert dlg.ais_pills.tags() == ["LEGACY_KI#4", "20.6.1"]
    assert dlg.wis_pills.tags() == ["LEGACY_WI"]
    assert dlg.sub_items_pills._suggestions == ["NAS Routing/transport", "Shared Sub"]

    # Select sub-topic under 20.6.1 and also add 20.6.2
    dlg.sub_items_pills.add_tag("NAS Routing/transport")
    dlg.sub_items_pills.add_tag("Shared Sub")
    dlg.ais_pills.add_tag("20.6.2")
    assert set(dlg.sub_items_pills._suggestions) == {"NAS Routing/transport", "KI#4 Bullet 3", "Shared Sub"}
    assert set(db.get_thread(tid)["linked_sub_items"]) == {"NAS Routing/transport", "Shared Sub"}

    # Removing parent AI 20.6.1 prunes "NAS Routing/transport" while keeping "Shared Sub" (still under 20.6.2)
    dlg.ais_pills.remove_tag("20.6.1")
    assert db.get_thread(tid)["linked_sub_items"] == ["Shared Sub"]
    assert dlg.sub_items_pills._suggestions == ["KI#4 Bullet 3", "Shared Sub"]

    # Removing remaining mapped AI 20.6.2 clears sub-topics and disables picker with "No sub-topics for AI"
    dlg.ais_pills.remove_tag("20.6.2")
    assert db.get_thread(tid)["linked_sub_items"] == []
    assert not dlg.sub_items_pills.pick_btn.isEnabled()
    assert "No sub-topics for AI" in dlg.sub_items_pills.pick_btn.text()

    # Removing the remaining legacy AI pill switches label to "Assign AI first"
    dlg.ais_pills.remove_tag("LEGACY_KI#4")
    assert "Assign AI first" in dlg.sub_items_pills.pick_btn.text()
    dlg.close()


def test_async_thread_sync_workers_and_progress_bar(tmp_path):
    from unittest.mock import patch
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])

    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.core.tracked_threads_workers import (
        OutlookFolderScanWorker,
        SingleThreadRefreshWorker,
        TrackedThreadsRefreshWorker,
    )
    from modules.emails.ui.sync_progress_bar import AsyncSyncProgressBar

    # 1. Test AsyncSyncProgressBar states
    bar = AsyncSyncProgressBar()
    bar.show()
    bar.start_busy("Syncing threads...", total=3, remaining=["Thread A", "Thread B", "Thread C"])
    assert not bar.isHidden()
    assert bar.progress_bar.maximum() == 3

    bar.update_progress(1, 3, "Thread A", ["Thread B", "Thread C"])
    assert "Thread A" in bar.status_label.text()
    assert "Next (2):" in bar.remaining_label.text()
    assert bar.progress_bar.value() == 1

    bar.show_error("Outlook RPC unavailable")
    assert not bar.isHidden()
    assert "Outlook RPC unavailable" in bar.status_label.text()
    assert bar.cancel_btn.text() == "Dismiss"
    bar.cancel_btn.click()
    assert bar.isHidden()

    bar.start_busy("Done test", total=1)
    bar.finish_ok()
    assert bar.isHidden()

    # 2. Test TrackedThreadsRefreshWorker + SingleThreadRefreshWorker + OutlookFolderScanWorker
    db_path = tmp_path / "threads.db"
    db = TrackedThreadsDatabase(db_path)
    t1 = db.create_or_get_thread("Thread One", "Thread One", meeting_id="SA2#178", outlook_folders=["Inbox"])
    t2 = db.create_or_get_thread("Thread Two", "Thread Two", meeting_id="SA2#178", outlook_folders=["Inbox"])

    progress_events = []
    finished_events = []

    def fake_refresh(_db, tid, progress_cb=None, cancel_check=None):
        if progress_cb:
            progress_cb(0, 1, "Inbox", [])
        return 2 if tid == t1 else 0

    with patch("modules.emails.core.tracked_threads_workers.refresh_thread", side_effect=fake_refresh):
        w = TrackedThreadsRefreshWorker(db_path, "SA2#178")
        w.progress.connect(lambda c, t, cur, rem: progress_events.append((c, t, cur, rem)))
        w.finished_sync.connect(lambda ok, upd, err: finished_events.append((ok, upd, err)))
        w.run()
        assert any(cur == "Thread One" and "Thread Two" in rem for _c, _t, cur, rem in progress_events)
        assert finished_events == [(True, {t1: 2}, "")]

        single_events = []
        sw = SingleThreadRefreshWorker(db_path, t1)
        sw.finished_sync.connect(lambda ok, cnt, err: single_events.append((ok, cnt, err)))
        sw.run()
        assert single_events == [(True, 2, "")]

    with patch(
        "modules.emails.core.tracked_threads_workers.list_recent_threads_in_folders",
        return_value=[{"title": "Sample", "normalized_subject": "Sample", "count": 3}],
    ):
        scan_events = []
        fw = OutlookFolderScanWorker("SA2")
        fw.finished_scan.connect(lambda ok, items, err: scan_events.append((ok, len(items), err)))
        fw.run()
        assert scan_events == [(True, 1, "")]
