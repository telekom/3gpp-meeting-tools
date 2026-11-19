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
    assert idx_nokia >= 1
    dlg.company_filter_combo.setCheckedItems(["Nokia"])
    dlg.company_filter_combo.selectionChanged.emit(dlg.company_filter_combo.getCheckedItems())
    assert "1 of 2" in dlg.msg_count_lbl.text()
    dlg.company_filter_combo.setCheckedItems(["Nokia", "Ericsson"])
    dlg.company_filter_combo.selectionChanged.emit(dlg.company_filter_combo.getCheckedItems())
    assert "2 message(s)" in dlg.msg_count_lbl.text()
    dlg.company_filter_combo.setCheckedItems([])
    dlg.company_filter_combo.selectionChanged.emit(dlg.company_filter_combo.getCheckedItems())
    assert "2 message(s)" in dlg.msg_count_lbl.text()
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

    # Also verify Ctrl+Wheel zoom on MeetingEmailInboxDialog.preview_browser even with pixel-sized QSS fonts
    from PyQt5.QtCore import QSettings
    from PyQt5.QtGui import QFont
    from modules.emails.core.general_email_db import GeneralEmailDatabase
    from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog

    gdb = GeneralEmailDatabase(tmp_path / "zoom_gen.db")
    gdb.save_emails_batch(
        [{"id": "E-Z1", "subject": "Zoom test", "sender_name": "Alice", "date_received": "2026-11-18 10:00:00", "body_text": "Line 1\nLine 2", "is_read": 0}],
        [{"email_id": "E-Z1", "tdoc_id": "S2-2612000", "match_location": "Subject"}],
    )
    idlg = MeetingEmailInboxDialog(
        db,
        gdb,
        meeting_id="SA2#178",
        settings=QSettings(str(tmp_path / "zoom_inbox.ini"), QSettings.IniFormat),
    )
    px_f = QFont(idlg.preview_browser.font())
    px_f.setPixelSize(16)
    idlg.preview_browser.setFont(px_f)
    idlg.show()
    ivp = idlg.preview_browser.viewport()
    QApplication.sendEvent(ivp, wheel_up)
    assert int(idlg.preview_browser.property("_reader_zoom_steps") or 0) == 1
    assert idlg.preview_browser.font().pointSizeF() > 7.0
    zoomed_up_pt = idlg.preview_browser.font().pointSizeF()
    QApplication.sendEvent(ivp, wheel_down)
    assert int(idlg.preview_browser.property("_reader_zoom_steps") or 0) == 0
    assert idlg.preview_browser.font().pointSizeF() < zoomed_up_pt
    idlg.close()


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
        edlg._on_anchor_clicked(QUrl("outlook://flag/EID-GEN-1"))
        assert fake_items["EID-GEN-1"].FlagStatus == 2
        assert "1 Flagged" in edlg.lbl_counts.text()
        assert "FLAGGED" in edlg.browser.toPlainText()

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

        edlg._on_anchor_clicked(QUrl("outlook://unflag/EID-GEN-1"))
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
    assert "Links && Tags (1)" in dlg.links_toggle_btn.text()
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

    # Large AI dropdown uses single-column scrollable popup with multi-token search (never multi-column QMenu)
    many_ais = [f"8.{i} — Legacy Maintenance Item {i}" for i in range(1, 75)] + [
        "20.6.1 — FS_6G_ARC: Paging and Mobility Aspects",
        "20.6.2 — FS_6G_ARC: Compute Offloading",
    ]
    dlg.ais_pills.set_suggestions(many_ais, value_extractor=lambda s: s.split(" — ", 1)[0].strip())
    dlg.ais_pills._populate_pick_menu()
    assert dlg.ais_pills.pick_list is not None
    assert dlg.ais_pills.pick_list.count() == len(many_ais)
    assert dlg.ais_pills.pick_list.maximumHeight() <= 700
    assert dlg.ais_pills.pick_search_edit is not None
    dlg.ais_pills.pick_search_edit.setText("20.6 paging")
    visible_rows = [
        dlg.ais_pills.pick_list.item(i)
        for i in range(dlg.ais_pills.pick_list.count())
        if not dlg.ais_pills.pick_list.item(i).isHidden()
    ]
    assert len(visible_rows) == 1
    assert "20.6.1" in visible_rows[0].text()
    dlg.ais_pills.pick_list.itemClicked.emit(visible_rows[0])
    assert "20.6.1" in dlg.ais_pills.tags()
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
    assert bar.cancel_btn.minimumHeight() >= 24

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


def test_html_table_injection_and_list_normalization():
    from PyQt5.QtWidgets import QApplication, QTextBrowser
    from modules.emails.core.thread_markdown_exporter import (
        export_thread_to_markdown,
        inject_html_tables_into_body,
    )

    app = QApplication.instance() or QApplication([])
    plain_body = (
        "Hi all,\n"
        "1.\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0Regarding the NF Set concept:\n"
        "-\xa0\xa0\xa0\xa0\xa0\xa0\xa0\xa0Use NF set for all NFs.\n\n"
        "nfSetIdList\n\narray(NfSetId)\n\nC\n\n1..N\n\nNF Set ID defined in clause 28.12.\n\n"
        "Table 6.1.6.3.3-1: Enumeration NFType\n\n"
        "Enumeration value\n\nDescription\n\n\"NRF\"\n\nNetwork Function: NRF\n"
    )
    html_body = (
        "<html><body>"
        "<table><tr><td>nfSetIdList</td><td>array(NfSetId)</td><td>C</td><td>1..N</td><td>NF Set ID defined in clause 28.12.</td></tr></table>"
        "<table><tr><th>Enumeration value</th><th>Description</th></tr><tr><td>\"NRF\"</td><td>Network Function: NRF</td></tr></table>"
        "</body></html>"
    )
    merged = inject_html_tables_into_body(plain_body, html_body)
    assert "| nfSetIdList | array(NfSetId) | C | 1..N | NF Set ID defined in clause 28.12. |" in merged
    assert "| Enumeration value | Description |" in merged
    assert "| \"NRF\" | Network Function: NRF |" in merged

    md = export_thread_to_markdown(
        {"title": "Table Thread", "normalized_subject": ""},
        [{"id": "EID-TBL", "sender_name": "Uri", "date_sent": "2026-11-19 00:38:19", "body_text": merged}],
        interactive_links=True,
        include_metadata=False,
    )
    browser = QTextBrowser()
    browser.setMarkdown(md)
    rendered = browser.toHtml()
    assert rendered.count("<table") == 2
    assert "<pre" not in rendered

    # Test table cells containing Safelinked <a href> hyperlinks (Outlook appends <https://safelink...> in Body)
    plain_linked_tbl = (
        "Inputs:\n\n"
        "20.6.1\n\n"
        "S2-2611551 <https://eur05.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.3gpp.org%2FDocs%2FS2-2611551.zip&data=1>\n\n"
        "P-CR\n"
    )
    html_linked_tbl = (
        "<table><tr>"
        "<td>20.6.1</td>"
        "<td><a href='https://eur05.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.3gpp.org%2FDocs%2FS2-2611551.zip&data=1'>S2-2611551</a></td>"
        "<td>P-CR</td>"
        "</tr></table>"
    )
    merged_links = inject_html_tables_into_body(plain_linked_tbl, html_linked_tbl)
    assert "| 20.6.1 | [S2-2611551](https://www.3gpp.org/Docs/S2-2611551.zip) | P-CR |" in merged_links


def test_reload_single_and_all_cached_emails_from_outlook(tmp_path):
    from unittest.mock import patch
    from PyQt5.QtCore import QUrl
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.core.tracked_threads_service import (
        reload_all_thread_messages_from_outlook,
    )
    from modules.emails.core.tracked_threads_workers import SingleThreadRefreshWorker
    from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog

    class FakeItem:
        def __init__(self, eid: str, body: str, html: str):
            self.EntryID = eid
            self.Subject = "RE: Table test"
            self.SenderName = "Alice"
            self.SenderEmailAddress = "alice@samsung.com"
            self.Body = body
            self.HTMLBody = html
            self.ConversationID = "CID-TBL"
            self.UnRead = False
            self.FlagStatus = 0

    fake_items = {
        "EID-R1": FakeItem(
            "EID-R1",
            "Header\nA\nB\n1\n2",
            "<table><tr><td>A</td><td>B</td></tr><tr><td>1</td><td>2</td></tr></table>",
        )
    }

    class FakeNS:
        def GetItemFromID(self, eid):
            return fake_items[eid]

    db_path = tmp_path / "threads.db"
    db = TrackedThreadsDatabase(db_path)
    tid = db.create_or_get_thread("Reload Table", "Reload Table", meeting_id="SA2#178")
    db.upsert_messages(
        tid,
        [{"id": "EID-R1", "sender_name": "Alice", "date_sent": "2026-11-19 01:00:00", "body_text": "stale flattened"}],
    )

    with patch.object(OutlookClient, "get_namespace", return_value=FakeNS()):
        dlg = TrackedThreadDialog(db, tid)
        dlg._on_anchor_clicked(QUrl("outlook://reload/EID-R1"))
        reloaded_msg = db.get_thread_messages(tid)[0]
        assert "| A | B |" in reloaded_msg["body_text"]
        assert "| 1 | 2 |" in reloaded_msg["body_text"]

        # Worker mode reload_cached=True and batch reload service also update cached messages
        assert reload_all_thread_messages_from_outlook(db, tid) == 1
        events = []
        sw = SingleThreadRefreshWorker(db_path, tid, reload_cached=True)
        sw.finished_sync.connect(lambda ok, cnt, err: events.append((ok, cnt, err)))
        sw.run()
        assert events == [(True, 1, "")]
        dlg.close()


def test_collect_matching_includes_sent_items_and_pulls_conversation():
    from unittest.mock import patch
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.tracked_threads_service import _collect_matching_from_folders

    class FakeRow:
        def __init__(self, data):
            self._data = data

        def __call__(self, key):
            return self._data[key]

    class FakeTable:
        def __init__(self, rows):
            self._rows = list(rows)
            self._idx = 0
            self.Columns = type("Cols", (), {"Add": lambda _s, _c: None})()

        @property
        def EndOfTable(self):
            return self._idx >= len(self._rows)

        def Sort(self, *_a):
            pass

        def GetNextRow(self):
            r = self._rows[self._idx]
            self._idx += 1
            return FakeRow(r)

    class FakeMailItem:
        def __init__(self, eid, subj, sender, conv_id, body, sent_item_table=None):
            self.EntryID = eid
            self.Subject = subj
            self.SenderName = sender
            self.SenderEmailAddress = f"{sender.lower()}@example.com"
            self.ConversationID = conv_id
            self.Body = body
            self.HTMLBody = ""
            self.UnRead = False
            self.FlagStatus = 0
            self._conv_rows = sent_item_table or []

        def GetConversation(self):
            rows = self._conv_rows
            return type("Conv", (), {"GetTable": lambda _s: FakeTable(rows)})()

    inbox_item = FakeMailItem(
        "EID-IN-1",
        "RE: [SA2#178] draft S2-2612711 was 12653",
        "Maxime",
        "CID-SENT-TEST",
        "Uploaded draft",
        sent_item_table=[{"EntryID": "EID-IN-1"}, {"EntryID": "EID-SENT-CONV"}],
    )
    sent_conv_item = FakeMailItem(
        "EID-SENT-CONV",
        "RE: [SA2#178] draft S2-2612711 was 12653",
        "Josep",
        "CID-SENT-TEST",
        "My sent reply via GetConversation",
    )
    sent_folder_item = FakeMailItem(
        "EID-SENT-FOLDER",
        "RE: [SA2#178] draft S2-2612711 was 12653",
        "Josep",
        "CID-DIFFERENT",
        "My sent reply only in Sent Items folder",
    )

    items_by_id = {
        "EID-IN-1": inbox_item,
        "EID-SENT-CONV": sent_conv_item,
        "EID-SENT-FOLDER": sent_folder_item,
    }

    inbox_folder = type(
        "Folder",
        (),
        {
            "FolderPath": "\\\\me@example.com\\Inbox\\SA2",
            "GetTable": lambda _s: FakeTable(
                [{"EntryID": "EID-IN-1", "Subject": inbox_item.Subject, "MessageClass": "IPM.Note", "ConversationID": "CID-SENT-TEST"}]
            ),
        },
    )()
    sent_folder = type(
        "Folder",
        (),
        {
            "FolderPath": "\\\\me@example.com\\Sent Items",
            "GetTable": lambda _s: FakeTable(
                [{"EntryID": "EID-SENT-FOLDER", "Subject": sent_folder_item.Subject, "MessageClass": "IPM.Note", "ConversationID": "CID-DIFFERENT"}]
            ),
        },
    )()

    class FakeNS:
        def GetDefaultFolder(self, idx):
            if idx == 5:
                return sent_folder
            raise KeyError(idx)

        def GetItemFromID(self, eid):
            return items_by_id[eid]

    def fake_get_folder(path):
        if "Sent Items" in path:
            return sent_folder
        return inbox_folder

    with patch.object(OutlookClient, "get_namespace", return_value=FakeNS()), \
         patch.object(OutlookClient, "get_folder_by_path", side_effect=fake_get_folder):
        # Even when CID-SENT-TEST is already on the thread, refresh still pulls GetConversation() AND scans Sent Items
        found = _collect_matching_from_folders(
            ["Inbox\\SA2"],
            conversation_ids=["CID-SENT-TEST"],
            subject_patterns=["draft S2-2612711 was 12653"],
        )
        found_ids = {m["id"] for m in found}
        assert found_ids == {"EID-IN-1", "EID-SENT-CONV", "EID-SENT-FOLDER"}


def test_tdocs_table_columns_fit_viewport_on_narrow_resize():
    from modules.meetings.ui.tdocs_window import TDocsWindow

    # Narrow viewport (e.g. ~1200px) must fit all 14 columns without pushing Emails off-screen
    w_narrow = TDocsWindow.compute_fitted_column_widths(1200, 14, scale_fn=lambda x: x)
    assert sum(w_narrow.values()) <= 1200
    assert w_narrow[13] >= 92

    # Wide viewport (e.g. 1650px) expands flexible columns and preserves Emails width
    w_wide = TDocsWindow.compute_fitted_column_widths(1650, 14, scale_fn=lambda x: x)
    assert sum(w_wide.values()) == 1650
    assert w_wide[13] >= 100
    assert w_wide[7] > w_narrow[7]
    assert w_wide[2] > w_narrow[2]


def test_thread_attachments_and_links_open_via_document_opener(tmp_path):
    from unittest.mock import patch
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.tracked_threads_service import (
        cache_and_open_thread_document_url,
        open_outlook_attachment,
    )

    class FakeAttachment:
        def __init__(self, fname: str):
            self.FileName = fname

        def SaveAsFile(self, path_str: str):
            Path(path_str).write_bytes(b"PK\x03\x04fake-docx")

    class FakeAttachments:
        def __init__(self, items):
            self._items = items
            self.Count = len(items)

        def Item(self, idx: int):
            return self._items[idx - 1]

    class FakeItem:
        def __init__(self):
            self.Attachments = FakeAttachments([FakeAttachment("draft_S2-2612929.docx")])

    class FakeNS:
        def GetItemFromID(self, eid):
            return FakeItem()

    opened_paths = []
    with patch.object(OutlookClient, "get_namespace", return_value=FakeNS()), \
         patch("modules.word_tools.core.document_opener.open_document", side_effect=lambda p: opened_paths.append(Path(p))):
        att_out = open_outlook_attachment("EID-ATT-1", "draft_S2-2612929.docx", save_dir=tmp_path)
        assert att_out is not None and att_out.is_file()
        assert opened_paths[-1].name == "draft_S2-2612929.docx"

        # Pre-cached URL document also opens via open_document() so sensitivity labels apply
        cached_doc = tmp_path / "S2-2612711_Conclusion.docx"
        cached_doc.write_bytes(b"PK\x03\x04cached")
        url_out = cache_and_open_thread_document_url(
            "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/Docs/S2-2612711_Conclusion.docx",
            tmp_path,
            open_after=True,
        )
        assert url_out == cached_doc
        assert opened_paths[-1].name == "S2-2612711_Conclusion.docx"


def test_discovery_classification_config_and_ignore(tmp_path):
    from PyQt5.QtWidgets import QApplication
    from modules.emails.core.general_email_db import GeneralEmailDatabase
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.core.tracked_threads_service import (
        classify_discovered_threads,
        load_meeting_discovery_config,
        save_meeting_discovery_config,
    )
    from modules.emails.ui.tracked_thread_dialog import PinFromOutlookDialog

    app = QApplication.instance() or QApplication([])
    mtg_dir = tmp_path / "SA2_178"
    (mtg_dir / "Agenda").mkdir(parents=True, exist_ok=True)

    # 1. Per-meeting config persistence
    cfg = load_meeting_discovery_config(
        mtg_dir, default_tag="SA2#178", default_start="2026-11-16", default_end="2026-11-20"
    )
    assert cfg["subject_tag"] == "SA2#178"
    assert cfg["coverage_filter"] == "UNTRACKED_AND_BODY"
    cfg["ignored_thread_keys"] = ["noisy off-topic thread"]
    save_meeting_discovery_config(mtg_dir, cfg)
    reloaded = load_meeting_discovery_config(mtg_dir)
    assert reloaded["ignored_thread_keys"] == ["noisy off-topic thread"]

    # 2. Coverage classification (PINNED, SUBJECT_TDOC, BODY_TDOC, STRICT_UNTRACKED)
    tdb = TrackedThreadsDatabase(tmp_path / "threads.db")
    tdb.create_or_get_thread("Pinned Topic", "Pinned Topic", meeting_id="SA2#178")

    gdb = GeneralEmailDatabase(tmp_path / "general.db")
    gdb.save_emails_batch(
        [
            {"id": "E-BODY", "subject": "[SA2#178] Body Ref Only", "date_received": "2026-11-18 10:00", "is_read": 0},
            {"id": "E-SUBJ", "subject": "[SA2#178] S2-2612000 Subject Ref", "date_received": "2026-11-18 11:00", "is_read": 0},
        ],
        [
            {"email_id": "E-BODY", "tdoc_id": "S2-2611999", "match_location": "Body"},
            {"email_id": "E-SUBJ", "tdoc_id": "S2-2612000", "match_location": "Subject"},
        ],
    )

    raw_threads = [
        {"title": "Pinned Topic", "normalized_subject": "Pinned Topic", "norm_key": "pinned topic", "entry_ids": ["E-PIN"], "last_date": "2026-11-18 12:00", "last_sender": "A", "count": 2, "unread_count": 1, "subject_tdocs": []},
        {"title": "Subject TDoc Discussion", "normalized_subject": "Subject TDoc Discussion", "norm_key": "subject tdoc discussion", "raw_subject": "Re: [SA2#178] S2-2612000 Subject TDoc Discussion", "entry_ids": ["E-SUBJ"], "last_date": "2026-11-18 11:30", "last_sender": "B", "count": 3, "unread_count": 1, "subject_tdocs": ["S2-2612000"]},
        {"title": "Body Ref Only", "normalized_subject": "Body Ref Only", "norm_key": "body ref only", "raw_subject": "Re: [SA2#178] Body Ref Only", "entry_ids": ["E-BODY"], "last_date": "2026-11-18 10:00", "last_sender": "C", "count": 4, "unread_count": 2, "subject_tdocs": []},
        {"title": "Pure Offline Breakout", "normalized_subject": "Pure Offline Breakout", "norm_key": "pure offline breakout", "raw_subject": "Re: [SA2#178] Pure Offline Breakout", "entry_ids": ["E-PURE"], "last_date": "2026-11-18 09:00", "last_sender": "D", "count": 5, "unread_count": 2, "subject_tdocs": []},
        {"title": "Wrong Tag Thread", "normalized_subject": "Wrong Tag Thread", "norm_key": "wrong tag thread", "raw_subject": "Re: [SA2#177] Wrong Tag Thread", "entry_ids": ["E-OLD"], "last_date": "2026-11-18 08:00", "last_sender": "E", "count": 1, "unread_count": 0, "subject_tdocs": []},
        {"title": "Too Old Thread", "normalized_subject": "Too Old Thread", "norm_key": "too old thread", "raw_subject": "Re: [SA2#178] Too Old Thread", "entry_ids": ["E-AGO"], "last_date": "2026-09-01 08:00", "last_sender": "F", "count": 1, "unread_count": 0, "subject_tdocs": []},
    ]

    classified = classify_discovered_threads(
        raw_threads,
        tdb,
        gdb,
        meeting_id="SA2#178",
        valid_tdocs={"S2-2611999", "S2-2612000"},
    )
    cov_by_title = {c["title"]: c["coverage"] for c in classified}
    assert cov_by_title["Pinned Topic"] == "PINNED"
    assert cov_by_title["Subject TDoc Discussion"] == "SUBJECT_TDOC"
    assert cov_by_title["Body Ref Only"] == "BODY_TDOC"
    assert cov_by_title["Pure Offline Breakout"] == "STRICT_UNTRACKED"

    # 3. PinFromOutlookDialog UI filtering, Ignore toggle, and Pinning
    dlg = PinFromOutlookDialog(
        tdb,
        "SA2#178",
        "SA2",
        meeting_dir=mtg_dir,
        general_email_db=gdb,
        valid_tdocs={"S2-2611999", "S2-2612000"},
        mtg_start_date="2026-11-16",
        mtg_end_date="2026-11-20",
        auto_scan=False,
    )
    dlg._on_scan_finished(True, raw_threads, "")
    # Option B default hides Pinned, Subject TDoc, wrong tag, and out-of-window date
    visible_titles = [dlg.table.item(r, 2).text() for r in range(dlg.table.rowCount())]
    assert visible_titles == ["Body Ref Only", "Pure Offline Breakout"]

    # Ignore row 1 ("Pure Offline Breakout") and confirm persisted & hidden
    dlg.table.selectRow(1)
    dlg._on_toggle_ignore_selected()
    visible_after_ignore = [dlg.table.item(r, 2).text() for r in range(dlg.table.rowCount())]
    assert visible_after_ignore == ["Body Ref Only"]
    assert "pure offline breakout" in load_meeting_discovery_config(mtg_dir)["ignored_thread_keys"]

    # Switch coverage to IGNORED -> shows ignored thread
    idx_ign = dlg.coverage_combo.findData("IGNORED")
    dlg.coverage_combo.setCurrentIndex(idx_ign)
    assert [dlg.table.item(r, 2).text() for r in range(dlg.table.rowCount())] == ["Pure Offline Breakout"]

    # Un-ignore and restore UNTRACKED_AND_BODY, then Pin "Body Ref Only"
    dlg.table.selectRow(0)
    dlg._on_toggle_ignore_selected()
    dlg.coverage_combo.setCurrentIndex(dlg.coverage_combo.findData("UNTRACKED_AND_BODY"))
    assert dlg.table.rowCount() == 2
    dlg.table.selectRow(0)
    dlg._on_pin_selected()
    assert len(dlg.pinned_thread_ids) == 1
    assert any(t["title"] == "Body Ref Only" for t in tdb.get_threads_for_meeting("SA2#178"))
    from PyQt5.QtCore import Qt as _Qt
    assert bool(dlg.windowFlags() & _Qt.Window)
    assert dlg.isSizeGripEnabled()
    dlg.close()

    # 4. list_recent_threads_in_folders GetTable() path (including None bucket & ConversationID column failure)
    from unittest.mock import patch
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.tracked_threads_service import list_recent_threads_in_folders

    class FakeCols:
        def Add(self, name):
            if name == "ConversationID":
                raise RuntimeError("Falscher Parameter")

    class FakeRow:
        def __init__(self, data):
            self._d = data

        def __call__(self, col):
            if col not in self._d:
                raise RuntimeError("Missing col")
            return self._d[col]

    class FakeTable:
        def __init__(self, rows):
            self._rows = list(rows)
            self._idx = 0
            self.Columns = FakeCols()

        @property
        def EndOfTable(self):
            return self._idx >= len(self._rows)

        def Sort(self, *_args):
            pass

        def GetNextRow(self):
            r = FakeRow(self._rows[self._idx])
            self._idx += 1
            return r

    class FakeFolder:
        def __init__(self, path, rows):
            self.FolderPath = path
            self._rows = rows

        def GetTable(self):
            return FakeTable(self._rows)

    f_inbox = FakeFolder("Inbox\\SA2", [
        {"EntryID": "E1", "Subject": "Re: [SA2#178] Breakout A", "MessageClass": "IPM.Note", "SenderName": "Alice", "ReceivedTime": "2026-11-18 10:00", "UnRead": True},
        {"EntryID": "E1", "Subject": "Re: [SA2#178] Breakout A", "MessageClass": "IPM.Note", "SenderName": "Alice", "ReceivedTime": "2026-11-18 10:00", "UnRead": True},
        {"EntryID": "E2", "Subject": "Fwd: [SA2#178] Breakout A", "MessageClass": "IPM.Note", "SenderName": "Bob", "ReceivedTime": "2026-11-18 11:00", "UnRead": False},
    ])
    with patch("modules.emails.core.tracked_threads_service.load_wg_email_config", return_value=[{"folder_path": "Inbox\\SA2"}]), \
         patch.object(OutlookClient, "get_namespace", return_value=None), \
         patch.object(OutlookClient, "get_folder_by_path", return_value=f_inbox):
        scanned = list_recent_threads_in_folders("SA2", max_items_per_folder=50)
        assert len(scanned) == 1
        assert scanned[0]["title"] == "Breakout A"
        assert scanned[0]["count"] == 2
        assert scanned[0]["unread_count"] == 1


def test_tdocs_email_state_quick_filters():
    from PyQt5.QtWidgets import QApplication
    from modules.meetings.ui.tdocs_models import TDocsFilterProxyModel, TDocsTableModel

    app = QApplication.instance() or QApplication([])
    rows = [
        {"TDoc": "S2-2610001", "Title": "Unread TDoc", "Source": "Nokia", "Type": "CR", "For": "Action", "Agenda Item": "20.1"},
        {"TDoc": "S2-2610002", "Title": "Flagged TDoc", "Source": "Samsung", "Type": "pCR", "For": "Discussion", "Agenda Item": "20.1"},
        {"TDoc": "S2-2610003", "Title": "Quiet TDoc", "Source": "Huawei", "Type": "pCR", "For": "Discussion", "Agenda Item": "20.1"},
    ]
    model = TDocsTableModel(data=rows)
    model.set_email_counts({
        "S2-2610001": {"total": 2, "unread": 1, "flagged": 0, "all_ids": {"E1", "E2"}, "unread_ids": {"E1"}, "flagged_ids": set()},
        "S2-2610002": {"total": 1, "unread": 0, "flagged": 1, "all_ids": {"E3"}, "unread_ids": set(), "flagged_ids": {"E3"}},
    })
    model.sync_tracked_thread_rows([
        {"thread_id": "THR-1", "title": "Active Thread", "unread_count": 2, "flagged_count": 1, "message_count": 5, "linked_ais": ["20.1"], "linked_wis": [], "linked_tdocs": [], "user_tags": [], "participants_companies": "Nokia"},
    ])
    proxy = TDocsFilterProxyModel()
    proxy.setSourceModel(model)
    proxy.setTypeFilters({str(r.get("Type", "")).strip() for r in model._data})
    proxy.setAIFilters({str(r.get("Agenda Item", "")).strip() for r in model._data})
    proxy.setWIFilters({w for r_idx in range(model.rowCount()) for w in model.get_related_wis(r_idx)})
    proxy.setStatusFilters({str(r.get("TDoc Status", "")).strip() for r in model._data})
    proxy.setMyStatusFilters({str(r.get("My Status", "")).strip() for r in model._data})
    proxy.setCompanyFilters({c for r in model._data for c in (r.get("_Sanitized_Companies") or ["Other"])})

    def _visible_tdocs():
        return [proxy.sourceModel()._data[proxy.mapToSource(proxy.index(r, 0)).row()]["TDoc"] for r in range(proxy.rowCount())]

    assert len(_visible_tdocs()) == 4

    proxy.setEmailStateFilter("UNREAD")
    assert set(_visible_tdocs()) == {"S2-2610001", "THR-1"}

    proxy.setEmailStateFilter("FLAGGED")
    assert set(_visible_tdocs()) == {"S2-2610002", "THR-1"}

    proxy.setEmailStateFilter("ANY")
    assert set(_visible_tdocs()) == {"S2-2610001", "S2-2610002", "THR-1"}

    proxy.setEmailStateFilter("ALL")
    assert len(_visible_tdocs()) == 4


def test_tracked_thread_auto_scroll_and_next_unread(tmp_path):
    import time
    from PyQt5.QtWidgets import QApplication
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog

    app = QApplication.instance() or QApplication([])
    db = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid = db.create_or_get_thread("Auto Scroll Thread", "Auto Scroll Thread", meeting_id="SA2#178")
    long_block = "\n\n".join(f"Paragraph {i}: detailed 3GPP discussion text" for i in range(25))
    db.upsert_messages(
        tid,
        [
            {"id": "M1", "sender_name": "Alice", "date_sent": "2026-11-18 09:00:00", "body_text": long_block, "is_read": 1},
            {"id": "M2", "sender_name": "Bob", "date_sent": "2026-11-18 10:00:00", "body_text": long_block, "is_read": 0},
            {"id": "M3", "sender_name": "Carol", "date_sent": "2026-11-18 11:00:00", "body_text": long_block, "is_read": 0},
        ],
    )
    dlg = TrackedThreadDialog(db, tid)
    dlg.resize(700, 420)
    dlg.order_combo.setCurrentIndex(1)  # Oldest First -> M1 (read) at top, M2/M3 below fold
    dlg._initial_scroll_done = False
    dlg.show()
    app.processEvents()
    deadline = time.time() + 0.6
    while time.time() < deadline and not dlg._initial_scroll_done:
        app.processEvents()

    assert len(dlg._unread_blocks) == 2
    assert dlg.next_unread_btn.isEnabled()
    assert dlg.next_unread_btn.text() == "🔵 Next Unread (2)"
    assert dlg._initial_scroll_done is True
    assert dlg._unread_cursor == 0
    first_pos = dlg.browser.verticalScrollBar().value()
    assert first_pos > 0

    dlg._on_next_unread_clicked()
    assert dlg._unread_cursor == 1
    second_pos = dlg.browser.verticalScrollBar().value()
    assert second_pos > first_pos
    dlg._on_next_unread_clicked()
    assert dlg._unread_cursor == 0
    assert dlg.browser.verticalScrollBar().value() < second_pos
    assert abs(dlg.browser.verticalScrollBar().value() - first_pos) <= 50
    dlg._settings.setValue("newest_first", True)
    dlg.close()


def test_meeting_email_inbox_dialog_triage(tmp_path):
    from datetime import datetime
    from unittest.mock import patch
    from PyQt5.QtWidgets import QApplication
    from modules.emails.core.general_email_db import GeneralEmailDatabase
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog

    app = QApplication.instance() or QApplication([])
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    tdb = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid = tdb.create_or_get_thread("Thread A", "Thread A", meeting_id="SA2#178", linked_ais=["20.6.1"])
    tdb.upsert_messages(
        tid,
        [
            {"id": "E-BOTH", "sender_name": "Alice", "company": "Nokia", "date_sent": now_str, "subject": "Shared Msg", "body_text": "Shared body", "is_read": 0, "flag_status": 0},
            {"id": "E-READ-FLAG", "sender_name": "Bob", "company": "Samsung", "date_sent": "2026-10-01 10:00:00", "subject": "Flagged Old", "body_text": "Flagged body", "is_read": 1, "flag_status": 2},
        ],
    )

    gdb = GeneralEmailDatabase(tmp_path / "general.db")
    gdb.save_emails_batch(
        [
            {"id": "E-BOTH", "subject": "Shared Msg", "sender_name": "Alice", "company": "Nokia", "date_received": now_str, "body_text": "Shared body", "is_read": 0, "flag_status": 0},
            {"id": "E-TDOC-ONLY", "subject": "Only TDoc Unread", "sender_name": "Carol", "company": "Ericsson", "date_received": "2026-11-17 09:00:00", "body_text": "See S2-2612345", "is_read": 0, "flag_status": 0},
        ],
        [
            {"email_id": "E-BOTH", "tdoc_id": "S2-2611111", "match_location": "Body"},
            {"email_id": "E-TDOC-ONLY", "tdoc_id": "S2-2612345", "match_location": "Subject"},
        ],
    )

    opened_threads = []
    opened_tdocs = []
    changes = []

    from PyQt5.QtCore import QSettings as _QSettings
    with patch.object(OutlookClient, "set_messages_read_status", return_value=1) as mock_read, \
         patch.object(OutlookClient, "set_messages_flag_status", return_value=1) as mock_flag:
        dlg = MeetingEmailInboxDialog(
            tdb,
            gdb,
            meeting_id="SA2#178",
            tdoc_ai_lookup={"S2-2612345": "20.2"},
            open_thread_callback=lambda t: opened_threads.append(t),
            open_tdoc_callback=lambda td: opened_tdocs.append(td),
            settings=_QSettings(str(tmp_path / "inbox_triage.ini"), _QSettings.IniFormat),
        )
        dlg.data_changed.connect(lambda: changes.append(True))

        # Deduplicated counts: 2 unread (E-BOTH + E-TDOC-ONLY), 1 flagged, 1 recent (E-BOTH), 3 total
        assert dlg.pill_unread.text() == "🔵 Unread Only (2)"
        assert dlg.pill_flagged.text() == "🚩 Flagged (1)"
        assert dlg.pill_recent.text() == "🕒 Last 24h (1)"
        assert dlg.pill_all.text() == "📋 All (3)"
        assert dlg.table.rowCount() == 2

        # Merged context on E-BOTH has both thread title & matched TDoc
        assert dlg._current_item["id"] == "E-BOTH"
        assert "S2-2611111" in dlg._current_item["linked_tdocs"]

        # Mark Read & Next Unread removes E-BOTH from Unread view, syncs BOTH DBs, selects E-TDOC-ONLY
        dlg._on_mark_read_and_next()
        mock_read.assert_called_with(["E-BOTH"], is_read=True)
        assert dlg.pill_unread.text() == "🔵 Unread Only (1)"
        assert dlg.table.rowCount() == 1
        assert dlg._current_item["id"] == "E-TDOC-ONLY"
        assert tdb.get_thread_messages(tid)[0]["is_read"] == 1
        assert gdb.get_all_emails_with_tdocs()[0]["is_read"] == 1 or any(
            e["id"] == "E-BOTH" and e["is_read"] == 1 for e in gdb.get_all_emails_with_tdocs()
        )
        assert len(changes) == 1

        # Flag E-TDOC-ONLY
        dlg._on_toggle_flag()
        mock_flag.assert_called_with(["E-TDOC-ONLY"], is_flagged=True)
        assert dlg.pill_flagged.text() == "🚩 Flagged (2)"

        # Open TDoc callback from E-TDOC-ONLY
        dlg._on_open_context()
        assert opened_tdocs == ["S2-2612345"]
        dlg.close()


def test_unified_tdoc_emails_dialog_timeline_and_actions(tmp_path):
    import time
    from unittest.mock import patch
    from PyQt5.QtCore import QUrl
    from PyQt5.QtWidgets import QApplication
    from modules.emails.core.general_email_db import GeneralEmailDatabase
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.ui.tdoc_emails_dialog import TDocEmailsDialog

    app = QApplication.instance() or QApplication([])
    db_path = tmp_path / "general.db"
    gdb = GeneralEmailDatabase(db_path)

    long_filler = "\n\n".join(f"Line {i}: Technical evaluation of S2-2612673 and S2-2612674." for i in range(25))
    gdb.save_emails_batch(
        [
            {
                "id": "E-OLD-READ",
                "folder_tag": "SA2",
                "subject": "[SA2#178] S2-2612673 Initial Draft",
                "sender_name": "Alice",
                "sender_email": "alice@nokia.com",
                "company": "Nokia",
                "date_received": "2026-11-18 09:00:00",
                "body_text": long_filler + "\n\n| Col1 | Col2 |\n| --- | --- |\n| V1 | V2 |",
                "clean_body": long_filler + "\n\n| Col1 | Col2 |\n| --- | --- |\n| V1 | V2 |",
                "attachments": ["S2-2612673_clean.docx"],
                "is_read": 1,
                "flag_status": 0,
            },
            {
                "id": "E-UNREAD-QUOTED",
                "folder_tag": "Sent",
                "subject": "RE: [SA2#178] Offline comment",
                "sender_name": "Bob",
                "sender_email": "bob@samsung.com",
                "company": "Samsung",
                "date_received": "2026-11-18 11:00:00",
                "body_text": long_filler + "\n\n-----Original Message-----\nSee S2-2612673.",
                "clean_body": long_filler,
                "attachments": [],
                "is_read": 0,
                "flag_status": 0,
            },
        ],
        [
            {"email_id": "E-OLD-READ", "tdoc_id": "S2-2612673", "rev_matched": "r00", "match_location": "Subject"},
            {"email_id": "E-UNREAD-QUOTED", "tdoc_id": "S2-2612673", "rev_matched": "", "match_location": "Quoted"},
        ],
    )

    class FakeItem:
        def __init__(self):
            self.EntryID = "E-OLD-READ"
            self.Subject = "[SA2#178] S2-2612673 Reloaded Subject"
            self.SenderName = "Alice"
            self.SenderEmailAddress = "alice@nokia.com"
            self.Body = "Reloaded Header\nA\nB\n1\n2"
            self.HTMLBody = "<table><tr><td>A</td><td>B</td></tr><tr><td>1</td><td>2</td></tr></table>"
            self.ConversationID = "CID-TDOC"
            self.UnRead = False
            self.FlagStatus = 0

    class FakeNS:
        def GetItemFromID(self, eid):
            return FakeItem()

    pinned_calls = []
    tdoc_jumps = []
    changes = []

    with patch.object(OutlookClient, "set_messages_read_status", return_value=1), \
         patch.object(OutlookClient, "set_messages_flag_status", return_value=1), \
         patch.object(OutlookClient, "get_namespace", return_value=FakeNS()):
        dlg = TDocEmailsDialog(
            "S2-2612673",
            ["S2-2612673"],
            db_path,
            wg="SA2",
            pin_thread_callback=lambda info: pinned_calls.append(info),
        )
        # Auto-read timer is removed
        assert not hasattr(dlg, "auto_read_timer")

        dlg.resize(760, 440)
        dlg.order_combo.setCurrentIndex(1)  # Oldest First -> E-OLD-READ at top, E-UNREAD-QUOTED below fold
        dlg._initial_scroll_done = False
        dlg.show()
        deadline = time.time() + 0.5
        while time.time() < deadline and not dlg._initial_scroll_done:
            app.processEvents()

        assert dlg.browser.verticalScrollBar().value() > 0
        assert len(dlg._unread_blocks) == 1
        assert dlg.next_unread_btn.text() == "🔵 Next Unread (1)"
        assert "<table" in dlg.browser.toHtml()

        dlg.data_changed.connect(lambda: changes.append(True))
        dlg.tdoc_selected.connect(lambda td: tdoc_jumps.append(td))

        # Clickable TDoc link inside timeline emits tdoc_selected
        dlg._on_anchor_clicked(QUrl("tdoc:S2-2612674"))
        assert tdoc_jumps == ["S2-2612674"]

        # Multi-select company filter on TDocEmailsDialog (CheckableComboBox)
        dlg.company_filter_combo.setCheckedItems(["Nokia"])
        dlg.company_filter_combo.selectionChanged.emit(dlg.company_filter_combo.getCheckedItems())
        assert "1 of 2" in dlg.msg_count_lbl.text()
        dlg.company_filter_combo.setCheckedItems(["Nokia", "Samsung"])
        dlg.company_filter_combo.selectionChanged.emit(dlg.company_filter_combo.getCheckedItems())
        assert "2 message(s)" in dlg.msg_count_lbl.text()
        dlg.company_filter_combo.setCheckedItems([])
        dlg.company_filter_combo.selectionChanged.emit(dlg.company_filter_combo.getCheckedItems())

        # Include Quoted toggle hides quoted match, restoring restores it
        dlg.chk_include_quoted.setChecked(False)
        assert len(dlg.emails) == 1
        dlg.chk_include_quoted.setChecked(True)
        assert len(dlg.emails) == 2

        # Inline mark read, ignore/show ignored, reload, pin, and CLI launch
        dlg._on_anchor_clicked(QUrl("outlook://read/E-UNREAD-QUOTED"))
        assert "0 Unread" in dlg.lbl_counts.text()

        dlg._on_anchor_clicked(QUrl("outlook://ignore/E-UNREAD-QUOTED"))
        assert len(dlg.emails) == 1
        dlg.chk_show_ignored.setChecked(True)
        assert len(dlg.emails) == 2
        assert "IGNORED" in dlg.browser.toPlainText()

        dlg._on_anchor_clicked(QUrl("outlook://reload/E-OLD-READ"))
        reloaded_rows = {r["id"]: r for r in gdb.get_emails_for_tdocs({"S2-2612673"}, show_ignored=True)}
        assert "| A | B |" in reloaded_rows["E-OLD-READ"]["body_text"]

        dlg._on_pin_as_tracked_thread()
        assert len(pinned_calls) == 1
        assert "S2-2612673" in pinned_calls[0]["linked_tdocs"]

        with patch("modules.emails.ui.tdoc_emails_dialog.launch_cli_in_terminal", return_value=True) as mock_cli:
            dlg._on_launch_cli("codex")
            mock_cli.assert_called_once()
            assert mock_cli.call_args.kwargs.get("auto_execute") is False
            assert (mock_cli.call_args.args[1] / "_CLI_CONTEXT.md").is_file()

        assert len(changes) >= 3
        dlg._settings.setValue("newest_first", True)
        dlg.close()


def test_inbox_unassigned_filters_grouping_and_inline_assignment(tmp_path):
    from unittest.mock import patch
    from PyQt5.QtWidgets import QApplication
    from modules.emails.core.general_email_db import GeneralEmailDatabase
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.core.tracked_threads_service import auto_suggest_links
    from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog

    app = QApplication.instance() or QApplication([])
    tdoc_meta = {
        "S2-2612711": {"ai": "20.6.1", "sub_item": "KI#4 Bullet 3", "wis": ["FS_6G_ARC"]},
        "S2-2612999": {"ai": "", "sub_item": "", "wis": []},
    }

    # 1. TDoc metadata inheritance in auto_suggest_links (subject only mentions TDoc revision)
    sugg = auto_suggest_links(
        "Re: [SA2#178] draft S2-2612711r01 comments",
        valid_wis=["FS_6G_ARC", "OTHER_WI"],
        tdoc_meta_map=tdoc_meta,
    )
    assert sugg["ais"] == ["20.6.1"]
    assert sugg["sub_items"] == ["KI#4 Bullet 3"]
    assert sugg["wis"] == ["FS_6G_ARC"]

    # 2. Setup mixed DB: (a) Tracked unassigned thread with 2 messages, (b) Tracked assigned thread, (c) Unpinned TDoc email
    tdb = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid_unassigned = tdb.create_or_get_thread("Breakout No AI", "Breakout No AI", meeting_id="SA2#178")
    tdb.upsert_messages(
        tid_unassigned,
        [
            {"id": "M-U1", "sender_name": "Alice", "company": "Nokia", "date_sent": "2026-11-18 09:00:00", "body_text": "First msg", "is_read": 1},
            {"id": "M-U2", "sender_name": "Bob", "company": "Samsung", "date_sent": "2026-11-18 10:00:00", "body_text": "Second msg", "is_read": 0},
        ],
    )
    tid_assigned = tdb.create_or_get_thread(
        "Assigned Thread", "Assigned Thread", meeting_id="SA2#178", linked_ais=["20.6.1"]
    )
    tdb.upsert_messages(
        tid_assigned,
        [{"id": "M-A1", "sender_name": "Carol", "company": "Ericsson", "date_sent": "2026-11-18 11:00:00", "body_text": "Assigned msg", "is_read": 0}],
    )

    gdb = GeneralEmailDatabase(tmp_path / "general.db")
    gdb.save_emails_batch(
        [
            {
                "id": "E-UNPIN",
                "subject": "[SA2#178] Unpinned TDoc S2-2612999",
                "sender_name": "Dave",
                "company": "Huawei",
                "date_received": "2026-11-18 12:00:00",
                "body_text": "Need to discuss S2-2612999",
                "is_read": 0,
            }
        ],
        [{"email_id": "E-UNPIN", "tdoc_id": "S2-2612999", "match_location": "Subject"}],
    )

    from PyQt5.QtCore import QSettings as _QSettings
    with patch.object(OutlookClient, "set_messages_read_status", return_value=1), \
         patch.object(OutlookClient, "get_namespace", return_value=None):
        dlg = MeetingEmailInboxDialog(
            tdb,
            gdb,
            meeting_id="SA2#178",
            ai_suggestions=["20.6.1 — FS_6G_ARC", "20.2 — Architecture"],
            wi_suggestions=["FS_6G_ARC"],
            ai_to_sub_items={"20.6.1": ["KI#4 Bullet 3"], "20.2": ["Mobility"]},
            tdoc_meta_map=tdoc_meta,
            settings=_QSettings(str(tmp_path / "inbox_unassigned.ini"), _QSettings.IniFormat),
        )
        dlg._set_view_mode("MESSAGES")
        # Default By Message has 3 unread rows (M-U2, M-A1, E-UNPIN)
        assert dlg.table.rowCount() == 3

        # Multi-select AI filter: __UNASSIGNED__ -> isolates M-U2 (no AI) and E-UNPIN (no AI)
        idx_unassigned = dlg.ai_combo.findData("__UNASSIGNED__")
        assert idx_unassigned >= 1
        assert "(2)" in dlg.ai_combo.itemText(idx_unassigned)
        dlg.ai_combo.setCheckedItems(["__UNASSIGNED__"])
        dlg.ai_combo.selectionChanged.emit(dlg.ai_combo.getCheckedItems())
        assert dlg.table.rowCount() == 2

        # Multi-select AI: __UNASSIGNED__ + 20.6.1 -> shows all 3 unread rows
        dlg.ai_combo.setCheckedItems(["__UNASSIGNED__", "20.6.1"])
        dlg.ai_combo.selectionChanged.emit(dlg.ai_combo.getCheckedItems())
        assert dlg.table.rowCount() == 3

        # Multi-select Company: Samsung + Huawei -> isolates M-U2 & E-UNPIN (2 rows), excluding Ericsson
        dlg.ai_combo.setCheckedItems([])
        dlg.ai_combo.selectionChanged.emit(dlg.ai_combo.getCheckedItems())
        dlg.company_combo.setCheckedItems(["Samsung", "Huawei"])
        dlg.company_combo.selectionChanged.emit(dlg.company_combo.getCheckedItems())
        assert dlg.table.rowCount() == 2
        dlg.company_combo.setCheckedItems([])
        dlg.company_combo.selectionChanged.emit(dlg.company_combo.getCheckedItems())

        # Filter AI = __UNASSIGNED__ & Source = UNPINNED -> isolates E-UNPIN only
        dlg.ai_combo.setCheckedItems(["__UNASSIGNED__"])
        dlg.ai_combo.selectionChanged.emit(dlg.ai_combo.getCheckedItems())
        dlg.source_combo.setCurrentIndex(dlg.source_combo.findData("UNPINNED"))
        assert dlg.table.rowCount() == 1
        assert dlg._current_item["id"] == "E-UNPIN"
        assert not dlg.btn_pin_inline.isHidden()

        # Assign AI inline directly on the unpinned email -> auto-pins in-place & links AI 20.2 WITHOUT blocking Outlook COM
        dlg.source_combo.setCurrentIndex(0)
        dlg.ai_combo.setCheckedItems([])
        dlg.ai_combo.selectionChanged.emit(dlg.ai_combo.getCheckedItems())
        assert dlg._current_item and dlg._current_item["id"] == "E-UNPIN"
        with patch("modules.emails.core.tracked_threads_service.refresh_thread") as mock_refresh_thread:
            dlg.assign_ais_pills.add_tag("20.2")
            mock_refresh_thread.assert_not_called()
        pinned_list = [t for t in tdb.get_threads_for_meeting("SA2#178") if "S2-2612999" in (t.get("linked_tdocs") or [])]
        assert len(pinned_list) == 1
        assert pinned_list[0]["linked_ais"] == ["20.2"]

        # Switch to By Thread grouped mode + All pill -> shows 3 distinct threads
        dlg.pill_all.click()
        dlg._set_view_mode("THREADS")
        assert dlg.table.rowCount() == 3

        # Select unassigned thread (tid_unassigned) and assign 20.6.1 + Sub-Topic inline
        target_row = next(r for r, it in enumerate(dlg._visible_items) if it.get("thread_id") == tid_unassigned)
        dlg.table.selectRow(target_row)
        assert dlg._current_item["_total_count"] == 2
        assert "First msg" in dlg.preview_browser.toPlainText()
        assert "Second msg" in dlg.preview_browser.toPlainText()

        dlg.assign_ais_pills.add_tag("20.6.1")
        dlg.assign_subs_pills.add_tag("KI#4 Bullet 3")
        dlg.assign_wis_pills.add_tag("FS_6G_ARC")
        updated_th = tdb.get_thread(tid_unassigned)
        assert updated_th["linked_ais"] == ["20.6.1"]
        assert updated_th["linked_sub_items"] == ["KI#4 Bullet 3"]
        assert updated_th["linked_wis"] == ["FS_6G_ARC"]

        # Clickable TDoc links in header & body + 2FS2-... cleanup
        opened_tdocs_from_links = []
        dlg._open_tdoc_callback = lambda td: opened_tdocs_from_links.append(td)
        dlg.reload_from_dbs()
        dlg.table.selectRow(0)
        if dlg._current_item and dlg._current_item.get("linked_tdocs"):
            dlg._render_detail(dlg._current_item)
        # Verify TrackedThreadDialog & Inbox linkActivated / anchorClicked emit clean TDoc IDs
        tdb.update_thread_links(tid_assigned, linked_tdocs=["2FS2-2612556", "S2-2612556", "S2-2611691"])
        dlg.reload_from_dbs(preserve_id=f"THR::{tid_assigned}" if dlg._view_mode == "THREADS" else "M-A1")
        assert "2FS2-" not in dlg.detail_links_lbl.text()
        assert 'href="tdoc:S2-2612556"' in dlg.detail_links_lbl.text()
        dlg.detail_links_lbl.linkActivated.emit("tdoc:S2-2612556")
        assert opened_tdocs_from_links == ["S2-2612556"]

        from PyQt5.QtCore import QUrl as _QUrl
        from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog
        tdlg_opens = []
        tdlg = TrackedThreadDialog(tdb, tid_assigned, open_tdoc_callback=lambda td: tdlg_opens.append(td))
        tdlg._on_anchor_clicked(_QUrl("tdoc:S2-2611691"))
        assert tdlg_opens == ["S2-2611691"]
        tdlg.close()

        # Verify table cell hover tooltips and Side-by-Side <-> Top/Bottom layout toggle
        from PyQt5.QtCore import Qt as _QtOrient
        assert dlg.table.rowCount() >= 1
        assert dlg.table.item(0, 3).toolTip() == dlg.table.item(0, 3).text()
        assert dlg.table.item(0, 1).toolTip()
        assert dlg.table.item(0, 2).toolTip()
        assert dlg.table.item(0, 4).toolTip()

        dlg._layout_orientation = "H"
        dlg._apply_layout_orientation(save_current=False)
        assert dlg.splitter.orientation() == _QtOrient.Horizontal
        assert "List on Top" in dlg.btn_layout_toggle.text()
        dlg._toggle_layout_orientation()
        assert dlg.splitter.orientation() == _QtOrient.Vertical
        assert "Side-by-Side" in dlg.btn_layout_toggle.text()
        dlg._toggle_layout_orientation()
        assert dlg.splitter.orientation() == _QtOrient.Horizontal

        dlg._settings.setValue("layout_orientation", "H")
        dlg._settings.setValue("view_mode", "MESSAGES")
        dlg.close()


def test_meeting_inbox_filter_persistence_per_meeting(tmp_path):
    from unittest.mock import patch
    from PyQt5.QtCore import QSettings
    from PyQt5.QtWidgets import QApplication

    app = QApplication.instance() or QApplication([])
    from modules.emails.core.general_email_db import GeneralEmailDatabase
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog

    tdb = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid1 = tdb.create_or_get_thread("Paging Discussion", "Paging Discussion", meeting_id="SA2#178", linked_ais=["20.6.1"])
    tdb.upsert_messages(
        tid1,
        [{"id": "M-1", "sender_name": "Alice", "company": "Nokia", "date_sent": "2026-11-18 09:00:00", "body_text": "paging", "is_read": 1}],
    )
    tid2 = tdb.create_or_get_thread("Breakout Unassigned", "Breakout Unassigned", meeting_id="SA2#178")
    tdb.upsert_messages(
        tid2,
        [{"id": "M-2", "sender_name": "Bob", "company": "Samsung", "date_sent": "2026-11-18 10:00:00", "body_text": "paging note", "is_read": 0}],
    )
    gdb = GeneralEmailDatabase(tmp_path / "general.db")
    gdb.save_emails_batch(
        [
            {
                "id": "E-1",
                "subject": "[SA2#178] Paging baseline",
                "sender_name": "Carol",
                "company": "Huawei",
                "date_received": "2026-11-18 11:00:00",
                "body_text": "paging comment",
                "is_read": 0,
            }
        ],
        [{"email_id": "E-1", "tdoc_id": "S2-2612999", "match_location": "Subject"}],
    )

    ini_path = str(tmp_path / "inbox_persist.ini")
    with patch.object(OutlookClient, "get_namespace", return_value=None):
        dlg1 = MeetingEmailInboxDialog(
            tdb,
            gdb,
            meeting_id="SA2#178",
            ai_suggestions=["20.6.1 — FS_6G_ARC", "20.2 — Architecture"],
            settings=QSettings(ini_path, QSettings.IniFormat),
        )
        dlg1.pill_all.click()
        dlg1.btn_view_threads.click()
        dlg1.source_combo.setCurrentIndex(dlg1.source_combo.findData("NEEDS_TRIAGE"))
        dlg1.company_combo.setCheckedItems(["Samsung", "Huawei"])
        dlg1.company_combo.selectionChanged.emit(dlg1.company_combo.getCheckedItems())
        dlg1.ai_combo.setCheckedItems(["__UNASSIGNED__"])
        dlg1.ai_combo.selectionChanged.emit(dlg1.ai_combo.getCheckedItems())
        dlg1.search_edit.setText("paging")
        assert dlg1.table.rowCount() == 2  # Breakout Unassigned + E-1

        # Simulate app restart WITHOUT closing dlg1 explicitly first -> state must already be persisted per-meeting
        dlg2 = MeetingEmailInboxDialog(
            tdb,
            gdb,
            meeting_id="SA2#178",
            ai_suggestions=["20.6.1 — FS_6G_ARC", "20.2 — Architecture"],
            settings=QSettings(ini_path, QSettings.IniFormat),
        )
        assert dlg2._active_mode == "ALL"
        assert dlg2.pill_all.isChecked()
        assert dlg2._view_mode == "THREADS"
        assert dlg2.btn_view_threads.isChecked()
        assert dlg2.source_combo.currentData() == "NEEDS_TRIAGE"
        assert dlg2.search_edit.text() == "paging"
        assert set(dlg2.company_combo.getCheckedItems()) == {"Samsung", "Huawei"}
        assert set(dlg2.ai_combo.getCheckedItems()) == {"__UNASSIGNED__"}
        assert dlg2.table.rowCount() == 2

        # Different meeting keeps clean independent defaults
        dlg_other = MeetingEmailInboxDialog(
            tdb,
            gdb,
            meeting_id="RAN2#132",
            ai_suggestions=["7.1 — Study"],
            settings=QSettings(ini_path, QSettings.IniFormat),
        )
        assert dlg_other._active_mode == "UNREAD"
        assert dlg_other.pill_unread.isChecked()
        assert dlg_other._view_mode == "MESSAGES"
        assert dlg_other.btn_view_messages.isChecked()
        assert dlg_other.source_combo.currentData() == ""
        assert dlg_other.search_edit.text() == ""
        assert not dlg_other.company_combo.isFiltered()
        assert not dlg_other.ai_combo.isFiltered()

        dlg1.close()
        dlg2.close()
        dlg_other.close()


def test_inbox_multi_select_batch_ai_wi_assignment(tmp_path):
    from unittest.mock import patch
    from PyQt5.QtCore import QItemSelectionModel, QSettings
    from PyQt5.QtWidgets import QAbstractItemView, QApplication

    app = QApplication.instance() or QApplication([])
    from modules.emails.core.general_email_db import GeneralEmailDatabase
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
    from modules.emails.ui.meeting_email_inbox_dialog import MeetingEmailInboxDialog

    tdb = TrackedThreadsDatabase(tmp_path / "threads.db")
    tid_pre = tdb.create_or_get_thread("Mobility Thread", "Mobility Thread", meeting_id="SA2#178", linked_ais=["20.2"])
    tdb.upsert_messages(
        tid_pre,
        [{"id": "M-PRE", "sender_name": "Alice", "company": "Nokia", "date_sent": "2026-11-18 09:00:00", "body_text": "pre", "is_read": 0}],
    )
    tid_blank = tdb.create_or_get_thread("Blank Tracked Thread", "Blank Tracked Thread", meeting_id="SA2#178")
    tdb.upsert_messages(
        tid_blank,
        [{"id": "M-BLK", "sender_name": "Bob", "company": "Samsung", "date_sent": "2026-11-18 10:00:00", "body_text": "blank", "is_read": 0}],
    )
    gdb = GeneralEmailDatabase(tmp_path / "general.db")
    gdb.save_emails_batch(
        [
            {
                "id": "E-UNPIN-B",
                "subject": "[SA2#178] Baseline paper S2-2612027",
                "sender_name": "Carol",
                "company": "Huawei",
                "date_received": "2026-11-18 11:00:00",
                "body_text": "Discuss baseline",
                "is_read": 0,
            }
        ],
        [{"email_id": "E-UNPIN-B", "tdoc_id": "S2-2612027", "match_location": "Subject"}],
    )

    with patch.object(OutlookClient, "set_messages_read_status", return_value=1), \
         patch.object(OutlookClient, "set_messages_flag_status", return_value=1), \
         patch.object(OutlookClient, "get_namespace", return_value=None):
        dlg = MeetingEmailInboxDialog(
            tdb,
            gdb,
            meeting_id="SA2#178",
            ai_suggestions=["20.6.1 — FS_6G_ARC", "20.2 — Architecture"],
            wi_suggestions=["FS_6G_ARC — 6G Study"],
            ai_to_sub_items={"20.6.1": ["KI#3 Bullet 2"], "20.2": ["Mobility"]},
            settings=QSettings(str(tmp_path / "batch.ini"), QSettings.IniFormat),
        )
        assert dlg.table.selectionMode() == QAbstractItemView.ExtendedSelection
        dlg._set_view_mode("THREADS")
        assert dlg.table.rowCount() == 3

        # Select all 3 threads at once
        dlg.table.selectAll()
        assert len(dlg._selected_items) == 3
        assert "3 Selected Threads" in dlg.detail_subject_lbl.text()
        assert "Batch Triage" in dlg.preview_browser.toPlainText()

        # Assign AI 20.6.1 + WI FS_6G_ARC to all 3 at once
        dlg.assign_ais_pills.add_tag("20.6.1")
        assert len(dlg._selected_items) == 3  # selection preserved across batch reload
        dlg.assign_wis_pills.add_tag("FS_6G_ARC")

        threads_by_title = {t["title"]: t for t in tdb.get_threads_for_meeting("SA2#178")}
        # Pre-existing 20.2 on Mobility Thread must be preserved alongside added 20.6.1
        assert threads_by_title["Mobility Thread"]["linked_ais"] == ["20.2", "20.6.1"]
        assert threads_by_title["Mobility Thread"]["linked_wis"] == ["FS_6G_ARC"]
        assert threads_by_title["Blank Tracked Thread"]["linked_ais"] == ["20.6.1"]
        assert threads_by_title["Blank Tracked Thread"]["linked_wis"] == ["FS_6G_ARC"]
        # Unpinned email auto-pinned & assigned
        unpin_thread = next(t for t in threads_by_title.values() if "S2-2612027" in (t.get("linked_tdocs") or []))
        assert unpin_thread["linked_ais"] == ["20.6.1"]
        assert unpin_thread["linked_wis"] == ["FS_6G_ARC"]

        # Remove shared WI FS_6G_ARC from first two selected rows only
        sel = dlg.table.selectionModel()
        dlg.table.clearSelection()
        sel.select(dlg.table.model().index(0, 0), QItemSelectionModel.Select | QItemSelectionModel.Rows)
        sel.select(dlg.table.model().index(1, 0), QItemSelectionModel.Select | QItemSelectionModel.Rows)
        assert len(dlg._selected_items) == 2

        # Batch Flag & Mark Read across selected rows
        dlg._on_toggle_flag()
        assert sum(1 for it in dlg._selected_items if it["is_flagged"]) == 2
        dlg._on_toggle_read()
        assert dlg.table.rowCount() == 1  # remaining unread thread still visible in Unread Only mode
        dlg.close()


def test_cross_meeting_tdoc_direct_open(tmp_path):
    from types import SimpleNamespace
    from unittest.mock import patch, MagicMock
    from modules.meetings.ui.tdocs_window import TDocsWindow

    fake_win = SimpleNamespace()
    fake_win.mtg_info = {"wg_name": "SA2", "meeting_number": "178", "mtg_id": 178, "first_tdoc": "S2-2611000"}
    fake_win.model = SimpleNamespace(valid_tdocs={"S2-2612556"})
    fake_win.docs_ftp_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178_Calgary_2026-11/Docs"
    fake_win.revisions_url = "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178_Calgary_2026-11/INBOX/Revisions"
    fake_win._get_tdoc_prefix = lambda: "S2-26"
    fake_win._get_meeting_id_label = lambda: "SA2#178"
    emitted_global = []
    local_downloads = []
    fake_win.global_action_requested = SimpleNamespace(emit=lambda tdoc, act: emitted_global.append((tdoc, act)))
    fake_win._trigger_download_thread = lambda base, td, url, is_silent_compare=False: local_downloads.append((base, td, url))
    fake_win._scroll_to_tdoc = lambda td: None

    # 1. Local meeting TDoc downloads via current meeting's FTP
    TDocsWindow._open_tdoc_document_directly(fake_win, "S2-2612556r01")
    assert local_downloads[-1] == ("S2-2612556", "S2-2612556R01", fake_win.revisions_url)

    # 2. Unindexed current-meeting draft (prefix S2-26 matches) still tries local Inbox when no other meeting owns it
    db_file = tmp_path / "db" / "3gpp_data.db"
    db_file.parent.mkdir(parents=True, exist_ok=True)
    db_file.write_bytes(b"")
    mock_db = MagicMock()
    mock_db.find_meeting_by_tdoc.side_effect = lambda td: {"mtg_id": 999, "wg_name": "RAN3"} if td.startswith("R3-") else {}

    with patch("modules.meetings.ui.tdocs_window.get_project_root", return_value=tmp_path), \
         patch("modules.meetings.core.meetings_db.MeetingsDatabase", return_value=mock_db):
        TDocsWindow._open_tdoc_document_directly(fake_win, "S2-2612999")
        assert local_downloads[-1] == ("S2-2612999", "S2-2612999", fake_win.docs_ftp_url)

        # 3. External TDoc (e.g. R3-263595 from RAN3) routes via global_action_requested("open_doc"), NOT SA2 FTP
        prev_local_len = len(local_downloads)
        TDocsWindow._open_tdoc_document_directly(fake_win, "R3-263595")
        assert len(local_downloads) == prev_local_len
        assert emitted_global == [("R3-263595", "open_doc")]

        # 4. Unsynced external TDoc (e.g. C1-269999 not in DB) notifies cleanly without hitting SA2 FTP
        with patch("modules.meetings.ui.tdocs_window.QMessageBox.information") as mock_info:
            TDocsWindow._open_tdoc_document_directly(fake_win, "C1-269999")
            mock_info.assert_called_once()
            assert len(local_downloads) == prev_local_len
