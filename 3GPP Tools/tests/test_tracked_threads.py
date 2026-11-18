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
