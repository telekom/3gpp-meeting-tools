"""Unified right-click command center context menu for TDoc table rows."""

from __future__ import annotations

from pathlib import Path

from PyQt5.QtCore import QPoint
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QMenu, QWidget

from core.ai.cli_launcher import populate_cli_agent_menu
from modules.meetings.ui.tdocs.menus.quick_copy import (
    is_tdoc_cached,
    open_3gu_portal,
    open_tdoc_local_folder,
    populate_quick_copy_menu,
)


def build_row_context_menu(
    parent: QWidget,
    row_data: dict,
    revisions: list,
    family: list,
    docs_ftp_url: str,
    revisions_url: str,
    meeting_dir: Path,
    download_callback,
    export_llm_callback,
    compose_email_callback,
    open_emails_callback,
    mark_read_callback,
    mark_unread_callback,
    open_details_callback,
    unread_emails_count: int = 0,
    pos: QPoint = None,
    triage_callback=None,
    draft_cr_callback=None,
    compare_active_word_callback=None,
    export_dossier_callback=None,
    load_session_plan_callback=None,
    launch_cli_callback=None,
    selected_rows=None,
    send_tdoc_email_callback=None,
    edit_notes_callback=None,
    open_ai_briefing_callback=None,
):
    """Complete command center menu for row-level operations."""
    menu = QMenu(parent)
    menu.setToolTipsVisible(True)
    revisions = revisions or []
    tdoc_id = str(row_data.get("TDoc", "")).strip()

    # --- 1. INSPECT FULL DETAILS & AI TRIAGE ---
    act_details = menu.addAction(f"ℹ️ View Full Details for {tdoc_id}...")
    act_details.setToolTip("Show all stored metadata and links for this TDoc.")
    act_details.triggered.connect(open_details_callback)

    if edit_notes_callback:
        act_notes = menu.addAction("🏷️ Edit Notes, Status && Tags...")
        act_notes.setToolTip("Edit personal stance (My Status), Tags, and Notes for this item.")
        act_notes.triggered.connect(edit_notes_callback)

    if triage_callback:
        act_triage = menu.addAction(f"🤖 AI Technical Triage for {tdoc_id}...")
        act_triage.setToolTip("Run the local-LLM technical triage for this TDoc.")
        act_triage.triggered.connect(lambda: triage_callback(tdoc_id))

    if launch_cli_callback:
        cli_rows = list(selected_rows) if (selected_rows and len(selected_rows) > 1) else [row_data]
        cli_label = (
            f"💬 Ask AI Agent (CLI) about {len(cli_rows)} Selected TDocs"
            if len(cli_rows) > 1
            else f"💬 Ask AI Agent (CLI) about {tdoc_id}"
        )
        cli_sub = menu.addMenu(cli_label)
        populate_cli_agent_menu(
            cli_sub,
            on_launch_callback=lambda aid, mode, rows=cli_rows: launch_cli_callback(rows, aid, mode),
            has_multiple_or_revs=(len(cli_rows) > 1 or bool(revisions)),
            multi_count=len(cli_rows),
            parent=parent,
        )

    if draft_cr_callback:
        act_draft_cr = menu.addAction("📝 Draft CR / pCR from Spec...")
        act_draft_cr.setToolTip("Compose a Change Request (CR) or Pseudo-CR (pCR) pre-populated from this TDoc.")
        act_draft_cr.triggered.connect(lambda: draft_cr_callback(row_data))

    # --- 1b. SESSION PLAN INGESTION SHORTCUT ---
    title_lower = str(row_data.get("Title", "")).lower()
    if "session plan" in title_lower or "time allocation" in title_lower:
        act_sp = menu.addAction(f"🗓️ Load {tdoc_id} into Session Plan...")
        act_sp.setToolTip("Parse and load this TDoc into the Meeting Session Plan viewer.")
        if load_session_plan_callback:
            act_sp.triggered.connect(lambda: load_session_plan_callback(tdoc_id))

    menu.addSeparator()

    # --- 2. OPEN DOCUMENT SUBMENU ---
    open_sub = menu.addMenu("📂 Open Document")
    open_sub.setToolTipsVisible(True)
    open_sub.setToolTip("Open the base TDoc or one of its revisions, using the local cache when available.")
    base_cached = is_tdoc_cached(meeting_dir, tdoc_id)

    if revisions:
        latest_rev = revisions[-1]
        latest_cached = is_tdoc_cached(meeting_dir, latest_rev)
        act_open_lat = open_sub.addAction(
            f"🚀 Open Latest: {latest_rev}" + (" (Local)" if latest_cached else "")
        )
        act_open_lat.setToolTip("Open the newest known revision of this TDoc.")
        act_open_lat.triggered.connect(
            lambda: download_callback(tdoc_id, latest_rev, is_silent_compare=False)
        )
        open_sub.addSeparator()

        act_open_base = open_sub.addAction(
            f"📄 Open Base: {tdoc_id}" + (" (Local)" if base_cached else "")
        )
        act_open_base.setToolTip("Open the baseline TDoc.")
        act_open_base.triggered.connect(
            lambda: download_callback(tdoc_id, tdoc_id, is_silent_compare=False)
        )

        for rev in revisions:
            rev_cached = is_tdoc_cached(meeting_dir, rev)
            act_open_r = open_sub.addAction(
                f"📄 Open Rev: {rev}" + (" (Local)" if rev_cached else "")
            )
            act_open_r.triggered.connect(
                lambda _, r=rev: download_callback(tdoc_id, r, is_silent_compare=False)
            )
    else:
        act_open_base = open_sub.addAction(
            f"📄 Open Base: {tdoc_id}" + (" (Local)" if base_cached else "")
        )
        act_open_base.setToolTip("Open the baseline TDoc.")
        act_open_base.triggered.connect(
            lambda: download_callback(tdoc_id, tdoc_id, is_silent_compare=False)
        )

    open_sub.addSeparator()
    act_folder = open_sub.addAction("📂 Open Local Folder")
    act_folder.setToolTip("Open the local folder used to cache this TDoc and its extracted files.")
    act_folder.triggered.connect(lambda: open_tdoc_local_folder(meeting_dir, tdoc_id))

    act_portal = open_sub.addAction("🌐 View TDoc on 3GU Portal")
    act_portal.setToolTip("Open this TDoc in the official 3GPP 3GU portal.")
    act_portal.triggered.connect(lambda: open_3gu_portal(tdoc_id))

    # --- 3. COMPARE WITH ACTIVE WORD & COMPARISON CART SUBMENU ---
    if compare_active_word_callback:
        if revisions:
            latest_rev = revisions[-1]
            act_cmp_lat = menu.addAction(f"⚖️ Compare Latest ({latest_rev}) vs Active Word")
            act_cmp_lat.setToolTip(
                "Immediately compare the latest revision of this TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_lat.triggered.connect(lambda: compare_active_word_callback(tdoc_id, latest_rev))

            cmp_rev_sub = menu.addMenu(f"⚖️ Compare {tdoc_id} (Base / Revs) vs Active Word")
            cmp_rev_sub.setToolTipsVisible(True)
            act_cmp_base = cmp_rev_sub.addAction(f"⚖️ Compare Base: {tdoc_id} vs Active Word")
            act_cmp_base.setToolTip("Compare baseline TDoc against active Word document.")
            act_cmp_base.triggered.connect(lambda: compare_active_word_callback(tdoc_id, tdoc_id))
            cmp_rev_sub.addSeparator()
            for rev in revisions:
                act_cmp_r = cmp_rev_sub.addAction(f"⚖️ Compare Rev: {rev} vs Active Word")
                act_cmp_r.triggered.connect(
                    lambda _, r=rev: compare_active_word_callback(tdoc_id, r)
                )
        else:
            act_cmp_active = menu.addAction(f"⚖️ Compare {tdoc_id} vs Active Word")
            act_cmp_active.setToolTip(
                "Immediately compare this TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_active.triggered.connect(lambda: compare_active_word_callback(tdoc_id, tdoc_id))

    cart_sub = menu.addMenu("⚖️ Comparison Cart")
    cart_sub.setToolTipsVisible(True)
    cart_sub.setToolTip("Add the base TDoc, a revision, or an open Word document to the comparison cart.")
    act_cart_base = cart_sub.addAction(f"➕ Add Base: {tdoc_id}")
    act_cart_base.setToolTip("Download/cache the baseline as needed and add it to the comparison cart.")
    act_cart_base.triggered.connect(
        lambda: download_callback(tdoc_id, tdoc_id, is_silent_compare=True)
    )

    if revisions:
        for rev in revisions:
            act_cart_rev = cart_sub.addAction(f"➕ Add Rev: {rev}")
            act_cart_rev.setToolTip("Download/cache this revision as needed and add it to the comparison cart.")
            act_cart_rev.triggered.connect(
                lambda _, r=rev: download_callback(tdoc_id, r, is_silent_compare=True)
            )

    cart_sub.addSeparator()
    from modules.meetings.ui.tdocs_components import populate_slot_word_menu

    word_slot_a_sub = cart_sub.addMenu("🪟 Stage Open Word Doc in Slot A (Original)")
    word_slot_a_sub.setToolTipsVisible(True)
    populate_slot_word_menu(word_slot_a_sub, "A", parent)

    word_slot_b_sub = cart_sub.addMenu("🪟 Stage Open Word Doc in Slot B (Revised)")
    word_slot_b_sub.setToolTipsVisible(True)
    populate_slot_word_menu(word_slot_b_sub, "B", parent)

    # --- 4. MARKDOWN & DOSSIER EXPORT ---
    if export_dossier_callback or export_llm_callback:
        if export_dossier_callback:
            act_dossier = menu.addAction("📑 Export to Navigable Dossier (HTML && PDF)...")
            act_dossier.setToolTip(
                "Compile this TDoc into an interactive HTML dossier and bookmarked vector PDF with companion AI citation prompt."
            )
            act_dossier.triggered.connect(lambda: export_dossier_callback([row_data]))
        if export_llm_callback:
            act_md = menu.addAction("📄 Export to Markdown (.md)")
            act_md.setToolTip("Export this TDoc as an individual Markdown document with metadata and extracted figures.")
            act_md.triggered.connect(lambda: export_llm_callback(tdoc_id))
        menu.addSeparator()

    # --- 5. EMAILS SUBMENU ---
    email_sub = menu.addMenu("📧 Emails")
    email_sub.setToolTipsVisible(True)
    email_sub.setToolTip("Inspect, draft, or update read state for emails associated with this TDoc family.")
    email_label = (
        f"👁️ View Related Emails ({unread_emails_count} unread)..."
        if unread_emails_count > 0
        else f"👁️ View Related Emails for {tdoc_id}..."
    )
    act_view_emails = email_sub.addAction(email_label)
    act_view_emails.setToolTip("Open locally indexed emails associated with this TDoc family.")
    act_view_emails.triggered.connect(lambda: open_emails_callback(tdoc_id))

    row_ai = str(row_data.get("Agenda Item") or "").strip()
    if open_ai_briefing_callback and row_ai and row_ai not in ("-", "N/A", "None"):
        act_ai_brief = email_sub.addAction(f"🤖 AI Briefing for AI {row_ai}...")
        act_ai_brief.setToolTip(f"Open the Per-AI email cross-thinking briefing for Agenda Item {row_ai}.")
        act_ai_brief.triggered.connect(lambda _=False, ai=row_ai: open_ai_briefing_callback(ai))

    act_draft_email = email_sub.addAction("✉️ Draft Email with Subject...")
    act_draft_email.setToolTip("Create a new email draft with meeting, AI and TDoc information in the subject.")
    act_draft_email.triggered.connect(lambda: compose_email_callback(tdoc_id))

    if send_tdoc_email_callback:
        if revisions:
            latest_rev = revisions[-1]
            latest_cached = is_tdoc_cached(meeting_dir, latest_rev)
            act_send_latest = email_sub.addAction(
                f"📎 Send Latest ({latest_rev}) via Email..." + (" (Local)" if latest_cached else "")
            )
            act_send_latest.setToolTip(
                "Create an Outlook email draft with the TDoc subject and attach the latest revision (downloads automatically if needed)."
            )
            act_send_latest.triggered.connect(lambda: send_tdoc_email_callback(latest_rev))

            send_ver_sub = email_sub.addMenu("📎 Send Specific Version via Email")
            send_ver_sub.setToolTipsVisible(True)
            act_send_base = send_ver_sub.addAction(
                f"📄 Send Base: {tdoc_id}" + (" (Local)" if base_cached else "")
            )
            act_send_base.setToolTip("Draft an email with the TDoc subject and attach the baseline TDoc.")
            act_send_base.triggered.connect(lambda: send_tdoc_email_callback(tdoc_id))
            send_ver_sub.addSeparator()
            for rev in revisions:
                rev_cached = is_tdoc_cached(meeting_dir, rev)
                act_send_r = send_ver_sub.addAction(
                    f"📄 Send Rev: {rev}" + (" (Local)" if rev_cached else "")
                )
                act_send_r.setToolTip("Draft an email with the TDoc subject and attach this revision.")
                act_send_r.triggered.connect(lambda _, r=rev: send_tdoc_email_callback(r))
        else:
            act_send_tdoc = email_sub.addAction(
                f"📎 Send {tdoc_id} via Email..." + (" (Local)" if base_cached else "")
            )
            act_send_tdoc.setToolTip(
                "Create an Outlook email draft with the TDoc subject and attach this TDoc (downloads automatically if needed)."
            )
            act_send_tdoc.triggered.connect(lambda: send_tdoc_email_callback(tdoc_id))

    email_sub.addSeparator()
    act_mark_read = email_sub.addAction(f"✔️ Mark all emails as read for {tdoc_id} family")
    act_mark_read.setToolTip("Mark indexed emails for this TDoc family as read.")
    act_mark_read.triggered.connect(mark_read_callback)

    act_mark_unread = email_sub.addAction(f"✉️ Mark all emails as unread for {tdoc_id} family")
    act_mark_unread.setToolTip("Mark indexed emails for this TDoc family as unread.")
    act_mark_unread.triggered.connect(mark_unread_callback)
    menu.addSeparator()

    # --- 6. QUICK COPY SUBMENU ---
    copy_sub = menu.addMenu("📋 Quick Copy")
    copy_sub.setToolTipsVisible(True)
    copy_sub.setToolTip("Copy commonly used TDoc metadata in plain text and, where useful, rich HTML form.")
    populate_quick_copy_menu(
        menu=copy_sub,
        tdoc_id=tdoc_id,
        row_data=row_data,
        docs_ftp_url=docs_ftp_url,
        revisions_url=revisions_url,
        parent=parent,
    )

    menu.exec_(pos or QCursor.pos())
