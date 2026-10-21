# --- File: src/modules/meetings/ui/tdocs_menu.py ---
from typing import Any

from PyQt5.QtCore import QPoint

from modules.meetings.ui.tdoc_info_dialog import TDocInfoDialog
from modules.meetings.ui.tdocs_menus import build_row_context_menu


def handle_table_context_menu(window: Any, pos: QPoint) -> None:
    """
    Renders context menu for a row in the TDocs table:
    - Open / download document and revisions
    - Copy TDoc number / title
    - View email discussion threads
    - Mark read / unread
    - View full metadata dialog and launch AI triage
    """
    index = window.table.indexAt(pos)
    if not index.isValid():
        return
    source_idx = window.proxy.mapToSource(index)
    row_data = window.model._data[source_idx.row()]
    tdoc_id = row_data.get("TDoc", "")
    if not tdoc_id:
        return

    revisions = window.model.revisions.get(tdoc_id, [])
    family = window.model.get_family_tdocs(tdoc_id)

    unread_count = 0
    if hasattr(window.model, "_email_counts") and isinstance(window.model._email_counts, dict):
        counts = window.model._email_counts.get(tdoc_id)
        if isinstance(counts, (tuple, list)) and len(counts) > 1:
            unread_count = counts[1]
        elif isinstance(counts, dict):
            unread_count = counts.get("unread", 0)

    build_row_context_menu(
        parent=window.table,
        row_data=row_data,
        revisions=revisions,
        family=family,
        docs_ftp_url=window.docs_ftp_url,
        revisions_url=window.revisions_url,
        meeting_dir=window.meeting_dir,
        download_callback=window._trigger_download_thread,
        export_llm_callback=window._export_llm_single,
        compose_email_callback=window._compose_email_draft,
        open_emails_callback=window._open_tdoc_emails,
        mark_read_callback=lambda: [
            window.general_email_db.set_tdocs_read_status(set(family), True),
            window._refresh_email_counts(),
        ],
        mark_unread_callback=lambda: [
            window.general_email_db.set_tdocs_read_status(set(family), False),
            window._refresh_email_counts(),
        ],
        open_details_callback=lambda: TDocInfoDialog(
            row_data,
            docs_ftp_url=window.docs_ftp_url,
            revisions=revisions,
            parent=window,
            triage_callback=window._open_tdoc_triage,
        ).exec_(),
        unread_emails_count=unread_count,
        pos=window.table.viewport().mapToGlobal(pos),
        triage_callback=window._open_tdoc_triage,
    )
