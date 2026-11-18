# --- File: src/modules/meetings/ui/tdocs_menu.py ---
from typing import Any

from PyQt5.QtCore import QPoint

from modules.meetings.ui.dialogs.tdocs_export_controller import get_package_export_rows
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

    if row_data.get("_is_tracked_thread"):
        from PyQt5.QtWidgets import QMenu
        from modules.emails.core.tracked_threads_service import refresh_thread

        tid = str(row_data.get("_thread_id") or tdoc_id).strip()
        menu = QMenu(window.table)
        menu.addAction("📧 Open Tracked Email Thread...", lambda: window._open_tracked_thread_dialog(tid))
        menu.addAction(
            "🔄 Refresh This Thread Now",
            lambda: (refresh_thread(window.tracked_threads_db, tid), window._on_tracked_thread_changed()),
        )
        menu.addAction(
            "✔️ Mark Thread as Read",
            lambda: (window.tracked_threads_db.mark_thread_read(tid), window._on_tracked_thread_changed()),
        )
        menu.addSeparator()
        menu.addAction(
            "📌 Unpin Thread",
            lambda: (window.tracked_threads_db.delete_thread(tid), window._on_tracked_thread_changed()),
        )
        menu.exec_(window.table.viewport().mapToGlobal(pos))
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

    selected_rows = [row_data]
    sel_model = window.table.selectionModel() if hasattr(window, "table") else None
    if sel_model and sel_model.selectedIndexes():
        pkg_rows = get_package_export_rows(window)
        if len(pkg_rows) > 1 and any(str(r.get("TDoc", "")).strip() == tdoc_id for r in pkg_rows):
            selected_rows = pkg_rows

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
        draft_cr_callback=window._open_draft_cr_dialog,
        compare_active_word_callback=getattr(window, "_compare_tdoc_with_active_word", None),
        export_dossier_callback=window._export_dossier,
        load_session_plan_callback=getattr(window, "_load_tdoc_as_session_plan", None),
        launch_cli_callback=getattr(window, "_launch_cli_for_tdocs", None),
        selected_rows=selected_rows,
        send_tdoc_email_callback=getattr(window, "_send_tdoc_via_email", None),
    )
