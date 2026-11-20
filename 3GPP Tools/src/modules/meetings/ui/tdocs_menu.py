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
        from PyQt5.QtWidgets import QDialog, QMenu
        from core.ui.toast import notify_info, notify_success
        from modules.emails.core.tracked_threads_service import (
            mark_thread_read_with_outlook,
            refresh_thread,
        )
        from modules.emails.ui.merge_threads_dialog import MergeThreadsDialog

        tid = str(row_data.get("_thread_id") or tdoc_id).strip()
        selected_thread_ids: list = []
        seen_tids = set()
        sel_model = window.table.selectionModel() if hasattr(window, "table") else None
        if sel_model and sel_model.selectedIndexes():
            for pkg_row in get_package_export_rows(window):
                if pkg_row.get("_is_tracked_thread"):
                    ptid = str(pkg_row.get("_thread_id") or pkg_row.get("TDoc") or "").strip()
                    if ptid and ptid not in seen_tids:
                        seen_tids.add(ptid)
                        selected_thread_ids.append(ptid)
        if tid and tid not in seen_tids:
            selected_thread_ids.insert(0, tid)

        def _launch_merge_threads():
            tdb = window.tracked_threads_db
            selected_records = []
            for sel_tid in selected_thread_ids:
                rec = tdb.get_thread(sel_tid)
                if rec:
                    selected_records.append(rec)
            if not selected_records:
                return
            mtg_id = str(selected_records[0].get("meeting_id") or "")
            all_meeting = tdb.get_threads_for_meeting(mtg_id) if mtg_id else []
            if len(selected_records) >= 2:
                dlg = MergeThreadsDialog(
                    selected_items=selected_records,
                    candidates=all_meeting,
                    parent=window,
                )
            else:
                others = [t for t in all_meeting if t["thread_id"] != selected_records[0]["thread_id"]]
                if not others:
                    notify_info(window, "No Other Threads", "No other tracked threads in this meeting to merge.")
                    return
                dlg = MergeThreadsDialog(
                    target_thread=selected_records[0],
                    candidates=others,
                    parent=window,
                )
            if dlg.exec_() == QDialog.Accepted:
                target_id = dlg.target_thread_id()
                source_ids = dlg.selected_source_ids()
                if target_id and source_ids:
                    tdb.merge_threads(target_id, source_ids)
                    window._on_tracked_thread_changed()
                    notify_success(
                        window,
                        "Threads Merged",
                        f"Merged {len(source_ids)} thread(s) into {target_id}.",
                    )

        menu = QMenu(window.table)
        menu.addAction("📧 Open Tracked Email Thread...", lambda: window._open_tracked_thread_dialog(tid))
        menu.addAction(
            "🏷️ Edit Notes, Status && Tags...",
            lambda: window._open_interactive_notes_for_row(row_data),
        )
        menu.addAction(
            "🔄 Refresh This Thread Now",
            lambda: (refresh_thread(window.tracked_threads_db, tid), window._on_tracked_thread_changed()),
        )
        menu.addAction(
            "✔️ Mark Thread as Read (Sync Outlook)",
            lambda: (mark_thread_read_with_outlook(window.tracked_threads_db, tid, sync_outlook=True), window._on_tracked_thread_changed()),
        )
        menu.addSeparator()
        if len(selected_thread_ids) >= 2:
            menu.addAction(
                f"🔗 Merge Selected Threads ({len(selected_thread_ids)})...",
                _launch_merge_threads,
            )
        else:
            menu.addAction(
                "🔗 Merge With Another Thread...",
                _launch_merge_threads,
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
        edit_notes_callback=lambda r=row_data: window._open_interactive_notes_for_row(r),
    )
