# --- File: src/modules/meetings/ui/meetings_menu.py ---
import os
import webbrowser
from pathlib import Path
from typing import Any, Dict

from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QMenu

from modules.meetings.ui.tdocs_window import _open_terminal


def populate_meeting_context_menu(tab: Any, menu: QMenu, row_data: Dict[str, Any], row_idx: int) -> None:
    """
    Populates dynamic context menu actions for a selected meeting row:
    - Multi-meeting sync and deletion
    - Meeting info and portal link
    - Reserve/Contribute TDoc link
    - 3GPP FTP and SYNC folder links
    - Local folder & terminal opening
    - TDocs cache and CSV agenda downloads
    """
    menu.clear()
    selected_rows = tab.table.selectionModel().selectedRows()

    if len(selected_rows) > 1 and any(r.row() == row_idx for r in selected_rows):
        menu.addAction(f"🔄 Sync selected ({len(selected_rows)} meetings)").triggered.connect(
            lambda _, rows=selected_rows: tab._emit_multi_sync(rows)
        )
        menu.addSeparator()
        menu.addAction(f"🗑️ Delete selected ({len(selected_rows)} meetings)").triggered.connect(
            lambda _, rows=selected_rows: tab._emit_multi_delete(rows)
        )
        return

    # Single meeting actions
    menu.addAction("ℹ️ Meeting Info").triggered.connect(lambda _, d=row_data: tab.show_meeting_info(d))

    mtg_id = row_data.get("mtg_id")
    if mtg_id:
        menu.addAction("🖥️ 3GU Meeting Portal").triggered.connect(
            lambda _, m=mtg_id: webbrowser.open(f"https://portal.3gpp.org/Home.aspx#/meeting?MtgId={m}")
        )

        end_date_str = row_data.get("end_date", "")
        end_date = QDate.fromString(end_date_str, "yyyy-MM-dd")
        current_date = QDate.currentDate()

        if not end_date.isValid() or end_date >= current_date:
            menu.addAction("📝 Reserve / Contribute TDoc").triggered.connect(
                lambda _, m=mtg_id: webbrowser.open(
                    f"https://portal.3gpp.org/ngppapp/CreateTdoc.Aspx?mode=create&meetingId={m}"
                )
            )

    menu.addAction("🔄 Sync this Meeting").triggered.connect(
        lambda: tab.update_specific_requested.emit(
            [{"wg": row_data.get("wg_name"), "meeting": row_data.get("meeting_number")}],
            tab.chk_wg.isChecked(),
            tab.chk_docs.isChecked(),
            tab.chk_dyna.isChecked(),
        )
    )
    menu.addSeparator()

    raw_url = row_data.get("url_key", "")
    if raw_url:
        full_ftp_url = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url}"
        menu.addAction("🌐 Open Main Folder (FTP)").triggered.connect(
            lambda _, u=full_ftp_url: webbrowser.open(u)
        )

    docs_url = row_data.get("docs_folder_url")
    if docs_url:
        menu.addAction("📂 Open Documents Folder").triggered.connect(
            lambda _, u=docs_url: webbrowser.open(u)
        )

    folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
    if folder_name:
        current_cache = (
            tab.dl_dir_input.text().strip()
            if hasattr(tab, "dl_dir_input")
            else tab.settings.cache_dir
        )
        local_path = Path(current_cache) / folder_name

        if local_path.exists() and local_path.is_dir():
            menu.addAction("📁 Open Local Cache Folder").triggered.connect(
                lambda _, p=str(local_path): os.startfile(p)
                if hasattr(os, "startfile")
                else webbrowser.open(f"file:///{p}")
            )

            terminal_action = menu.addAction("💻 Open Terminal in Meeting Cache")
            terminal_action.setToolTip(
                "Open Windows Terminal with this meeting's local cache folder as its working directory."
            )
            terminal_action.triggered.connect(
                lambda _, p=local_path: _open_terminal(p, tab)
            )

            menu.addAction("📋 Download Agenda CSV").triggered.connect(
                lambda _, d=row_data: tab._download_agenda_csv(d)
            )

        if mtg_id:
            status = row_data.get("tdoc_btn_status", "na")
            if status == "open":
                menu.addAction("📗 Open TDocs List").triggered.connect(
                    lambda _, d=row_data: tab._open_tdocs_window(d, d.get("tdoc_filepath"))
                )
            elif status == "get":
                menu.addAction("⬇️ Get TDocs List").triggered.connect(
                    lambda _, d=row_data, r=row_idx: tab._download_and_open_tdocs(d, r)
                )

        if docs_url:
            menu.addAction("📥 Cache TDocs (Docs/)").triggered.connect(
                lambda _, u=docs_url, p=local_path: tab._start_tdocs_caching(u, p)
            )

    wg_name = row_data.get("wg_name", "")
    meeting_name = row_data.get("name", "")
    start_date = row_data.get("start_date", "")
    end_date = row_data.get("end_date", "")
    is_elec = row_data.get("is_electronic", 0)

    if tab.db.is_active_sync_meeting(wg_name, start_date, end_date, is_elec):
        menu.addSeparator()
        sync_wg = "SA3LI" if wg_name == "SA3" and "LI" in meeting_name.upper() else wg_name
        sync_base_url = f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}"

        menu.addAction("🔄 Open SYNC folder (FTP)").triggered.connect(
            lambda _, u=sync_base_url: webbrowser.open(u)
        )
        menu.addAction("📂 Open SYNC Documents folder").triggered.connect(
            lambda _, u=f"{sync_base_url}/Docs": webbrowser.open(u)
        )

    menu.addSeparator()
    menu.addAction("🗑️ Delete this Meeting").triggered.connect(
        lambda: tab._confirm_delete_specific(
            [{"wg": row_data.get("wg_name"), "meeting": row_data.get("meeting_number")}]
        )
    )
