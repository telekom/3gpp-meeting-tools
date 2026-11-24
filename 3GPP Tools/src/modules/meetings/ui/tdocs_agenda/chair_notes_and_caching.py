"""Chairman's Notes probing/download/scanning and Docs/ TDocs batch caching actions."""

import logging
import os
import webbrowser
from pathlib import Path

from PyQt5.QtWidgets import QDialog

from core.network.network_state import NetworkState
from core.utils.folder_opener import open_folder as _open_folder
from modules.meetings.core.chair_notes_downloader import (
    ChairNotesDownloaderThread,
    ChairNotesListingThread,
)
from modules.meetings.core.url_router import URLRouter

from .refresh_state import (
    _ensure_refresh_btn_unlocked,
    _qmessagebox,
    _restore_refresh_btn,
    _set_refresh_btn_busy,
    _update_refresh_btn_progress,
)

logger = logging.getLogger(__name__)


def download_chair_notes(window):
    """Probe /Inbox/Chair_Notes and present the multi-chair picker dialog."""
    wg_name = window.mtg_info.get("wg_name", "").upper()
    main_url = window.mtg_info.get("url_key", "")
    if main_url and not main_url.startswith("http"):
        main_url = f"https://www.3gpp.org/ftp/{main_url.lstrip('/')}"

    is_active = window.mtg_info.get("is_active_sync", False)
    candidate_urls = []
    if NetworkState.get_instance().is_local_active():
        candidate_urls.append(URLRouter._get_local_server_base(wg_name))

    if is_active:
        sync_wg = "SA3LI" if wg_name == "SA3LI" else wg_name
        candidate_urls.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}")

    if main_url:
        candidate_urls.append(main_url)

    if not candidate_urls:
        return _qmessagebox().warning(
            window, "No URL Available", "Cannot download Chairman's Notes: No valid meeting URL found."
        )

    seq = _set_refresh_btn_busy(window, "⏳ Checking Notes...")

    window.chair_notes_listing_thread = ChairNotesListingThread(candidate_urls, window)
    window.chair_notes_listing_thread.progress.connect(
        lambda msg: _update_refresh_btn_progress(window, msg)
    )
    window.chair_notes_listing_thread.listing_finished.connect(
        lambda success, files, dir_url, msg: on_chair_notes_listing_finished(
            window, success, files, dir_url, msg, candidate_urls
        )
    )
    window.chair_notes_listing_thread.finished.connect(
        lambda: _ensure_refresh_btn_unlocked(window, seq)
    )
    window.chair_notes_listing_thread.start()


def on_chair_notes_listing_finished(
    window, success: bool, files: list, dir_url: str, msg: str, candidate_urls: list
):
    """Handle completion of server directory scanning and present the picker dialog."""
    _restore_refresh_btn(window)
    mb = _qmessagebox()

    if not success:
        mb.warning(window, "Notes Not Found", msg)
        return

    if not files:
        mb.information(
            window,
            "No Chairman's Notes Found",
            msg or "Connected to Chair_Notes folder, but no matching Chairman's Notes files were found on the server.",
        )
        return

    chair_notes_dir = window.meeting_dir / "Agenda" / "ChairNotes"

    from modules.meetings.ui.chair_notes_picker_dialog import ChairNotesPickerDialog

    dlg = ChairNotesPickerDialog(
        parent=window,
        file_tuples=files,
        local_dir=chair_notes_dir,
        mtg_info=getattr(window, "mtg_info", None),
        target_dir_url=dir_url,
    )

    if dlg.exec_() == QDialog.Accepted:
        selected_files = dlg.get_selected_files()
        if not selected_files:
            return
        overwrite = dlg.should_overwrite()
        start_chair_notes_download(window, candidate_urls, chair_notes_dir, selected_files, overwrite)


def start_chair_notes_download(
    window, candidate_urls: list, chair_notes_dir: Path, selected_files: list, overwrite: bool
):
    """Launch the download thread for user-selected Chairman's Notes files."""
    seq = _set_refresh_btn_busy(window, "⏳ Downloading Notes...")

    window.chair_notes_thread = ChairNotesDownloaderThread(
        candidate_base_urls=candidate_urls,
        target_dir=chair_notes_dir,
        parent=window,
        selected_files=selected_files,
        overwrite=overwrite,
    )
    window.chair_notes_thread.progress.connect(
        lambda msg: _update_refresh_btn_progress(window, msg)
    )
    window.chair_notes_thread.download_finished.connect(
        lambda success, count, files, msg: on_chair_notes_finished(
            window, success, count, files, msg
        )
    )
    window.chair_notes_thread.finished.connect(
        lambda: _ensure_refresh_btn_unlocked(window, seq)
    )
    window.chair_notes_thread.start()


def on_chair_notes_finished(window, success: bool, count: int, files: list, msg: str):
    """Handle completion of Chairman's Notes download and present direct open actions."""
    mb = _qmessagebox()
    if success and count > 0:
        _restore_refresh_btn(window, f"✅ {count} Notes", reset_after_ms=4000)
        chair_notes_dir = window.meeting_dir / "Agenda" / "ChairNotes"

        box = mb(window)
        box.setWindowTitle("Chairman's Notes Downloaded")
        box.setText(f"Successfully downloaded {count} Chairman's Notes file(s) to:\n{chair_notes_dir}")

        if files:
            file_sample = "\n".join(f"• {f}" for f in files[:8])
            if len(files) > 8:
                file_sample += f"\n...and {len(files) - 8} more"
            box.setInformativeText(file_sample)

        btn_word = box.addButton("📝 Open in Word", mb.ActionRole)
        btn_folder = box.addButton("📂 Open Folder", mb.ActionRole)
        box.addButton("Close", mb.RejectRole)
        box.setDefaultButton(btn_word)
        box.exec_()

        clicked = box.clickedButton()
        if clicked == btn_word:
            if files:
                doc_files = [f for f in files if Path(f).suffix.lower() in (".docx", ".doc")]
                target_file = chair_notes_dir / (doc_files[0] if doc_files else files[0])
                _open_file_in_word(target_file, window)
        elif clicked == btn_folder:
            _open_folder(chair_notes_dir)

    elif success and count == 0:
        _restore_refresh_btn(window)
        mb.information(window, "Chairman's Notes", msg)
    else:
        _restore_refresh_btn(window)
        mb.warning(window, "Download Failed", msg)


def _open_file_in_word(filepath: Path, parent=None):
    """Open a document in Microsoft Word using the system default handler."""
    try:
        if hasattr(os, "startfile"):
            os.startfile(str(filepath))
        else:
            webbrowser.open(f"file:///{filepath}")
    except Exception as e:
        if parent:
            _qmessagebox().warning(parent, "Open Failed", f"Could not open document:\n{e}")


def open_chair_notes_folder(window):
    """Open the local folder containing Chairman's Notes."""
    _open_folder(window.meeting_dir / "Agenda" / "ChairNotes")


def scan_existing_chair_notes(window):
    """Scan local Chairman's Notes files on disk, extract agreements, and merge into the table."""
    mb = _qmessagebox()
    if not getattr(window, "is_ran", False):
        mb.information(
            window,
            "RAN Feature Only",
            "Scanning Chairman's Notes for inline text agreements is designed for RAN working groups (RAN1, RAN2, RAN3, etc.).\n\n"
            f"Meeting '{getattr(window, 'mtg_info', {}).get('wg_name', '')}' follows the SA/CT operational model where agreements "
            "are approved directly as TDoc status in the official meeting table or TdocsByAgenda.htm.",
        )
        return

    from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser

    _set_refresh_btn_busy(window, "Scanning Notes...")

    try:
        agenda_data, agreements, files_scanned = ChairNotesAgreementsParser.scan_and_cache_chair_notes(
            window.meeting_dir,
            mtg_info=getattr(window, "mtg_info", None),
            agenda_map=getattr(window, "agenda_map", None),
            force_rescan=True,
        )

        non_recap_agreements = [a for a in (agreements or []) if not a.get("is_previous_meeting_recap")]
        ag_count = len(non_recap_agreements)
        tdoc_count = len(agenda_data)

        if agreements:
            window.chair_agreements = agreements
            window.model.set_chair_agreements(agreements)

        if agenda_data:
            window.model.merge_agenda_data(agenda_data)
            window._refresh_comboboxes()

        if hasattr(window, "_update_agreements_counter"):
            window._update_agreements_counter()

        if files_scanned > 0:
            _restore_refresh_btn(window, f"✅ {ag_count} Agreements", reset_after_ms=4000)

            wa_cnt = sum(1 for a in non_recap_agreements if "working assumption" in str(a.get("kind", "")).lower())
            ffs_cnt = sum(
                1 for a in non_recap_agreements
                if "ffs" in str(a.get("kind", "")).lower() or "open issue" in str(a.get("kind", "")).lower()
            )
            pure_ag = ag_count - wa_cnt - ffs_cnt

            ais_covered = len({str(a.get("ai", "")).strip() for a in non_recap_agreements if a.get("ai")})
            wis_covered = len(
                {str(a.get("wi", "")).strip() for a in non_recap_agreements if a.get("wi") and a.get("wi") != "General"}
            )

            msg = (
                f"Successfully scanned {files_scanned} Chairman's Notes file(s) in:\n"
                f"{window.meeting_dir / 'Agenda' / 'ChairNotes'}\n\n"
                f"• Agreements & Decisions Extracted: {ag_count}\n"
                f"  - 🟢 Technical Agreements: {pure_ag}\n"
                f"  - 🟡 Working Assumptions: {wa_cnt}\n"
                f"  - 🔵 FFS / Open Issues: {ffs_cnt}\n"
                f"• Agenda Items Covered: {ais_covered}\n"
                f"• Work Items Covered: {wis_covered}\n"
                f"• Meeting Table Rows Updated: {tdoc_count}\n\n"
                f"All agreements are cached and ready for post-meeting and Markdown reports."
            )
            mb.information(window, "Chairman's Notes Scanned", msg)
        else:
            _restore_refresh_btn(window)
            chair_notes_dir = window.meeting_dir / "Agenda" / "ChairNotes"
            msg = (
                f"No Chairman's Notes files (.docx, .doc, .zip) were found in:\n{chair_notes_dir}\n\n"
                "Would you like to check the 3GPP server and download Chairman's Notes now?"
            )

            reply = mb.question(
                window,
                "No Local Chairman's Notes Found",
                msg,
                mb.Yes | mb.No,
                mb.Yes,
            )
            if reply == mb.Yes:
                download_chair_notes(window)
    except Exception as e:
        _restore_refresh_btn(window)
        logger.error(f"[scan_existing_chair_notes] Error: {e}", exc_info=True)
        mb.warning(window, "Scan Failed", f"Failed to scan Chairman's Notes:\n{e}")


def start_window_tdocs_caching(window):
    """Initiate asynchronous parallel caching of all meeting TDocs from Docs/."""
    from modules.meetings.core.tdocs.tdocs_cacher import TDocsCacherThread

    mb = _qmessagebox()
    if getattr(window, "cacher_thread", None) and window.cacher_thread.isRunning():
        mb.information(
            window,
            "Caching in Progress",
            "TDocs caching is already running in the background. Please wait for it to complete.",
        )
        return

    docs_url = getattr(window, "docs_ftp_url", "")
    if not docs_url and getattr(window, "main_ftp_url", ""):
        docs_url = f"{window.main_ftp_url.rstrip('/')}/Docs/"

    if not docs_url:
        mb.warning(
            window,
            "No Docs Folder URL",
            "Cannot cache TDocs: No valid Docs folder URL found for this meeting.",
        )
        return

    _set_refresh_btn_busy(window, "Caching TDocs...")

    parent_obj = window if hasattr(window, "metaObject") else None
    window.cacher_thread = TDocsCacherThread(docs_url, window.meeting_dir, parent=parent_obj)
    window.cacher_thread.progress.connect(
        lambda msg: _update_refresh_btn_progress(window, msg)
    )
    window.cacher_thread.finished.connect(
        lambda success, msg: on_window_tdocs_caching_finished(window, success, msg)
    )
    window.cacher_thread.start()


def on_window_tdocs_caching_finished(window, success: bool, msg: str):
    """Reset UI and display the result of meeting TDocs caching."""
    _restore_refresh_btn(window)
    mb = _qmessagebox()
    if success:
        mb.information(window, "Caching Complete", msg)
    else:
        mb.warning(window, "Caching Failed", msg)
