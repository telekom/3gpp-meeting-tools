# --- File: src/modules/specifications/core/spec_tdoc_opener.py ---
"""
TDoc Document Opener & Cross-Subsystem Dispatcher.

Handles locating, downloading, unzipping, and opening 3GPP TDoc contributions
(.docx / .doc) referenced in specification change history. Connects directly
to Meetings database and 3GPP servers.
"""

import logging
import os
import re
import webbrowser
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from PyQt5.QtCore import QObject, QThread, pyqtSignal
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QMessageBox, QProgressDialog, QToolTip, QWidget

from core.network.session import NetworkSession
from core.utils.paths import get_app_data_root, get_cache_root, get_project_root, to_extended_path
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdoc_url_resolver import normalize_3gpp_url, resolve_tdoc_download_url
from modules.word_tools.core.document_opener import open_document

logger = logging.getLogger(__name__)


def find_cached_tdoc_file(tdoc_id: str, cache_root: Optional[Path] = None) -> Optional[Path]:
    """
    Scans the local meeting cache for an already extracted or downloaded TDoc file.
    Matches {tdoc_id}.docx, {tdoc_id}.doc, or {tdoc_id}.zip.
    """
    clean_id = tdoc_id.strip()
    root = Path(cache_root or get_cache_root())
    if not root.exists():
        return None

    # Search for matching docx or doc files
    patterns = [
        f"{clean_id}.docx",
        f"{clean_id}.doc",
        f"*{clean_id}*.docx",
        f"{clean_id}.zip",
    ]

    for pat in patterns:
        matches = list(root.rglob(pat))
        for m in matches:
            if not m.name.startswith("~$") and m.is_file():
                return m

    return None


class TDocDownloadWorker(QThread):
    """Background worker for downloading a TDoc ZIP archive and extracting its Word document."""

    progress = pyqtSignal(int, str)
    finished_success = pyqtSignal(Path)
    error = pyqtSignal(str)

    def __init__(self, tdoc_id: str, download_url: str, target_dir: Path):
        super().__init__()
        self.tdoc_id = tdoc_id.strip()
        self.download_url = download_url.strip()
        self.target_dir = Path(target_dir)

    def run(self):
        try:
            self.target_dir.mkdir(parents=True, exist_ok=True)
            zip_dest = self.target_dir / f"{self.tdoc_id}.zip"

            self.progress.emit(15, f"Downloading {self.tdoc_id} from 3GPP server...")

            def _on_progress(pct: int):
                self.progress.emit(15 + int(pct * 0.65), f"Downloading {self.tdoc_id}... ({pct}%)")

            NetworkSession.download_file(
                url=self.download_url,
                dest_path=zip_dest,
                progress_cb=_on_progress,
                cancel_cb=self.isInterruptionRequested,
                atomic=True,
            )

            if self.isInterruptionRequested():
                return

            self.progress.emit(85, f"Extracting {self.tdoc_id} archive...")

            extracted_docs: List[Path] = []
            if zip_dest.exists():
                with zipfile.ZipFile(to_extended_path(zip_dest), "r") as zf:
                    for member in zf.namelist():
                        if "__MACOSX" in member or member.startswith("._"):
                            continue
                        if member.lower().endswith((".docx", ".doc")):
                            dest_file = self.target_dir / Path(member).name
                            with open(to_extended_path(dest_file), "wb") as f:
                                f.write(zf.read(member))
                            extracted_docs.append(dest_file)

            if extracted_docs:
                # Prefer .docx over legacy .doc
                docx_files = [p for p in extracted_docs if p.suffix.lower() == ".docx"]
                target = docx_files[0] if docx_files else extracted_docs[0]
                self.progress.emit(100, f"Ready: {target.name}")
                self.finished_success.emit(target)
            elif zip_dest.exists():
                self.finished_success.emit(zip_dest)
            else:
                self.error.emit(f"No document files found inside {self.tdoc_id}.zip")

        except Exception as e:
            logger.error(f"Error downloading TDoc {self.tdoc_id}: {e}", exc_info=True)
            self.error.emit(str(e))


class SpecTDocOpener:
    """Dispatches TDoc document opening requests across local cache, Meetings database, and FTP."""

    @staticmethod
    def resolve_local_tdoc_file(tdoc_id: str, cache_root: Optional[Path] = None) -> Optional[Path]:
        """Locates any existing cached document (.docx, .doc, or .zip) for the TDoc."""
        return find_cached_tdoc_file(tdoc_id, cache_root=cache_root)

    @staticmethod
    def open_tdoc(
        tdoc_id: str,
        meeting_hint: str = "",
        parent: Optional[QWidget] = None,
        db_path: Optional[Path] = None,
    ) -> bool:
        """
        Locates and opens the actual TDoc Word or ZIP document.
        If cached locally, launches it immediately in Microsoft Word (<50ms).
        If not cached, downloads from 3GPP server in background with a progress dialog.
        """
        clean_tdoc = tdoc_id.strip()
        if not clean_tdoc or clean_tdoc == "-":
            return False

        # 1. Check local cache first
        local_hit = find_cached_tdoc_file(clean_tdoc)
        if local_hit and local_hit.exists():
            if local_hit.suffix.lower() in (".docx", ".doc"):
                logger.info(f"📂 [TDoc Opener] Opening cached contribution: {local_hit.name}")
                open_document(local_hit)
                return True
            elif local_hit.suffix.lower() == ".zip":
                # Check if extracted Word file exists inside or next to it
                parent_dir = local_hit.parent
                candidates = list(parent_dir.glob(f"*{clean_tdoc}*.doc*"))
                if candidates:
                    logger.info(f"📂 [TDoc Opener] Opening extracted contribution: {candidates[0].name}")
                    open_document(candidates[0])
                    return True
                else:
                    logger.info(f"📂 [TDoc Opener] Opening archive: {local_hit.name}")
                    open_document(local_hit)
                    return True

        # 2. Resolve download URL via Meetings database
        logger.info(f"🔍 [TDoc Opener] Resolving download URL for TDoc {clean_tdoc}...")
        data_db_path = Path(db_path or (get_project_root() / "db" / "3gpp_data.db"))
        meetings_db = MeetingsDatabase(data_db_path)

        meeting = meetings_db.find_meeting_by_tdoc(clean_tdoc)
        download_url = ""
        target_dir = get_cache_root() / "tdocs"

        if meeting:
            docs_url = meeting.get("docs_folder_url") or ""
            folder_name = meeting.get("folder_name") or meeting.get("meeting_number", "unknown")
            target_dir = get_cache_root() / folder_name / "TDocs"

            is_rev = bool(re.search(r"[rR]\d{1,2}$", clean_tdoc))
            if is_rev:
                raw_url = meeting.get("url_key", "")
                main_ftp = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}"
                download_url = f"{main_ftp.rstrip('/')}/INBOX/Revisions/{clean_tdoc}.zip"
            elif docs_url:
                norm_docs = normalize_3gpp_url(docs_url)
                download_url = f"{norm_docs.rstrip('/')}/{clean_tdoc}.zip"

        if not download_url:
            # Fallback 1: construct standard 3GPP URL from TDoc prefix
            # e.g. S2-xxxxxxx -> SA2, RP-xxxxxxx -> RAN, CP-xxxxxxx -> CT, SP-xxxxxxx -> SA
            download_url = resolve_tdoc_download_url(clean_tdoc)

        if not download_url:
            # Fallback 2: Open official 3GU portal view
            logger.info(f"Opening 3GU portal for {clean_tdoc}")
            portal_url = f"https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId={clean_tdoc.split('r')[0].split('R')[0]}"
            webbrowser.open(portal_url)
            return True

        # 3. Asynchronous download with modal progress dialog
        progress_dlg = QProgressDialog(f"Fetching {clean_tdoc}...", "Cancel", 0, 100, parent)
        progress_dlg.setWindowTitle(f"Downloading 3GPP Document: {clean_tdoc}")
        progress_dlg.setMinimumDuration(200)

        worker = TDocDownloadWorker(clean_tdoc, download_url, target_dir)

        def _on_progress(pct: int, msg: str):
            if not progress_dlg.wasCanceled():
                progress_dlg.setValue(pct)
                progress_dlg.setLabelText(msg)

        def _on_success(doc_path: Path):
            progress_dlg.close()
            logger.info(f"✅ [TDoc Opener] Successfully downloaded and opening {clean_tdoc}: {doc_path.name}")
            try:
                open_document(doc_path)
            except Exception as e:
                logger.error(f"Error opening downloaded TDoc: {e}", exc_info=True)
                if hasattr(os, "startfile"):
                    os.startfile(str(doc_path))

        def _on_error(err_msg: str):
            progress_dlg.close()
            # If download fails, offer opening 3GU portal
            reply = QMessageBox.question(
                parent,
                f"Download Failed: {clean_tdoc}",
                f"Could not automatically download {clean_tdoc} from:\n{download_url}\n\nError: {err_msg}\n\n"
                f"Would you like to open the official 3GU Portal view instead?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                base_id = clean_tdoc.split("r")[0].split("R")[0]
                webbrowser.open(f"https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId={base_id}")

        worker.progress.connect(_on_progress)
        worker.finished_success.connect(_on_success)
        worker.error.connect(_on_error)
        progress_dlg.canceled.connect(worker.requestInterruption)

        # Keep a reference alive during execution
        worker.finished.connect(worker.deleteLater)
        worker.start()
        progress_dlg.exec_()
        return True

    @staticmethod
    def open_tdoc_in_portal(tdoc_id: str) -> bool:
        """Opens the official 3GU web portal view for a TDoc in default browser."""
        clean_tdoc = tdoc_id.strip()
        base_id = clean_tdoc.split("r")[0].split("R")[0]
        portal_url = f"https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId={base_id}"
        logger.info(f"🌐 [TDoc Opener] Opening 3GU portal: {portal_url}")
        return webbrowser.open(portal_url)

    @staticmethod
    def jump_to_meetings_tab(tdoc_id: str, main_window: Optional[QWidget] = None) -> bool:
        """
        Navigates to the Meetings tab in the main window and populates the Quick TDoc Jump box.
        """
        clean_tdoc = tdoc_id.strip()
        if not main_window:
            # Locate active main window
            for w in QApplication.topLevelWidgets():
                if hasattr(w, "meetings_tab"):
                    main_window = w
                    break

        if not main_window or not hasattr(main_window, "meetings_tab"):
            logger.warning(f"Could not locate Meetings tab to jump to {clean_tdoc}")
            return False

        try:
            logger.info(f"🚀 [Meetings] Jumping to Meetings tab for TDoc {clean_tdoc}...")
            # Switch tab if QTabWidget exists
            if hasattr(main_window, "tab_widget"):
                idx = main_window.tab_widget.indexOf(main_window.meetings_tab)
                if idx >= 0:
                    main_window.tab_widget.setCurrentIndex(idx)

            # Populate jump box
            meetings_tab = main_window.meetings_tab
            if hasattr(meetings_tab, "global_tdoc_input"):
                meetings_tab.global_tdoc_input.setText(clean_tdoc)
                if hasattr(meetings_tab, "search_controller"):
                    meetings_tab.search_controller.on_tdoc_input_changed(clean_tdoc)
            return True
        except Exception as e:
            logger.error(f"Error jumping to meetings tab for {clean_tdoc}: {e}", exc_info=True)
            return False

    @staticmethod
    def format_tdoc_citation(
        tdoc_id: str,
        spec_number: str = "",
        version: str = "",
        meeting: str = "",
        subject: str = "",
    ) -> Tuple[str, str]:
        """Generates (plain_text, rich_html) citation strings for a TDoc and specification version."""
        clean_tdoc = tdoc_id.strip()
        base_id = clean_tdoc.split("r")[0].split("R")[0]
        portal_url = f"https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId={base_id}"

        details = []
        if spec_number:
            details.append(f"TS {spec_number}" if not spec_number.startswith("TR") else spec_number)
        if version:
            details.append(f"v{version}")
        if meeting:
            details.append(meeting)

        detail_str = f" ({', '.join(details)})" if details else ""
        subject_str = f" — {subject}" if subject else ""

        plain = f"{clean_tdoc}{detail_str}{subject_str} [{portal_url}]"
        html = f"""<a href="{portal_url}"><b>{clean_tdoc}</b></a>{detail_str}{f' <i>{subject}</i>' if subject else ''}"""
        return plain, html

    @staticmethod
    def copy_tdoc_citation(
        tdoc_id: str,
        spec_number: str = "",
        version: str = "",
        meeting: str = "",
        subject: str = "",
        parent: Optional[QWidget] = None
    ):
        """Copies formatted plain text and rich HTML citation to clipboard."""
        from modules.meetings.ui.tdocs_menus import _copy_to_clipboard
        plain, html = SpecTDocOpener.format_tdoc_citation(
            tdoc_id=tdoc_id,
            spec_number=spec_number,
            version=version,
            meeting=meeting,
            subject=subject,
        )
        _copy_to_clipboard(plain, html, label=f"TDoc {tdoc_id.strip()}", parent=parent)
        logger.info(f"📋 [Clipboard] Copied citation for TDoc {clean_tdoc}")
