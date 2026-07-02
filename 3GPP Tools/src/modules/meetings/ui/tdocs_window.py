# --- File: modules/meetings/ui/tdocs_window.py ---
import datetime
import json
import os
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QPoint
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableView,
                             QHeaderView, QLabel, QLineEdit, QFrame,
                             QPushButton, QMessageBox, QMenu, QApplication, QToolTip, QCheckBox, QTextEdit, QDialog)

from modules.meetings.core.compare_manager import ComparisonManager
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.core.tdocs_threads import TDocsRevisionsFetcherThread, TDocActionThread, TdocsByAgendaThread
from modules.meetings.ui.tdoc_delegates import HtmlDelegate, TDocActionDelegate
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel, natural_sort_key
from modules.emails.ui.email_window import EmailManagerWindow


# ==========================================
# --- TDOCS WINDOW ---
# ==========================================
class TDocActionMenu(QMenu):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def mouseReleaseEvent(self, event):
        # Intercept right-clicks on the menu
        if event.button() == Qt.RightButton:
            action = self.actionAt(event.pos())
            if action and action.data():
                # Grab the hidden URL, copy to clipboard, and show a temporary tooltip
                url = action.data()
                QApplication.clipboard().setText(url)
                QToolTip.showText(event.globalPos(), "📋 URL Copied to Clipboard!", self)
                self.close()
                return
        # Pass normal left-clicks through
        super().mouseReleaseEvent(event)

class TDocsWindow(QWidget):
    # ---> NEW: Bridge to ask the main application to handle global searches
    global_action_requested = pyqtSignal(str, str)

    def __init__(self, mtg_info: dict, tdocs_data: list, filepath: str):
        super().__init__()
        self.mtg_info = mtg_info
        self.filepath = filepath
        self.meeting_dir = Path(filepath).parent.parent
        self.active_threads = {}

        # 1. Safely initialize and sanitize all URLs
        wg_name = str(self.mtg_info.get('wg_name', '')).upper()
        self.is_sa2_electronic = ('SA2' in wg_name) and bool(self.mtg_info.get('is_electronic', 0))

        main_ftp = self.mtg_info.get("url_key", "")
        if main_ftp and not main_ftp.startswith("http"):
            main_ftp = "https://www.3gpp.org/ftp/" + main_ftp.lstrip('/')
        self.main_ftp_url = main_ftp

        docs_ftp = self.mtg_info.get("docs_folder_url", "")
        if docs_ftp and not docs_ftp.startswith("http"):
            docs_ftp = "https://www.3gpp.org/ftp/" + docs_ftp.lstrip('/')
        self.docs_ftp_url = docs_ftp

        if self.is_sa2_electronic and self.main_ftp_url:
            self.revisions_url = self.main_ftp_url.rstrip('/') + '/INBOX/Revisions/'

        # 2. Build UI
        title = f"TDocs: {mtg_info.get('wg_name', '')} {mtg_info.get('meeting_number', '')}"
        self.setWindowTitle(title)
        self.resize(1400, 750)
        self.setStyleSheet("QWidget { background-color: #FAFAFA; }")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # --- HEADER & COUNT ---
        header_layout = QHBoxLayout()
        title_lbl = QLabel(f"<b>{title}</b>")
        title_lbl.setStyleSheet("font-size: 18px; color: #333;")

        self.last_mod_lbl = QLabel(self._get_mod_date_str())
        self.last_mod_lbl.setStyleSheet("font-size: 11px; color: #999999; margin-right: 15px; font-style: italic;")

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setCursor(Qt.PointingHandCursor)
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold;
                border-radius: 6px; padding: 5px 12px;
                color: #555555; background-color: #F0F0F0; border: 1px solid #CCCCCC;
            }
            QPushButton:hover, QPushButton::menu-indicator { background-color: #E0E0E0; border: 1px solid #AAAAAA; }
        """)

        refresh_menu = QMenu(self)
        refresh_menu.setStyleSheet("QMenu { font-size: 12px; }")
        refresh_menu.addAction("📗 Refresh Excel List", self._refresh_excel)
        wg_name = str(self.mtg_info.get('wg_name', '')).upper()
        if wg_name == "SA2":
            refresh_menu.addAction("📄 Import TdocsByAgenda.htm", self._fetch_tdocs_by_agenda)

        if self.is_sa2_electronic:
            refresh_menu.addAction("📝 Refresh Revisions", lambda: self._refresh_revisions(silent=False))
            refresh_menu.addAction("🔄 Refresh Excel && Revisions", self._refresh_both)
        self.refresh_btn.setMenu(refresh_menu)

        # ---> FIXED: New Multi-Action Folders Menu!
        self.folder_btn = QPushButton("🗂️ Resources")
        self.folder_btn.setCursor(Qt.PointingHandCursor)
        self.folder_btn.setStyleSheet("""
                    QPushButton {
                        font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold;
                        border-radius: 6px; padding: 5px 12px;
                        color: #005A9E; background-color: #E1F0FF; border: 1px solid #99C9FF;
                    }
                    QPushButton:hover, QPushButton::menu-indicator { background-color: #CCE4FF; border: 1px solid #005A9E; }
                """)

        folder_menu = QMenu(self)
        folder_menu.setStyleSheet("QMenu { font-size: 12px; }")
        folder_menu.addAction("📁 Local: Meeting Folder", self._open_meeting_folder)
        wg_name = str(self.mtg_info.get('wg_name', '')).upper()
        if wg_name == "SA2":
            folder_menu.addAction("📄 Local: TdocsByAgenda.htm", self._open_agenda_file)
        folder_menu.addSeparator()
        if hasattr(self, 'main_ftp_url') and self.main_ftp_url:
            folder_menu.addAction("🌐 FTP: Main Folder", lambda: webbrowser.open(self.main_ftp_url))
        if hasattr(self, 'docs_ftp_url') and self.docs_ftp_url:
            folder_menu.addAction("🌐 FTP: Docs Folder", lambda: webbrowser.open(self.docs_ftp_url))
        if hasattr(self, 'revisions_url') and self.revisions_url:
            folder_menu.addAction("🌐 FTP: Revisions Folder", lambda: webbrowser.open(self.revisions_url))
        self.folder_btn.setMenu(folder_menu)

        self.excel_btn = QPushButton("📗 Open in Excel")
        self.excel_btn.setCursor(Qt.PointingHandCursor)
        self.excel_btn.setStyleSheet("""
            QPushButton {
                font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold;
                border-radius: 6px; padding: 5px 12px;
                color: #0C6B0C; background-color: #E6F4E6; border: 1px solid #A3DDA3;
            }
            QPushButton:hover { background-color: #D1EED1; border: 1px solid #0C6B0C; }
        """)
        self.excel_btn.clicked.connect(self._open_excel)

        # ---> NEW: eMeeting Email Manager Button
        self.email_btn = QPushButton("📧 Emails")
        self.email_btn.setCursor(Qt.PointingHandCursor)
        self.email_btn.setStyleSheet("""
                 QPushButton {
                     font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold;
                     border-radius: 6px; padding: 5px 12px;
                     color: #B85C00; background-color: #FFF4CE; border: 1px solid #F3C74C;
                 }
                 QPushButton:hover { background-color: #FFE0B2; }
             """)
        self.email_btn.clicked.connect(self._open_email_manager)
        self.email_btn.setVisible(self.is_sa2_electronic)  # Only show for SA2 eMeetings!

        self.count_lbl = QLabel(f"Showing {len(tdocs_data)} of {len(tdocs_data)} TDocs")
        self.count_lbl.setStyleSheet("font-size: 13px; color: #666;")

        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.last_mod_lbl)
        header_layout.addWidget(self.refresh_btn)
        header_layout.addWidget(self.folder_btn)
        header_layout.addWidget(self.excel_btn)
        header_layout.addWidget(self.email_btn)  # <--- Add it here
        header_layout.addSpacing(15)
        header_layout.addWidget(self.count_lbl)
        main_layout.addLayout(header_layout)

        # --- MODERN FILTER BAR ---
        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; }
            QLabel { font-weight: bold; color: #555; border: none; }
            QLineEdit, QComboBox { padding: 6px; border: 1px solid #CCC; border-radius: 4px; background: #FFF; }
            QLineEdit:focus { border: 1px solid #0078D7; }
        """)
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(15, 10, 15, 10)

        filter_layout.addWidget(QLabel("🔍 Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search TDoc number, title, source, or abstract...")
        self.search_input.setMinimumWidth(250)
        self.search_input.textChanged.connect(self._on_search_changed)
        filter_layout.addWidget(self.search_input)

        def sanitize(val):
            return str(val).strip() if val is not None else ""

        self.type_combo = CheckableComboBox("Type")
        self.type_combo.setMinimumWidth(150)
        unique_types = sorted(list(set(sanitize(r.get("Type", "")) for r in tdocs_data)))
        self.type_combo.addItems(unique_types)
        self.type_combo.selectionChanged.connect(self._on_type_changed)
        filter_layout.addWidget(self.type_combo)

        self.ai_combo = CheckableComboBox("AI")
        self.ai_combo.setMinimumWidth(150)
        unique_ais = sorted(list(set(sanitize(r.get("Agenda Item", "")) for r in tdocs_data)), key=natural_sort_key)
        self.ai_combo.addItems(unique_ais)
        self.ai_combo.selectionChanged.connect(self._on_ai_changed)
        filter_layout.addWidget(self.ai_combo)

        self.status_combo = CheckableComboBox("Status")
        self.status_combo.setMinimumWidth(150)
        unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in tdocs_data)))
        self.status_combo.addItems(unique_statuses)
        self.status_combo.selectionChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

        if self.is_sa2_electronic:
            self.chk_no_comments = QCheckBox("No Comments Only")
            self.chk_no_comments.setStyleSheet("margin-left: 10px;")
            self.chk_no_comments.toggled.connect(self._on_no_comments_toggled)
            filter_layout.addWidget(self.chk_no_comments)

        main_layout.addWidget(filter_frame)

        # --- TABLE SETUP ---
        self.table = QTableView()
        self.model = TDocsTableModel(self.meeting_dir, tdocs_data)

        self.proxy = TDocsFilterProxyModel()
        self.proxy.setSourceModel(self.model)
        self.proxy.layoutChanged.connect(self._update_count_label)

        self.proxy.setTypeFilters(unique_types)
        self.proxy.setAIFilters(unique_ais)
        self.proxy.setStatusFilters(unique_statuses)

        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectItems)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.setFocusPolicy(Qt.WheelFocus)

        # ---> NEW: Trigger our popup when a cell is double-clicked
        self.table.doubleClicked.connect(self._show_cell_popup)

        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setStyleSheet("""
                    QTableView { gridline-color: #E0E0E0; border: 1px solid #E0E0E0; background-color: #FFFFFF; }
                    QHeaderView::section { background-color: #F5F5F5; padding: 4px; font-weight: bold; border: 1px solid #E0E0E0; }
                """)

        # ---> FIX: Add Ctrl+C Support to the Table
        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self._copy_table_selection)

        self.table.setWordWrap(True)
        # ---> OPTIMIZATION 2: ResizeToContents forces PyQt to calculate the height of 3,000+ rows,
        # which severely freezes the UI. We lock it to a comfortable fixed height and rely on ToolTips!
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(48)

        self.action_delegate = TDocActionDelegate(self.table)
        self.action_delegate.actionClicked.connect(self._handle_tdoc_action)
        self.table.setItemDelegateForColumn(0, self.action_delegate)

        self.html_delegate = HtmlDelegate(self.table)
        self.html_delegate.linkClicked.connect(self._scroll_to_tdoc)
        self.html_delegate.linkRightClicked.connect(self._show_related_menu)
        self.table.setItemDelegateForColumn(10, self.html_delegate)

        self.table.setItemDelegateForColumn(10, self.html_delegate)
        self.table.viewport().setMouseTracking(True)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 110)
        header.resizeSection(1, 100)
        header.resizeSection(2, 200)
        header.resizeSection(3, 100)
        header.setSectionResizeMode(6, QHeaderView.Stretch)

        # ---> UPDATE: Lock the Abstract column to exactly 28 pixels!
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.resizeSection(6, 28)

        # Stretch Secretary Remarks to fill the remaining space
        header.setSectionResizeMode(7, QHeaderView.Stretch)

        header.resizeSection(10, 160)

        main_layout.addWidget(self.table)

        # --- LOCAL CACHE AUTO-LOAD ---
        agenda_dir = self.meeting_dir / "Agenda"

        if self.is_sa2_electronic:
            # 1. Auto-load TdocsByAgenda.htm if it exists locally
            local_agenda = agenda_dir / "TdocsByAgenda.htm"
            if local_agenda.exists():
                agenda_data = TDocsParser.parse_tdocs_by_agenda(str(local_agenda))
                if agenda_data:
                    self.model.merge_agenda_data(agenda_data)

            # 2. Auto-load Revisions if they exist, otherwise fetch them
            local_revs = agenda_dir / "revisions.json"
            if local_revs.exists():
                try:
                    with open(local_revs, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        self.model.revisions = data
                        self.model.dataChanged.emit(self.model.index(0, 0),
                                                    self.model.index(self.model.rowCount() - 1, 0))
                except Exception as e:
                    print(f"Failed to load cached revisions: {e}")
                    if hasattr(self, 'revisions_url') and self.revisions_url:
                        self._refresh_revisions(silent=True)
            else:
                # If there's no cache, fall back to fetching them dynamically
                if hasattr(self, 'revisions_url') and self.revisions_url:
                    self._refresh_revisions(silent=True)

    # --- ACTIONS & TRIGGERS ---
    def _get_mod_date_str(self):
        try:
            mod_time = os.path.getmtime(self.filepath)
            return f"List last updated: {datetime.datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M')}"
        except Exception:
            return "List last updated: Unknown"

    def _refresh_both(self):
        self._refresh_excel()
        self._refresh_revisions(silent=True)

    def _refresh_revisions(self, silent=False):
        if not hasattr(self, 'revisions_url'): return

        # ---> UPDATE: Pass self.meeting_dir so the thread knows where to save the json!
        self.rev_thread = TDocsRevisionsFetcherThread(self.revisions_url, self.meeting_dir)

        self.rev_thread.finished.connect(lambda s, d, m: self._on_revisions_fetched(s, d, m, silent))
        self.rev_thread.start()

    def _on_revisions_fetched(self, success: bool, data: dict, msg: str, silent: bool):
        if success:
            self.model.revisions = data
            topLeft = self.model.index(0, 0)
            bottomRight = self.model.index(self.model.rowCount() - 1, 0)
            self.model.dataChanged.emit(topLeft, bottomRight)

            if not silent:
                # ---> OPTIMIZATION: Non-blocking success notification!
                self.refresh_btn.setText(f"✅ {len(data)} Revs")
                QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))
        else:
            if not silent:
                QMessageBox.warning(self, "Revisions Error", f"Failed to sync revisions:\n{msg}")
                self.refresh_btn.setText("🔄 Refresh")

    def _refresh_excel(self):
        mtg_id = self.mtg_info.get("mtg_id")
        if not mtg_id:
            QMessageBox.warning(self, "Missing ID", "Cannot refresh: Missing 3GPP Portal ID for this meeting.")
            return

        self.refresh_btn.setText("⏳ Downloading...")
        self.refresh_btn.setEnabled(False)

        self.dl_thread = TDocsDownloaderThread(mtg_id, self.meeting_dir, self)
        self.dl_thread.finished.connect(self._on_refresh_excel_finished)
        self.dl_thread.start()

    def _on_refresh_excel_finished(self, success: bool, result: str, mtg_id: str):
        self.refresh_btn.setText("🔄 Refresh")
        self.refresh_btn.setEnabled(True)

        if success:
            self.filepath = result
            new_data = TDocsParser.parse_tdocs_excel(self.filepath)
            if not new_data:
                QMessageBox.warning(self, "Parse Error", "Successfully downloaded, but could not parse the Excel file.")
                return

            self.model.update_data(new_data)

            def sanitize(val):
                return str(val).strip() if val is not None else ""

            unique_types = sorted(list(set(sanitize(r.get("Type", "")) for r in new_data)))
            unique_ais = sorted(list(set(sanitize(r.get("Agenda Item", "")) for r in new_data)), key=natural_sort_key)
            unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in new_data)))

            self.type_combo.updateItems(unique_types)
            self.ai_combo.updateItems(unique_ais)
            self.status_combo.updateItems(unique_statuses)

            self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
            self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
            self.proxy.setStatusFilters(self.status_combo.getCheckedItems())

            self.last_mod_lbl.setText(self._get_mod_date_str())
            self._update_count_label()
        else:
            QMessageBox.critical(self, "Download Error", f"Failed to refresh TDocs List:\n{result}")

    def _handle_tdoc_action(self, base_tdoc: str):
        if base_tdoc in self.model.loading_tdocs: return
        if not hasattr(self, 'docs_ftp_url') or not self.docs_ftp_url: return

        docs_url = self.docs_ftp_url if self.docs_ftp_url.startswith(
            "http") else "https://www.3gpp.org/ftp/" + self.docs_ftp_url.lstrip('/')
        revisions = self.model.revisions.get(base_tdoc, [])

        # ---> FIX: Use our new custom right-click capable menu!
        menu = TDocActionMenu(self.table)
        menu.setStyleSheet("QMenu { font-size: 13px; }")
        menu.setToolTipsVisible(True)

        # --- 1. OPEN ACTIONS ---
        base_zip = self.meeting_dir / base_tdoc / f"{base_tdoc}.zip"
        act_base = menu.addAction(f"🗎 Open Base: {base_tdoc}" + ("  (Local)" if base_zip.exists() else ""))

        # ---> Embed the URL so the right-click event can extract it
        act_base.setData(docs_url.rstrip('/') + f"/{base_tdoc}.zip")
        act_base.setToolTip("Left-click to open. Right-click to copy FTP link.")
        act_base.triggered.connect(lambda _, t=base_tdoc: self._trigger_download_thread(base_tdoc, t, docs_url, False))

        if revisions:
            menu.addSeparator()
            for rev in revisions:
                target_filename = f"{base_tdoc}{rev}"
                rev_zip = self.meeting_dir / base_tdoc / f"{target_filename}.zip"
                act_rev = menu.addAction(
                    f"📝 Open Revision: {target_filename}" + ("  (Local)" if rev_zip.exists() else ""))

                act_rev.setData(self.revisions_url.rstrip('/') + f"/{target_filename}.zip")
                act_rev.setToolTip("Left-click to open. Right-click to copy FTP link.")
                act_rev.triggered.connect(
                    lambda _, t=target_filename: self._trigger_download_thread(base_tdoc, t, self.revisions_url, False))

        # --- 2. LOCAL FOLDER ---
        menu.addSeparator()
        act_folder = menu.addAction("📂 Open Local Folder")
        act_folder.triggered.connect(lambda _, d=(self.meeting_dir / base_tdoc): self._open_specific_folder(d))

        # --- 3. COMPARISON CART SUBMENU ---
        menu.addSeparator()
        compare_menu = TDocActionMenu("⚖️ Add to Comparison Cart...", self.table)
        compare_menu.setToolTipsVisible(True)
        menu.addMenu(compare_menu)

        act_cmp_base = compare_menu.addAction(
            f"🗎 Base Version: {base_tdoc}" + ("  (Local)" if base_zip.exists() else ""))
        act_cmp_base.setData(docs_url.rstrip('/') + f"/{base_tdoc}.zip")
        act_cmp_base.setToolTip("Right-click to copy FTP link.")
        act_cmp_base.triggered.connect(
            lambda _, t=base_tdoc: self._trigger_download_thread(base_tdoc, t, docs_url, True))

        for rev in revisions:
            target_filename = f"{base_tdoc}{rev}"
            rev_zip = self.meeting_dir / base_tdoc / f"{target_filename}.zip"
            act_cmp_rev = compare_menu.addAction(
                f"📝 Revision: {target_filename}" + ("  (Local)" if rev_zip.exists() else ""))
            act_cmp_rev.setData(self.revisions_url.rstrip('/') + f"/{target_filename}.zip")
            act_cmp_rev.setToolTip("Right-click to copy FTP link.")
            act_cmp_rev.triggered.connect(
                lambda _, t=target_filename: self._trigger_download_thread(base_tdoc, t, self.revisions_url, True))

        menu.exec_(QCursor.pos())

    def _trigger_download_thread(self, base_tdoc: str, target_filename: str, base_url: str,
                                 is_silent_compare: bool = False):
        self.model.set_loading(base_tdoc, True)

        # ---> FIX 2: Pass open_file=not is_silent_compare to suppress the auto-open behavior
        thread = TDocActionThread(base_tdoc, target_filename, base_url, self.meeting_dir,
                                  open_file=not is_silent_compare)

        thread.is_silent_compare = is_silent_compare
        thread.target_filename = target_filename

        thread.finished_action.connect(lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th))
        self.active_threads[base_tdoc] = thread
        thread.start()

    def _on_tdoc_action_finished(self, tdoc: str, success: bool, msg: str, thread: TDocActionThread):
        if tdoc in self.active_threads:
            del self.active_threads[tdoc]
        self.model.set_loading(tdoc, False)

        if not success:
            QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)
            return

        if getattr(thread, "is_silent_compare", False):
            # ---> THE FIX: Grab the EXACT file path directly from the extraction thread!
            # No more guessing with glob("*.doc*")
            extracted_files = getattr(thread, "extracted_doc_paths", [])

            if extracted_files:
                # Add the exact file to the global cart
                ComparisonManager.get_instance().add_to_cart(thread.target_filename, str(extracted_files[0]))
            else:
                QMessageBox.warning(self, "Compare Failed", "No Word document found inside this TDoc ZIP.")

    def _scroll_to_tdoc(self, target_tdoc: str):
        """Handles Left-Clicks on Related TDocs."""
        import re

        # 1. Strip the revision suffix (e.g., S2-2606287r11 -> S2-2606287)
        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', target_tdoc, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else target_tdoc.upper()

        if base_tdoc in self.model.valid_tdocs:
            for row in range(self.proxy.rowCount()):
                idx = self.proxy.index(row, 1)
                if self.proxy.data(idx, Qt.UserRole) == base_tdoc:
                    self.table.scrollTo(idx, QTableView.PositionAtCenter)
                    self.table.selectRow(row)  # Highlight it for the user
                    return
            QMessageBox.information(self, "Hidden", f"TDoc '{base_tdoc}' is currently hidden by your active filters.")
        else:
            # ---> NEW: It's from a different meeting! Ask to jump globally.
            reply = QMessageBox.question(
                self, "External TDoc",
                f"{base_tdoc} is not from this meeting.\n\nWould you like to search the global database and open its parent meeting?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.global_action_requested.emit(base_tdoc, 'open_meeting')

    def _show_related_menu(self, target_tdoc: str, pos: QPoint):
        """Handles Right-Clicks on Related TDocs."""
        import re

        menu = QMenu(self)
        menu.setStyleSheet("QMenu { font-size: 13px; }")

        # Strip the revision to check if the BASE TDoc is in this meeting
        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', target_tdoc, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else target_tdoc.upper()

        is_local = base_tdoc in self.model.valid_tdocs

        # Smart FTP Routing: If it has a revision suffix, pull from Revisions URL; otherwise pull from Docs URL
        dl_url = self.revisions_url if (match and hasattr(self, 'revisions_url')) else self.docs_ftp_url

        if is_local:
            menu.addAction("⬇️ Go to Row").triggered.connect(lambda: self._scroll_to_tdoc(target_tdoc))

            # ---> Notice how we pass the FULL target_tdoc (e.g. S2-2606287R11) into the download thread!
            menu.addAction(f"📄 Open Document: {target_tdoc}").triggered.connect(
                lambda: self._trigger_download_thread(base_tdoc, target_tdoc, dl_url, False)
            )
            menu.addAction(f"⚖️ Add to Comparison Cart: {target_tdoc}").triggered.connect(
                lambda: self._trigger_download_thread(base_tdoc, target_tdoc, dl_url, True)
            )
        else:
            menu.addAction("🌐 Search && Open Meeting").triggered.connect(
                lambda: self.global_action_requested.emit(base_tdoc, 'open_meeting')
            )
            menu.addAction(f"📄 Search && Open Document: {target_tdoc}").triggered.connect(
                lambda: self.global_action_requested.emit(target_tdoc, 'open_doc')
            )
            menu.addAction(f"⚖️ Search && Add to Comparison Cart: {target_tdoc}").triggered.connect(
                lambda: self.global_action_requested.emit(target_tdoc, 'add_to_cart')
            )

        menu.exec_(pos)

    def _open_meeting_folder(self):
        if self.meeting_dir.exists():
            if hasattr(os, 'startfile'):
                os.startfile(str(self.meeting_dir))
            else:
                webbrowser.open(f"file:///{self.meeting_dir}")
        else:
            QMessageBox.warning(self, "Not Found", "The root meeting folder has not been created yet.")

    def _open_agenda_file(self):
        """Safely opens the downloaded TdocsByAgenda.htm file."""
        # self.meeting_dir points directly to the root meeting folder (e.g., C:/Temp/WG2_Arch_175)
        agenda_path = self.meeting_dir / "Agenda" / "TdocsByAgenda.htm"

        if agenda_path.exists():
            if hasattr(os, 'startfile'):
                os.startfile(str(agenda_path))
            else:
                webbrowser.open(f"file:///{agenda_path}")
        else:
            QMessageBox.information(self, "Not Found",
                                    "TdocsByAgenda.htm has not been downloaded yet.\n\nPlease use the 'Refresh' menu to import it first!")

    def _open_specific_folder(self, folder_path: Path):
        """Safely opens a specific TDoc's local folder."""
        if folder_path.exists():
            if hasattr(os, 'startfile'):
                os.startfile(str(folder_path))
            else:
                webbrowser.open(f"file:///{folder_path}")
        else:
            QMessageBox.information(self, "Not Found",
                                    "This TDoc has not been downloaded yet, so its local folder does not exist.")

    def _open_excel(self):
        try:
            if hasattr(os, 'startfile'):
                os.startfile(self.filepath)
            else:
                webbrowser.open(f"file:///{self.filepath}")
        except Exception as e:
            QMessageBox.warning(self, "Open Error", f"Could not open the Excel file:\n{e}")

    def _on_search_changed(self, text):
        self.proxy.setGlobalFilter(text)

    def _on_type_changed(self, types):
        self.proxy.setTypeFilters(types)

    def _on_ai_changed(self, ais):
        self.proxy.setAIFilters(ais)

    def _on_status_changed(self, statuses):
        self.proxy.setStatusFilters(statuses)

    def _on_no_comments_toggled(self, checked):
        self.proxy.setNoCommentsFilter(checked)

    def _update_count_label(self):
        visible = self.proxy.rowCount()
        total = self.model.rowCount()
        self.count_lbl.setText(f"Showing {visible} of {total} TDocs")

    def _fetch_tdocs_by_agenda(self):
        url_key = self.mtg_info.get("url_key", "")
        if not url_key:
            return

        ftp_url = url_key if url_key.startswith("http") else f"https://www.3gpp.org/ftp/{url_key.lstrip('/')}"

        self.refresh_btn.setText("⏳ Parsing HTML...")

        # ---> FIXED: We delete the manual JSON path reconstruction and just pass
        # the guaranteed correct root directory (self.meeting_dir) directly to the thread!
        self.agenda_thread = TdocsByAgendaThread(ftp_url, self.meeting_dir)
        self.agenda_thread.ui_log_msg.connect(self._handle_thread_log)
        self.agenda_thread.finished.connect(self._on_agenda_fetched)
        self.agenda_thread.start()

    def _on_agenda_fetched(self, success: bool, agenda_data: dict):
        if success and agenda_data:
            self.model.merge_agenda_data(agenda_data)

            # ---> THE FIX: Refresh comboboxes to include new on-the-fly injected values!
            def sanitize(val):
                return str(val).strip() if val is not None else ""

            unique_types = sorted(list(set(sanitize(r.get("Type", "")) for r in self.model._data)))
            unique_ais = sorted(list(set(sanitize(r.get("Agenda Item", "")) for r in self.model._data)),
                                key=natural_sort_key)
            unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in self.model._data)))

            # Push the updated lists to the dropdown menus
            self.type_combo.updateItems(unique_types)
            self.ai_combo.updateItems(unique_ais)
            self.status_combo.updateItems(unique_statuses)

            # Ensure the proxy recognizes the newly added data types so they aren't hidden
            self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
            self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
            self.proxy.setStatusFilters(self.status_combo.getCheckedItems())

            self._update_count_label()

            # OPTIMIZATION: Non-blocking success notification!
            self.refresh_btn.setText(f"✅ {len(agenda_data)} Merged")
            QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))
        else:
            self.refresh_btn.setText("🔄 Refresh")
            QMessageBox.warning(self, "Extraction Error",
                                "Failed to download or parse TdocsByAgenda.htm. It might not exist on the FTP server yet.")

    def _handle_thread_log(self, msg: str, level: int):
        # A simple print catcher so the thread's logs don't crash the UI window
        print(f"Agenda Sync: {msg}")

    def _open_email_manager(self):
        # Aggressively strip spaces to prevent dictionary misses
        ai_lookup = {
            str(r.get("TDoc", "")).strip().upper(): str(r.get("Agenda Item", "N/A")).strip()
            for r in self.model._data if r.get("TDoc")
        }

        # ---> NEW: Extract the meeting dates from your current TDocs context
        # Note: If your variables are named differently in TDocsWindow
        # (e.g., self.meeting_start_date), simply update the names below!
        m_start = self.mtg_info.get("start_date", "")
        m_end = self.mtg_info.get("end_date", "")

        # Pass the dates into the new initialization parameters
        self.email_window = EmailManagerWindow(self.meeting_dir, ai_lookup, m_start, m_end)
        self.email_window.show()

    def _copy_table_selection(self):
        indexes = self.table.selectionModel().selectedIndexes()
        if not indexes:
            return

        # Sort indexes by row and then column to maintain the visual grid structure
        indexes = sorted(indexes, key=lambda x: (x.row(), x.column()))

        text_lines = []
        current_row = indexes[0].row()
        current_line = []

        for idx in indexes:
            # If we've moved to a new row, save the current line and start a new one
            if idx.row() != current_row:
                text_lines.append("\t".join(current_line))
                current_line = []
                current_row = idx.row()

            # Extract plain text safely (Using UserRole automatically bypasses the HTML tags in your Related TDocs column!)
            cell_text = str(idx.data(Qt.UserRole) or "").strip()

            # Fallback to standard DisplayRole if UserRole is empty
            if not cell_text:
                cell_text = str(idx.data(Qt.DisplayRole) or "").strip()

            current_line.append(cell_text)

        # Append the final row
        text_lines.append("\t".join(current_line))

        # Push to the Windows Clipboard
        QApplication.clipboard().setText("\n".join(text_lines))

        # Show a quick tooltip confirming the copy at the mouse cursor
        QToolTip.showText(QCursor.pos(), "📋 Copied to clipboard!", self.table)

    def _show_cell_popup(self, index):
        """Shows a popup with the full text of the double-clicked cell."""
        if not index.isValid(): return

        # Extract the exact column name
        col_name = self.model._headers[index.column()]

        # We only want popups for columns that hold long text
        if col_name not in ["Secretary Remarks", "Title", "Source", "Abstract"]:
            return

        # Use the hidden UserRole to grab the pristine, un-truncated text
        full_text = str(index.data(Qt.UserRole) or "").strip()
        if not full_text: return

        # Build the Dialog window
        dialog = QDialog(self)
        dialog.setWindowTitle(f"📄 Viewing: {col_name}")
        dialog.resize(600, 450)
        dialog.setStyleSheet("QDialog { background-color: #FAFAFA; }")

        layout = QVBoxLayout(dialog)

        # A Read-Only Text editor so you can easily highlight/select specific words
        text_edit = QTextEdit()
        text_edit.setPlainText(full_text)
        text_edit.setReadOnly(True)
        text_edit.setStyleSheet("""
            font-size: 13px; font-family: 'Segoe UI', Arial; 
            padding: 10px; background-color: white; border: 1px solid #CCC;
        """)
        layout.addWidget(text_edit)

        # Bottom Action Buttons
        btn_layout = QHBoxLayout()

        copy_btn = QPushButton("📋 Copy All")
        copy_btn.setStyleSheet(
            "padding: 6px 15px; font-weight: bold; background-color: #0078D7; color: white; border-radius: 4px;")
        # Copies the text and gracefully closes the window so you can immediately paste it elsewhere
        copy_btn.clicked.connect(lambda: [QApplication.clipboard().setText(full_text), dialog.accept()])

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet("padding: 6px 15px;")
        close_btn.clicked.connect(dialog.accept)

        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        btn_layout.addWidget(copy_btn)
        layout.addLayout(btn_layout)

        dialog.exec_()
