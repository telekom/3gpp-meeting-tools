# --- File: modules/meetings/ui/tdocs_window.py ---
import os
import webbrowser
import datetime
from pathlib import Path

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableView,
                             QHeaderView, QLabel, QLineEdit, QFrame,
                             QPushButton, QMessageBox, QMenu)
from PyQt5.QtGui import QCursor
from PyQt5.QtCore import Qt, QTimer

from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.core.tdocs_threads import TDocsRevisionsFetcherThread, TDocActionThread
from modules.meetings.ui.tdoc_delegates import HtmlDelegate, TDocActionDelegate
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel


# ==========================================
# --- TDOCS WINDOW ---
# ==========================================
class TDocsWindow(QWidget):
    def __init__(self, mtg_info: dict, tdocs_data: list, filepath: str):
        super().__init__()
        self.mtg_info = mtg_info
        self.filepath = filepath
        self.meeting_dir = Path(filepath).parent.parent
        self.active_threads = {}

        # Identifies SA2 Electronic meetings for Revisions scraping
        wg_name = str(self.mtg_info.get('wg_name', '')).upper()
        self.is_sa2_electronic = ('SA2' in wg_name) and bool(self.mtg_info.get('is_electronic', 0))

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

        # Multi-Action Refresh Menu
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
        refresh_menu.addAction("Refresh Excel List", self._refresh_excel)

        if self.is_sa2_electronic:
            refresh_menu.addAction("Refresh Revisions", lambda: self._refresh_revisions(silent=False))
            refresh_menu.addAction("Refresh Both", self._refresh_both)

        self.refresh_btn.setMenu(refresh_menu)

        self.folder_btn = QPushButton("📂 Meeting Folder")
        self.folder_btn.setCursor(Qt.PointingHandCursor)
        self.folder_btn.setStyleSheet("""
            QPushButton {
                font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold;
                border-radius: 6px; padding: 5px 12px;
                color: #005A9E; background-color: #E1F0FF; border: 1px solid #99C9FF;
            }
            QPushButton:hover { background-color: #CCE4FF; border: 1px solid #005A9E; }
        """)
        self.folder_btn.clicked.connect(self._open_meeting_folder)

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

        self.count_lbl = QLabel(f"Showing {len(tdocs_data)} of {len(tdocs_data)} TDocs")
        self.count_lbl.setStyleSheet("font-size: 13px; color: #666;")

        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.last_mod_lbl)
        header_layout.addWidget(self.refresh_btn)
        header_layout.addWidget(self.folder_btn)
        header_layout.addWidget(self.excel_btn)
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
        unique_ais = sorted(list(set(sanitize(r.get("Agenda Item", "")) for r in tdocs_data)))
        self.ai_combo.addItems(unique_ais)
        self.ai_combo.selectionChanged.connect(self._on_ai_changed)
        filter_layout.addWidget(self.ai_combo)

        self.status_combo = CheckableComboBox("Status")
        self.status_combo.setMinimumWidth(150)
        unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in tdocs_data)))
        self.status_combo.addItems(unique_statuses)
        self.status_combo.selectionChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

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
        self.table.setSelectionMode(QTableView.NoSelection)
        self.table.setFocusPolicy(Qt.NoFocus)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setStyleSheet("""
            QTableView { gridline-color: #E0E0E0; border: 1px solid #E0E0E0; background-color: #FFFFFF; }
            QHeaderView::section { background-color: #F5F5F5; padding: 4px; font-weight: bold; border: 1px solid #E0E0E0; }
        """)

        self.table.setWordWrap(True)
        self.table.verticalHeader().setDefaultSectionSize(20)
        self.table.resizeRowsToContents()

        self.action_delegate = TDocActionDelegate(self.table)
        self.action_delegate.actionClicked.connect(self._handle_tdoc_action)
        self.table.setItemDelegateForColumn(0, self.action_delegate)

        self.html_delegate = HtmlDelegate(self.table)
        self.html_delegate.linkClicked.connect(self._scroll_to_tdoc)
        self.table.setItemDelegateForColumn(10, self.html_delegate)
        self.table.viewport().setMouseTracking(True)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 110)  # Expanded to fit "(+Rev)"
        header.resizeSection(1, 100)
        header.resizeSection(2, 200)
        header.resizeSection(3, 100)
        header.setSectionResizeMode(6, QHeaderView.Stretch)
        header.resizeSection(10, 160)

        main_layout.addWidget(self.table)

        # Fire off an initial silent background fetch for revisions if applicable
        if self.is_sa2_electronic and self.mtg_info.get("url_key"):
            self.revisions_url = self.mtg_info.get("url_key").rstrip('/') + '/INBOX/Revisions/'
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

        self.rev_thread = TDocsRevisionsFetcherThread(self.revisions_url)
        self.rev_thread.finished.connect(lambda s, d, m: self._on_revisions_fetched(s, d, m, silent))
        self.rev_thread.start()

    def _on_revisions_fetched(self, success: bool, data: dict, msg: str, silent: bool):
        if success:
            self.model.revisions = data
            # Force the Action Column to redraw to instantly show (+Rev)
            topLeft = self.model.index(0, 0)
            bottomRight = self.model.index(self.model.rowCount() - 1, 0)
            self.model.dataChanged.emit(topLeft, bottomRight)

            if not silent:
                QMessageBox.information(self, "Revisions Sync",
                                        f"Successfully synced available revisions for {len(data)} TDocs.")
        else:
            if not silent:
                QMessageBox.warning(self, "Revisions Error", f"Failed to sync revisions:\n{msg}")

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
            unique_ais = sorted(list(set(sanitize(r.get("Agenda Item", "")) for r in new_data)))
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

        docs_url = self.mtg_info.get("docs_folder_url")
        if not docs_url:
            QMessageBox.warning(self, "Missing URL", "This meeting does not have a Docs/ URL mapped in the database.")
            return

        revisions = self.model.revisions.get(base_tdoc, [])
        if not revisions:
            # Standard Download for files with NO Revisions
            self._trigger_download_thread(base_tdoc, base_tdoc, docs_url)
        else:
            # Construct a dynamic Context Menu
            menu = QMenu(self.table)
            menu.setStyleSheet("QMenu { font-size: 13px; }")

            base_zip = self.meeting_dir / base_tdoc / f"{base_tdoc}.zip"
            lbl = f"🗎 Base Version: {base_tdoc}" + ("  (Local)" if base_zip.exists() else "")

            act_base = menu.addAction(lbl)
            act_base.triggered.connect(lambda _, t=base_tdoc: self._trigger_download_thread(base_tdoc, t, docs_url))
            menu.addSeparator()

            # Add all known revisions
            for rev in revisions:
                target_filename = f"{base_tdoc}{rev}"
                rev_zip = self.meeting_dir / base_tdoc / f"{target_filename}.zip"
                lbl = f"📝 Revision: {target_filename}" + ("  (Local)" if rev_zip.exists() else "")

                act_rev = menu.addAction(lbl)
                act_rev.triggered.connect(
                    lambda _, t=target_filename: self._trigger_download_thread(base_tdoc, t, self.revisions_url))

            menu.exec_(QCursor.pos())

    def _trigger_download_thread(self, base_tdoc: str, target_filename: str, base_url: str):
        self.model.set_loading(base_tdoc, True)
        QTimer.singleShot(0, self.table.resizeRowsToContents)

        thread = TDocActionThread(base_tdoc, target_filename, base_url, self.meeting_dir)
        thread.finished_action.connect(self._on_tdoc_action_finished)
        self.active_threads[base_tdoc] = thread
        thread.start()

    def _on_tdoc_action_finished(self, tdoc: str, success: bool, msg: str):
        if tdoc in self.active_threads:
            del self.active_threads[tdoc]

        self.model.set_loading(tdoc, False)
        QTimer.singleShot(0, self.table.resizeRowsToContents)

        if not success:
            QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)

    def _scroll_to_tdoc(self, target_tdoc: str):
        for row in range(self.proxy.rowCount()):
            idx = self.proxy.index(row, 1)
            if self.proxy.data(idx, Qt.UserRole) == target_tdoc:
                self.table.scrollTo(idx, QTableView.PositionAtCenter)
                return
        QMessageBox.information(self, "Hidden", f"TDoc '{target_tdoc}' is currently hidden by your filters.")

    def _open_meeting_folder(self):
        if self.meeting_dir.exists():
            if hasattr(os, 'startfile'):
                os.startfile(str(self.meeting_dir))
            else:
                webbrowser.open(f"file:///{self.meeting_dir}")
        else:
            QMessageBox.warning(self, "Not Found", "The root meeting folder has not been created yet.")

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

    def _update_count_label(self):
        visible = self.proxy.rowCount()
        total = self.model.rowCount()
        self.count_lbl.setText(f"Showing {visible} of {total} TDocs")
        if hasattr(self, 'table'):
            QTimer.singleShot(0, self.table.resizeRowsToContents)