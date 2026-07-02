# --- File: src/modules/meetings/ui/tdocs_window.py ---
import datetime
import json
import os
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QPoint
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableView,
                             QHeaderView, QLabel, QLineEdit, QFrame,
                             QPushButton, QMessageBox, QMenu, QApplication,
                             QToolTip, QCheckBox)

from modules.meetings.core.compare_manager import ComparisonManager
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.core.tdocs_threads import TDocsRevisionsFetcherThread, TDocActionThread, TdocsByAgendaThread
from modules.meetings.core.tdocs_db import TDocsDatabase
from modules.meetings.core.markdown_exporter import MarkdownExporterThread
from modules.meetings.core.statistics_exporter import StatisticsExporterThread

from modules.meetings.ui.tdoc_delegates import HtmlDelegate, TDocActionDelegate
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.meetings.ui.tdocs_models import TDocsTableModel, TDocsFilterProxyModel, natural_sort_key
from modules.meetings.ui.tdocs_menus import build_action_menu, build_related_menu
from modules.meetings.ui.tdocs_dialogs import ReadOnlyViewerDialog, InteractiveNotesDialog, StatisticsSettingsDialog
from modules.emails.ui.email_window import EmailManagerWindow


class TDocsWindow(QWidget):
    global_action_requested = pyqtSignal(str, str)

    def __init__(self, mtg_info: dict, tdocs_data: list, filepath: str):
        super().__init__()
        self.mtg_info = mtg_info
        self.filepath = filepath
        self.meeting_dir = Path(filepath).parent.parent
        self.active_threads = {}

        self.db = TDocsDatabase(self.meeting_dir / "Agenda" / "user_tdocs.db")
        user_data = self.db.get_all()

        wg_name = str(self.mtg_info.get('wg_name', '')).upper()
        self.is_sa2 = ('SA2' in wg_name)
        is_electronic = bool(self.mtg_info.get('is_electronic', 0))
        self.is_sa2_electronic = self.is_sa2 and is_electronic

        main_ftp = self.mtg_info.get("url_key", "")
        if main_ftp and not main_ftp.startswith("http"): main_ftp = "https://www.3gpp.org/ftp/" + main_ftp.lstrip('/')
        self.main_ftp_url = main_ftp

        docs_ftp = self.mtg_info.get("docs_folder_url", "")
        if docs_ftp and not docs_ftp.startswith("http"): docs_ftp = "https://www.3gpp.org/ftp/" + docs_ftp.lstrip('/')
        self.docs_ftp_url = docs_ftp

        self.revisions_url = self.main_ftp_url.rstrip('/') + '/INBOX/Revisions/' if (
                    self.is_sa2_electronic and self.main_ftp_url) else ""

        mtg_icon = "💻" if is_electronic else "🤝"
        title = f"TDocs: {mtg_info.get('wg_name', '')} {mtg_info.get('meeting_number', '')} {mtg_icon}"
        self.setWindowTitle(title)
        self.resize(1400, 750)
        self.setStyleSheet("QWidget { background-color: #FAFAFA; }")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        self._setup_header(main_layout, title, is_electronic, len(tdocs_data))
        self._setup_filters(main_layout)
        self._setup_table(main_layout, tdocs_data, user_data)
        self._setup_cache()

    def _setup_header(self, layout, title, is_electronic, count):
        header_layout = QHBoxLayout()
        title_lbl = QLabel(f"<b>{title}</b>")
        title_lbl.setStyleSheet("font-size: 18px; color: #333;")
        title_lbl.setToolTip("Electronic Meeting (eMeeting)" if is_electronic else "In-Person Meeting (Face-to-Face)")
        title_lbl.setCursor(Qt.WhatsThisCursor)

        self.last_mod_lbl = QLabel(self._get_mod_date_str())
        self.last_mod_lbl.setStyleSheet("font-size: 11px; color: #999999; margin-right: 15px; font-style: italic;")

        # --- UNIFIED CLEAN BUTTON STYLING ---
        def style_btn():
            return """
            QPushButton { 
                font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold; 
                border-radius: 6px; padding: 6px 12px; 
                color: #333333; background-color: #FFFFFF; border: 1px solid #CCCCCC; 
            }
            QPushButton:hover, QPushButton::menu-indicator { 
                background-color: #F0F4F8; border: 1px solid #005A9E; color: #005A9E;
            }
            """

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setStyleSheet(style_btn())

        refresh_menu = QMenu(self)
        refresh_menu.addAction("📗 Refresh Excel List", self._refresh_excel)
        if self.is_sa2: refresh_menu.addAction("📄 Import TdocsByAgenda.htm", self._fetch_tdocs_by_agenda)
        if self.is_sa2_electronic:
            refresh_menu.addAction("📝 Refresh Revisions", lambda: self._refresh_revisions(silent=False))
            refresh_menu.addAction("🔄 Refresh Excel && Revisions", self._refresh_both)
        self.refresh_btn.setMenu(refresh_menu)

        self.folder_btn = QPushButton("🗂️ Resources")
        self.folder_btn.setStyleSheet(style_btn())

        folder_menu = QMenu(self)
        folder_menu.addAction("📁 Local: Meeting Folder", self._open_meeting_folder)
        folder_menu.addSeparator()
        folder_menu.addAction("📝 Export Markdown Reports", self._export_reports)
        folder_menu.addSeparator()
        if self.is_sa2: folder_menu.addAction("📄 Local: TdocsByAgenda.htm", self._open_agenda_file)
        folder_menu.addSeparator()
        if self.main_ftp_url: folder_menu.addAction("🌐 FTP: Main Folder", lambda: webbrowser.open(self.main_ftp_url))
        if self.docs_ftp_url: folder_menu.addAction("🌐 FTP: Docs Folder", lambda: webbrowser.open(self.docs_ftp_url))
        if self.revisions_url: folder_menu.addAction("🌐 FTP: Revisions Folder",
                                                     lambda: webbrowser.open(self.revisions_url))
        self.folder_btn.setMenu(folder_menu)

        self.excel_btn = QPushButton("📗 Excel")
        self.excel_btn.setStyleSheet(style_btn())
        self.excel_btn.clicked.connect(self._open_excel)

        self.export_btn = QPushButton("📝 Export")
        self.export_btn.setStyleSheet(style_btn())
        self.export_btn.clicked.connect(self._export_reports)

        self.stats_btn = QPushButton("📊 Statistics")
        self.stats_btn.setStyleSheet(style_btn())
        self.stats_btn.clicked.connect(self._generate_statistics)

        self.stats_cfg_btn = QPushButton("⚙️")
        self.stats_cfg_btn.setStyleSheet(style_btn())
        self.stats_cfg_btn.setFixedWidth(35)
        self.stats_cfg_btn.setToolTip("Configure Statistics Parameters")
        self.stats_cfg_btn.clicked.connect(self._open_stats_config)

        self.email_btn = QPushButton("📧 Emails")
        self.email_btn.setStyleSheet(style_btn())
        self.email_btn.clicked.connect(self._open_email_manager)
        self.email_btn.setVisible(self.is_sa2_electronic)

        self.count_lbl = QLabel(f"Showing {count} of {count} TDocs")
        self.count_lbl.setStyleSheet("font-size: 13px; color: #666;")

        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.last_mod_lbl)
        header_layout.addWidget(self.refresh_btn)
        header_layout.addWidget(self.folder_btn)
        header_layout.addWidget(self.excel_btn)
        header_layout.addWidget(self.export_btn)
        header_layout.addWidget(self.stats_btn)
        header_layout.addWidget(self.stats_cfg_btn)
        header_layout.addWidget(self.email_btn)
        header_layout.addSpacing(15)
        header_layout.addWidget(self.count_lbl)
        layout.addLayout(header_layout)

    def _setup_filters(self, layout):
        filter_frame = QFrame()
        filter_frame.setStyleSheet(
            "QFrame { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; } QLabel { font-weight: bold; color: #555; border: none; } QLineEdit, QComboBox { padding: 6px; border: 1px solid #CCC; border-radius: 4px; background: #FFF; }")
        filter_layout = QHBoxLayout(filter_frame)

        filter_layout.addWidget(QLabel("🔍 Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search TDoc number, title, source, or abstract...")
        self.search_input.textChanged.connect(self._on_search_changed)
        filter_layout.addWidget(self.search_input)

        self.type_combo = CheckableComboBox("Type")
        self.type_combo.selectionChanged.connect(self._on_type_changed)
        filter_layout.addWidget(self.type_combo)

        self.ai_combo = CheckableComboBox("AI")
        self.ai_combo.selectionChanged.connect(self._on_ai_changed)
        filter_layout.addWidget(self.ai_combo)

        self.status_combo = CheckableComboBox("TDoc Status")
        self.status_combo.selectionChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

        if self.is_sa2:
            self.chk_no_comments = QCheckBox("No Comments Only")
            self.chk_no_comments.toggled.connect(self._on_no_comments_toggled)
            filter_layout.addWidget(self.chk_no_comments)

        layout.addWidget(filter_frame)

    def _setup_table(self, layout, data, user_data):
        self.table = QTableView()
        self.model = TDocsTableModel(self.meeting_dir, data, user_data)
        self.proxy = TDocsFilterProxyModel()
        self.proxy.setSourceModel(self.model)
        self.proxy.layoutChanged.connect(self._update_count_label)

        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectItems)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.doubleClicked.connect(self._show_cell_popup)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setStyleSheet(
            "QTableView { gridline-color: #E0E0E0; border: 1px solid #E0E0E0; background-color: #FFFFFF; } QHeaderView::section { background-color: #F5F5F5; padding: 4px; font-weight: bold; border: 1px solid #E0E0E0; }")
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(48)

        self.action_delegate = TDocActionDelegate(self.table)
        self.action_delegate.actionClicked.connect(self._handle_tdoc_action)
        self.table.setItemDelegateForColumn(0, self.action_delegate)

        self.html_delegate = HtmlDelegate(self.table)
        self.html_delegate.linkClicked.connect(self._scroll_to_tdoc)
        self.html_delegate.linkRightClicked.connect(self._show_related_menu)
        self.table.setItemDelegateForColumn(7, self.html_delegate)
        self.table.setItemDelegateForColumn(12, self.html_delegate)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 110)
        header.resizeSection(1, 100)
        header.resizeSection(2, 200)
        header.resizeSection(3, 100)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.resizeSection(6, 28)
        header.setSectionResizeMode(7, QHeaderView.Stretch)
        header.resizeSection(8, 90)
        header.setSectionResizeMode(9, QHeaderView.Fixed)
        header.resizeSection(9, 28)
        header.resizeSection(10, 80)
        header.resizeSection(12, 160)

        layout.addWidget(self.table)
        self._refresh_comboboxes()

        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self._copy_table_selection)

    def _setup_cache(self):
        agenda_dir = self.meeting_dir / "Agenda"
        if self.is_sa2:
            local_agenda = agenda_dir / "TdocsByAgenda.htm"
            if local_agenda.exists():
                agenda_data = TDocsParser.parse_tdocs_by_agenda(str(local_agenda))
                if agenda_data:
                    self.model.merge_agenda_data(agenda_data)
                    self._refresh_comboboxes()

        if self.is_sa2_electronic:
            local_revs = agenda_dir / "revisions.json"
            if local_revs.exists():
                try:
                    with open(local_revs, "r", encoding="utf-8") as f:
                        self.model.revisions = json.load(f)
                        self.model.dataChanged.emit(self.model.index(0, 0),
                                                    self.model.index(self.model.rowCount() - 1, 0))
                except:
                    if self.revisions_url: self._refresh_revisions(silent=True)
            else:
                if self.revisions_url: self._refresh_revisions(silent=True)

    def _get_mod_date_str(self):
        try:
            return f"List last updated: {datetime.datetime.fromtimestamp(os.path.getmtime(self.filepath)).strftime('%Y-%m-%d %H:%M')}"
        except:
            return "List last updated: Unknown"

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
        self.count_lbl.setText(f"Showing {self.proxy.rowCount()} of {self.model.rowCount()} TDocs")

    def _refresh_comboboxes(self):
        def sanitize(val): return str(val).strip() if val is not None else ""

        unique_types = sorted(list(set(sanitize(r.get("Type", "")) for r in self.model._data)))
        unique_ais = sorted(list(set(sanitize(r.get("Agenda Item", "")) for r in self.model._data)),
                            key=natural_sort_key)
        unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in self.model._data)))

        self.type_combo.updateItems(unique_types)
        self.ai_combo.updateItems(unique_ais)
        self.status_combo.updateItems(unique_statuses)
        self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
        self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
        self.proxy.setStatusFilters(self.status_combo.getCheckedItems())
        self._update_count_label()

    def _clear_all_filters(self):
        self.search_input.blockSignals(True);
        self.search_input.clear();
        self.search_input.blockSignals(False)
        self.proxy.setGlobalFilter("")
        if self.is_sa2:
            self.chk_no_comments.blockSignals(True);
            self.chk_no_comments.setChecked(False);
            self.chk_no_comments.blockSignals(False)
            self.proxy.setNoCommentsFilter(False)
        for combo in [self.type_combo, self.ai_combo, self.status_combo]:
            combo.blockSignals(True)
            combo.model().item(0).setCheckState(Qt.Checked)
            for i in range(1, combo.model().rowCount()): combo.model().item(i).setCheckState(Qt.Checked)
            combo.updateText()
            combo.blockSignals(False)
        self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
        self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
        self.proxy.setStatusFilters(self.status_combo.getCheckedItems())

    def _handle_tdoc_action(self, base_tdoc: str):
        if base_tdoc in self.model.loading_tdocs or not self.docs_ftp_url: return
        revisions = self.model.revisions.get(base_tdoc, [])
        build_action_menu(self.table, base_tdoc, self.docs_ftp_url, self.revisions_url, revisions, self.meeting_dir,
                          self._trigger_download_thread, QCursor.pos())

    def _show_related_menu(self, target_tdoc: str, pos: QPoint):
        build_related_menu(self, target_tdoc, self.model.valid_tdocs, self.docs_ftp_url, self.revisions_url,
                           self._scroll_to_tdoc, self._trigger_download_thread, self.global_action_requested.emit, pos)

    def _show_cell_popup(self, index):
        if not index.isValid(): return
        col_name = self.model._headers[index.column()]
        if col_name not in ["Secretary Remarks", "Title", "Source", "Abstract", "My Notes", "My Status"]: return

        row_data = self.model._data[self.proxy.mapToSource(index).row()]
        tdoc_id = row_data.get("TDoc", "")

        if col_name in ["Title", "Source", "Abstract"]:
            val = next((str(v) for k, v in row_data.items() if str(k).strip().lower() == col_name.lower() and v), "")
            ReadOnlyViewerDialog(self, f"📄 Viewing: {col_name} ({tdoc_id})", val).exec_()
        else:
            InteractiveNotesDialog(self, tdoc_id, row_data, self._save_user_data).exec_()

    def _save_user_data(self, tdoc_id: str, status: str, notes: str):
        self.db.upsert(tdoc_id, status, notes)
        self.model.user_data = self.db.get_all()
        self.model.apply_user_data_refresh()

    def _export_reports(self):
        self.export_thread = MarkdownExporterThread(self.meeting_dir, self.model._data, self.docs_ftp_url,
                                                    self.mtg_info)
        self.export_thread.finished.connect(lambda s, m: self._on_export_finished(s, m, False))
        self.export_thread.start()

    def _open_stats_config(self):
        StatisticsSettingsDialog(self).exec_()

    def _generate_statistics(self):
        self.stats_btn.setText("⏳ Generating...")
        self.stats_btn.setEnabled(False)

        config = StatisticsSettingsDialog().load_config()
        self.stats_thread = StatisticsExporterThread(self.meeting_dir, self.model._data, self.mtg_info, config)
        self.stats_thread.finished.connect(lambda s, m: self._on_export_finished(s, m, True))
        self.stats_thread.start()

    def _on_export_finished(self, success: bool, msg: str, is_stats: bool):
        if is_stats:
            self.stats_btn.setText("📊 Statistics")
            self.stats_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Export Complete", f"Successfully generated:\n{msg}")
            if hasattr(os, 'startfile'): os.startfile(str(msg))
        else:
            QMessageBox.warning(self, "Export Failed", msg)

    def _scroll_to_tdoc(self, target_tdoc: str):
        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', target_tdoc, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else target_tdoc.upper()
        if base_tdoc in self.model.valid_tdocs:
            for row in range(self.proxy.rowCount()):
                if self.proxy.data(self.proxy.index(row, 1), Qt.UserRole) == base_tdoc:
                    self.table.scrollTo(self.proxy.index(row, 1), QTableView.PositionAtCenter)
                    self.table.selectRow(row)
                    return
            if QMessageBox.question(self, "Hidden TDoc",
                                    f"TDoc '{base_tdoc}' is hidden by active filters.\nClear filters to view?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self._clear_all_filters()
                self._scroll_to_tdoc(target_tdoc)
        else:
            if QMessageBox.question(self, "External TDoc",
                                    f"{base_tdoc} is not from this meeting.\nSearch global database?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self.global_action_requested.emit(base_tdoc, 'open_meeting')

    def _trigger_download_thread(self, base_tdoc: str, target_filename: str, base_url: str,
                                 is_silent_compare: bool = False):
        self.model.set_loading(base_tdoc, True)
        thread = TDocActionThread(base_tdoc, target_filename, base_url, self.meeting_dir,
                                  open_file=not is_silent_compare)
        thread.is_silent_compare = is_silent_compare
        thread.target_filename = target_filename
        thread.finished_action.connect(lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th))
        self.active_threads[base_tdoc] = thread
        thread.start()

    def _on_tdoc_action_finished(self, tdoc: str, success: bool, msg: str, thread: TDocActionThread):
        if tdoc in self.active_threads: del self.active_threads[tdoc]
        self.model.set_loading(tdoc, False)
        if not success: return QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)
        if getattr(thread, "is_silent_compare", False):
            files = getattr(thread, "extracted_doc_paths", [])
            if files:
                ComparisonManager.get_instance().add_to_cart(thread.target_filename, str(files[0]))
            else:
                QMessageBox.warning(self, "Compare Failed", "No Word document found in TDoc ZIP.")

    def _refresh_both(self):
        self._refresh_excel(); self._refresh_revisions(silent=True)

    def _refresh_excel(self):
        if not self.mtg_info.get("mtg_id"): return QMessageBox.warning(self, "Missing ID",
                                                                       "Cannot refresh: Missing 3GPP Portal ID.")
        self.refresh_btn.setText("⏳ Downloading...");
        self.refresh_btn.setEnabled(False)
        self.dl_thread = TDocsDownloaderThread(self.mtg_info.get("mtg_id"), self.meeting_dir, self)
        self.dl_thread.finished.connect(self._on_refresh_excel_finished)
        self.dl_thread.start()

    def _on_refresh_excel_finished(self, success: bool, result: str, mtg_id: str):
        self.refresh_btn.setText("🔄 Refresh");
        self.refresh_btn.setEnabled(True)
        if success:
            self.filepath = result
            if new_data := TDocsParser.parse_tdocs_excel(self.filepath):
                self.model.update_data(new_data);
                self._refresh_comboboxes();
                self.last_mod_lbl.setText(self._get_mod_date_str())
            else:
                QMessageBox.warning(self, "Parse Error", "Downloaded, but could not parse the Excel file.")
        else:
            QMessageBox.critical(self, "Download Error", f"Failed to refresh TDocs:\n{result}")

    def _refresh_revisions(self, silent=False):
        if not self.revisions_url: return
        self.rev_thread = TDocsRevisionsFetcherThread(self.revisions_url, self.meeting_dir)
        self.rev_thread.finished.connect(lambda s, d, m: self._on_revisions_fetched(s, d, m, silent))
        self.rev_thread.start()

    def _on_revisions_fetched(self, success: bool, data: dict, msg: str, silent: bool):
        if success:
            self.model.revisions = data
            self.model.dataChanged.emit(self.model.index(0, 0), self.model.index(self.model.rowCount() - 1, 0))
            if not silent:
                self.refresh_btn.setText(f"✅ {len(data)} Revs");
                QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))
        elif not silent:
            QMessageBox.warning(self, "Revisions Error", f"Failed to sync revisions:\n{msg}")

    def _fetch_tdocs_by_agenda(self):
        url_key = self.mtg_info.get("url_key", "")
        if not url_key: return
        self.refresh_btn.setText("⏳ Parsing HTML...")
        self.agenda_thread = TdocsByAgendaThread(
            url_key if url_key.startswith("http") else f"https://www.3gpp.org/ftp/{url_key.lstrip('/')}",
            self.meeting_dir)
        self.agenda_thread.finished.connect(self._on_agenda_fetched)
        self.agenda_thread.start()

    def _on_agenda_fetched(self, success: bool, agenda_data: dict):
        if success and agenda_data:
            self.model.merge_agenda_data(agenda_data);
            self._refresh_comboboxes()
            self.refresh_btn.setText(f"✅ {len(agenda_data)} Merged");
            QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))
        else:
            self.refresh_btn.setText("🔄 Refresh");
            QMessageBox.warning(self, "Error", "Failed to parse TdocsByAgenda.htm.")

    def _open_meeting_folder(self):
        __open_folder(self.meeting_dir)

    def _open_agenda_file(self):
        __open_folder(self.meeting_dir / "Agenda" / "TdocsByAgenda.htm")

    def _open_excel(self):
        __open_folder(Path(self.filepath))

    def _open_email_manager(self):
        self.email_window = EmailManagerWindow(self.meeting_dir, {
            str(r.get("TDoc", "")).strip().upper(): str(r.get("Agenda Item", "N/A")).strip() for r in self.model._data
            if r.get("TDoc")}, self.mtg_info.get("start_date", ""), self.mtg_info.get("end_date", ""))
        self.email_window.show()

    def _copy_table_selection(self):
        indexes = sorted(self.table.selectionModel().selectedIndexes(), key=lambda x: (x.row(), x.column()))
        if not indexes: return
        lines, current_line, current_row = [], [], indexes[0].row()
        for idx in indexes:
            if idx.row() != current_row:
                lines.append("\t".join(current_line));
                current_line = [];
                current_row = idx.row()
            cell_text = str(idx.data(Qt.UserRole) or "").strip()
            if not cell_text: cell_text = str(idx.data(Qt.DisplayRole) or "").strip()
            current_line.append(cell_text)
        lines.append("\t".join(current_line))
        QApplication.clipboard().setText("\n".join(lines));
        QToolTip.showText(QCursor.pos(), "📋 Copied to clipboard!", self.table)


def __open_folder(p: Path):
    if p.exists():
        os.startfile(str(p)) if hasattr(os, 'startfile') else webbrowser.open(f"file:///{p}")
    else:
        QMessageBox.warning(None, "Not Found", "Target file/folder does not exist yet.")