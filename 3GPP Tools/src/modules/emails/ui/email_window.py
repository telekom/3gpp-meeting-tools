# --- File: modules/emails/ui/email_window.py ---
import datetime
import json
import os
import webbrowser
from pathlib import Path
from PyQt5.QtCore import Qt, pyqtSignal, QDate
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QLineEdit, QTableView, QHeaderView, QSplitter, QTextBrowser,
                             QMessageBox, QInputDialog, QDialog, QDateEdit, QMenu, QApplication)

from modules.emails.core.email_db import EmailDatabase
from modules.emails.core.email_threads import EmailSyncThread, EmailMoveThread, EmailTargetRescanThread
from modules.emails.ui.email_models import EmailTableModel, EmailProxyModel
from modules.meetings.ui.tdocs_components import CheckableComboBox


class EmailManagerWindow(QWidget):
    # Emit this signal to open the TDoc in your main window
    tdoc_open_requested = pyqtSignal(str)

    def __init__(self, meeting_dir: Path, ai_lookup: dict, meeting_start: str = "", meeting_end: str = ""):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.ai_lookup = ai_lookup

        self.db = EmailDatabase(self.meeting_dir / "Agenda" / "emails.db")
        self.config_path = self.meeting_dir / "Agenda" / "email_config.json"

        config_data = self._load_config()
        self.source_folder = config_data.get("source_folder", "")
        self.target_folder = config_data.get("target_folder", "")

        saved_start = config_data.get("start_date", "")
        saved_end = config_data.get("end_date", "")

        if saved_start and saved_end:
            self.start_date = saved_start
            self.end_date = saved_end
        else:
            try:
                if meeting_start and meeting_end:
                    s_dt = datetime.datetime.strptime(meeting_start, "%Y-%m-%d") - datetime.timedelta(days=3)
                    e_dt = datetime.datetime.strptime(meeting_end, "%Y-%m-%d") + datetime.timedelta(days=3)
                    self.start_date = s_dt.strftime("%Y-%m-%d")
                    self.end_date = e_dt.strftime("%Y-%m-%d")
                else:
                    self.start_date = ""
                    self.end_date = ""
            except Exception:
                self.start_date, self.end_date = "", ""

        self.setWindowTitle("📧 eMeeting Email Manager")
        self.resize(1200, 800)
        self.setStyleSheet("QWidget { background-color: #FAFAFA; }")

        self._setup_ui()
        self._refresh_table()

    def _load_config(self) -> dict:
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    return json.load(f)
            except:
                pass
        return {"source_folder": "", "target_folder": ""}

    def _save_config(self, source_folder: str, target_folder: str, sd: str, ed: str):
        self.source_folder = source_folder
        self.target_folder = target_folder
        self.start_date = sd
        self.end_date = ed
        with open(self.config_path, 'w') as f:
            json.dump(
                {"source_folder": source_folder, "target_folder": target_folder, "start_date": sd, "end_date": ed}, f)

    def _configure_folders(self):
        from modules.emails.ui.config_dialog import EmailConfigDialog
        dialog = EmailConfigDialog(self.source_folder, self.target_folder, self)
        if dialog.exec_() == QDialog.Accepted:
            src, tgt = dialog.get_paths()
            self.source_folder = src
            self.target_folder = tgt
            self._save_config(src, tgt, self.start_date, self.end_date)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)

        # ---> NEW: Cohesive, Professional Button Styling Helper
        def get_btn_style(primary=False):
            if primary:
                return """
                QPushButton { font-weight: bold; background-color: #0078D7; color: white; padding: 6px 12px; border-radius: 4px; border: 1px solid #005A9E; }
                QPushButton:hover { background-color: #005A9E; }
                """
            return """
                QPushButton { font-weight: bold; background-color: #FFFFFF; color: #333333; padding: 6px 12px; border-radius: 4px; border: 1px solid #CCCCCC; }
                QPushButton:hover { background-color: #F0F4F8; border-color: #005A9E; color: #005A9E; }
                """

        # --- TOOLBAR ---
        toolbar = QHBoxLayout()
        self.btn_sync = QPushButton("🔄 Sync Source")
        self.btn_sync.setStyleSheet(get_btn_style(primary=True))  # Only Sync gets the bold primary blue
        self.btn_sync.clicked.connect(self._run_sync)

        self.btn_move = QPushButton("➡️ Move Selected")
        self.btn_move.setStyleSheet(get_btn_style())
        self.btn_move.clicked.connect(self._run_move)

        self.btn_move_all = QPushButton("⏭️ Move All")
        self.btn_move_all.setStyleSheet(get_btn_style())
        self.btn_move_all.clicked.connect(self._run_move_all)

        self.btn_rescan = QPushButton("🔁 Scan Target")
        self.btn_rescan.setStyleSheet(get_btn_style())
        self.btn_rescan.clicked.connect(self._run_target_rescan)

        self.btn_stats = QPushButton("📊 Statistics")
        self.btn_stats.setStyleSheet(get_btn_style())
        self.btn_stats.clicked.connect(self._generate_statistics)

        self.btn_config = QPushButton("⚙️ Folders")
        self.btn_config.setStyleSheet(get_btn_style())
        self.btn_config.clicked.connect(self._configure_folders)

        self.lbl_status = QLabel("Ready.")
        self.lbl_status.setStyleSheet("color: #666; font-style: italic; margin-left: 10px;")

        toolbar.addWidget(self.btn_sync)
        toolbar.addWidget(self.btn_move)
        toolbar.addWidget(self.btn_move_all)
        toolbar.addWidget(self.btn_rescan)
        toolbar.addWidget(self.btn_stats)
        toolbar.addWidget(self.btn_config)
        toolbar.addWidget(self.lbl_status)
        toolbar.addStretch()
        main_layout.addLayout(toolbar)

        # --- FILTERS & DATES ---
        filter_layout = QHBoxLayout()

        # Date Pickers
        self.dt_start = QDateEdit()
        self.dt_start.setCalendarPopup(True)
        if getattr(self, 'start_date', None):
            self.dt_start.setDate(QDate.fromString(self.start_date, Qt.ISODate))
        else:
            self.dt_start.setDate(QDate.currentDate().addDays(-7))

        self.dt_end = QDateEdit()
        self.dt_end.setCalendarPopup(True)
        if getattr(self, 'end_date', None):
            self.dt_end.setDate(QDate.fromString(self.end_date, Qt.ISODate))
        else:
            self.dt_end.setDate(QDate.currentDate().addDays(7))

        self.dt_start.dateChanged.connect(lambda d: setattr(self, 'start_date', d.toString(Qt.ISODate)))
        self.dt_end.dateChanged.connect(lambda d: setattr(self, 'end_date', d.toString(Qt.ISODate)))
        self.dt_start.dateChanged.connect(
            lambda: self._save_config(self.source_folder, self.target_folder, self.start_date, self.end_date))
        self.dt_end.dateChanged.connect(
            lambda: self._save_config(self.source_folder, self.target_folder, self.start_date, self.end_date))

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.textChanged.connect(self._apply_filters)

        self.cb_ai = CheckableComboBox("AI")
        self.cb_company = CheckableComboBox("Company")
        self.cb_sender = CheckableComboBox("Sender")

        for cb in [self.cb_ai, self.cb_company, self.cb_sender]:
            cb.setMinimumWidth(120)
            cb.selectionChanged.connect(self._apply_filters)

        self.btn_filter_star = QPushButton("⭐ Starred")
        self.btn_filter_star.setCheckable(True)
        self.btn_filter_star.setStyleSheet(
            "QPushButton { padding: 4px; border: 1px solid #CCC; background: white; border-radius: 3px; } QPushButton:checked { background-color: #FFF4CE; font-weight: bold; border-color: #E2C08D; }")
        self.btn_filter_star.clicked.connect(self._apply_filters)

        self.btn_filter_follow = QPushButton("👀 Followed AIs")
        self.btn_filter_follow.setCheckable(True)
        self.btn_filter_follow.setStyleSheet(
            "QPushButton { padding: 4px; border: 1px solid #CCC; background: white; border-radius: 3px; } QPushButton:checked { background-color: #E6F4E6; font-weight: bold; color: #0C6B0C; border-color: #0C6B0C; }")
        self.btn_filter_follow.clicked.connect(self._apply_filters)

        # ---> NEW: The Live Email Count Label
        self.lbl_count = QLabel("Showing 0 of 0 Emails")
        self.lbl_count.setStyleSheet("font-size: 13px; color: #555; font-weight: bold; padding-left: 10px;")

        filter_layout.addWidget(QLabel("📅 Filter:"))
        filter_layout.addWidget(self.dt_start)
        filter_layout.addWidget(QLabel("-"))
        filter_layout.addWidget(self.dt_end)
        filter_layout.addSpacing(15)

        filter_layout.addWidget(self.btn_filter_star)
        filter_layout.addWidget(self.btn_filter_follow)
        filter_layout.addWidget(QLabel("🔍:"))
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.cb_ai)
        filter_layout.addWidget(self.cb_company)
        filter_layout.addWidget(self.cb_sender)
        filter_layout.addWidget(self.lbl_count)  # Added to the end of the filter bar
        main_layout.addLayout(filter_layout)

        # --- SPLITTER ---
        splitter = QSplitter(Qt.Vertical)

        # Table
        self.table = QTableView()
        self.model = EmailTableModel()
        self.proxy = EmailProxyModel()
        self.proxy.setSourceModel(self.model)
        self.table.setModel(self.proxy)

        # ---> NEW: Connect the proxy filter and model resets to our dynamic count updater
        self.proxy.layoutChanged.connect(self._update_count_label)
        self.model.modelReset.connect(self._update_count_label)

        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet("QTableView { background: white; gridline-color: #EEE; }")

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 30)
        header.resizeSection(1, 85)
        header.resizeSection(2, 75)
        header.resizeSection(3, 120)
        header.resizeSection(4, 90)
        header.resizeSection(5, 60)
        header.resizeSection(6, 70)
        header.resizeSection(7, 110)
        header.resizeSection(8, 140)
        header.setSectionResizeMode(9, QHeaderView.Stretch)

        self.table.selectionModel().selectionChanged.connect(self._on_email_selected)

        # ---> NEW: Click handling for Hyperlinks & Right-click Context Menu
        self.table.setCursor(Qt.PointingHandCursor)
        self.table.clicked.connect(self._on_table_clicked)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)

        splitter.addWidget(self.table)

        # Reading Pane
        pane_widget = QWidget()
        pane_layout = QVBoxLayout(pane_widget)
        pane_layout.setContentsMargins(0, 10, 0, 0)

        self.btn_open_msg = QPushButton("📄 Open .msg File in Outlook")
        self.btn_open_msg.setEnabled(False)
        self.btn_open_msg.clicked.connect(self._open_current_msg)

        self.btn_toggle_star = QPushButton("⭐ Star TDoc")
        self.btn_toggle_star.setEnabled(False)
        self.btn_toggle_star.clicked.connect(self._toggle_current_star)

        self.btn_toggle_follow = QPushButton("👀 Follow AI")
        self.btn_toggle_follow.setEnabled(False)
        self.btn_toggle_follow.clicked.connect(self._toggle_current_ai_follow)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.btn_open_msg)
        btn_layout.addWidget(self.btn_toggle_star)
        btn_layout.addWidget(self.btn_toggle_follow)
        btn_layout.addStretch()

        self.reading_pane = QTextBrowser()
        self.reading_pane.setStyleSheet("background: white; border: 1px solid #CCC; border-radius: 4px; padding: 10px;")
        pane_layout.addLayout(btn_layout)
        pane_layout.addWidget(self.reading_pane)
        splitter.addWidget(pane_widget)

        splitter.setSizes([400, 300])
        main_layout.addWidget(splitter)

    # ---> NEW: Click Logic
    def _on_table_clicked(self, index):
        if not index.isValid(): return
        source_idx = self.proxy.mapToSource(index)
        col_name = self.model._headers[index.column()]
        row_data = self.model.get_row_data(source_idx.row())

        if col_name == "Sender":
            email = row_data.get("sender_email", "")
            if email:
                webbrowser.open(f"mailto:{email}")

        elif col_name == "Rev":
            revs = row_data.get("revisions_mentioned", "")
            if revs:
                # The DB stores the FULL TDoc string internally (S2-261234r05), even though we shortened it for display.
                first_rev = revs.split(',')[0].strip()
                # Broadcast signal so TDocsWindow can catch it!
                self.tdoc_open_requested.emit(first_rev)
                self.lbl_status.setText(f"Requesting open for: {first_rev}")

    def _show_context_menu(self, pos):
        index = self.table.indexAt(pos)
        if not index.isValid(): return

        col_name = self.model._headers[index.column()]
        if col_name == "Sender":
            row_data = self.model.get_row_data(self.proxy.mapToSource(index).row())
            email = row_data.get("sender_email", "")

            menu = QMenu(self)
            copy_action = menu.addAction("📋 Copy Email Address")
            action = menu.exec_(self.table.viewport().mapToGlobal(pos))

            if action == copy_action and email:
                QApplication.clipboard().setText(email)
                self.lbl_status.setText(f"Copied {email} to clipboard.")

    def _run_sync(self):
        if not self.source_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Source Folder first.")
            return

        self._set_buttons_enabled(False)
        self.btn_sync.setText("⏳ Syncing...")

        sd = self.dt_start.date().toString(Qt.ISODate)
        ed = self.dt_end.date().toString(Qt.ISODate)

        self.sync_thread = EmailSyncThread(self.source_folder, self.meeting_dir, self.ai_lookup, self.db, sd, ed)
        self.sync_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))
        self.sync_thread.progress_update.connect(
            lambda c, t: self.lbl_status.setText(f"Scanning Source: {c} / {t} items..."))
        self.sync_thread.finished.connect(self._on_sync_finished)
        self.sync_thread.start()

    def _on_sync_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_sync.setText("🔄 Sync Source")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _refresh_table(self):
        import sqlite3
        with sqlite3.connect(self.db.db_path) as conn:
            conn.row_factory = sqlite3.Row
            data = [dict(row) for row in conn.execute('SELECT * FROM emails ORDER BY date_received DESC').fetchall()]

        starred_tdocs = self.db.get_starred_tdocs()
        followed_ais = self.db.get_followed_ais()
        self.model.update_data(data, starred_tdocs, followed_ais)

        def clean(val): return str(val).strip() if val else ""

        # ---> FIX 4: Regex-powered Natural Sorting Key
        import re
        def natural_sort_key(s):
            # Splits the string by numbers, allowing 20.6.2 to correctly sort before 20.6.19
            return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

        # Extract unique, cleaned values
        unique_ais = set(clean(r.get("agenda_item")) for r in data)
        unique_companies = set(clean(r.get("company")) for r in data)
        unique_senders = set(clean(r.get("sender_name")) for r in data)

        # Apply the natural sort key to the Agenda Items
        self.cb_ai.updateItems(sorted(unique_ais, key=natural_sort_key))
        self.cb_company.updateItems(sorted(unique_companies))
        self.cb_sender.updateItems(sorted(unique_senders))

        self._apply_filters()

    def _apply_filters(self):
        self.proxy.set_filters(
            self.search_input.text(),
            self.cb_ai.getCheckedItems(),
            self.cb_company.getCheckedItems(),
            self.cb_sender.getCheckedItems(),
            self.btn_filter_star.isChecked(),
            self.btn_filter_follow.isChecked()
        )

        self._update_count_label()

    def _on_email_selected(self, selected, deselected):
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            self.reading_pane.clear()
            self.btn_open_msg.setEnabled(False)
            self.btn_toggle_star.setEnabled(False)
            self.btn_toggle_follow.setEnabled(False)
            return

        source_idx = self.proxy.mapToSource(indexes[0])
        row_data = self.model.get_row_data(source_idx.row())

        self.current_tdoc_id = row_data.get("tdoc_id", "")
        self.current_agenda_item = row_data.get("agenda_item", "")
        self.current_msg_path = row_data.get("msg_path", "")

        self.btn_open_msg.setEnabled(bool(self.current_msg_path))
        self.btn_toggle_star.setEnabled(bool(self.current_tdoc_id))
        self.btn_toggle_follow.setEnabled(bool(self.current_agenda_item))

        is_starred = self.current_tdoc_id in self.model.starred_tdocs
        self.btn_toggle_star.setText("❌ Unstar TDoc" if is_starred else "⭐ Star TDoc")

        is_followed = self.current_agenda_item in self.model.followed_ais
        ai_display = self.current_agenda_item or "Unknown AI"
        self.btn_toggle_follow.setText(f"❌ Unfollow AI: {ai_display}" if is_followed else f"👀 Follow AI: {ai_display}")

        # Display rich text
        real_email = row_data.get('sender_email', '')
        html = f"""
        <h2 style='color:#005A9E; margin-bottom: 2px;'>{row_data.get('subject', '')}</h2>
        <p style='color:#555; margin-top:0px;'><b>From:</b> {row_data.get('sender_name')} &lt;{real_email}&gt; ({row_data.get('company')}) | <b>TDoc:</b> {row_data.get('tdoc_id')} | <b>Rev:</b> {row_data.get('revisions_mentioned')}</p>
        <hr>
        <h3 style='color:#D83B01;'>Short Text / Comments</h3>
        <p style='background-color:#FFF3E0; padding:10px; border-left: 4px solid #FFB74D;'>{row_data.get('short_text', '').replace(chr(10), '<br>')}</p>
        <h3>Free Text</h3>
        <p style='color:#333;'>{row_data.get('free_text', '').replace(chr(10), '<br>')}</p>
        """
        self.reading_pane.setHtml(html)

    def _toggle_current_ai_follow(self):
        if not getattr(self, "current_agenda_item", ""): return
        selected_indexes = self.table.selectionModel().selectedRows()
        selected_id = None
        if selected_indexes:
            source_idx = self.proxy.mapToSource(selected_indexes[0])
            selected_id = self.model.get_row_data(source_idx.row()).get("id")

        is_followed = self.current_agenda_item in self.model.followed_ais
        self.db.toggle_ai_follow(self.current_agenda_item, not is_followed)
        self._refresh_table()

        if selected_id:
            for i, row in enumerate(self.model._data):
                if row.get("id") == selected_id:
                    proxy_idx = self.proxy.mapFromSource(self.model.index(i, 0))
                    if proxy_idx.isValid(): self.table.selectRow(proxy_idx.row())
                    break

    def _toggle_current_star(self):
        if not getattr(self, "current_tdoc_id", ""): return
        selected_indexes = self.table.selectionModel().selectedRows()
        selected_id = None
        if selected_indexes:
            source_idx = self.proxy.mapToSource(selected_indexes[0])
            selected_id = self.model.get_row_data(source_idx.row()).get("id")

        is_starred = self.current_tdoc_id in self.model.starred_tdocs
        self.db.toggle_tdoc_star(self.current_tdoc_id, not is_starred)
        self._refresh_table()

        if selected_id:
            for i, row in enumerate(self.model._data):
                if row.get("id") == selected_id:
                    proxy_idx = self.proxy.mapFromSource(self.model.index(i, 0))
                    if proxy_idx.isValid(): self.table.selectRow(proxy_idx.row())
                    break

    def _open_current_msg(self):
        if self.current_msg_path and Path(self.current_msg_path).exists():
            try:
                os.startfile(self.current_msg_path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    def _run_move(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return

        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            QMessageBox.information(self, "Selection Required", "Please select at least one email to move.")
            return

        items_to_move = []
        for idx in indexes:
            source_idx = self.proxy.mapToSource(idx)
            row_data = self.model.get_row_data(source_idx.row())
            if row_data.get("outlook_location") == "Source":
                items_to_move.append((row_data.get("id"), row_data.get("agenda_item")))

        if not items_to_move:
            QMessageBox.information(self, "Notice", "Selected emails have already been moved to the Target.")
            return

        self._set_buttons_enabled(False)
        self.btn_move.setText("⏳ Moving...")
        self.move_thread = EmailMoveThread(items_to_move, self.target_folder, self.db)
        self.move_thread.progress_update.connect(lambda c, t: self.lbl_status.setText(f"Moving {c}/{t}..."))
        self.move_thread.finished.connect(self._on_move_finished)
        self.move_thread.start()

    def _run_move_all(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return

        items_to_move = []
        for row_data in self.model._data:
            if row_data.get("outlook_location") == "Source":
                items_to_move.append((row_data.get("id"), row_data.get("agenda_item")))

        if not items_to_move:
            QMessageBox.information(self, "Notice", "There are no emails in the Source folder to move.")
            return

        reply = QMessageBox.question(self, 'Confirm Move All',
                                     f"Move ALL {len(items_to_move)} source emails to the target?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No: return

        self._set_buttons_enabled(False)
        self.btn_move_all.setText("⏳ Moving All...")
        self.move_thread = EmailMoveThread(items_to_move, self.target_folder, self.db)
        self.move_thread.progress_update.connect(lambda c, t: self.lbl_status.setText(f"Moving {c}/{t}..."))
        self.move_thread.finished.connect(self._on_move_all_finished)
        self.move_thread.start()

    def _on_move_all_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_move_all.setText("⏭️ Move All")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _on_move_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_move.setText("➡️ Move Selected")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _run_target_rescan(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return

        self._set_buttons_enabled(False)
        self.btn_rescan.setText("⏳ Scanning...")

        sd = self.dt_start.date().toString(Qt.ISODate)
        ed = self.dt_end.date().toString(Qt.ISODate)

        self.rescan_thread = EmailTargetRescanThread(self.target_folder, self.meeting_dir, self.ai_lookup, self.db, sd,
                                                     ed)
        self.rescan_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))
        self.rescan_thread.progress_update.connect(
            lambda c, t: self.lbl_status.setText(f"Scanning Target: {c} / {t} items..."))
        self.rescan_thread.finished.connect(self._on_rescan_finished)
        self.rescan_thread.start()

    def _on_rescan_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_rescan.setText("🔁 Scan Target")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _set_buttons_enabled(self, state: bool):
        self.btn_sync.setEnabled(state)
        self.btn_move.setEnabled(state)
        self.btn_move_all.setEnabled(state)
        self.btn_rescan.setEnabled(state)

    def _generate_statistics(self):
        # ---> THE FIX: Use VISIBLE emails based on active filters, not the entire DB!
        visible_emails = []
        for r in range(self.proxy.rowCount()):
            index = self.proxy.index(r, 0)
            source_index = self.proxy.mapToSource(index)
            row_data = self.model._data[source_index.row()]
            visible_emails.append(row_data)

        if not visible_emails:
            QMessageBox.warning(self, "No Data", "There are no visible emails to analyze. Clear filters and try again.")
            return

        self._set_buttons_enabled(False)
        self.btn_stats.setText("⏳ Generating...")

        from modules.emails.core.statistics_threads import EmailStatsExporterThread
        meeting_name = self.meeting_dir.name if self.meeting_dir else "Meeting"

        self.stats_thread = EmailStatsExporterThread(self.meeting_dir, visible_emails, meeting_name)
        self.stats_thread.finished.connect(self._on_stats_finished)
        self.stats_thread.start()

    def _on_stats_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_stats.setText("📊 Statistics")

        if success:
            self.lbl_status.setText("✅ Analytics Report Generated.")
            if hasattr(os, 'startfile'):
                os.startfile(msg)
            else:
                webbrowser.open(f"file:///{msg}")
        else:
            self.lbl_status.setText("❌ Analytics Generation Failed.")
            QMessageBox.warning(self, "Error", f"Could not generate statistics:\n{msg}")

    # Don't forget to update _set_buttons_enabled to include the new button!
    def _set_buttons_enabled(self, state: bool):
        self.btn_sync.setEnabled(state)
        self.btn_move.setEnabled(state)
        self.btn_move_all.setEnabled(state)
        self.btn_rescan.setEnabled(state)
        self.btn_stats.setEnabled(state)  # <--- Added

    def _update_count_label(self):
        """Dynamically updates the UI label based on active grid filters."""
        self.lbl_count.setText(f"Showing {self.proxy.rowCount()} of {self.model.rowCount()} Emails")