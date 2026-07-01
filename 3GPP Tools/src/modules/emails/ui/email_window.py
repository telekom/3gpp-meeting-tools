# --- File: modules/emails/ui/email_window.py ---
import datetime
import json
import os
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QLineEdit, QTableView, QHeaderView, QSplitter, QTextBrowser,
                             QMessageBox, QInputDialog, QDialog)

from modules.emails.core.email_db import EmailDatabase
from modules.emails.core.email_threads import EmailSyncThread, EmailMoveThread, EmailTargetRescanThread
from modules.emails.ui.email_models import EmailTableModel, EmailProxyModel
from modules.meetings.ui.tdocs_components import CheckableComboBox  # Reusing your awesome combo box!


class EmailManagerWindow(QWidget):
    def __init__(self, meeting_dir: Path, ai_lookup: dict, meeting_start: str = "", meeting_end: str = ""):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.ai_lookup = ai_lookup

        self.db = EmailDatabase(self.meeting_dir / "Agenda" / "emails.db")
        self.config_path = self.meeting_dir / "Agenda" / "email_config.json"

        config_data = self._load_config()
        self.source_folder = config_data.get("source_folder", "")
        self.target_folder = config_data.get("target_folder", "")

        # ---> NEW: Intelligent Date Pre-filling
        saved_start = config_data.get("start_date", "")
        saved_end = config_data.get("end_date", "")

        # If the user previously saved custom dates to JSON, use them.
        if saved_start and saved_end:
            self.start_date = saved_start
            self.end_date = saved_end
        else:
            # Otherwise, calculate the default +/- 3 day buffer from the exact meeting dates!
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
                self.start_date = ""
                self.end_date = ""

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

    def _save_config(self, source_folder: str, target_folder: str):
        self.source_folder = source_folder
        self.target_folder = target_folder
        with open(self.config_path, 'w') as f:
            json.dump({"source_folder": source_folder, "target_folder": target_folder}, f)

    def _configure_folders(self):
        from modules.emails.ui.config_dialog import EmailConfigDialog
        dialog = EmailConfigDialog(self.source_folder, self.target_folder, self)
        if dialog.exec_() == QDialog.Accepted:
            src, tgt = dialog.get_paths()
            self._save_config(src, tgt)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)

        # --- TOOLBAR ---
        toolbar = QHBoxLayout()
        self.btn_sync = QPushButton("🔄 Sync Source")
        self.btn_sync.setStyleSheet(
            "padding: 6px 12px; font-weight: bold; background-color: #0078D7; color: white; border-radius: 4px;")
        self.btn_sync.clicked.connect(self._run_sync)

        self.btn_move = QPushButton("➡️ Move Selected")
        self.btn_move.setStyleSheet(
            "padding: 6px 12px; font-weight: bold; background-color: #0C6B0C; color: white; border-radius: 4px;")
        self.btn_move.clicked.connect(self._run_move)

        self.btn_move_all = QPushButton("⏭️ Move All")
        self.btn_move_all.setStyleSheet(
            "padding: 6px 12px; font-weight: bold; background-color: #084D08; color: white; border-radius: 4px;")
        self.btn_move_all.clicked.connect(self._run_move_all)

        # ---> NEW: RESCAN TARGET BUTTON
        self.btn_rescan = QPushButton("🔁 Scan Target")
        self.btn_rescan.setStyleSheet(
            "padding: 6px 12px; font-weight: bold; background-color: #E1F0FF; color: #005A9E; border: 1px solid #99C9FF; border-radius: 4px;")
        self.btn_rescan.clicked.connect(self._run_target_rescan)

        self.btn_config = QPushButton("⚙️ Folders")
        self.btn_config.clicked.connect(self._configure_folders)

        self.lbl_status = QLabel("Ready.")
        self.lbl_status.setStyleSheet("color: #666; font-style: italic;")

        toolbar.addWidget(self.btn_sync)
        toolbar.addWidget(self.btn_move)
        toolbar.addWidget(self.btn_move_all)
        toolbar.addWidget(self.btn_rescan)
        toolbar.addWidget(self.btn_config)
        toolbar.addWidget(self.lbl_status)
        toolbar.addStretch()
        main_layout.addLayout(toolbar)

        # --- FILTERS ---
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search Subjects, Short Text, Free Text, or Revisions...")
        self.search_input.textChanged.connect(self._apply_filters)

        self.cb_ai = CheckableComboBox("AI")
        self.cb_company = CheckableComboBox("Company")
        self.cb_sender = CheckableComboBox("Sender")

        for cb in [self.cb_ai, self.cb_company, self.cb_sender]:
            cb.setMinimumWidth(150)
            cb.selectionChanged.connect(self._apply_filters)

        # Starred Toggle Filter
        self.btn_filter_star = QPushButton("⭐ Starred")
        self.btn_filter_star.setCheckable(True)
        self.btn_filter_star.setStyleSheet("""
                    QPushButton { padding: 4px 10px; border-radius: 4px; border: 1px solid #CCC; background: white; }
                    QPushButton:checked { background-color: #FFF4CE; border: 1px solid #F3C74C; font-weight: bold; }
                """)
        self.btn_filter_star.clicked.connect(self._apply_filters)

        # Followed AI Toggle Filter
        self.btn_filter_follow = QPushButton("👀 Followed AIs")
        self.btn_filter_follow.setCheckable(True)
        self.btn_filter_follow.setStyleSheet("""
                    QPushButton { padding: 4px 10px; border-radius: 4px; border: 1px solid #CCC; background: white; }
                    QPushButton:checked { background-color: #E6F4E6; border: 1px solid #A3DDA3; font-weight: bold; color: #0C6B0C; }
                """)
        self.btn_filter_follow.clicked.connect(self._apply_filters)

        filter_layout.addWidget(self.btn_filter_star)
        filter_layout.addWidget(self.btn_filter_follow)  # <--- ADD HERE
        filter_layout.addWidget(QLabel("🔍:"))
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.cb_ai)
        filter_layout.addWidget(self.cb_company)
        filter_layout.addWidget(self.cb_sender)
        main_layout.addLayout(filter_layout)

        # --- SPLITTER (TABLE / READING PANE) ---
        splitter = QSplitter(Qt.Vertical)

        # Table
        self.table = QTableView()
        self.model = EmailTableModel()
        self.proxy = EmailProxyModel()
        self.proxy.setSourceModel(self.model)
        self.table.setModel(self.proxy)

        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.ExtendedSelection) # <--- Change to ExtendedSelection
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet("QTableView { background: white; gridline-color: #EEE; }")

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 30)  # ⭐
        header.resizeSection(1, 85)  # Status
        header.resizeSection(2, 75)  # Local
        header.resizeSection(3, 120)  # Date
        header.resizeSection(4, 90)  # TDoc
        header.resizeSection(5, 60)  # Rev
        header.resizeSection(6, 70)  # AI
        header.resizeSection(7, 110)  # Company
        header.resizeSection(8, 140)  # Sender
        header.setSectionResizeMode(9, QHeaderView.Stretch)

        self.table.selectionModel().selectionChanged.connect(self._on_email_selected)
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

    def _run_sync(self):
        if not self.source_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Source Folder first.")
            return

        self._set_buttons_enabled(False)  # <--- Disables all buttons!
        self.btn_sync.setText("⏳ Syncing...")

        self.sync_thread = EmailSyncThread(self.source_folder, self.meeting_dir, self.ai_lookup, self.db, self.start_date, self.end_date)
        self.sync_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))

        self.sync_thread.progress_update.connect(
            lambda c, t: self.lbl_status.setText(f"Scanning Source: {c} / {t} items..."))

        self.sync_thread.finished.connect(self._on_sync_finished)
        self.sync_thread.start()

    def _on_sync_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)  # <--- Enables all buttons!
        self.btn_sync.setText("🔄 Sync Source")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _refresh_table(self):
        import sqlite3
        with sqlite3.connect(self.db.db_path) as conn:
            conn.row_factory = sqlite3.Row
            data = [dict(row) for row in conn.execute('SELECT * FROM emails ORDER BY date_received DESC').fetchall()]

        # Fetch both state sets
        starred_tdocs = self.db.get_starred_tdocs()
        followed_ais = self.db.get_followed_ais()
        self.model.update_data(data, starred_tdocs, followed_ais)

        def clean(val): return str(val).strip() if val else ""

        self.cb_ai.updateItems(sorted(set(clean(r.get("agenda_item")) for r in data)))
        self.cb_company.updateItems(sorted(set(clean(r.get("company")) for r in data)))
        self.cb_sender.updateItems(sorted(set(clean(r.get("sender_name")) for r in data)))

        self._apply_filters()

    def _apply_filters(self):
        self.proxy.set_filters(
            self.search_input.text(),
            self.cb_ai.getCheckedItems(),
            self.cb_company.getCheckedItems(),
            self.cb_sender.getCheckedItems(),
            self.btn_filter_star.isChecked(),
            self.btn_filter_follow.isChecked()  # <--- Pass the Follow filter state!
        )

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

        # Update the Star & Follow button text dynamically
        is_starred = self.current_tdoc_id in self.model.starred_tdocs
        self.btn_toggle_star.setText("❌ Unstar TDoc" if is_starred else "⭐ Star TDoc")

        is_followed = self.current_agenda_item in self.model.followed_ais
        ai_display = self.current_agenda_item or "Unknown AI"
        self.btn_toggle_follow.setText(f"❌ Unfollow AI: {ai_display}" if is_followed else f"👀 Follow AI: {ai_display}")

        # Format Reading Pane
        html = f"""
        <h2 style='color:#005A9E; margin-bottom: 2px;'>{row_data.get('subject', '')}</h2>
        <p style='color:#555; margin-top:0px;'><b>From:</b> {row_data.get('sender_name')} ({row_data.get('company')}) | <b>TDoc:</b> {row_data.get('tdoc_id')} | <b>Rev:</b> {row_data.get('revisions_mentioned')}</p>
        <hr>
        <h3 style='color:#D83B01;'>Short Text / Comments</h3>
        <p style='background-color:#FFF3E0; padding:10px; border-left: 4px solid #FFB74D;'>{row_data.get('short_text', '').replace(chr(10), '<br>')}</p>
        <h3>Free Text</h3>
        <p style='color:#333;'>{row_data.get('free_text', '').replace(chr(10), '<br>')}</p>
        """
        self.reading_pane.setHtml(html)

    # ---> NEW: Toggle Follow Method
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

        # Restore selection
        if selected_id:
            for i, row in enumerate(self.model._data):
                if row.get("id") == selected_id:
                    proxy_idx = self.proxy.mapFromSource(self.model.index(i, 0))
                    if proxy_idx.isValid():
                        self.table.selectRow(proxy_idx.row())
                    break

    def _toggle_current_star(self):
        if not getattr(self, "current_tdoc_id", ""): return

        # Save current selection to restore it after table refresh
        selected_indexes = self.table.selectionModel().selectedRows()
        selected_id = None
        if selected_indexes:
            source_idx = self.proxy.mapToSource(selected_indexes[0])
            selected_id = self.model.get_row_data(source_idx.row()).get("id")

        # Toggle DB
        is_starred = self.current_tdoc_id in self.model.starred_tdocs
        self.db.toggle_tdoc_star(self.current_tdoc_id, not is_starred)

        self._refresh_table()

        # Safely restore selection so your reading pane doesn't vanish!
        if selected_id:
            for i, row in enumerate(self.model._data):
                if row.get("id") == selected_id:
                    proxy_idx = self.proxy.mapFromSource(self.model.index(i, 0))
                    if proxy_idx.isValid():
                        self.table.selectRow(proxy_idx.row())
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

        # Extract (EntryID, AI) from selected rows
        items_to_move = []
        for idx in indexes:
            source_idx = self.proxy.mapToSource(idx)
            row_data = self.model.get_row_data(source_idx.row())

            # Only move items that are currently in the Source folder
            if row_data.get("outlook_location") == "Source":
                items_to_move.append((row_data.get("id"), row_data.get("agenda_item")))

        if not items_to_move:
            QMessageBox.information(self, "Notice", "Selected emails have already been moved to the Target.")
            return

        self._set_buttons_enabled(False)  # <--- Disables all buttons!
        self.btn_move.setText("⏳ Moving...")

        self.move_thread = EmailMoveThread(items_to_move, self.target_folder, self.db)
        self.move_thread.progress_update.connect(lambda c, t: self.lbl_status.setText(f"Moving {c}/{t}..."))
        self.move_thread.finished.connect(self._on_move_finished)
        self.move_thread.start()

    def _run_move_all(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return

        # Extract (EntryID, AI) from ALL rows that are currently in the Source folder
        items_to_move = []
        for row_data in self.model._data:
            if row_data.get("outlook_location") == "Source":
                items_to_move.append((row_data.get("id"), row_data.get("agenda_item")))

        if not items_to_move:
            QMessageBox.information(self, "Notice", "There are no emails in the Source folder to move.")
            return

        # Add a safety confirmation so you don't accidentally click it!
        reply = QMessageBox.question(self, 'Confirm Move All',
                                     f"Are you sure you want to move ALL {len(items_to_move)} source emails to the target folder?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No:
            return

        self._set_buttons_enabled(False)  # <--- Disables all buttons!
        self.btn_move_all.setText("⏳ Moving All...")

        self.move_thread = EmailMoveThread(items_to_move, self.target_folder, self.db)
        self.move_thread.progress_update.connect(lambda c, t: self.lbl_status.setText(f"Moving {c}/{t}..."))
        self.move_thread.finished.connect(self._on_move_all_finished)
        self.move_thread.start()

    def _on_move_all_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)  # <--- Enables all buttons!
        self.btn_move_all.setText("⏭️ Move All")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _on_move_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)  # <--- Enables all buttons!
        self.btn_move.setText("➡️ Move Selected")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _run_target_rescan(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return

        self._set_buttons_enabled(False)  # <--- Disables all buttons!
        self.btn_rescan.setText("⏳ Scanning...")

        self.rescan_thread = EmailTargetRescanThread(self.target_folder, self.meeting_dir, self.ai_lookup, self.db, self.start_date, self.end_date)
        self.rescan_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))

        self.rescan_thread.progress_update.connect(
            lambda c, t: self.lbl_status.setText(f"Scanning Target: {c} / {t} items..."))

        self.rescan_thread.finished.connect(self._on_rescan_finished)
        self.rescan_thread.start()

    def _on_rescan_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)  # <--- Enables all buttons!
        self.btn_rescan.setText("🔁 Scan Target")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _set_buttons_enabled(self, state: bool):
        """Helper to cleanly disable/enable all action buttons during background tasks."""
        self.btn_sync.setEnabled(state)
        self.btn_move.setEnabled(state)
        self.btn_move_all.setEnabled(state)
        self.btn_rescan.setEnabled(state)