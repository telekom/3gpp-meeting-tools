# --- File: modules/emails/ui/email_window.py ---
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
    def __init__(self, meeting_dir: Path, ai_lookup: dict):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.ai_lookup = ai_lookup

        self.db = EmailDatabase(self.meeting_dir / "Agenda" / "emails.db")
        self.config_path = self.meeting_dir / "Agenda" / "email_config.json"

        config_data = self._load_config()
        self.source_folder = config_data.get("source_folder", "")
        self.target_folder = config_data.get("target_folder", "")

        self.setWindowTitle("📧 eMeeting Email Manager")

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
        toolbar.addWidget(self.btn_rescan)  # Add it to the layout here!
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

        # ---> FIX: Updated column widths for the 9 new columns
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 85)  # Status
        header.resizeSection(1, 75)  # Local Disk
        header.resizeSection(2, 120)  # Date
        header.resizeSection(3, 90)  # TDoc
        header.resizeSection(4, 60)  # Rev
        header.resizeSection(5, 70)  # AI
        header.resizeSection(6, 110)  # Company
        header.resizeSection(7, 140)  # Sender
        header.setSectionResizeMode(8, QHeaderView.Stretch)  # Short Text gets remaining space

        self.table.selectionModel().selectionChanged.connect(self._on_email_selected)
        splitter.addWidget(self.table)

        # Reading Pane
        pane_widget = QWidget()
        pane_layout = QVBoxLayout(pane_widget)
        pane_layout.setContentsMargins(0, 10, 0, 0)

        self.btn_open_msg = QPushButton("📄 Open .msg File in Outlook")
        self.btn_open_msg.setEnabled(False)
        self.btn_open_msg.clicked.connect(self._open_current_msg)

        self.reading_pane = QTextBrowser()
        self.reading_pane.setStyleSheet("background: white; border: 1px solid #CCC; border-radius: 4px; padding: 10px;")

        pane_layout.addWidget(self.btn_open_msg)
        pane_layout.addWidget(self.reading_pane)
        splitter.addWidget(pane_widget)

        splitter.setSizes([400, 300])
        main_layout.addWidget(splitter)

    def _run_sync(self):
        if not self.source_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Source Folder first.")
            return

        self.btn_sync.setEnabled(False)
        self.btn_sync.setText("⏳ Syncing...")

        self.sync_thread = EmailSyncThread(self.source_folder, self.meeting_dir, self.ai_lookup, self.db)
        self.sync_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))
        self.sync_thread.finished.connect(self._on_sync_finished)
        self.sync_thread.start()

    def _on_sync_finished(self, success: bool, msg: str):
        self.btn_sync.setEnabled(True)
        self.btn_sync.setText("🔄 Sync from Outlook")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _refresh_table(self):
        # Fetch all emails from the DB (TDoc isolation guarantees this isn't too heavy)
        import sqlite3
        with sqlite3.connect(self.db.db_path) as conn:
            conn.row_factory = sqlite3.Row
            data = [dict(row) for row in conn.execute('SELECT * FROM emails ORDER BY date_received DESC').fetchall()]

        self.model.update_data(data)

        # Populate Comboboxes
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
            self.cb_sender.getCheckedItems()
        )

    def _on_email_selected(self, selected, deselected):
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            self.reading_pane.clear()
            self.btn_open_msg.setEnabled(False)
            return

        source_idx = self.proxy.mapToSource(indexes[0])
        row_data = self.model.get_row_data(source_idx.row())

        self.current_msg_path = row_data.get("msg_path", "")
        self.btn_open_msg.setEnabled(bool(self.current_msg_path))

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

        self.btn_move.setEnabled(False)
        self.btn_sync.setEnabled(False)
        self.btn_move.setText("⏳ Moving...")

        self.move_thread = EmailMoveThread(items_to_move, self.target_folder, self.db)
        self.move_thread.progress_update.connect(lambda c, t: self.lbl_status.setText(f"Moving {c}/{t}..."))
        self.move_thread.finished.connect(self._on_move_finished)
        self.move_thread.start()

    def _on_move_finished(self, success: bool, msg: str):
        self.btn_move.setEnabled(True)
        self.btn_sync.setEnabled(True)
        self.btn_move.setText("➡️ Move Selected to Target")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _run_target_rescan(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return

        self.btn_sync.setEnabled(False)
        self.btn_move.setEnabled(False)
        self.btn_rescan.setEnabled(False)
        self.btn_rescan.setText("⏳ Scanning...")

        self.rescan_thread = EmailTargetRescanThread(self.target_folder, self.meeting_dir, self.ai_lookup, self.db)
        self.rescan_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))
        self.rescan_thread.progress_update.connect(
            lambda c, t: self.lbl_status.setText(f"Scanning Target... ({c} items)"))
        self.rescan_thread.finished.connect(self._on_rescan_finished)
        self.rescan_thread.start()

    def _on_rescan_finished(self, success: bool, msg: str):
        self.btn_sync.setEnabled(True)
        self.btn_move.setEnabled(True)
        self.btn_rescan.setEnabled(True)
        self.btn_rescan.setText("🔁 Scan Target")
        self.lbl_status.setText(msg)
        self._refresh_table()