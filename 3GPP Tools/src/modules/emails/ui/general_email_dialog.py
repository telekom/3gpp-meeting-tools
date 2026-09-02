# --- File: src/modules/emails/ui/general_email_dialog.py ---
import json
import logging
import re
from pathlib import Path
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QDate
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableView,
    QHeaderView, QTextBrowser, QCheckBox, QAbstractItemView, QFrame,
    QLineEdit, QSpinBox, QMessageBox, QDateEdit, QTableWidget, QTableWidgetItem
)
from PyQt5.QtGui import QColor, QFont, QStandardItemModel, QStandardItem

from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.general_email_sync import GeneralEmailSyncThread
from modules.emails.ui.config_dialog import OutlookFolderPickerDialog

CONFIG_PATH = Path(__file__).resolve().parents[3] / "emails_config.json"


def load_wg_email_config(wg: str) -> list:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get(wg.upper(), [])
        except Exception as e:
            logging.error(f"Failed to load emails_config.json: {e}")
    return []


def save_wg_email_config(wg: str, folder_list: list):
    data = {}
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
    data[wg.upper()] = folder_list
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        logging.error(f"Failed to save emails_config.json: {e}")


class GeneralEmailFoldersDialog(QDialog):
    def __init__(self, wg: str, parent=None):
        super().__init__(parent)
        self.wg = wg.upper()
        self.setWindowTitle(f"⚙️ Outlook Folders Configuration: {self.wg}")
        self.resize(650, 420)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; }")

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"<b>Configured Outlook Folders for {self.wg}:</b>"))
        desc = QLabel(
            "Add distribution lists or subfolders to scan for this Working Group. Tag them for quick identification.")
        desc.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(desc)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Outlook Folder Path", "Tag (e.g. WG, Disc, Offline)"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        layout.addWidget(self.table)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ Add Folder via Outlook...")
        btn_add.clicked.connect(self._add_folder)
        btn_del = QPushButton("🗑️ Remove Selected")
        btn_del.clicked.connect(self._remove_selected)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        bottom_row = QHBoxLayout()
        btn_save = QPushButton("💾 Save && Close")
        btn_save.setStyleSheet(
            "font-weight: bold; background-color: #0078D7; color: white; padding: 6px 15px; border-radius: 4px;")
        btn_save.clicked.connect(self._save_and_close)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        bottom_row.addStretch()
        bottom_row.addWidget(btn_cancel)
        bottom_row.addWidget(btn_save)
        layout.addLayout(bottom_row)

        self._populate_table()

    def _populate_table(self):
        folders = load_wg_email_config(self.wg)
        self.table.setRowCount(len(folders))
        for row, item in enumerate(folders):
            self.table.setItem(row, 0, QTableWidgetItem(item.get("folder_path", "")))
            self.table.setItem(row, 1, QTableWidgetItem(item.get("tag", "WG")))

    def _add_folder(self):
        dialog = OutlookFolderPickerDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            path = dialog.get_selected_path()
            if path:
                r = self.table.rowCount()
                self.table.insertRow(r)
                self.table.setItem(r, 0, QTableWidgetItem(path))
                self.table.setItem(r, 1, QTableWidgetItem("WG"))

    def _remove_selected(self):
        for idx in sorted([idx.row() for idx in self.table.selectionModel().selectedRows()], reverse=True):
            self.table.removeRow(idx)

    def _save_and_close(self):
        folders = []
        for r in range(self.table.rowCount()):
            p = self.table.item(r, 0).text().strip()
            t = self.table.item(r, 1).text().strip()
            if p:
                folders.append({"folder_path": p, "tag": t or "WG"})
        save_wg_email_config(self.wg, folders)
        self.accept()


class GeneralEmailSyncDialog(QDialog):
    def __init__(self, wg: str, start_date: str, end_date: str, parent=None):
        super().__init__(parent)
        self.wg = wg.upper()
        self.setWindowTitle(f"🔄 Sync Emails: {self.wg}")
        self.resize(450, 260)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; }")

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"<b>Sync Outlook Emails for {self.wg}</b>"))

        folders = load_wg_email_config(self.wg)
        lbl_folders = QLabel(
            f"Configured Folders: <b>{len(folders)}</b> ({', '.join(f.get('tag', '') for f in folders) or 'None'})")
        layout.addWidget(lbl_folders)

        form = QVBoxLayout()
        dt_layout = QHBoxLayout()
        dt_layout.addWidget(QLabel("Date Range:"))
        self.dt_start = QDateEdit()
        self.dt_start.setCalendarPopup(True)
        self.dt_end = QDateEdit()
        self.dt_end.setCalendarPopup(True)

        if start_date:
            self.dt_start.setDate(QDate.fromString(start_date, Qt.ISODate))
        else:
            self.dt_start.setDate(QDate.currentDate().addDays(-14))

        if end_date:
            self.dt_end.setDate(QDate.fromString(end_date, Qt.ISODate))
        else:
            self.dt_end.setDate(QDate.currentDate().addDays(7))

        dt_layout.addWidget(self.dt_start)
        dt_layout.addWidget(QLabel("-"))
        dt_layout.addWidget(self.dt_end)
        form.addLayout(dt_layout)

        buf_layout = QHBoxLayout()
        buf_layout.addWidget(QLabel("Buffer Window (Days):"))
        self.spin_buf = QSpinBox()
        self.spin_buf.setRange(0, 30)
        self.spin_buf.setValue(3)
        buf_layout.addWidget(self.spin_buf)
        buf_layout.addStretch()
        form.addLayout(buf_layout)
        layout.addLayout(form)

        btn_row = QHBoxLayout()
        btn_sync = QPushButton("🚀 Start Sync")
        btn_sync.setStyleSheet(
            "font-weight: bold; background-color: #0078D7; color: white; padding: 6px 15px; border-radius: 4px;")
        btn_sync.clicked.connect(self.accept)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addStretch()
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_sync)
        layout.addLayout(btn_row)

    def get_params(self):
        return {
            "start_date": self.dt_start.date().toString(Qt.ISODate),
            "end_date": self.dt_end.date().toString(Qt.ISODate),
            "buffer": self.spin_buf.value()
        }


class TDocEmailsDialog(QDialog):
    data_changed = pyqtSignal()

    def __init__(self, target_tdoc: str, family_tdocs: list, db_path: Path, parent=None):
        super().__init__(parent)
        self.target_tdoc = target_tdoc.upper()
        self.family_tdocs = [t.upper() for t in family_tdocs]
        if self.target_tdoc not in self.family_tdocs:
            self.family_tdocs.insert(0, self.target_tdoc)

        self.db = GeneralEmailDatabase(db_path)
        self.setWindowTitle(f"📧 Related Emails: {self.target_tdoc}")
        self.resize(1050, 680)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; }")

        self.auto_read_timer = QTimer(self)
        self.auto_read_timer.setSingleShot(True)
        self.auto_read_timer.setInterval(800)
        self.auto_read_timer.timeout.connect(self._mark_current_read)

        self._init_ui()
        self._load_emails()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # Top Bar: Lineage & Breadcrumbs
        top_card = QFrame()
        top_card.setStyleSheet(
            "QFrame { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 6px; padding: 6px; }")
        top_layout = QHBoxLayout(top_card)

        chips_text = " ➔ ".join([f"<b>{t}</b>" if t == self.target_tdoc else t for t in self.family_tdocs])
        lbl_chips = QLabel(f"<b>Document Family:</b> {chips_text}")
        top_layout.addWidget(lbl_chips)
        top_layout.addStretch()

        self.chk_family = QCheckBox("Show Family Revisions")
        self.chk_family.setChecked(True)
        self.chk_family.toggled.connect(self._load_emails)
        top_layout.addWidget(self.chk_family)

        self.lbl_counts = QLabel("0 Emails")
        self.lbl_counts.setStyleSheet("font-weight: bold; color: #0078D7; margin-left: 10px;")
        top_layout.addWidget(self.lbl_counts)
        layout.addWidget(top_card)

        # Action Toolbar
        act_row = QHBoxLayout()
        self.btn_mark_all = QPushButton("✔️ Mark All Read")
        self.btn_mark_all.clicked.connect(self._mark_all_family_read)
        self.btn_toggle_unread = QPushButton("✉️ Mark Selected Unread")
        self.btn_toggle_unread.clicked.connect(self._mark_selected_unread)
        self.btn_open_outlook = QPushButton("🚀 Open in Outlook")
        self.btn_open_outlook.setStyleSheet(
            "font-weight: bold; background-color: #0078D7; color: white; border-radius: 4px; padding: 4px 12px;")
        self.btn_open_outlook.clicked.connect(self._open_in_outlook)

        act_row.addWidget(self.btn_mark_all)
        act_row.addWidget(self.btn_toggle_unread)
        act_row.addWidget(self.btn_open_outlook)
        act_row.addStretch()
        layout.addLayout(act_row)

        # Main Table
        self.table = QTableView()
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(
            "QTableView { background: white; border: 1px solid #CCC; } QHeaderView::section { background: #F0F0F0; font-weight: bold; }")

        self.model = QStandardItemModel()
        self.headers = ["Status", "Tag", "Match In", "Rev", "Sender", "Company", "Date", "Subject"]
        self.model.setHorizontalHeaderLabels(self.headers)
        self.table.setModel(self.model)
        self.table.selectionModel().selectionChanged.connect(self._on_email_selected)
        self.table.doubleClicked.connect(self._open_in_outlook)

        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.resizeSection(0, 60)
        hdr.resizeSection(1, 65)
        hdr.resizeSection(2, 75)
        hdr.resizeSection(3, 75)
        hdr.resizeSection(4, 130)
        hdr.resizeSection(5, 110)
        hdr.resizeSection(6, 120)
        hdr.setSectionResizeMode(7, QHeaderView.Stretch)
        layout.addWidget(self.table, stretch=2)

        # Reading Pane
        self.reading_pane = QTextBrowser()
        self.reading_pane.setStyleSheet("background: white; border: 1px solid #CCC; padding: 8px;")
        layout.addWidget(self.reading_pane, stretch=1)

    def _load_emails(self):
        query_set = set(self.family_tdocs) if self.chk_family.isChecked() else {self.target_tdoc}
        self.emails = self.db.get_emails_for_tdocs(query_set)

        self.model.removeRows(0, self.model.rowCount())
        unread_count = 0

        for r_idx, e in enumerate(self.emails):
            is_read = bool(e.get("is_read", 0))
            if not is_read:
                unread_count += 1

            status_item = QStandardItem("⚪ Read" if is_read else "🔵 Unread")
            status_item.setTextAlignment(Qt.AlignCenter)
            if not is_read:
                status_item.setForeground(QColor("#0078D7"))
                status_item.setFont(QFont("Segoe UI", weight=QFont.Bold))

            tag_item = QStandardItem(f"[{e.get('folder_tag', 'WG')}]")
            tag_item.setTextAlignment(Qt.AlignCenter)
            loc_item = QStandardItem(e.get("match_location", "Body"))
            loc_item.setTextAlignment(Qt.AlignCenter)
            rev_item = QStandardItem(e.get("rev_matched") or "-")
            rev_item.setTextAlignment(Qt.AlignCenter)
            sender_item = QStandardItem(e.get("sender_name", ""))
            company_item = QStandardItem(e.get("company", ""))
            date_item = QStandardItem(str(e.get("date_received", ""))[:16])
            subj_item = QStandardItem(e.get("subject", ""))

            row = [status_item, tag_item, loc_item, rev_item, sender_item, company_item, date_item, subj_item]
            self.model.appendRow(row)

        total = len(self.emails)
        self.lbl_counts.setText(f"{total} Total ({unread_count} Unread)")
        if self.emails:
            self.table.selectRow(0)

    def _on_email_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            self.reading_pane.clear()
            return
        idx = rows[0].row()
        e = self.emails[idx]

        # Highlight matches in yellow
        body = e.get("body_text", "").replace("<", "&lt;").replace(">", "&gt;")
        pattern = re.compile(rf"\b({'|'.join(re.escape(t) for t in self.family_tdocs)})\b", re.IGNORECASE)
        body_hl = pattern.sub(r"<span style='background-color: #FFF176; font-weight: bold;'>\1</span>", body)

        html = f"""
        <h3 style='margin:0; color:#005A9E;'>{e.get('subject', '')}</h3>
        <p style='color:#555; margin:4px 0;'><b>From:</b> {e.get('sender_name')} &lt;{e.get('sender_email')}&gt; ({e.get('company')}) | <b>Date:</b> {e.get('date_received')}</p>
        <hr>
        <div style='white-space: pre-wrap; font-family: Segoe UI, sans-serif;'>{body_hl}</div>
        """
        self.reading_pane.setHtml(html)

        # Trigger auto-read
        if not e.get("is_read", 0):
            self.auto_read_timer.start()

    def _mark_current_read(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return
        idx = rows[0].row()
        e = self.emails[idx]
        if not e.get("is_read", 0):
            self.db.set_email_read_status(e["id"], True)
            e["is_read"] = 1
            self.model.item(idx, 0).setText("⚪ Read")
            self.model.item(idx, 0).setForeground(QColor("#333"))
            self.model.item(idx, 0).setFont(QFont("Segoe UI"))
            self.data_changed.emit()

    def _mark_all_family_read(self):
        query_set = set(self.family_tdocs) if self.chk_family.isChecked() else {self.target_tdoc}
        self.db.set_tdocs_read_status(query_set, True)
        self.data_changed.emit()
        self._load_emails()

    def _mark_selected_unread(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return
        idx = rows[0].row()
        e = self.emails[idx]
        self.db.set_email_read_status(e["id"], False)
        e["is_read"] = 0
        self.data_changed.emit()
        self._load_emails()

    def _open_in_outlook(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return
        idx = rows[0].row()
        e = self.emails[idx]
        entry_id = e.get("id")

        try:
            namespace = OutlookClient.get_namespace()
            if namespace and entry_id:
                mail = namespace.GetItemFromID(entry_id)
                mail.Display(False)
                self._mark_current_read()
            else:
                QMessageBox.warning(self, "Outlook Error", "Could not connect to Microsoft Outlook MAPI.")
        except Exception as err:
            QMessageBox.warning(self, "Item Not Found",
                                f"Could not open email in Outlook:\n{err}\n\n(It may have been permanently deleted or moved to a different PST store).")