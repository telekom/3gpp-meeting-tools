# --- File: src/modules/emails/ui/general_email_dialog.py ---
import html
import json
import logging
import re
from pathlib import Path
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QDate
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableView,
    QHeaderView, QTextBrowser, QCheckBox, QAbstractItemView, QFrame,
    QLineEdit, QSpinBox, QMessageBox, QDateEdit, QTableWidget, QTableWidgetItem,
    QColorDialog, QMenu, QApplication
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
        self.resize(720, 440)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; }")

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"<b>Configured Outlook Folders for {self.wg}:</b>"))
        desc = QLabel("Configure folders to scan, assign display tags, and pick custom tag colors.")
        desc.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(desc)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Outlook Folder Path", "Tag", "Tag Color"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
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
        btn_save = QPushButton("💾 Save & Close")
        btn_save.setStyleSheet("font-weight: bold; background-color: #0078D7; color: white; padding: 6px 15px; border-radius: 4px;")
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
            color = item.get("color", "#0078D7")
            self._set_color_widget(row, color)

    def _set_color_widget(self, row: int, hex_color: str):
        btn_color = QPushButton(hex_color)
        btn_color.setStyleSheet(f"background-color: {hex_color}; color: white; font-weight: bold; border-radius: 3px; padding: 3px 8px;")
        btn_color.clicked.connect(lambda _, r=row, b=btn_color: self._pick_color(r, b))
        self.table.setCellWidget(row, 2, btn_color)

    def _pick_color(self, row: int, btn: QPushButton):
        curr = QColor(btn.text()) if QColor(btn.text()).isValid() else QColor("#0078D7")
        col = QColorDialog.getColor(curr, self, "Pick Tag Color")
        if col.isValid():
            hex_code = col.name().upper()
            btn.setText(hex_code)
            lum = (col.red() * 0.299 + col.green() * 0.587 + col.blue() * 0.114)
            fg = "#FFFFFF" if lum < 150 else "#000000"
            btn.setStyleSheet(f"background-color: {hex_code}; color: {fg}; font-weight: bold; border-radius: 3px; padding: 3px 8px;")

    def _add_folder(self):
        dialog = OutlookFolderPickerDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            path = dialog.get_selected_path()
            if path:
                r = self.table.rowCount()
                self.table.insertRow(r)
                self.table.setItem(r, 0, QTableWidgetItem(path))
                self.table.setItem(r, 1, QTableWidgetItem("WG"))
                self._set_color_widget(r, "#0078D7")

    def _remove_selected(self):
        for idx in sorted([idx.row() for idx in self.table.selectionModel().selectedRows()], reverse=True):
            self.table.removeRow(idx)

    def _save_and_close(self):
        folders = []
        for r in range(self.table.rowCount()):
            p = self.table.item(r, 0).text().strip()
            t = self.table.item(r, 1).text().strip()
            btn = self.table.cellWidget(r, 2)
            c = btn.text().strip() if btn else "#0078D7"
            if p:
                folders.append({"folder_path": p, "tag": t or "WG", "color": c})
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

    def __init__(self, target_tdoc: str, family_tdocs: list, db_path: Path, wg: str = "SA2", parent=None):
        super().__init__(parent)
        self.target_tdoc = target_tdoc.upper()
        self.family_tdocs = [t.upper() for t in family_tdocs]
        if self.target_tdoc not in self.family_tdocs:
            self.family_tdocs.insert(0, self.target_tdoc)

        self.wg = wg
        self.db = GeneralEmailDatabase(db_path)
        self.tag_colors = {f.get("tag", "").upper(): f.get("color", "#0078D7") for f in load_wg_email_config(self.wg)}

        # Modeless dialog setup: does not block parent UI
        self.setWindowModality(Qt.NonModal)
        self.setWindowTitle(f"📧 Related Emails: {self.target_tdoc}")
        self.resize(1100, 700)
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
        top_card.setStyleSheet("QFrame { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 6px; padding: 6px; }")
        top_layout = QHBoxLayout(top_card)

        chips_text = " ➔ ".join([f"<b>{t}</b>" if t == self.target_tdoc else t for t in self.family_tdocs])
        lbl_chips = QLabel(f"<b>Document Family:</b> {chips_text}")
        top_layout.addWidget(lbl_chips)
        top_layout.addStretch()

        self.chk_family = QCheckBox("Show Family Revisions")
        self.chk_family.setChecked(True)
        self.chk_family.toggled.connect(self._load_emails)
        top_layout.addWidget(self.chk_family)

        self.chk_include_quoted = QCheckBox("Include Quoted Matches")
        self.chk_include_quoted.setToolTip("Uncheck to hide emails where the TDoc was only cited in previous email quotations.")
        self.chk_include_quoted.setChecked(True)
        self.chk_include_quoted.toggled.connect(self._load_emails)
        top_layout.addWidget(self.chk_include_quoted)

        self.chk_show_ignored = QCheckBox("Show Ignored")
        self.chk_show_ignored.setChecked(False)
        self.chk_show_ignored.toggled.connect(self._load_emails)
        top_layout.addWidget(self.chk_show_ignored)

        self.lbl_counts = QLabel("0 Emails")
        self.lbl_counts.setStyleSheet("font-weight: bold; color: #0078D7; margin-left: 10px;")
        top_layout.addWidget(self.lbl_counts)
        layout.addWidget(top_card)

        # Action Toolbar
        act_row = QHBoxLayout()
        self.btn_mark_read = QPushButton("✔️ Mark Read")
        self.btn_mark_read.clicked.connect(lambda: self._set_selected_read(True))

        self.btn_mark_unread = QPushButton("✉️ Mark Unread")
        self.btn_mark_unread.clicked.connect(lambda: self._set_selected_read(False))

        self.btn_ignore = QPushButton("🚫 Ignore")
        self.btn_ignore.setToolTip("Exclude selected email(s) from counts without deleting them.")
        self.btn_ignore.clicked.connect(self._toggle_selected_ignore)

        self.btn_delete = QPushButton("🗑️ Delete")
        self.btn_delete.setToolTip("Permanently delete selected email records from database.")
        self.btn_delete.clicked.connect(self._delete_selected)

        self.btn_mark_all_read = QPushButton("✔️ Mark All Read")
        self.btn_mark_all_read.clicked.connect(self._mark_all_family_read)

        self.btn_open_outlook = QPushButton("🚀 Open in Outlook")
        self.btn_open_outlook.setStyleSheet("font-weight: bold; background-color: #0078D7; color: white; border-radius: 4px; padding: 4px 12px;")
        self.btn_open_outlook.clicked.connect(self._open_in_outlook)

        act_row.addWidget(self.btn_mark_read)
        act_row.addWidget(self.btn_mark_unread)
        act_row.addWidget(self.btn_ignore)
        act_row.addWidget(self.btn_delete)
        act_row.addSpacing(10)
        act_row.addWidget(self.btn_mark_all_read)
        act_row.addStretch()
        act_row.addWidget(self.btn_open_outlook)
        layout.addLayout(act_row)

        # Main Table (ExtendedSelection enabled)
        self.table = QTableView()
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("QTableView { background: white; border: 1px solid #CCC; } QHeaderView::section { background: #F0F0F0; font-weight: bold; }")

        self.model = QStandardItemModel()
        self.headers = ["Status", "Tag", "Match In", "Rev", "Sender", "Company", "Date", "Subject"]
        self.model.setHorizontalHeaderLabels(self.headers)
        self.table.setModel(self.model)
        self.table.selectionModel().selectionChanged.connect(self._on_email_selection_changed)
        self.table.doubleClicked.connect(self._open_in_outlook)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)

        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.resizeSection(0, 75)
        hdr.resizeSection(1, 80)
        hdr.resizeSection(2, 85)
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
        raw_emails = self.db.get_emails_for_tdocs(query_set, show_ignored=self.chk_show_ignored.isChecked())

        # Filter out quoted matches if unchecked
        if not self.chk_include_quoted.isChecked():
            self.emails = [e for e in raw_emails if e.get("match_location") != "Quoted"]
        else:
            self.emails = raw_emails

        self.model.removeRows(0, self.model.rowCount())
        unread_count = 0

        for r_idx, e in enumerate(self.emails):
            is_read = (int(e.get("is_read", 0)) == 1)
            is_ignored = (int(e.get("is_ignored", 0)) == 1)

            if not is_read and not is_ignored:
                unread_count += 1

            if is_ignored:
                status_text = "🚫 Ignored"
            elif is_read:
                status_text = "⚪ Read"
            else:
                status_text = "🔵 Unread"

            status_item = QStandardItem(status_text)
            status_item.setTextAlignment(Qt.AlignCenter)
            if is_ignored:
                status_item.setForeground(QColor("#888888"))
            elif not is_read:
                status_item.setForeground(QColor("#0078D7"))
                status_item.setFont(QFont("Segoe UI", weight=QFont.Bold))

            tag_str = e.get("folder_tag", "WG")
            tag_item = QStandardItem(f"[{tag_str}]")
            tag_item.setTextAlignment(Qt.AlignCenter)

            color_hex = self.tag_colors.get(tag_str.upper(), "#0078D7")
            tag_item.setForeground(QColor(color_hex))
            tag_item.setFont(QFont("Segoe UI", weight=QFont.Bold))

            loc_str = e.get("match_location", "Body")
            loc_item = QStandardItem(loc_str)
            loc_item.setTextAlignment(Qt.AlignCenter)
            if loc_str == "Quoted":
                loc_item.setForeground(QColor("#888888"))
                loc_item.setToolTip("TDoc was cited inside previous email quotations below.")
            elif loc_str == "Subject":
                loc_item.setForeground(QColor("#0C6B0C"))
                loc_item.setFont(QFont("Segoe UI", weight=QFont.Bold))

            rev_item = QStandardItem(e.get("rev_matched") or "-")
            rev_item.setTextAlignment(Qt.AlignCenter)

            sender_item = QStandardItem(e.get("sender_name", ""))
            company_item = QStandardItem(e.get("company", ""))
            date_item = QStandardItem(str(e.get("date_received", ""))[:16])
            subj_item = QStandardItem(e.get("subject", ""))

            if is_ignored:
                for item in [loc_item, rev_item, sender_item, company_item, date_item, subj_item]:
                    item.setForeground(QColor("#999999"))

            row = [status_item, tag_item, loc_item, rev_item, sender_item, company_item, date_item, subj_item]
            self.model.appendRow(row)

        total = len(self.emails)
        self.lbl_counts.setText(f"{total} Total ({unread_count} Unread)")
        if self.emails:
            self.table.selectRow(0)

    def _get_selected_indices(self) -> list:
        return [idx.row() for idx in self.table.selectionModel().selectedRows() if idx.isValid()]

    def _on_email_selection_changed(self):
        rows = self._get_selected_indices()
        if not rows:
            self.reading_pane.clear()
            return

        primary_idx = rows[0]
        e = self.emails[primary_idx]

        raw_body = e.get("body_text", "")
        # Normalize and clean excessive blank lines
        cleaned = raw_body.replace("\x00", "").replace("\r\n", "\n").replace("\r", "\n")
        cleaned = re.sub(r"\n{3,}", "\n\n", cleaned.strip())

        # Safe HTML conversion
        body_escaped = html.escape(cleaned)

        # Highlight TDoc references
        pattern = re.compile(rf"\b({'|'.join(re.escape(t) for t in self.family_tdocs)})\b", re.IGNORECASE)
        body_hl = pattern.sub(r"<span style='background-color: #FFF176; color: #000; font-weight: bold;'>\1</span>", body_escaped)
        body_html = body_hl.replace("\n", "<br>")

        # Match Context Excerpt
        match_excerpt = ""
        found_matches = list(pattern.finditer(cleaned))
        if found_matches:
            m = found_matches[0]
            start = max(0, m.start() - 60)
            end = min(len(cleaned), m.end() + 60)
            snippet = html.escape(cleaned[start:end]).replace("\n", " ")
            snippet_hl = pattern.sub(r"<span style='background-color: #FFF176; color: #000; font-weight: bold;'>\1</span>", snippet)
            loc_label = e.get("match_location", "Body")
            match_excerpt = f"""
            <div style='background-color: #FFF8E1; border: 1px solid #FFE082; border-radius: 4px; padding: 6px; margin: 6px 0; font-size: 11px;'>
                <b>💡 Match Found ({loc_label}):</b> ...{snippet_hl}...
            </div>
            """

        ignored_banner = "<p style='color: #D83B01; font-weight: bold; margin: 4px 0;'>⚠️ This email is currently IGNORED from TDoc counts.</p>" if e.get("is_ignored") else ""

        full_html = f"""
        {ignored_banner}
        <h3 style='margin: 0 0 4px 0; color: #005A9E;'>{html.escape(e.get('subject', ''))}</h3>
        <p style='color: #555; margin: 0 0 6px 0; font-size: 12px;'>
            <b>From:</b> {html.escape(e.get('sender_name', ''))} &lt;{html.escape(e.get('sender_email', ''))}&gt; ({html.escape(e.get('company', ''))}) | 
            <b>Date:</b> {html.escape(e.get('date_received', ''))}
        </p>
        {match_excerpt}
        <hr style='border: 0; border-top: 1px solid #E0E0E0; margin: 6px 0;'>
        <div style='font-family: Segoe UI, sans-serif; font-size: 12px; color: #222; line-height: 1.4;'>{body_html}</div>
        """
        self.reading_pane.setHtml(full_html)

        if len(rows) == 1 and int(e.get("is_read", 0)) == 0 and int(e.get("is_ignored", 0)) == 0:
            self.auto_read_timer.start()

    def _mark_current_read(self):
        rows = self._get_selected_indices()
        if len(rows) == 1:
            idx = rows[0]
            e = self.emails[idx]
            if int(e.get("is_read", 0)) == 0 and int(e.get("is_ignored", 0)) == 0:
                self.db.set_emails_read_status([e["id"]], True)
                e["is_read"] = 1
                self.model.item(idx, 0).setText("⚪ Read")
                self.model.item(idx, 0).setForeground(QColor("#333"))
                self.model.item(idx, 0).setFont(QFont("Segoe UI"))
                self.data_changed.emit()
                self._update_unread_count_label()

    def _set_selected_read(self, is_read: bool):
        rows = self._get_selected_indices()
        if not rows:
            return
        email_ids = [self.emails[r]["id"] for r in rows]
        self.db.set_emails_read_status(email_ids, is_read)
        for r in rows:
            self.emails[r]["is_read"] = 1 if is_read else 0
            if not self.emails[r].get("is_ignored"):
                item = self.model.item(r, 0)
                item.setText("⚪ Read" if is_read else "🔵 Unread")
                item.setForeground(QColor("#333") if is_read else QColor("#0078D7"))
                item.setFont(QFont("Segoe UI", weight=QFont.Normal if is_read else QFont.Bold))
        self.data_changed.emit()
        self._update_unread_count_label()

    def _toggle_selected_ignore(self):
        rows = self._get_selected_indices()
        if not rows:
            return
        any_active = any(int(self.emails[r].get("is_ignored", 0)) == 0 for r in rows)
        target_state = True if any_active else False

        email_ids = [self.emails[r]["id"] for r in rows]
        self.db.set_emails_ignored_status(email_ids, target_state)
        self.data_changed.emit()
        self._load_emails()

    def _delete_selected(self):
        rows = self._get_selected_indices()
        if not rows:
            return
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to permanently delete {len(rows)} email record(s)?\n\n"
            "Note: Re-syncing Outlook will re-import them unless they are Ignored instead.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        email_ids = [self.emails[r]["id"] for r in rows]
        self.db.delete_emails(email_ids)
        self.data_changed.emit()
        self._load_emails()

    def _mark_all_family_read(self):
        query_set = set(self.family_tdocs) if self.chk_family.isChecked() else {self.target_tdoc}
        self.db.set_tdocs_read_status(query_set, True)
        self.data_changed.emit()
        self._load_emails()

    def _update_unread_count_label(self):
        unread = sum(1 for e in self.emails if int(e.get("is_read", 0)) == 0 and int(e.get("is_ignored", 0)) == 0)
        self.lbl_counts.setText(f"{len(self.emails)} Total ({unread} Unread)")

    def _open_in_outlook(self):
        rows = self._get_selected_indices()
        if not rows:
            return
        e = self.emails[rows[0]]
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
                                f"Could not open email in Outlook:\n{err}\n\n(It may have been moved across PSTs or deleted).")

    def _show_context_menu(self, pos):
        rows = self._get_selected_indices()
        if not rows:
            return

        menu = QMenu(self)
        act_open = menu.addAction("🚀 Open in Outlook")
        act_open.triggered.connect(self._open_in_outlook)
        menu.addSeparator()

        act_read = menu.addAction("✔️ Mark as Read")
        act_read.triggered.connect(lambda: self._set_selected_read(True))

        act_unread = menu.addAction("✉️ Mark as Unread")
        act_unread.triggered.connect(lambda: self._set_selected_read(False))

        menu.addSeparator()
        any_active = any(int(self.emails[r].get("is_ignored", 0)) == 0 for r in rows)
        act_ignore = menu.addAction("🚫 Ignore Email(s)" if any_active else "↩️ Un-ignore Email(s)")
        act_ignore.triggered.connect(self._toggle_selected_ignore)

        act_del = menu.addAction("🗑️ Delete Email(s)")
        act_del.triggered.connect(self._delete_selected)

        menu.exec_(self.table.viewport().mapToGlobal(pos))