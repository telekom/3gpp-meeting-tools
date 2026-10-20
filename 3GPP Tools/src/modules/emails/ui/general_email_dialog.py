# --- File: src/modules/emails/ui/general_email_dialog.py ---
import html
import json
import logging
import re
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QDate, QUrl
from PyQt5.QtGui import QColor, QFont, QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableView,
    QHeaderView, QTextBrowser, QCheckBox, QAbstractItemView, QFrame,
    QSpinBox, QMessageBox, QDateEdit, QTableWidget, QTableWidgetItem,
    QColorDialog, QMenu, QSplitter, QWidget
)

from core.utils.paths import get_project_root
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.ui.config_dialog import OutlookFolderPickerDialog
from modules.emails.ui.email_summary_dialog import EmailSummaryDialog

# Configuration helpers re-exported from core for backward compatibility
from modules.emails.core.email_config import load_wg_email_config, save_wg_email_config

# Backward compatibility re-exports for dialogs extracted to dedicated module
from modules.emails.ui.tdoc_emails_dialog import StandaloneEmailReaderWindow, TDocEmailsDialog


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


