"""Background portal fetcher thread and manual Add/Fetch Work Item dialog."""

from __future__ import annotations

import logging
import re
from typing import Optional

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from core.ui.theme import C
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_scraper import extract_wi_code, fetch_wi_details_from_portal


class ManualWIFetcherThread(QThread):
    """
    Background worker that fetches and parses Work Item details directly from
    the 3GPP Portal without blocking the GUI thread.
    """

    fetch_finished = pyqtSignal(bool, dict, str)
    log_msg = pyqtSignal(str, int)

    def __init__(self, query: str):
        super().__init__()
        self.query = query.strip()

    def run(self):
        try:
            if self.isInterruptionRequested():
                return

            code = extract_wi_code(self.query)
            if not code:
                self.fetch_finished.emit(
                    False, {}, "Could not determine a valid Work Item ID from the input."
                )
                return

            self.log_msg.emit(f"Connecting to 3GPP Portal for WI #{code}...", logging.INFO)
            metadata = fetch_wi_details_from_portal(code, log_cb=self.log_msg.emit)

            if self.isInterruptionRequested():
                return

            if metadata and (metadata.get("name") or metadata.get("acronym")):
                acr = metadata.get("acronym") or metadata.get("name")
                remarks_cnt = len(metadata.get("remarks", []))
                msg = f"Fetched metadata for WI #{code} ({acr}) successfully ({remarks_cnt} remarks)."
                self.fetch_finished.emit(True, metadata, msg)
            else:
                err_cause = metadata.get("error") if metadata else ""
                reason = f": {err_cause}" if err_cause else " (Check WI ID or network connection)"
                fail_msg = f"Could not extract details for WI #{code}{reason}"
                self.fetch_finished.emit(False, metadata or {}, fail_msg)
        except Exception as e:
            self.fetch_finished.emit(False, {}, f"Error fetching details: {e}")


class AddWorkItemDialog(QDialog):
    """Dialog allowing users to query, preview, manually edit, and register an individual Work Item."""

    def __init__(self, db: WorkItemsDatabase, initial_query: str = "", parent=None):
        super().__init__(parent)
        self.db = db
        self.fetch_thread: Optional[ManualWIFetcherThread] = None
        self.saved_wi_code = ""
        self.is_starred_requested = False
        self.is_following_requested = False
        self._fetched_remarks = []

        self.setWindowTitle("➕ Add / Fetch Work Item")
        self.setMinimumWidth(580)
        self.setStyleSheet(f"""
            QDialog {{ background-color: {C.BG_ALT_LIGHT}; }}
            QGroupBox {{
                font-weight: bold;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: {C.BG_CARD};
            }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 4px; }}
            QLineEdit, QComboBox {{ padding: 5px; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; }}
            QLineEdit:focus, QComboBox:focus {{ border: 1px solid {C.ACCENT_BRIGHT}; }}
        """)

        self._setup_ui()

        if initial_query:
            self.query_input.setText(initial_query)
            self._start_fetch()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)

        # 1. Query Section
        query_group = QGroupBox("1. Query 3GPP Work Item")
        query_layout = QVBoxLayout(query_group)

        row_layout = QHBoxLayout()
        self.query_input = QLineEdit()
        self.query_input.setPlaceholderText(
            "e.g. 1121123, #1121123, or https://portal.3gpp.org/...workitemId=1121123"
        )
        self.query_input.setToolTip("Enter a 3GPP Work Item ID or Portal URL")
        self.query_input.returnPressed.connect(self._start_fetch)

        self.btn_fetch = QPushButton("🔍 Fetch Details")
        self.btn_fetch.setCursor(Qt.PointingHandCursor)
        self.btn_fetch.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.ACCENT_BRIGHT};
                color: {C.TEXT_ON_ACCENT};
                font-weight: bold;
                padding: 6px 14px;
                border-radius: 4px;
            }}
            QPushButton:hover {{ background-color: {C.OFFICE_BLUE}; }}
            QPushButton:disabled {{ background-color: {C.ACCENT_RING}; }}
        """)
        self.btn_fetch.clicked.connect(self._start_fetch)

        row_layout.addWidget(self.query_input, 1)
        row_layout.addWidget(self.btn_fetch)
        query_layout.addLayout(row_layout)

        self.lbl_status = QLabel("Enter a Work Item ID or URL and click 'Fetch Details'.")
        self.lbl_status.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; margin-top: 2px;")
        self.lbl_status.setWordWrap(True)
        query_layout.addWidget(self.lbl_status)
        main_layout.addWidget(query_group)

        # 2. Form Preview (Editable)
        self.preview_group = QGroupBox("2. Work Item Details (Editable)")
        form = QFormLayout(self.preview_group)
        form.setLabelAlignment(Qt.AlignRight)

        self.edit_code = QLineEdit()
        self.edit_code.setPlaceholderText("e.g. 1121123")
        form.addRow("WI Code *:", self.edit_code)

        self.edit_acronym = QLineEdit()
        self.edit_acronym.setPlaceholderText("e.g. NR_Sensing_bis-Core")
        form.addRow("Acronym *:", self.edit_acronym)

        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText(
            "e.g. Core part: Integrated Sensing and communication (ISAC) for NR"
        )
        form.addRow("Title / Name *:", self.edit_name)

        self.edit_release = QLineEdit()
        self.edit_release.setPlaceholderText("e.g. Rel-20, Rel-19, Rel-18...")
        form.addRow("Release:", self.edit_release)

        self.edit_groups = QLineEdit()
        self.edit_groups.setPlaceholderText("e.g. RAN3 (comma-separated for multiple)")
        form.addRow("Responsible WG(s):", self.edit_groups)

        self.edit_status = QLineEdit()
        self.edit_status.setPlaceholderText("e.g. PCG approved, In progress, Finished...")
        form.addRow("Status:", self.edit_status)

        self.edit_start_date = QLineEdit()
        self.edit_start_date.setPlaceholderText("YYYY-MM-DD")
        form.addRow("Start Date:", self.edit_start_date)

        self.edit_end_date = QLineEdit()
        self.edit_end_date.setPlaceholderText("YYYY-MM-DD")
        form.addRow("End Date:", self.edit_end_date)

        self.edit_latest_wid = QLineEdit()
        self.edit_latest_wid.setPlaceholderText("e.g. RP-261901")
        form.addRow("Latest WID:", self.edit_latest_wid)

        tracking_layout = QVBoxLayout()
        tracking_layout.setSpacing(4)
        self.chk_starred = QCheckBox("⭐ Add to Starred (Personal Watchlist)")
        self.chk_following = QCheckBox("🎯 Add to Following (Delegation Responsibilities)")
        tracking_layout.addWidget(self.chk_starred)
        tracking_layout.addWidget(self.chk_following)
        form.addRow("Tracking:", tracking_layout)

        main_layout.addWidget(self.preview_group)

        # 3. Actions
        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("💾 Save Work Item")
        self.btn_save.setCursor(Qt.PointingHandCursor)
        self.btn_save.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.SUCCESS};
                color: {C.TEXT_ON_ACCENT};
                font-weight: bold;
                padding: 7px 18px;
                border-radius: 4px;
            }}
            QPushButton:hover {{ background-color: {C.SUCCESS_DARK}; }}
        """)
        self.btn_save.clicked.connect(self._save_wi)

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_cancel)
        main_layout.addLayout(btn_layout)

    def _start_fetch(self):
        if self.fetch_thread and self.fetch_thread.isRunning():
            return

        query = self.query_input.text().strip()
        code = extract_wi_code(query)
        if not code:
            QMessageBox.warning(
                self, "Input Required", "Please enter a valid Work Item ID or 3GPP Portal URL."
            )
            return

        self.btn_fetch.setEnabled(False)
        self.btn_fetch.setText("⏳ Fetching...")
        self.lbl_status.setText(f"⏳ Connecting to 3GPP Portal for WI #{code}...")
        self.lbl_status.setStyleSheet(f"color: {C.ACCENT_BRIGHT}; font-weight: bold;")

        self.fetch_thread = ManualWIFetcherThread(query)
        self.fetch_thread.log_msg.connect(self._on_fetch_log)
        self.fetch_thread.fetch_finished.connect(self._on_fetch_finished)
        self.fetch_thread.start()

    def _on_fetch_log(self, msg: str, level: int):
        self.lbl_status.setText(msg)
        if level >= logging.ERROR:
            self.lbl_status.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")
        elif level >= logging.WARNING:
            self.lbl_status.setStyleSheet(f"color: {C.WARNING_AMBER}; font-weight: bold;")
        else:
            self.lbl_status.setStyleSheet(f"color: {C.ACCENT_BRIGHT};")

    def _on_fetch_finished(self, success: bool, data: dict, msg: str):
        self.btn_fetch.setEnabled(True)
        self.btn_fetch.setText("🔍 Fetch Details")

        if success:
            self.lbl_status.setText(f"✅ {msg}")
            self.lbl_status.setStyleSheet(f"color: {C.SUCCESS}; font-weight: bold;")

            self.edit_code.setText(data.get("code", ""))
            self.edit_acronym.setText(data.get("acronym", ""))
            self.edit_name.setText(data.get("name", ""))
            self.edit_release.setText(data.get("release", ""))
            groups = data.get("responsible_groups", [])
            self.edit_groups.setText(", ".join(groups) if isinstance(groups, list) else str(groups))
            self.edit_status.setText(data.get("status", ""))
            self.edit_start_date.setText(data.get("start_date", ""))
            self.edit_end_date.setText(data.get("end_date", ""))
            self.edit_latest_wid.setText(data.get("latest_wid", ""))
            self._fetched_remarks = data.get("remarks", [])
        else:
            self.lbl_status.setText(f"⚠️ {msg}")
            self.lbl_status.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")

    def _save_wi(self):
        code = self.edit_code.text().strip()
        acronym = self.edit_acronym.text().strip()
        name = self.edit_name.text().strip()

        if not code:
            QMessageBox.warning(self, "Validation Error", "Please provide a valid Work Item Code.")
            self.edit_code.setFocus()
            return

        if not acronym and not name:
            QMessageBox.warning(
                self, "Validation Error", "Please provide at least a Work Item Acronym or Title."
            )
            self.edit_acronym.setFocus()
            return

        groups_str = self.edit_groups.text().strip()
        groups = [g.strip() for g in re.split(r"[,;/|\s]+", groups_str) if g.strip()]

        wi_data = {
            "code": code,
            "acronym": acronym,
            "name": name,
            "release": self.edit_release.text().strip(),
            "latest_wid": self.edit_latest_wid.text().strip(),
            "start_date": self.edit_start_date.text().strip(),
            "end_date": self.edit_end_date.text().strip(),
            "responsible_groups": groups,
            "remarks": self._fetched_remarks,
        }

        if self.db.upsert_single_work_item(wi_data):
            self.saved_wi_code = code
            self.is_starred_requested = self.chk_starred.isChecked()
            self.is_following_requested = self.chk_following.isChecked()
            self.accept()
        else:
            QMessageBox.critical(
                self, "Database Error", f"Failed to save Work Item #{code} to the database."
            )

    def closeEvent(self, event):
        if self.fetch_thread and self.fetch_thread.isRunning():
            self.fetch_thread.requestInterruption()
            self.fetch_thread.wait(1000)
        super().closeEvent(event)
