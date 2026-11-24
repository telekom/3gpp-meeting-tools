"""DynaReport manual specification fetcher thread and AddSpecDialog."""

import logging
import re
from typing import Optional

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
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
from modules.specifications.core.database import SpecsDatabase
from modules.specifications.core.scraper import fetch_metadata_from_dynareport


class ManualSpecFetcherThread(QThread):
    """
    Background worker that queries 3GPP DynaReport HTML using the shared
    scraper engine with real-time log signaling for UI troubleshooting.
    """

    fetch_finished = pyqtSignal(bool, dict, str)
    log_msg = pyqtSignal(str, int)

    def __init__(self, spec_number: str):
        super().__init__()
        self.spec_number = spec_number.strip()

    def run(self):
        try:
            if self.isInterruptionRequested():
                return

            metadata = fetch_metadata_from_dynareport(self.spec_number, log_cb=self.log_msg.emit)

            if self.isInterruptionRequested():
                return

            if metadata and metadata.get("title"):
                spec_type = metadata.get("type") or "TS"
                msg = f"Fetched metadata for {spec_type} {metadata['number']} successfully."
                self.fetch_finished.emit(True, metadata, msg)
            else:
                err_cause = metadata.get("error") if metadata else ""
                reason = f": {err_cause}" if err_cause else " (Check spec number or network)"
                fail_msg = f"Could not extract metadata for {self.spec_number}{reason}"
                self.fetch_finished.emit(False, metadata or {}, fail_msg)
        except Exception as e:
            self.fetch_finished.emit(False, {}, f"Error fetching details: {e}")


class AddSpecDialog(QDialog):
    """Dialog allowing users to query, preview, manually edit, and register an individual specification."""

    def __init__(self, db: SpecsDatabase, parent=None):
        super().__init__(parent)
        self.db = db
        self.fetch_thread: Optional[ManualSpecFetcherThread] = None
        self.sync_requested = False
        self.saved_spec_number = ""

        self.setWindowTitle("➕ Add / Fetch Specification")
        self.setMinimumWidth(560)
        self.setStyleSheet(f"""
            QDialog {{ background-color: {C.BG_ALT_LIGHT}; }}
            QGroupBox {{
                font-weight: bold;
                border: 1px solid {C.BORDER_GRAY};
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

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)

        # 1. Query Section
        query_group = QGroupBox("1. Query 3GPP Specification")
        query_layout = QVBoxLayout(query_group)

        row_layout = QHBoxLayout()
        self.query_input = QLineEdit()
        self.query_input.setPlaceholderText("e.g. 23.501, 38.331, 23.700-01...")
        self.query_input.setToolTip("Enter 3GPP specification number")
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

        self.lbl_status = QLabel("Enter a specification number and click 'Fetch Details'.")
        self.lbl_status.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; margin-top: 2px;")
        self.lbl_status.setWordWrap(True)
        query_layout.addWidget(self.lbl_status)
        main_layout.addWidget(query_group)

        # 2. Form Preview
        self.preview_group = QGroupBox("2. Specification Details (Editable)")
        form = QFormLayout(self.preview_group)
        form.setLabelAlignment(Qt.AlignRight)

        self.edit_number = QLineEdit()
        self.edit_number.setPlaceholderText("e.g. 23.501")
        form.addRow("Spec Number *:", self.edit_number)

        self.edit_title = QLineEdit()
        self.edit_title.setPlaceholderText("e.g. System architecture for the 5G System (5GS)")
        form.addRow("Title *:", self.edit_title)

        self.type_combo = QComboBox()
        self.type_combo.addItems(["TS", "TR"])
        form.addRow("Type:", self.type_combo)

        self.edit_group = QLineEdit()
        self.edit_group.setPlaceholderText("e.g. SA2, RAN2, CT1")
        form.addRow("Primary Group:", self.edit_group)

        self.edit_init_rel = QLineEdit()
        self.edit_init_rel.setPlaceholderText("e.g. Rel-15")
        form.addRow("Initial Release:", self.edit_init_rel)

        self.edit_tech = QLineEdit()
        self.edit_tech.setPlaceholderText("e.g. 5G, LTE")
        form.addRow("Radio Technology:", self.edit_tech)

        self.chk_sync_now = QCheckBox("Immediately sync files and releases from 3GPP FTP")
        self.chk_sync_now.setChecked(True)
        form.addRow("", self.chk_sync_now)

        main_layout.addWidget(self.preview_group)

        # 3. Actions
        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("💾 Save Specification")
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
        self.btn_save.clicked.connect(self._save_spec)

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
        if not query:
            QMessageBox.warning(self, "Input Required", "Please enter a specification number.")
            return

        self.btn_fetch.setEnabled(False)
        self.btn_fetch.setText("⏳ Fetching...")
        self.lbl_status.setText(f"⏳ Connecting to 3GPP DynaReport for '{query}'...")
        self.lbl_status.setStyleSheet(f"color: {C.ACCENT_BRIGHT}; font-weight: bold;")

        self.fetch_thread = ManualSpecFetcherThread(query)
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

            self.edit_number.setText(data.get("number", self.query_input.text().strip()))
            self.edit_title.setText(data.get("title", ""))
            self.type_combo.setCurrentText(data.get("type", "TS"))
            self.edit_group.setText(data.get("primary_group", ""))
            self.edit_init_rel.setText(data.get("initial_release", ""))
            self.edit_tech.setText(data.get("radio_technology", ""))
        else:
            self.lbl_status.setText(f"⚠️ {msg}")
            self.lbl_status.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")

            raw_num = self.query_input.text().strip()
            if raw_num:
                self.edit_number.setText(raw_num)
                if re.search(r"\b\d{2}\.[89]\d{2}\b", raw_num):
                    self.type_combo.setCurrentText("TR")
                else:
                    self.type_combo.setCurrentText("TS")

    def _save_spec(self):
        spec_num = self.edit_number.text().strip()
        title = self.edit_title.text().strip()

        if not spec_num:
            QMessageBox.warning(self, "Validation Error", "Specification Number is required.")
            return

        data = {
            "number": spec_num,
            "title": title,
            "type": self.type_combo.currentText(),
            "primary_group": self.edit_group.text().strip(),
            "initial_release": self.edit_init_rel.text().strip(),
            "radio_technology": self.edit_tech.text().strip(),
        }

        if self.db.upsert_manual_spec(data):
            self.saved_spec_number = spec_num
            self.sync_requested = self.chk_sync_now.isChecked()
            QMessageBox.information(
                self, "Success", f"Specification {data['type']} {spec_num} saved to database."
            )
            self.accept()
        else:
            QMessageBox.critical(
                self, "Error", f"Failed to save specification {data['type']} {spec_num} to database."
            )

    def reject(self):
        """Cleanly abort worker thread if dialog is dismissed while fetching."""
        if self.fetch_thread and self.fetch_thread.isRunning():
            self.fetch_thread.requestInterruption()
            self.fetch_thread.quit()
            self.fetch_thread.wait(200)
        super().reject()
