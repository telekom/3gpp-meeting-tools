# --- File: modules/meetings/ui/dialogs.py ---
import webbrowser
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                             QLineEdit, QComboBox, QCheckBox, QGroupBox, QFormLayout,
                             QMessageBox, QApplication)

from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.scraper import ManualMeetingFetcherThread, MEETING_SOURCES


def _format_meeting_info(data: dict) -> str:
    if not data: return ""

    FIELD_MAP = {
        "wg_name": "Working Group",
        "meeting_number": "Meeting Number",
        "mtg_id": "3GPP Portal ID",
        "name": "Meeting Name",
        "location": "Location",
        "start_date": "Start Date",
        "end_date": "End Date",
        "is_ad_hoc": "Ad-Hoc / BIS",
        "is_electronic": "Meeting Type",
        "first_tdoc": "First TDoc",
        "last_tdoc": "Last TDoc",
        "url_key": "Main FTP Link",
        "docs_folder_url": "Docs Folder Link",
        "id": "Database ID"
    }

    html_parts = []
    for key, display_name in FIELD_MAP.items():
        val = data.get(key)
        if key == "is_ad_hoc":
            val_str = "✅ Yes" if val else "❌ No"
        elif key == "is_electronic":
            val_str = "✅ Yes (Electronic)" if val else "❌ No (In-Person)"
        elif key == "mtg_id":
            if val:
                val_str = f'<a href="https://portal.3gpp.org/Home.aspx#/meeting?MtgId={val}">{val}</a>'
            else:
                val_str = "N/A"
        elif key == "url_key":
            if val and not str(val).startswith('http'):
                val = f"https://www.3gpp.org/ftp/{str(val).lstrip('/')}"
            val_str = f'<a href="{val}">{val}</a>' if val else "N/A"
        elif key == "docs_folder_url":
            val_str = f'<a href="{val}">{val}</a>' if val else "N/A"
        else:
            val_str = str(val) if val else "N/A"

        html_parts.append(f"<b>{display_name}:</b> {val_str}")
        if key in ["end_date", "is_electronic", "last_tdoc"]:
            html_parts.append("<hr>")

    future_keys = [k for k in data.keys() if k not in FIELD_MAP and k not in ["wg_id", "sort_number"]]
    if future_keys:
        html_parts.append("<hr><b>--- Advanced Metrics ---</b>")
        for k in future_keys:
            clean_name = k.replace("_", " ").title()
            html_parts.append(f"<b>{clean_name}:</b> {data[k]}")

    return "<br>".join(html_parts)


class MeetingInfoDialog(QDialog):
    def __init__(self, data: dict, parent=None):
        super().__init__(parent)
        title_str = f"{data.get('wg_name', '')} {data.get('meeting_number', '')}".strip()
        self.setWindowTitle(f"Meeting Details: {title_str}")
        self.setMinimumWidth(500)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; } QLabel { font-size: 13px; }")

        layout = QVBoxLayout(self)
        info_label = QLabel(_format_meeting_info(data))
        info_label.setWordWrap(True)
        info_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        info_label.linkActivated.connect(webbrowser.open)
        layout.addWidget(info_label)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)


# ==========================================
# --- DIALOG: ADD / FETCH MEETING ---
# ==========================================
class AddMeetingDialog(QDialog):
    def __init__(self, db: MeetingsDatabase, parent=None):
        super().__init__(parent)
        self.db = db
        self.fetch_thread = None
        self.current_fetched_data = {}

        self.setWindowTitle("➕ Add / Fetch Meeting Manually")
        self.setMinimumWidth(560)
        self.setStyleSheet("""
            QDialog { background-color: #FAFAFA; }
            QGroupBox { font-weight: bold; border: 1px solid #D0D0D0; border-radius: 6px; margin-top: 10px; padding-top: 10px; background-color: white; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            QLineEdit, QComboBox { padding: 5px; border: 1px solid #CCC; border-radius: 4px; }
            QLineEdit:focus, QComboBox:focus { border: 1px solid #0078D7; }
        """)

        self._setup_ui()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)

        # 1. Search Box
        search_group = QGroupBox("1. Query 3GPP Meeting")
        search_layout = QVBoxLayout(search_group)

        row_layout = QHBoxLayout()
        self.wg_combo = QComboBox()
        self.wg_combo.addItem("All / Auto-Detect")
        self.wg_combo.addItems(list(MEETING_SOURCES.keys()))
        self.wg_combo.setToolTip("Target Working Group (or leave on Auto-Detect)")

        self.query_input = QLineEdit()
        self.query_input.setPlaceholderText("e.g. 33120, SA3-130, 130, or FTP URL")
        self.query_input.setToolTip("Enter 3GPP Portal MtgId, WG Meeting tag, or meeting number")
        self.query_input.returnPressed.connect(self._start_fetch)

        self.btn_fetch = QPushButton("🔍 Fetch Details")
        self.btn_fetch.setStyleSheet("QPushButton { background-color: #0078D7; color: white; font-weight: bold; padding: 6px 14px; border-radius: 4px; } QPushButton:hover { background-color: #005A9E; }")
        self.btn_fetch.clicked.connect(self._start_fetch)

        row_layout.addWidget(self.wg_combo)
        row_layout.addWidget(self.query_input, 1)
        row_layout.addWidget(self.btn_fetch)
        search_layout.addLayout(row_layout)

        self.lbl_status = QLabel("Enter a Meeting ID, Number, or FTP URL and click 'Fetch Details'.")
        self.lbl_status.setStyleSheet("color: #666; font-size: 11px; margin-top: 2px;")
        search_layout.addWidget(self.lbl_status)
        main_layout.addWidget(search_group)

        # 2. Preview and Editable Form
        self.preview_group = QGroupBox("2. Meeting Details Preview (Editable)")
        form = QFormLayout(self.preview_group)
        form.setLabelAlignment(Qt.AlignRight)

        self.edit_wg = QLineEdit()
        self.edit_wg.setPlaceholderText("e.g. SA3")
        form.addRow("Working Group *:", self.edit_wg)

        self.edit_num = QLineEdit()
        self.edit_num.setPlaceholderText("e.g. 130 or 130-e")
        form.addRow("Meeting Number *:", self.edit_num)

        self.edit_mtg_id = QLineEdit()
        self.edit_mtg_id.setPlaceholderText("e.g. 33120")
        form.addRow("3GPP Portal ID (MtgId):", self.edit_mtg_id)

        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("e.g. 3GPP SA WG3 #130")
        form.addRow("Meeting Name / Title:", self.edit_name)

        self.edit_location = QLineEdit()
        self.edit_location.setPlaceholderText("e.g. Jeju, KR")
        form.addRow("Location:", self.edit_location)

        date_layout = QHBoxLayout()
        self.edit_start = QLineEdit()
        self.edit_start.setPlaceholderText("YYYY-MM-DD")
        self.edit_end = QLineEdit()
        self.edit_end.setPlaceholderText("YYYY-MM-DD")
        date_layout.addWidget(self.edit_start)
        date_layout.addWidget(QLabel("to"))
        date_layout.addWidget(self.edit_end)
        form.addRow("Dates (Start / End):", date_layout)

        self.edit_url_key = QLineEdit()
        self.edit_url_key.setPlaceholderText("e.g. tsg_sa/WG3_Security/TSGS3_130_Jeju/")
        form.addRow("FTP Main Path:", self.edit_url_key)

        self.edit_docs_url = QLineEdit()
        self.edit_docs_url.setPlaceholderText("e.g. https://www.3gpp.org/ftp/tsg_sa/WG3_Security/TSGS3_130_Jeju/Docs/")
        form.addRow("Docs/ URL:", self.edit_docs_url)

        tdoc_layout = QHBoxLayout()
        self.edit_first_tdoc = QLineEdit()
        self.edit_first_tdoc.setPlaceholderText("e.g. S3-241001")
        self.edit_last_tdoc = QLineEdit()
        self.edit_last_tdoc.setPlaceholderText("e.g. S3-241850")
        tdoc_layout.addWidget(self.edit_first_tdoc)
        tdoc_layout.addWidget(QLabel("to"))
        tdoc_layout.addWidget(self.edit_last_tdoc)
        form.addRow("TDocs Range:", tdoc_layout)

        flags_layout = QHBoxLayout()
        self.chk_adhoc = QCheckBox("Ad-Hoc / BIS Meeting")
        self.chk_electronic = QCheckBox("Electronic Meeting (eMeeting)")
        flags_layout.addWidget(self.chk_adhoc)
        flags_layout.addWidget(self.chk_electronic)
        flags_layout.addStretch()
        form.addRow("Flags:", flags_layout)

        main_layout.addWidget(self.preview_group)

        # 3. Actions
        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("💾 Save to Database")
        self.btn_save.setEnabled(False)
        self.btn_save.setStyleSheet("QPushButton { background-color: #107C41; color: white; font-weight: bold; padding: 7px 18px; border-radius: 4px; } QPushButton:hover { background-color: #0B5A30; } QPushButton:disabled { background-color: #A6D6B8; }")
        self.btn_save.clicked.connect(self._save_to_db)

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_cancel)
        main_layout.addLayout(btn_layout)

    def _start_fetch(self):
        # Prevent starting a new thread if one is already running
        if self.fetch_thread and self.fetch_thread.isRunning():
            return

        query = self.query_input.text().strip()
        if not query:
            QMessageBox.warning(self, "Input Required", "Please enter a 3GPP Meeting ID, number, or URL.")
            return

        selected_wg = self.wg_combo.currentText()
        self.btn_fetch.setEnabled(False)
        self.btn_fetch.setText("⏳ Fetching...")
        self.lbl_status.setText("⏳ Querying 3GPP Portal and FTP archive...")
        self.lbl_status.setStyleSheet("color: #005A9E; font-weight: bold;")
        self.btn_save.setEnabled(False)

        self.fetch_thread = ManualMeetingFetcherThread(self.db.db_path, query, selected_wg)
        self.fetch_thread.progress_msg.connect(lambda msg: self.lbl_status.setText(msg))
        self.fetch_thread.fetch_finished.connect(self._on_fetch_finished)
        self.fetch_thread.start()

    def reject(self):
        """Cleanly abort thread if dialog is dismissed."""
        if self.fetch_thread and self.fetch_thread.isRunning():
            self.fetch_thread.requestInterruption()
            self.fetch_thread.quit()
            self.fetch_thread.wait(150)
        super().reject()

    def _on_fetch_finished(self, success: bool, data: dict, msg: str):
        self.btn_fetch.setEnabled(True)
        self.btn_fetch.setText("🔍 Fetch Details")

        if success:
            self.lbl_status.setText(f"✅ {msg}")
            self.lbl_status.setStyleSheet("color: #107C41; font-weight: bold;")
            self.current_fetched_data = data

            # Pre-populate preview fields
            self.edit_wg.setText(data.get("wg_name", ""))
            self.edit_num.setText(data.get("meeting_number", ""))
            self.edit_mtg_id.setText(str(data.get("mtg_id", "")))
            self.edit_name.setText(data.get("name", ""))
            self.edit_location.setText(data.get("location", ""))
            self.edit_start.setText(data.get("start_date", ""))
            self.edit_end.setText(data.get("end_date", ""))
            self.edit_url_key.setText(data.get("url_key", ""))
            self.edit_docs_url.setText(data.get("docs_folder_url", ""))
            self.edit_first_tdoc.setText(data.get("first_tdoc", ""))
            self.edit_last_tdoc.setText(data.get("last_tdoc", ""))
            self.chk_adhoc.setChecked(bool(data.get("is_ad_hoc", 0)))
            self.chk_electronic.setChecked(bool(data.get("is_electronic", 0)))

            self.btn_save.setEnabled(True)
        else:
            self.lbl_status.setText(f"⚠️ {msg}")
            self.lbl_status.setStyleSheet("color: #D83B01; font-weight: bold;")
            QMessageBox.warning(self, "Fetch Failed", f"{msg}\n\nYou can still fill in the fields manually.")
            self.btn_save.setEnabled(True)

    def _save_to_db(self):
        wg = self.edit_wg.text().strip()
        num = self.edit_num.text().strip()

        if not wg or not num:
            QMessageBox.warning(self, "Missing Fields", "Working Group and Meeting Number are required to save.")
            return

        save_dict = {
            "wg_name": wg,
            "meeting_number": num,
            "mtg_id": self.edit_mtg_id.text().strip(),
            "name": self.edit_name.text().strip(),
            "location": self.edit_location.text().strip(),
            "start_date": self.edit_start.text().strip(),
            "end_date": self.edit_end.text().strip(),
            "url_key": self.edit_url_key.text().strip(),
            "docs_folder_url": self.edit_docs_url.text().strip(),
            "first_tdoc": self.edit_first_tdoc.text().strip(),
            "last_tdoc": self.edit_last_tdoc.text().strip(),
            "is_ad_hoc": 1 if self.chk_adhoc.isChecked() else 0,
            "is_electronic": 1 if self.chk_electronic.isChecked() else 0,
            "folder_name": self.edit_url_key.text().strip().split('/')[-1] if self.edit_url_key.text().strip() else num
        }

        try:
            self.db.upsert_single_meeting(save_dict)
            QMessageBox.information(self, "Meeting Saved", f"Meeting {wg} #{num} was successfully saved to your database.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Could not save meeting to database:\n{e}")