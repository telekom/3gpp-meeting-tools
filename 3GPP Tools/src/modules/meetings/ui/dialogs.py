import webbrowser

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout


# ==========================================
# --- HELPER & DIALOG: MEETING INFO ---
# ==========================================

def _format_meeting_info(data: dict) -> str:
    if not data: return ""

    # 1. Map database column names to pretty UI labels
    FIELD_MAP = {
        "wg_name": "Working Group",
        "meeting_number": "Meeting Number",
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

    # 2. Iterate through our known fields in order
    for key, display_name in FIELD_MAP.items():
        val = data.get(key)

        # --- Custom Formatting for specific data types ---
        if key == "is_ad_hoc":
            val_str = "✅ Yes" if val else "❌ No"
        elif key == "is_electronic":
            val_str = "✅ Yes (Electronic)" if val else "❌ No (In-Person)"

        elif key == "url_key":
            if val and not str(val).startswith('http'):
                val = f"https://www.3gpp.org/ftp/{str(val).lstrip('/')}"
            val_str = f'<a href="{val}">{val}</a>' if val else "N/A"

        elif key == "docs_folder_url":
            val_str = f'<a href="{val}">{val}</a>' if val else "N/A"

        else:
            val_str = str(val) if val else "N/A"

        # Append the formatted line
        html_parts.append(f"<b>{display_name}:</b> {val_str}")

        # Add visual separators at logical sections
        if key in ["end_date", "is_electronic", "last_tdoc"]:
            html_parts.append("<hr>")

    # 3. FUTURE-PROOFING: Catch any new database columns we add later!
    # (We ignore internal columns like 'wg_id' and 'sort_number')
    future_keys = [k for k in data.keys() if k not in FIELD_MAP and k not in ["wg_id", "sort_number"]]

    if future_keys:
        html_parts.append("<hr><b>--- Additional Data ---</b>")
        for k in future_keys:
            clean_name = k.replace("_", " ").title() # Turns 'new_cool_column' into 'New Cool Column'
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
