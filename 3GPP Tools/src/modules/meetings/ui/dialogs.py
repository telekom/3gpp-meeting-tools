import webbrowser

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout

from modules.meetings.ui.ui_tabs import _format_meeting_info


class MeetingInfoDialog(QDialog):
    """A silent QDialog to show meeting info with clickable links."""

    def __init__(self, data: dict, parent=None):
        super().__init__(parent)
        title_str = f"{data.get('wg_name', '')} {data.get('meeting_number', '')}".strip()
        self.setWindowTitle(f"Meeting Details: {title_str}")
        self.setMinimumWidth(500)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; } QLabel { font-size: 13px; }")

        layout = QVBoxLayout(self)

        # --- IMPROVED: Use a QLabel that supports interaction ---
        info_label = QLabel(_format_meeting_info(data))
        info_label.setWordWrap(True)
        # TextBrowserInteraction allows clicking links, TextSelectableByMouse allows copying
        info_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        # This ensures the browser opens when the link is clicked
        info_label.linkActivated.connect(webbrowser.open)

        layout.addWidget(info_label)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
