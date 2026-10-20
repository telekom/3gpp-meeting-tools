# --- File: src/modules/meetings/ui/tdoc_info_dialog.py ---
import logging
import webbrowser

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

logger = logging.getLogger(__name__)

class TDocInfoDialog(QDialog):
    """Modernized TDoc Details Dialog displaying all stored database attributes."""

    def __init__(
        self,
        details: dict,
        docs_ftp_url: str = "",
        revisions: list = None,
        parent=None,
        triage_callback=None,  # <--- Added parameter
    ):
        super().__init__(parent)
        self.details = details
        self.docs_ftp_url = docs_ftp_url.rstrip('/') if docs_ftp_url else ""
        self.revisions = revisions or []
        self.triage_callback = triage_callback

        tdoc_id = str(details.get("TDoc", "Unknown")).strip()
        tdoc_type = str(details.get("Type", "TDoc")).strip() or "TDoc"
        tdoc_status = str(details.get("TDoc Status", "")).strip()
        title = str(details.get("Title", "No Title Available")).strip()

        self.setWindowTitle(f"TDoc Details: {tdoc_id}")
        self.resize(650, 720)
        self.setMinimumWidth(580)
        self.setStyleSheet("""
            QDialog {
                background-color: #F8F9FA;
            }
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
            }
            QLabel {
                font-size: 13px;
                color: #2D3748;
            }
            QPushButton {
                padding: 6px 14px;
                font-size: 12px;
                border-radius: 4px;
                border: 1px solid #CBD5E0;
                background-color: #FFFFFF;
                color: #2D3748;
            }
            QPushButton:hover {
                background-color: #EDF2F7;
                border-color: #A0AEC0;
            }
            QPushButton#primaryActionBtn {
                background-color: #005A9E;
                color: #FFFFFF;
                border: 1px solid #004578;
                font-weight: bold;
            }
            QPushButton#primaryActionBtn:hover {
                background-color: #004578;
            }
            QPushButton#revChipBtn {
                background-color: #F0F4F8;
                border: 1px solid #D2E3FC;
                border-radius: 12px;
                padding: 3px 10px;
                font-size: 11px;
                color: #005A9E;
                font-weight: bold;
            }
            QPushButton#revChipBtn:hover {
                background-color: #E1EFFF;
                border-color: #005A9E;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # --- 1. HEADER CARD ---
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(8)

        title_row = QHBoxLayout()
        title_row.setSpacing(8)

        # Document Type Badge
        type_badge = QLabel(f"<b>{tdoc_type}</b>")
        type_badge.setStyleSheet("""
            background-color: #EBF8FF;
            color: #2B6CB0;
            border: 1px solid #BEE3F8;
            border-radius: 4px;
            padding: 2px 8px;
            font-size: 12px;
            font-weight: bold;
        """)
        title_row.addWidget(type_badge)

        # Status Badge (if populated)
        if tdoc_status:
            status_badge = QLabel(f"<b>{tdoc_status}</b>")
            status_style = self._get_status_badge_style(tdoc_status)
            status_badge.setStyleSheet(status_style)
            title_row.addWidget(status_badge)

        tdoc_label = QLabel(f"<b>{tdoc_id}</b>")
        tdoc_label.setStyleSheet("font-size: 18px; color: #1A202C; font-weight: bold;")
        tdoc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        title_row.addWidget(tdoc_label)

        title_row.addStretch()
        header_layout.addLayout(title_row)

        desc_label = QLabel(title)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #4A5568; font-size: 13px; line-height: 1.4;")
        desc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        header_layout.addWidget(desc_label)

        layout.addWidget(header_card)

        # --- 2. SCROLLABLE DETAILS CARD ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; }")

        details_card = QFrame()
        details_card.setObjectName("cardFrame")
        form = QFormLayout(details_card)
        form.setContentsMargins(14, 14, 14, 14)
        form.setSpacing(10)
        form.setLabelAlignment(Qt.AlignRight)

        # Primary Metadata
        self._add_row(form, "Agenda Item", details.get("Agenda Item") or "-")
        self._add_row(form, "Source", details.get("Source") or "-")
        self._add_row(form, "Intended For", details.get("For") or "-")

        # Personal Status & Notes
        my_status = str(details.get("My Status", "")).replace("🔄 ", "").strip()
        if my_status and my_status != "⚪ Neutral":
            self._add_row(form, "My Status", my_status)

        my_notes = str(details.get("My Notes", "")).strip()
        if my_notes:
            self._add_row(form, "My Notes", my_notes)

        # Secretary Remarks
        sec_remarks = str(details.get("Secretary Remarks", "")).strip()
        if sec_remarks:
            self._add_row(form, "Secretary Remarks", sec_remarks)

        # Abstract
        abstract = str(details.get("Abstract", "")).strip()
        if abstract:
            self._add_row(form, "Abstract", abstract)

        # Related TDocs
        related = str(details.get("Related TDocs", "")).strip()
        if related:
            self._add_row(form, "Related TDocs", related)

        # Revision Lineage Chips
        if self.revisions:
            rev_container = QWidget()
            rev_layout = QHBoxLayout(rev_container)
            rev_layout.setContentsMargins(0, 0, 0, 0)
            rev_layout.setSpacing(6)

            for rev_id in self.revisions:
                chip = QPushButton(rev_id)
                chip.setObjectName("revChipBtn")
                chip.setCursor(Qt.PointingHandCursor)
                chip.setToolTip(f"Revision: {rev_id}")
                if self.docs_ftp_url:
                    chip.clicked.connect(lambda _, r=rev_id: webbrowser.open(f"{self.docs_ftp_url}/{r}.zip"))
                rev_layout.addWidget(chip)

            rev_layout.addStretch()
            form.addRow(self._make_key_label("Revisions:"), rev_container)

        # FTP Download URL
        if self.docs_ftp_url and tdoc_id != "Unknown":
            doc_url = f"{self.docs_ftp_url}/{tdoc_id}.zip"
            ftp_label = QLabel(f'<a href="{doc_url}" style="color: #005A9E; text-decoration: none;">{doc_url}</a>')
            ftp_label.setOpenExternalLinks(True)
            ftp_label.setTextInteractionFlags(Qt.TextBrowserInteraction | Qt.TextSelectableByMouse)
            form.addRow(self._make_key_label("FTP Download:"), ftp_label)

        # Catch-all: render any remaining attributes not explicitly covered
        excluded_keys = {
            "TDoc", "Type", "Title", "Source", "For", "Agenda Item",
            "TDoc Status", "Secretary Remarks", "My Status", "My Notes",
            "Abstract", "Related TDocs", "Emails"
        }
        for key, value in details.items():
            if key not in excluded_keys and not key.startswith("_") and value:
                display_key = key.replace("_", " ").title()
                self._add_row(form, display_key, str(value))

        scroll.setWidget(details_card)
        layout.addWidget(scroll)

        # --- 3. ACTION BUTTONS ---
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        copy_btn = QPushButton("📋 Copy Info")
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_btn.setToolTip("Copy formatted summary to clipboard")
        copy_btn.clicked.connect(self._copy_tdoc_summary)
        btn_layout.addWidget(copy_btn)

        if self.triage_callback and tdoc_id != "Unknown":
            triage_btn = QPushButton("🤖 AI Triage")
            triage_btn.setCursor(Qt.PointingHandCursor)
            triage_btn.setToolTip("Stream technical summary and impact analysis using local LLM")
            triage_btn.clicked.connect(lambda: [self.triage_callback(tdoc_id), self.accept()])
            btn_layout.addWidget(triage_btn)

        if self.docs_ftp_url and tdoc_id != "Unknown":
            ftp_btn = QPushButton("📂 Open FTP Archive")
            ftp_btn.setToolTip('Open the ZIP archive for this TDoc in the configured 3GPP document location.')
            ftp_btn.setObjectName("primaryActionBtn")
            ftp_btn.setCursor(Qt.PointingHandCursor)
            ftp_btn.clicked.connect(lambda: webbrowser.open(f"{self.docs_ftp_url}/{tdoc_id}.zip"))
            btn_layout.addWidget(ftp_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setToolTip('Close this dialog.')
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _make_key_label(self, text: str) -> QLabel:
        lbl = QLabel(f"<b>{text}</b>")
        lbl.setStyleSheet("color: #718096; font-size: 12px;")
        return lbl

    def _add_row(self, form: QFormLayout, label_text: str, value_text: str):
        val_label = QLabel(value_text)
        val_label.setWordWrap(True)
        val_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow(self._make_key_label(f"{label_text}:"), val_label)

    def _get_status_badge_style(self, status: str) -> str:
        s = status.lower()
        if "agreed" in s or "approved" in s:
            return "background-color: #E6F4EA; color: #137333; border: 1px solid #CEEAD6; border-radius: 4px; padding: 2px 8px; font-size: 12px; font-weight: bold;"
        if "noted" in s or "postponed" in s:
            return "background-color: #FEF7E0; color: #B06000; border: 1px solid #FEEFC3; border-radius: 4px; padding: 2px 8px; font-size: 12px; font-weight: bold;"
        if "revised" in s or "withdrawn" in s:
            return "background-color: #FCE8E6; color: #C5221F; border: 1px solid #FAD2CF; border-radius: 4px; padding: 2px 8px; font-size: 12px; font-weight: bold;"
        return "background-color: #F1F3F4; color: #3C4043; border: 1px solid #DADCE0; border-radius: 4px; padding: 2px 8px; font-size: 12px; font-weight: bold;"

    def _copy_tdoc_summary(self):
        lines = [
            f"TDoc: {self.details.get('TDoc', '')}",
            f"Title: {self.details.get('Title', '')}",
            f"Source: {self.details.get('Source', '')}",
            f"Type: {self.details.get('Type', '')}",
            f"Agenda Item: {self.details.get('Agenda Item', '')}",
            f"Status: {self.details.get('TDoc Status', '')}",
            f"For: {self.details.get('For', '')}",
        ]
        if self.details.get("Secretary Remarks"):
            lines.append(f"Remarks: {self.details.get('Secretary Remarks')}")
        if self.details.get("My Status"):
            lines.append(f"My Status: {self.details.get('My Status')}")
        if self.details.get("My Notes"):
            lines.append(f"My Notes: {self.details.get('My Notes')}")

        QApplication.clipboard().setText("\n".join(lines))
        QMessageBox.information(self, "Copied", "TDoc details copied to clipboard.")
