# --- File: src/modules/emails/ui/email_dialogs.py ---
import logging

from PyQt5.QtWidgets import (
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
)

logger = logging.getLogger(__name__)

class EditEmailDialog(QDialog):
    def __init__(self, current_data: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("✏️ Edit Email Metadata")
        self.resize(500, 400)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; color: #333; }")

        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        self.txt_tdoc = QLineEdit(current_data.get("tdoc_id", ""))
        self.txt_ai = QLineEdit(current_data.get("agenda_item", ""))
        self.txt_company = QLineEdit(current_data.get("company", ""))
        self.txt_sender_name = QLineEdit(current_data.get("sender_name", ""))
        self.txt_sender_email = QLineEdit(current_data.get("sender_email", ""))
        self.txt_subject = QLineEdit(current_data.get("subject", ""))
        self.txt_short_text = QTextEdit(current_data.get("short_text", ""))
        self.txt_short_text.setMaximumHeight(100)

        for widget in [self.txt_tdoc, self.txt_ai, self.txt_company, self.txt_sender_name, self.txt_sender_email,
                       self.txt_subject, self.txt_short_text]:
            widget.setStyleSheet("background-color: #FFFFFF; border: 1px solid #CCC; color: #333;")

        form_layout.addRow("TDoc ID:", self.txt_tdoc)
        form_layout.addRow("Agenda Item:", self.txt_ai)
        form_layout.addRow("Company:", self.txt_company)
        form_layout.addRow("Sender Name:", self.txt_sender_name)
        form_layout.addRow("Sender Email:", self.txt_sender_email)
        form_layout.addRow("Subject:", self.txt_subject)
        form_layout.addRow("Short Text:", self.txt_short_text)

        layout.addLayout(form_layout)

        btn_layout = QHBoxLayout()
        btn_save = QPushButton("💾 Save Changes")
        btn_save.setStyleSheet(
            "font-weight: bold; background-color: #0078D7; color: white; padding: 6px 15px; border-radius: 4px;")
        btn_cancel = QPushButton("Cancel")
        btn_cancel.setStyleSheet("color: #333;")

        btn_save.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_save)
        layout.addLayout(btn_layout)

    def get_data(self):
        return {
            "tdoc_id": self.txt_tdoc.text().strip(),
            "agenda_item": self.txt_ai.text().strip(),
            "company": self.txt_company.text().strip(),
            "sender_name": self.txt_sender_name.text().strip(),
            "sender_email": self.txt_sender_email.text().strip(),
            "subject": self.txt_subject.text().strip(),
            "short_text": self.txt_short_text.toPlainText().strip()
        }
