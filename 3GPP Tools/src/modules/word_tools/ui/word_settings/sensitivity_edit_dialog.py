"""Modal sub-dialog to create or edit an enterprise sensitivity label preset."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from core.ui.ui_components import SEARCH_INPUT_STYLE


class SensitivityLabelEditDialog(QDialog):
    """Dialog to edit or create a sensitivity label configuration."""

    def __init__(
        self,
        parent=None,
        name: str = "",
        guid: str = "",
        description: str = "",
        display_name: str = "",
        is_encrypted: bool = False,
    ):
        super().__init__(parent)
        self.setWindowTitle("✏️ Edit Sensitivity Label" if name else "➕ Add Sensitivity Label")
        self.setMinimumWidth(480)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self._display_name = display_name

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        form = QFormLayout()
        form.setSpacing(8)

        self.edit_name = QLineEdit(name)
        self.edit_name.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_name.setPlaceholderText("e.g. PUBLIC or INTERNAL")
        form.addRow("Label Name:", self.edit_name)

        self.edit_guid = QLineEdit(guid)
        self.edit_guid.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_guid.setPlaceholderText("e.g. 00000000-0000-0000-0000-000000000000")
        form.addRow("Label GUID / ID:", self.edit_guid)

        self.edit_desc = QLineEdit(description)
        self.edit_desc.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_desc.setPlaceholderText("e.g. Public / Unclassified")
        form.addRow("Description:", self.edit_desc)

        self.chk_encrypted = QCheckBox("RMS Encrypted label (🔒 warns when selected)")
        self.chk_encrypted.setChecked(bool(is_encrypted))
        form.addRow("", self.chk_encrypted)

        layout.addLayout(form)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(self.btn_cancel)

        self.btn_save = QPushButton("Save Label")
        self.btn_save.setObjectName("primaryBtn")
        self.btn_save.clicked.connect(self._validate_and_accept)
        btn_row.addWidget(self.btn_save)

        layout.addLayout(btn_row)

    def _validate_and_accept(self):
        if not self.edit_name.text().strip():
            QMessageBox.warning(self, "Invalid Name", "Label display name cannot be empty.")
            return
        if not self.edit_guid.text().strip():
            QMessageBox.warning(self, "Invalid GUID", "Label GUID cannot be empty.")
            return
        self.accept()

    def get_data(self) -> dict:
        name = self.edit_name.text().strip()
        return {
            "name": name,
            "display_name": self._display_name or name,
            "id": self.edit_guid.text().strip(),
            "description": self.edit_desc.text().strip(),
            "is_encrypted": self.chk_encrypted.isChecked(),
            "source": "custom",
        }
