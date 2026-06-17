# --- File: modules/specs_db/ui_tabs.py ---
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QLabel, QCheckBox
from PyQt5.QtCore import pyqtSignal


class SpecificationsTab(QWidget):
    update_db_requested = pyqtSignal(bool)

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        self.info_label = QLabel(
            "📚 3GPP Specifications Database\nSynchronize the local database with the 3GPP archive.")
        layout.addWidget(self.info_label)

        self.force_meta_checkbox = QCheckBox("Force Update Metadata (Scrape DynaReports for ALL specs)")
        self.force_meta_checkbox.setToolTip(
            "If unchecked, it will only scrape metadata for newly discovered specifications.")
        layout.addWidget(self.force_meta_checkbox)

        self.update_btn = QPushButton("🔄 Start Synchronization")
        self.update_btn.clicked.connect(self._on_update_clicked)
        layout.addWidget(self.update_btn)

        layout.addStretch()
        self.setLayout(layout)

    def _on_update_clicked(self):
        force_meta = self.force_meta_checkbox.isChecked()
        self.update_db_requested.emit(force_meta)