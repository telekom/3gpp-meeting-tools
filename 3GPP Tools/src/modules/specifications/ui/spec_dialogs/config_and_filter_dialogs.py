"""Configuration, sync, and filter dialogs for Specifications module."""

import os
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from core.ui.theme import C
from modules.specifications.core.database import SpecsDatabase


class SpecsConfigDialog(QDialog):
    """Configuration dialog for setting and verifying the specifications download directory."""

    def __init__(self, current_path: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Specifications Settings")
        self.setModal(True)
        self.resize(520, 160)
        self.selected_path = current_path

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        info_label = QLabel(
            "Specify the local directory where downloaded 3GPP specifications, "
            "ZIP archives, and converted documents are stored:"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")
        layout.addWidget(info_label)

        path_layout = QHBoxLayout()
        path_layout.setSpacing(6)

        self.path_input = QLineEdit(current_path)
        self.path_input.setPlaceholderText("Select folder...")

        browse_btn = QPushButton("📂 Browse...")
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.clicked.connect(self._browse)

        open_btn = QPushButton("↗️ Open")
        open_btn.setCursor(Qt.PointingHandCursor)
        open_btn.setToolTip("Open this directory in Windows Explorer")
        open_btn.clicked.connect(self._open_folder)

        path_layout.addWidget(self.path_input)
        path_layout.addWidget(browse_btn)
        path_layout.addWidget(open_btn)
        layout.addLayout(path_layout)

        layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.clicked.connect(self.reject)

        save_btn = QPushButton("💾 Save Settings")
        save_btn.setCursor(Qt.PointingHandCursor)
        save_btn.setStyleSheet(f"""
            QPushButton {{
                font-weight: bold;
                background-color: {C.ACCENT_BRIGHT};
                color: {C.TEXT_ON_ACCENT};
                padding: 6px 16px;
                border-radius: 4px;
                border: 1px solid {C.ACCENT_BRIGHT_BORDER};
            }}
            QPushButton:hover {{
                background-color: {C.ACCENT_BRIGHT_BORDER};
            }}
        """)
        save_btn.clicked.connect(self._save_and_accept)

        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def _browse(self):
        new_dir = QFileDialog.getExistingDirectory(
            self, "Select Download Directory", self.path_input.text().strip()
        )
        if new_dir:
            self.path_input.setText(new_dir)

    def _open_folder(self):
        p = Path(self.path_input.text().strip())
        if not p.exists():
            try:
                p.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                QMessageBox.warning(self, "Directory Error", f"Could not create directory:\n{e}")
                return
        try:
            os.startfile(str(p))
        except Exception as e:
            QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")

    def _save_and_accept(self):
        self.selected_path = self.path_input.text().strip()
        self.accept()

    def get_download_path(self) -> str:
        return self.selected_path


class AdvancedSyncDialog(QDialog):
    """Network Database Sync Dialog with Strict Drop-Down Menus."""

    def __init__(self, db: SpecsDatabase, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Advanced Filtered Sync")
        self.setModal(True)
        self.resize(450, 250)
        self.matching_specs = []

        options = db.get_filter_options()

        layout = QVBoxLayout(self)
        info_label = QLabel(
            "Note: Filters apply to specifications already discovered in your local database. "
            "To discover brand new specifications, run a 'Full Sync' first."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet(f"color: {C.TEXT_MUTED}; font-style: italic; margin-bottom: 10px;")
        layout.addWidget(info_label)

        form = QFormLayout()

        self.series_combo = QComboBox()
        self.series_combo.addItem("Any")
        self.series_combo.addItems(options["series"])

        self.tech_combo = QComboBox()
        self.tech_combo.addItem("Any")
        self.tech_combo.addItems(options["techs"])

        self.group_combo = QComboBox()
        self.group_combo.addItem("Any")
        self.group_combo.addItems(options["groups"])

        self.type_combo = QComboBox()
        self.type_combo.addItem("Any")
        self.type_combo.addItems(options["types"])

        form.addRow("Series:", self.series_combo)
        form.addRow("Radio Tech:", self.tech_combo)
        form.addRow("Working Group:", self.group_combo)
        form.addRow("Type:", self.type_combo)
        layout.addLayout(form)

        self.count_label = QLabel("Matching specifications: 0")
        self.count_label.setStyleSheet(f"font-weight: bold; color: {C.ACCENT_INDIGO_BLUE}; margin-top: 10px;")
        layout.addWidget(self.count_label)

        btn_layout = QHBoxLayout()
        self.sync_btn = QPushButton("🚀 Start Sync")
        self.sync_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(self.sync_btn)
        layout.addLayout(btn_layout)

        self.series_combo.currentTextChanged.connect(self.update_count)
        self.tech_combo.currentTextChanged.connect(self.update_count)
        self.group_combo.currentTextChanged.connect(self.update_count)
        self.type_combo.currentTextChanged.connect(self.update_count)

        self.update_count()

    def update_count(self):
        series = "" if self.series_combo.currentText() == "Any" else self.series_combo.currentText()
        tech = "" if self.tech_combo.currentText() == "Any" else self.tech_combo.currentText()
        group = "" if self.group_combo.currentText() == "Any" else self.group_combo.currentText()
        spec_type = self.type_combo.currentText()

        self.matching_specs = self.db.get_filtered_specs(series, tech, group, spec_type)
        count = len(self.matching_specs)
        self.count_label.setText(f"Matching specifications in local DB: {count}")
        self.sync_btn.setEnabled(count > 0)


class TableFilterDialog(QDialog):
    """Local Table Filter Dialog with Strict Drop-Down Menus."""

    def __init__(self, db: SpecsDatabase, current_filters: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Filter Specifications")
        self.setModal(True)
        self.resize(350, 200)

        options = db.get_filter_options()

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.series_combo = QComboBox()
        self.series_combo.addItem("Any")
        self.series_combo.addItems(options["series"])
        self.series_combo.setCurrentText(current_filters.get("series", "Any") or "Any")

        self.tech_combo = QComboBox()
        self.tech_combo.addItem("Any")
        self.tech_combo.addItems(options["techs"])
        self.tech_combo.setCurrentText(current_filters.get("tech", "Any") or "Any")

        self.group_combo = QComboBox()
        self.group_combo.addItem("Any")
        self.group_combo.addItems(options["groups"])
        self.group_combo.setCurrentText(current_filters.get("group", "Any") or "Any")

        self.type_combo = QComboBox()
        self.type_combo.addItem("Any")
        self.type_combo.addItems(options["types"])
        self.type_combo.setCurrentText(current_filters.get("spec_type", "Any") or "Any")

        form.addRow("Series:", self.series_combo)
        form.addRow("Radio Tech:", self.tech_combo)
        form.addRow("Working Group:", self.group_combo)
        form.addRow("Type:", self.type_combo)
        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        apply_btn = QPushButton("✅ Apply Filters")
        apply_btn.clicked.connect(self.accept)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_and_accept)

        btn_layout.addStretch()
        btn_layout.addWidget(clear_btn)
        btn_layout.addWidget(apply_btn)
        layout.addLayout(btn_layout)

    def _clear_and_accept(self):
        self.series_combo.setCurrentText("Any")
        self.tech_combo.setCurrentText("Any")
        self.group_combo.setCurrentText("Any")
        self.type_combo.setCurrentText("Any")
        self.accept()

    def get_filters(self) -> dict:
        return {
            "series": (
                "" if self.series_combo.currentText() == "Any" else self.series_combo.currentText()
            ),
            "tech": (
                "" if self.tech_combo.currentText() == "Any" else self.tech_combo.currentText()
            ),
            "group": (
                "" if self.group_combo.currentText() == "Any" else self.group_combo.currentText()
            ),
            "spec_type": self.type_combo.currentText(),
        }


class TargetedSyncDialog(QDialog):
    """Dialog for fetching brand new specifications directly by number or series."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🎯 Quick Fetch Specification")
        self.setModal(True)
        self.resize(350, 150)

        layout = QVBoxLayout(self)

        info_label = QLabel(
            "Enter a specific specification (e.g., <b>23.801-01</b>) or an entire series (e.g., <b>23</b>) "
            "to fetch directly from 3GPP.<br><br><i>You can separate multiple targets with commas.</i>"
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("e.g., 23, 38.331, 23.501")
        layout.addWidget(self.input_field)

        btn_layout = QHBoxLayout()
        self.fetch_btn = QPushButton("🚀 Fetch Now")
        self.fetch_btn.clicked.connect(self.accept)
        self.fetch_btn.setEnabled(False)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(self.fetch_btn)
        layout.addLayout(btn_layout)

        self.input_field.textChanged.connect(lambda t: self.fetch_btn.setEnabled(bool(t.strip())))

    def get_targets(self) -> list:
        raw_text = self.input_field.text()
        return [t.strip() for t in raw_text.split(",") if t.strip()]
