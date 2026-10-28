# --- File: src/modules/puml2visio/ui/preset_dialog.py ---
"""Dialog for creating, editing, and managing user-defined PlantUML presets/snippets.

User presets are persisted to `src/config/plantuml_presets_user.json` (untracked in Git)
or `APP_DATA_ROOT/plantuml_presets_user.json`, allowing delegates to add company-specific
network functions, architectures, and signaling blocks without touching python code.
"""

import logging
import os
from pathlib import Path
import re
from typing import Optional

from PyQt5.QtCore import Qt, QUrl, pyqtSignal
from PyQt5.QtGui import QDesktopServices, QFont
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
)

BUTTON_STYLE_PRIMARY = """
    QPushButton {
        background-color: #2563EB;
        color: #FFFFFF;
        font-weight: bold;
        border: 1px solid #1D4ED8;
        border-radius: 4px;
        padding: 5px 14px;
        font-size: 12px;
    }
    QPushButton:hover {
        background-color: #1D4ED8;
    }
    QPushButton:pressed {
        background-color: #1E40AF;
    }
"""
from modules.puml2visio.templates.plantuml_snippets import (
    PlantUMLSnippet,
    add_user_snippet,
    delete_user_snippet,
    get_default_presets_path,
    get_user_presets_path,
    reload_snippets,
)


def open_presets_folder():
    """Opens the configuration folder containing plantuml_presets.json and plantuml_presets_user.json."""
    config_dir = get_default_presets_path().parent
    config_dir.mkdir(parents=True, exist_ok=True)
    try:
        if os.name == "nt":
            os.startfile(str(config_dir))
        else:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(config_dir)))
    except Exception as e:
        logging.error(f"Failed to open presets folder: {e}")


def open_user_presets_file():
    """Opens the user presets file in the system default editor (e.g. VS Code, Notepad)."""
    user_path = get_user_presets_path()
    if not user_path.exists():
        user_path.parent.mkdir(parents=True, exist_ok=True)
        with open(user_path, "w", encoding="utf-8") as f:
            f.write('{\n    "snippets": []\n}\n')
    try:
        if os.name == "nt":
            os.startfile(str(user_path))
        else:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(user_path)))
    except Exception as e:
        logging.error(f"Failed to open user presets file: {e}")


class UserPresetDialog(QDialog):
    """Modal dialog to add or edit user-defined PlantUML presets."""
    preset_saved = pyqtSignal(object)  # Emits saved PlantUMLSnippet

    def __init__(self, existing_snippet: Optional[PlantUMLSnippet] = None, parent=None):
        super().__init__(parent)
        self.snippet = existing_snippet
        self.is_edit_mode = existing_snippet is not None
        self.setWindowTitle("PlantUML Preset Configuration" if self.is_edit_mode else "➕ Add Custom PlantUML Preset")
        self.resize(620, 560)
        self._setup_ui()
        if self.snippet:
            self._populate_fields(self.snippet)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        # Header Info Banner
        header_frame = QFrame()
        header_frame.setStyleSheet("""
            QFrame {
                background-color: #F1F5F9;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        h_layout = QVBoxLayout(header_frame)
        h_layout.setContentsMargins(4, 4, 4, 4)
        h_layout.setSpacing(4)

        title_lbl = QLabel(
            "<b>✏️ Edit User Preset</b>" if self.is_edit_mode else "<b>➕ Define Custom PlantUML Preset</b>"
        )
        title_lbl.setStyleSheet("color: #1E293B; font-size: 13px;")
        desc_lbl = QLabel(
            "User presets are saved to <code>src/config/plantuml_presets_user.json</code> "
            "(excluded from Git commits) and immediately available across the PlantUML editor, "
            "elements palette, context menu, and stencils table."
        )
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet("color: #475569; font-size: 11px;")

        h_layout.addWidget(title_lbl)
        h_layout.addWidget(desc_lbl)
        main_layout.addWidget(header_frame)

        # Form Fields
        form_layout = QFormLayout()
        form_layout.setSpacing(10)
        form_layout.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)

        # Preset Label
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("e.g. My Company Private 5G Core, gNB-DU, Sidelink Relay...")
        self.label_edit.setStyleSheet("padding: 5px 8px; font-size: 12px;")
        form_layout.addRow("<b>Preset Name:</b>", self.label_edit)

        # Category
        self.category_combo = QComboBox()
        self.category_combo.setEditable(True)
        categories = [
            "3GPP Network Entities & NFs",
            "3GPP Architecture Presets",
            "Participants & Grouping",
            "Messages & Arrows",
            "Logic & Grouping",
            "Notes & Lifelines",
            "User Defined Presets",
            "State Diagram",
            "Activity Diagram",
            "Timing Diagram",
            "Packet Diagram",
            "Formatting & Directives",
        ]
        self.category_combo.addItems(categories)
        self.category_combo.setCurrentText("3GPP Network Entities & NFs")
        form_layout.addRow("<b>Category:</b>", self.category_combo)

        # Diagram Scope
        self.scope_combo = QComboBox()
        self.scope_combo.setEditable(True)
        scopes = ["Sequence", "All", "State", "Activity (Beta)", "Timing", "Packet (packetdiag)", "Component", "Class"]
        self.scope_combo.addItems(scopes)
        self.scope_combo.setCurrentText("Sequence")
        form_layout.addRow("<b>Diagram Scope:</b>", self.scope_combo)

        # Description
        self.desc_edit = QLineEdit()
        self.desc_edit.setPlaceholderText("Short description of what this preset defines or inserts")
        self.desc_edit.setStyleSheet("padding: 5px 8px; font-size: 12px;")
        form_layout.addRow("<b>Description:</b>", self.desc_edit)

        main_layout.addLayout(form_layout)

        # Code Editor
        code_lbl = QLabel("<b>PlantUML Template Code:</b>")
        code_lbl.setStyleSheet("color: #1E293B; font-size: 12px;")
        main_layout.addWidget(code_lbl)

        self.code_edit = QPlainTextEdit()
        self.code_edit.setPlaceholderText(
            "participant \"MyEntity\" as ME\n"
            "{cursor}\n\n"
            "' Dynamic tokens supported:\n"
            "' {cursor}     : Cursor destination after insertion\n"
            "' {selection}  : Selected text block if wrapping\n"
            "' {p1}, {p2}   : Auto-detected first/second participant"
        )
        self.code_edit.setFont(QFont("Consolas", 10))
        self.code_edit.setStyleSheet("""
            QPlainTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 6px;
                font-family: Consolas, 'Courier New', monospace;
                font-size: 12px;
                color: #0F172A;
            }
        """)
        main_layout.addWidget(self.code_edit, stretch=1)

        # Checkbox: is_block_wrapper
        self.wrapper_check = QCheckBox("Wraps selected text (replaces {selection} with indented user selection)")
        self.wrapper_check.setStyleSheet("color: #334155; font-size: 11px;")
        main_layout.addWidget(self.wrapper_check)

        # Bottom Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        # Config folder link button
        self.open_folder_btn = QPushButton("📂 Open Config Folder")
        self.open_folder_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.open_folder_btn.setToolTip("Open src/config in Windows Explorer to view or edit JSON directly")
        self.open_folder_btn.clicked.connect(open_presets_folder)
        btn_layout.addWidget(self.open_folder_btn)

        if self.is_edit_mode and self.snippet and self.snippet.is_user_defined:
            self.delete_btn = QPushButton("🗑️ Delete Preset")
            self.delete_btn.setStyleSheet("background-color: #FEE2E2; color: #DC2626; border: 1px solid #FCA5A5; padding: 5px 12px; border-radius: 4px;")
            self.delete_btn.clicked.connect(self._on_delete)
            btn_layout.addWidget(self.delete_btn)

        btn_layout.addStretch()

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.cancel_btn)

        self.save_btn = QPushButton("💾 Save Preset")
        self.save_btn.setStyleSheet(BUTTON_STYLE_PRIMARY)
        self.save_btn.clicked.connect(self._on_save)
        btn_layout.addWidget(self.save_btn)

        main_layout.addLayout(btn_layout)

    def _populate_fields(self, snippet: PlantUMLSnippet):
        self.label_edit.setText(snippet.label)
        self.category_combo.setCurrentText(snippet.category)
        scope_str = snippet.diagram_scope[0] if snippet.diagram_scope else "All"
        self.scope_combo.setCurrentText(scope_str)
        self.desc_edit.setText(snippet.description)
        self.code_edit.setPlainText(snippet.code)
        self.wrapper_check.setChecked(snippet.is_block_wrapper)

    def _on_save(self):
        label = self.label_edit.text().strip()
        code = self.code_edit.toPlainText().strip()
        if not label:
            QMessageBox.warning(self, "Validation Error", "Please provide a name for this preset.")
            self.label_edit.setFocus()
            return

        if not code:
            QMessageBox.warning(self, "Validation Error", "Please provide the PlantUML template code.")
            self.code_edit.setFocus()
            return

        category = self.category_combo.currentText().strip() or "User Defined Presets"
        scope = self.scope_combo.currentText().strip() or "All"
        description = self.desc_edit.text().strip() or f"User defined preset: {label}"
        is_wrapper = self.wrapper_check.isChecked()

        # Generate stable ID
        if self.snippet:
            snip_id = self.snippet.id
        else:
            safe_suffix = re.sub(r"[^a-zA-Z0-9_]+", "_", label.lower()).strip("_")
            snip_id = f"user_{safe_suffix}"

        new_snippet = PlantUMLSnippet(
            id=snip_id,
            label=label,
            category=category,
            diagram_scope=(scope,),
            description=description,
            code=code,
            is_block_wrapper=is_wrapper,
            is_user_defined=True,
        )

        try:
            add_user_snippet(new_snippet)
            self.preset_saved.emit(new_snippet)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", f"Could not save user preset: {e}")

    def _on_delete(self):
        if not self.snippet:
            return
        confirm = QMessageBox.question(
            self,
            "Confirm Deletion",
            f"Are you sure you want to delete user preset '{self.snippet.label}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm == QMessageBox.Yes:
            try:
                delete_user_snippet(self.snippet.id)
                self.accept()
            except Exception as e:
                QMessageBox.critical(self, "Deletion Failed", f"Could not delete user preset: {e}")
