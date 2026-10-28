"""Card 4: 3GPP Stencils & Architecture Reference."""

import logging
from typing import Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor, QCursor, QFont
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    TABLE_STYLE_CLEAN,
)
from modules.puml2visio.templates.plantuml_snippets import (
    PLANTUML_SNIPPETS,
    PlantUMLSnippet,
    register_snippets_reloaded_callback,
    reload_snippets,
)
from modules.puml2visio.ui.preset_dialog import (
    UserPresetDialog,
    open_presets_folder,
)

logger = logging.getLogger(__name__)


class StencilsCard(QWidget):
    """Card 4: Interactive catalog of standardized 5GC, NG-RAN, EPS, and IMS presets."""

    insert_snippet_requested = pyqtSignal(object)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Header description
        desc_lbl = QLabel(
            "<b>3GPP Architecture Presets & Stencils:</b> Quick-insert standardized 5G Core, NG-RAN, EPS, and IMS "
            "Network Function definitions into your PlantUML diagrams. Supports custom user-defined presets "
            "(saved to Git-ignored config)."
        )
        desc_lbl.setStyleSheet("color: #334155; font-size: 12px;")
        layout.addWidget(desc_lbl)

        # Toolbar: Filter search, category combo, add preset, reload, open folder
        tb_layout = QHBoxLayout()
        tb_layout.setSpacing(8)

        self.stencils_search_edit = QLineEdit()
        self.stencils_search_edit.setPlaceholderText("Filter stencils (e.g. UE, AMF, gNB, 5GS)...")
        self.stencils_search_edit.setStyleSheet("padding: 4px 8px; font-size: 11px; border: 1px solid #CBD5E1; border-radius: 4px;")
        self.stencils_search_edit.textChanged.connect(self._populate_stencils_table)
        tb_layout.addWidget(self.stencils_search_edit, stretch=2)

        self.stencils_cat_combo = QComboBox()
        self.stencils_cat_combo.addItems([
            "All Categories",
            "3GPP Network Entities & NFs",
            "3GPP Architecture Presets",
            "Participants & Grouping",
            "⭐ User Defined Presets",
        ])
        self.stencils_cat_combo.setStyleSheet("padding: 4px 8px; font-size: 11px;")
        self.stencils_cat_combo.currentIndexChanged.connect(self._populate_stencils_table)
        tb_layout.addWidget(self.stencils_cat_combo, stretch=1)

        btn_add = QPushButton("➕ Add User Preset...")
        btn_add.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_add.setToolTip("Create a custom PlantUML preset saved to Git-ignored configuration")
        btn_add.clicked.connect(self._open_add_stencil_dialog)
        tb_layout.addWidget(btn_add)

        btn_reload = QPushButton("🔄 Reload")
        btn_reload.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_reload.setToolTip("Reload presets from disk (plantuml_presets.json and user presets)")
        btn_reload.clicked.connect(self._reload_stencils)
        tb_layout.addWidget(btn_reload)

        btn_folder = QPushButton("📂 Open Config")
        btn_folder.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_folder.setToolTip("Open configuration folder in Windows Explorer")
        btn_folder.clicked.connect(open_presets_folder)
        tb_layout.addWidget(btn_folder)

        layout.addLayout(tb_layout)

        # Stencils Table / List
        self.stencils_table = QTableWidget(0, 4)
        self.stencils_table.setHorizontalHeaderLabels(["Category", "Preset Name", "Description", "Actions"])
        self.stencils_table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.stencils_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.stencils_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        s_header = self.stencils_table.horizontalHeader()
        s_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(2, QHeaderView.Stretch)
        s_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        # Populate with 3GPP snippets
        self._populate_stencils_table()

        # Listen for global reload
        register_snippets_reloaded_callback(self._populate_stencils_table)

        layout.addWidget(self.stencils_table, stretch=1)

    def _open_add_stencil_dialog(self):
        dialog = UserPresetDialog(parent=self)
        if dialog.exec_():
            self._populate_stencils_table()

    def _reload_stencils(self):
        reload_snippets()
        self._populate_stencils_table()

    def _edit_stencil_dialog(self, snippet: PlantUMLSnippet):
        dialog = UserPresetDialog(parent=self, edit_snippet=snippet)
        if dialog.exec_():
            self._populate_stencils_table()

    def _populate_stencils_table(self):
        self.stencils_table.setRowCount(0)
        query = self.stencils_search_edit.text().strip().lower() if hasattr(self, 'stencils_search_edit') else ""
        selected_cat = self.stencils_cat_combo.currentText() if hasattr(self, 'stencils_cat_combo') else "All Categories"

        relevant_snippets = []
        for s in PLANTUML_SNIPPETS:
            # Filter by category combo
            if selected_cat == "⭐ User Defined Presets":
                if not s.is_user_defined:
                    continue
            elif selected_cat != "All Categories":
                if s.category != selected_cat:
                    continue
            else:
                # In All Categories mode, show 3GPP and participant stencils plus any user-defined
                if not ("3gpp" in s.category.lower() or "participant" in s.category.lower() or s.is_user_defined):
                    continue

            # Filter by search text
            if query:
                combined = f"{s.label} {s.category} {s.description} {s.code}".lower()
                if query not in combined:
                    continue

            relevant_snippets.append(s)

        for s in relevant_snippets:
            row = self.stencils_table.rowCount()
            self.stencils_table.insertRow(row)

            cat_item = QTableWidgetItem(s.category)
            display_name = f"⭐ {s.label} [User]" if s.is_user_defined else s.label
            name_item = QTableWidgetItem(display_name)
            name_item.setFont(QFont("Segoe UI", 9, QFont.Bold))
            if s.is_user_defined:
                name_item.setForeground(QColor("#B45309"))

            desc_item = QTableWidgetItem(s.description)
            desc_item.setToolTip(s.code)

            self.stencils_table.setItem(row, 0, cat_item)
            self.stencils_table.setItem(row, 1, name_item)
            self.stencils_table.setItem(row, 2, desc_item)

            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(2, 2, 2, 2)
            action_layout.setSpacing(4)

            btn_copy = QPushButton("📋 Copy")
            btn_copy.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            btn_copy.setToolTip("Copy this PlantUML snippet to clipboard")
            btn_copy.clicked.connect(lambda checked, snip=s: self._copy_snippet_to_clipboard(snip))
            action_layout.addWidget(btn_copy)

            btn_insert = QPushButton("📝 Insert in Editor")
            btn_insert.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            btn_insert.setToolTip("Switch to PlantUML editor and insert this element at the cursor")
            btn_insert.clicked.connect(lambda checked, snip=s: self._insert_snippet_into_editor(snip))
            action_layout.addWidget(btn_insert)

            if s.is_user_defined:
                btn_edit = QPushButton("✏️ Edit")
                btn_edit.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
                btn_edit.setToolTip("Edit this user-defined preset")
                btn_edit.clicked.connect(lambda checked, snip=s: self._edit_stencil_dialog(snip))
                action_layout.addWidget(btn_edit)

            self.stencils_table.setCellWidget(row, 3, action_widget)

    def _copy_snippet_to_clipboard(self, snippet: PlantUMLSnippet):
        QApplication.clipboard().setText(snippet.code)
        QToolTip.showText(QCursor.pos(), f"📋 Copied '{snippet.label}' snippet!", self, msecShowTime=1800)

    def _insert_snippet_into_editor(self, snippet: PlantUMLSnippet):
        self.insert_snippet_requested.emit(snippet)
