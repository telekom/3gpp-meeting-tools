# --- File: src/modules/puml2visio/ui/elements_palette.py ---
"""Collapsible Elements & Syntax Palette for the PlantUML Mini-IDE.

Provides a visual, searchable catalog of PlantUML syntax elements, logic blocks,
3GPP architecture presets, user-defined presets, and arrows with live syntax preview
and one-click insertion.
"""

from typing import List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAction,
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    TREE_STYLE_CLEAN,
)
from modules.puml2visio.templates.plantuml_snippets import (
    PLANTUML_SNIPPETS,
    PlantUMLSnippet,
    delete_user_snippet,
    register_snippets_reloaded_callback,
    reload_snippets,
)
from modules.puml2visio.ui.preset_dialog import (
    UserPresetDialog,
    open_presets_folder,
)


class ElementsPaletteWidget(QFrame):
    """Visual searchable sidebar palette for PlantUML elements and snippets."""
    snippet_triggered = pyqtSignal(object)  # Emits PlantUMLSnippet

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("cardFrame")
        self.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #F8FAFC; 
                border: 1px solid #CBD5E1; 
                border-radius: 6px;
            }
        """)
        self._current_scope = "Sequence"
        self._setup_ui()
        self._populate_tree()

        # Listen for global presets reload events
        register_snippets_reloaded_callback(self._populate_tree)

    def _setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # Header Title and Action Buttons
        title_layout = QHBoxLayout()
        title_lbl = QLabel("📚 Syntax & Element Palette")
        title_lbl.setStyleSheet("font-weight: bold; font-size: 13px; color: #1E293B;")
        title_layout.addWidget(title_lbl)
        title_layout.addStretch()

        # Header action buttons: Add User Preset, Reload, Open Config Folder
        btn_add = QPushButton("➕")
        btn_add.setToolTip("Add custom user-defined preset (saved to Git-ignored config)")
        btn_add.setStyleSheet("padding: 2px 6px; font-size: 11px;")
        btn_add.clicked.connect(self._open_add_preset_dialog)
        title_layout.addWidget(btn_add)

        btn_reload = QPushButton("🔄")
        btn_reload.setToolTip("Reload presets from disk (plantuml_presets.json and user presets)")
        btn_reload.setStyleSheet("padding: 2px 6px; font-size: 11px;")
        btn_reload.clicked.connect(self._reload_presets)
        title_layout.addWidget(btn_reload)

        btn_folder = QPushButton("📂")
        btn_folder.setToolTip("Open presets configuration folder in Explorer")
        btn_folder.setStyleSheet("padding: 2px 6px; font-size: 11px;")
        btn_folder.clicked.connect(open_presets_folder)
        title_layout.addWidget(btn_folder)

        main_layout.addLayout(title_layout)

        # Search Bar
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter elements (e.g. UE, AMF, note, arrow)...")
        self.search_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #CBD5E1; 
                border-radius: 4px; 
                padding: 5px 8px; 
                background: white; 
                font-size: 12px;
            }
            QLineEdit:focus {
                border-color: #395396;
            }
        """)
        self.search_input.textChanged.connect(self._filter_items)
        main_layout.addWidget(self.search_input)

        # Scope Checkbox
        self.scope_check = QCheckBox("Filter by current diagram type")
        self.scope_check.setChecked(True)
        self.scope_check.setStyleSheet("color: #475569; font-size: 11px; font-weight: 500;")
        self.scope_check.toggled.connect(self._populate_tree)
        main_layout.addWidget(self.scope_check)

        # Splitter between Tree and Preview
        splitter = QSplitter(Qt.Vertical)
        splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #E2E8F0;
                height: 3px;
                margin: 2px 0px;
            }
        """)

        # Tree Widget of Snippets
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setStyleSheet(TREE_STYLE_CLEAN + """
            QTreeWidget {
                border: 1px solid #E2E8F0;
                font-size: 12px;
            }
            QTreeWidget::item {
                padding: 4px 6px;
            }
        """)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)
        self.tree.itemSelectionChanged.connect(self._on_item_selected)
        self.tree.itemDoubleClicked.connect(self._on_item_double_clicked)
        splitter.addWidget(self.tree)

        # Bottom Preview & Tip Card
        preview_container = QWidget()
        preview_layout = QVBoxLayout()
        preview_layout.setContentsMargins(0, 4, 0, 0)
        preview_layout.setSpacing(4)

        preview_header = QLabel("Preview & Usage:")
        preview_header.setStyleSheet("font-weight: bold; font-size: 11px; color: #64748B;")
        preview_layout.addWidget(preview_header)

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(110)
        self.preview_text.setStyleSheet("""
            QTextEdit {
                font-family: Consolas, 'Courier New', monospace;
                font-size: 11px;
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 4px;
                color: #334155;
            }
        """)
        preview_layout.addWidget(self.preview_text)

        # Insert Button
        self.insert_btn = QPushButton("➕ Insert at Cursor")
        self.insert_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.insert_btn.setEnabled(False)
        self.insert_btn.clicked.connect(self._insert_selected)
        preview_layout.addWidget(self.insert_btn)

        preview_container.setLayout(preview_layout)
        splitter.addWidget(preview_container)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)

        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

    def set_diagram_scope(self, diagram_type: str):
        """Updates the current diagram scope (e.g. 'Sequence', 'State', etc.)"""
        self._current_scope = diagram_type
        if self.scope_check.isChecked():
            self._populate_tree()

    def _matches_scope(self, snippet: PlantUMLSnippet) -> bool:
        if not self.scope_check.isChecked():
            return True
        if "All" in snippet.diagram_scope:
            return True
        return self._current_scope in snippet.diagram_scope

    def _populate_tree(self):
        self.tree.clear()
        query = self.search_input.text().strip().lower()

        categories = {}
        for snip in PLANTUML_SNIPPETS:
            if not self._matches_scope(snip):
                continue
            if query:
                content = f"{snip.label} {snip.description} {snip.category} {snip.code}".lower()
                if query not in content:
                    continue
            categories.setdefault(snip.category, []).append(snip)

        for cat_name, snips in categories.items():
            cat_item = QTreeWidgetItem(self.tree)
            cat_item.setText(0, f"📁 {cat_name} ({len(snips)})")
            cat_item.setFlags(cat_item.flags() & ~Qt.ItemIsSelectable)
            cat_item.setData(0, Qt.UserRole, None)

            font = cat_item.font(0)
            font.setBold(True)
            cat_item.setFont(0, font)

            for snip in snips:
                child = QTreeWidgetItem(cat_item)
                if snip.is_user_defined:
                    icon = "⭐"
                    display_label = f"{snip.label} (Custom)"
                elif "3gpp" in snip.category.lower():
                    icon = "📦" if "Architecture" in snip.category else "📡"
                    display_label = snip.label
                else:
                    icon = "🧩"
                    display_label = snip.label

                child.setText(0, f"{icon} {display_label}")
                child.setData(0, Qt.UserRole, snip)
                child.setToolTip(0, snip.description)

            cat_item.setExpanded(True)

        self._on_item_selected()

    def _filter_items(self):
        self._populate_tree()

    def _on_item_selected(self):
        selected = self.tree.selectedItems()
        if not selected:
            self.preview_text.setPlainText("Select an element above to see syntax preview.")
            self.insert_btn.setEnabled(False)
            return

        snip: Optional[PlantUMLSnippet] = selected[0].data(0, Qt.UserRole)
        if not snip:
            self.preview_text.setPlainText("Category selected. Choose a specific element to preview.")
            self.insert_btn.setEnabled(False)
            return

        clean_code = snip.code.replace("{cursor}", "").replace("{/cursor}", "")
        desc = snip.description
        wrapper_info = "\n[Supports wrapping selected lines]" if snip.is_block_wrapper else ""
        user_info = "\n[⭐ User-Defined Preset - stored in plantuml_presets_user.json]" if snip.is_user_defined else ""
        self.preview_text.setPlainText(f"/* {desc}{wrapper_info}{user_info} */\n{clean_code}")
        self.insert_btn.setEnabled(True)

    def _on_item_double_clicked(self, item: QTreeWidgetItem, column: int):
        snip: Optional[PlantUMLSnippet] = item.data(0, Qt.UserRole)
        if snip:
            self.snippet_triggered.emit(snip)

    def _insert_selected(self):
        selected = self.tree.selectedItems()
        if selected:
            snip: Optional[PlantUMLSnippet] = selected[0].data(0, Qt.UserRole)
            if snip:
                self.snippet_triggered.emit(snip)

    def _open_add_preset_dialog(self):
        dialog = UserPresetDialog(parent=self)
        if dialog.exec_():
            self._populate_tree()

    def _reload_presets(self):
        reload_snippets()
        self._populate_tree()

    def _show_tree_context_menu(self, pos):
        item = self.tree.itemAt(pos)
        menu = QMenu(self)

        if item:
            snip: Optional[PlantUMLSnippet] = item.data(0, Qt.UserRole)
            if snip:
                insert_act = menu.addAction("➕ Insert at Cursor")
                insert_act.triggered.connect(lambda: self.snippet_triggered.emit(snip))

                menu.addSeparator()

                if snip.is_user_defined:
                    edit_act = menu.addAction("✏️ Edit User Preset...")
                    edit_act.triggered.connect(lambda: self._edit_user_snippet(snip))

                    del_act = menu.addAction("🗑️ Delete User Preset")
                    del_act.triggered.connect(lambda: self._delete_user_snippet(snip))
                    menu.addSeparator()

        add_act = menu.addAction("➕ Add Custom Preset...")
        add_act.triggered.connect(self._open_add_preset_dialog)

        reload_act = menu.addAction("🔄 Reload Presets from Disk")
        reload_act.triggered.connect(self._reload_presets)

        folder_act = menu.addAction("📂 Open Config Folder")
        folder_act.triggered.connect(open_presets_folder)

        menu.exec_(self.tree.mapToGlobal(pos))

    def _edit_user_snippet(self, snip: PlantUMLSnippet):
        dialog = UserPresetDialog(existing_snippet=snip, parent=self)
        if dialog.exec_():
            self._populate_tree()

    def _delete_user_snippet(self, snip: PlantUMLSnippet):
        confirm = QMessageBox.question(
            self,
            "Delete Preset",
            f"Are you sure you want to delete preset '{snip.label}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if confirm == QMessageBox.Yes:
            delete_user_snippet(snip.id)
            self._populate_tree()
