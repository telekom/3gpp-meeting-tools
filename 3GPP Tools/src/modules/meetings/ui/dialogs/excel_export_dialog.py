# --- File: src/modules/meetings/ui/excel_export_dialog.py ---
import json
import logging

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
)

from core.utils.paths import get_project_root

logger = logging.getLogger(__name__)

ALL_AVAILABLE_COLUMNS = [
    "TDoc",
    "Title",
    "Source",
    "Type",
    "Agenda Item",
    "Sub-Item",
    "TDoc Status",
    "For",
    "Secretary Remarks",
    "My Status",
    "My Notes",
    "Abstract",
    "Related TDocs"
]

DEFAULT_SELECTED_COLUMNS = [
    "TDoc",
    "Title",
    "Source",
    "Type",
    "Agenda Item",
    "TDoc Status"
]

REVIEW_MATRIX_COLUMNS = [
    "TDoc",
    "Title",
    "Source",
    "Type",
    "Agenda Item",
    "TDoc Status",
    "My Status",
    "My Notes"
]


def get_default_company_name() -> str:
    """Attempts to retrieve the user's default company profile from WordConfig."""
    try:
        from modules.word_tools.core.word_config import WordConfig
        comp = WordConfig.load().get("default_source_company", "")
        if comp:
            return comp.strip()
    except Exception:
        pass
    return "Deutsche Telekom"


class ExcelExportDialog(QDialog):
    def __init__(self, parent, visible_count: int, total_count: int):
        super().__init__(parent)
        self.setWindowTitle("📊 Export TDocs to Excel")
        self.resize(540, 680)
        self.setStyleSheet("QDialog { background-color: #FAFAFA; } QLabel { font-size: 12px; color: #333; }")

        self.visible_count = visible_count
        self.total_count = total_count

        self.config_path = get_project_root() / "config" / "export_config.json"
        self.saved_config = self._load_saved_config()

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # --- Scope Selection ---
        scope_group = QGroupBox("1. Select Export Scope")
        scope_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        scope_layout = QVBoxLayout(scope_group)

        self.radio_visible = QRadioButton(f"Visible / Filtered Rows only ({self.visible_count} TDocs)")
        self.radio_all = QRadioButton(f"All Meeting Rows ({self.total_count} TDocs)")

        self.radio_visible.setChecked(True if self.visible_count > 0 else False)
        self.radio_all.setChecked(True if self.visible_count == 0 else False)

        scope_layout.addWidget(self.radio_visible)
        scope_layout.addWidget(self.radio_all)
        layout.addWidget(scope_group)

        # --- Columns Selection & Reordering ---
        col_group = QGroupBox("2. Choose Columns && Drag to Reorder")
        col_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        col_main_layout = QHBoxLayout(col_group)

        self.col_list = QListWidget()
        self.col_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.col_list.setDefaultDropAction(Qt.MoveAction)
        self.col_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.col_list.setStyleSheet("""
            QListWidget { 
                background: white; border: 1px solid #CCC; border-radius: 4px; padding: 4px; font-size: 12px; 
            }
            QListWidget::item { padding: 4px; }
            QListWidget::item:hover { background-color: #F0F4F8; }
        """)
        self._populate_column_list()
        col_main_layout.addWidget(self.col_list)

        # Action buttons for reordering & presets
        btn_box = QVBoxLayout()
        btn_box.setSpacing(6)

        btn_style = "QPushButton { padding: 5px 10px; font-size: 11px; background: white; border: 1px solid #CCC; border-radius: 4px; } QPushButton:hover { background: #EAEAEA; }"
        btn_preset_style = "QPushButton { padding: 5px 8px; font-size: 11px; font-weight: bold; background: #FDF5F8; color: #E20074; border: 1px solid #E20074; border-radius: 4px; } QPushButton:hover { background: #FCE8F0; }"

        btn_up = QPushButton("▲ Move Up")
        btn_up.setToolTip('Move the selected export column one position upward.')
        btn_up.setStyleSheet(btn_style)
        btn_up.clicked.connect(self._move_item_up)

        btn_down = QPushButton("▼ Move Down")
        btn_down.setToolTip('Move the selected export column one position downward.')
        btn_down.setStyleSheet(btn_style)
        btn_down.clicked.connect(self._move_item_down)

        btn_all = QPushButton("☑️ Select All")
        btn_all.setToolTip('Select all available columns for Excel export.')
        btn_all.setStyleSheet(btn_style)
        btn_all.clicked.connect(lambda: self._set_all_checked(True))

        btn_none = QPushButton("◻️ Clear All")
        btn_none.setToolTip('Clear all selected Excel export columns.')
        btn_none.setStyleSheet(btn_style)
        btn_none.clicked.connect(lambda: self._set_all_checked(False))

        btn_review_preset = QPushButton("📋 Review Matrix")
        btn_review_preset.setToolTip('Preset: Select standard agenda columns plus My Status & My Notes with company attribution.')
        btn_review_preset.setStyleSheet(btn_preset_style)
        btn_review_preset.clicked.connect(self._apply_review_preset)

        btn_standard_preset = QPushButton("📑 Standard")
        btn_standard_preset.setToolTip('Restore the standard default meeting agenda column selection.')
        btn_standard_preset.setStyleSheet(btn_style)
        btn_standard_preset.clicked.connect(self._apply_standard_preset)

        btn_box.addWidget(btn_up)
        btn_box.addWidget(btn_down)
        btn_box.addSpacing(6)
        btn_box.addWidget(btn_all)
        btn_box.addWidget(btn_none)
        btn_box.addSpacing(6)
        btn_box.addWidget(btn_review_preset)
        btn_box.addWidget(btn_standard_preset)
        btn_box.addStretch()
        col_main_layout.addLayout(btn_box)

        layout.addWidget(col_group)

        # --- Delegation & Review Options ---
        review_group = QGroupBox("3. Delegation && Review Matrix Options")
        review_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        review_layout = QVBoxLayout(review_group)
        review_layout.setSpacing(6)

        self.chk_company_attribution = QCheckBox("Format review columns with company attribution")
        self.chk_company_attribution.setToolTip(
            "When checked, 'My Status' and 'My Notes' will be exported as company review columns,\n"
            "e.g. 'Status (Deutsche Telekom)' and 'Comments (Deutsche Telekom)'."
        )
        self.chk_company_attribution.setStyleSheet("font-size: 12px; font-weight: normal;")

        comp_row = QHBoxLayout()
        comp_label = QLabel("Company Name:")
        comp_label.setStyleSheet("font-weight: normal; font-size: 12px;")

        default_company = self.saved_config.get("company_name") or get_default_company_name()
        self.edit_company = QLineEdit(default_company)
        self.edit_company.setPlaceholderText("e.g. Deutsche Telekom, Qualcomm, Ericsson")
        self.edit_company.setStyleSheet("padding: 4px 8px; border: 1px solid #CCC; border-radius: 4px; font-size: 12px;")

        comp_row.addWidget(comp_label)
        comp_row.addWidget(self.edit_company)
        review_layout.addWidget(self.chk_company_attribution)
        review_layout.addLayout(comp_row)

        self.lbl_preview = QLabel()
        self.lbl_preview.setStyleSheet("font-size: 11px; color: #666; font-style: italic; margin-left: 20px;")
        review_layout.addWidget(self.lbl_preview)

        layout.addWidget(review_group)

        is_attr_enabled = self.saved_config.get("company_attribution_enabled", True)
        self.chk_company_attribution.setChecked(is_attr_enabled)
        self.chk_company_attribution.toggled.connect(self._update_company_ui)
        self.edit_company.textChanged.connect(self._update_company_ui)
        self._update_company_ui()

        # --- Options ---
        self.chk_auto_open = QCheckBox("Automatically open spreadsheet after export")
        self.chk_auto_open.setToolTip('Open the generated Excel workbook automatically after a successful export.')
        self.chk_auto_open.setChecked(self.saved_config.get("auto_open", True))
        self.chk_auto_open.setStyleSheet("font-size: 12px;")
        layout.addWidget(self.chk_auto_open)

        # --- Dialog Buttons ---
        btn_layout = QHBoxLayout()
        export_btn = QPushButton("📊 Export to Excel")
        export_btn.setToolTip('Export the chosen rows and columns to a formatted Excel workbook.')
        export_btn.setStyleSheet("padding: 7px 20px; font-weight: bold; background-color: #005A9E; color: white; border-radius: 4px;")
        export_btn.clicked.connect(self._on_export_clicked)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet("padding: 7px 15px; background-color: #E0E0E0; border: 1px solid #CCC; border-radius: 4px;")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(export_btn)
        layout.addLayout(btn_layout)

    def _update_company_ui(self):
        enabled = self.chk_company_attribution.isChecked()
        self.edit_company.setEnabled(enabled)
        company = self.edit_company.text().strip() or "Company"
        if enabled:
            self.lbl_preview.setText(f"Export headers: 'Status ({company})' and 'Comments ({company})'")
            self.lbl_preview.setStyleSheet("font-size: 11px; color: #E20074; font-style: italic; margin-left: 20px;")
        else:
            self.lbl_preview.setText("Export headers: 'My Status' and 'My Notes' (no company attribution)")
            self.lbl_preview.setStyleSheet("font-size: 11px; color: #666; font-style: italic; margin-left: 20px;")

    def _reorder_and_check(self, target_cols: list):
        new_order = list(target_cols)
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in new_order:
                new_order.append(c)

        self.col_list.clear()
        for col_name in new_order:
            item = QListWidgetItem(col_name)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable | Qt.ItemIsDragEnabled)
            item.setCheckState(Qt.Checked if col_name in target_cols else Qt.Unchecked)
            self.col_list.addItem(item)

    def _apply_review_preset(self):
        self._reorder_and_check(REVIEW_MATRIX_COLUMNS)
        self.chk_company_attribution.setChecked(True)
        self._update_company_ui()

    def _apply_standard_preset(self):
        self._reorder_and_check(DEFAULT_SELECTED_COLUMNS)

    def _populate_column_list(self):
        self.col_list.clear()
        saved_order = self.saved_config.get("column_order", ALL_AVAILABLE_COLUMNS)
        saved_checked = set(self.saved_config.get("checked_columns", DEFAULT_SELECTED_COLUMNS))

        # Add all saved columns preserving order
        ordered_cols = [c for c in saved_order if c in ALL_AVAILABLE_COLUMNS]
        # Append any missing columns that might have been added later
        for c in ALL_AVAILABLE_COLUMNS:
            if c not in ordered_cols:
                ordered_cols.append(c)

        for col_name in ordered_cols:
            item = QListWidgetItem(col_name)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable | Qt.ItemIsDragEnabled)
            item.setCheckState(Qt.Checked if col_name in saved_checked else Qt.Unchecked)
            self.col_list.addItem(item)

    def _move_item_up(self):
        row = self.col_list.currentRow()
        if row > 0:
            item = self.col_list.takeItem(row)
            self.col_list.insertItem(row - 1, item)
            self.col_list.setCurrentRow(row - 1)

    def _move_item_down(self):
        row = self.col_list.currentRow()
        if row >= 0 and row < self.col_list.count() - 1:
            item = self.col_list.takeItem(row)
            self.col_list.insertItem(row + 1, item)
            self.col_list.setCurrentRow(row + 1)

    def _set_all_checked(self, checked: bool):
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.col_list.count()):
            self.col_list.item(i).setCheckState(state)

    def _reset_defaults(self):
        self.saved_config = {
            "column_order": ALL_AVAILABLE_COLUMNS,
            "checked_columns": DEFAULT_SELECTED_COLUMNS,
            "company_attribution_enabled": True,
            "company_name": get_default_company_name()
        }
        self._populate_column_list()
        self.chk_company_attribution.setChecked(True)
        self.edit_company.setText(get_default_company_name())
        self._update_company_ui()

    def _load_saved_config(self):
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f).get("excel_export", {})
            except Exception:
                pass
        return {}

    def _save_current_config(self):
        order = []
        checked = []
        for i in range(self.col_list.count()):
            item = self.col_list.item(i)
            order.append(item.text())
            if item.checkState() == Qt.Checked:
                checked.append(item.text())

        full_config = {}
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    full_config = json.load(f)
            except Exception:
                pass

        full_config["excel_export"] = {
            "column_order": order,
            "checked_columns": checked,
            "auto_open": self.chk_auto_open.isChecked(),
            "company_attribution_enabled": self.chk_company_attribution.isChecked(),
            "company_name": self.get_company_name(),
        }

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(full_config, f, indent=4)
        except Exception as e:
            logging.error(f"Failed to persist export configuration: {e}")

    def is_company_attribution_enabled(self) -> bool:
        return self.chk_company_attribution.isChecked()

    def get_company_name(self) -> str:
        return self.edit_company.text().strip()

    def get_column_headers(self) -> dict:
        headers = {}
        if self.chk_company_attribution.isChecked():
            company = self.get_company_name() or "Company"
            headers["My Status"] = f"Status ({company})"
            headers["My Notes"] = f"Comments ({company})"
        else:
            headers["My Status"] = "My Status"
            headers["My Notes"] = "My Notes"
        return headers

    def get_selected_columns(self):
        selected = []
        for i in range(self.col_list.count()):
            item = self.col_list.item(i)
            if item.checkState() == Qt.Checked:
                selected.append(item.text())
        return selected

    def is_visible_only(self):
        return self.radio_visible.isChecked()

    def should_auto_open(self):
        return self.chk_auto_open.isChecked()

    def _on_export_clicked(self):
        if not self.get_selected_columns():
            QMessageBox.warning(self, "No Columns Selected", "Please check at least one column to export.")
            return
        self._save_current_config()
        self.accept()
