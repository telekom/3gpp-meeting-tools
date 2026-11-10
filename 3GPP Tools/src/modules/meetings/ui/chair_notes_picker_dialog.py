# --- File: src/modules/meetings/ui/chair_notes_picker_dialog.py ---
"""
Multi-chair / Multi-track Chairman's Notes Picker Dialog (`ChairNotesPickerDialog`).

Allows delegates to inspect candidate notes from /Inbox/Chair_Notes before downloading:
- Pre-selects the newest/most recent file for each chair / track (e.g. Andy, Hector, Ruicheng, Jian)
- Marks which files are already downloaded locally vs new remote files
- Supports quick filters, search filtering, and batch selection
"""
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEARCH_INPUT_STYLE,
)
from modules.meetings.core.chair_notes_classifier import analyze_candidate_files

BUTTON_STYLE_PRIMARY = """
QPushButton {
    background-color: #2563EB;
    color: #FFFFFF;
    border: 1px solid #1D4ED8;
    border-radius: 4px;
    padding: 6px 14px;
    font-size: 12px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #1D4ED8;
    border-color: #1E40AF;
}
QPushButton:pressed {
    background-color: #1E40AF;
}
QPushButton:disabled {
    background-color: #E2E8F0;
    color: #94A3B8;
    border-color: #CBD5E1;
}
"""


class ChairNotesPickerDialog(QDialog):
    """
    Dialog displaying candidate Chairman's Notes files found on the 3GPP server.
    Pre-picks the most recent file per chair/track and highlights local download status.
    """

    def __init__(
        self,
        parent=None,
        file_tuples: Optional[List[Tuple[str, str]]] = None,
        local_dir: Optional[Path] = None,
        mtg_info: Optional[dict] = None,
        target_dir_url: str = "",
    ):
        super().__init__(parent)
        self.file_tuples = file_tuples or []
        self.local_dir = Path(local_dir) if local_dir else Path(".")
        self.mtg_info = mtg_info or {}
        self.target_dir_url = target_dir_url

        self.analyzed_files: List[Dict[str, Any]] = analyze_candidate_files(
            self.file_tuples, self.local_dir
        )
        self._row_to_record: Dict[int, Dict[str, Any]] = {}

        wg_name = str(self.mtg_info.get("wg_name", "3GPP")).strip()
        mtg_num = str(self.mtg_info.get("meeting_number", "")).strip()
        meeting_label = f"{wg_name} #{mtg_num}" if mtg_num else wg_name

        self.setWindowTitle(f"📥 Download Chairman's Notes — {meeting_label}")
        self.setMinimumSize(780, 520)
        self.resize(840, 580)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self._init_ui(meeting_label)
        self._populate_table()
        self._update_selection_summary()

    def _init_ui(self, meeting_label: str):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        # 1. Header Frame
        header_frame = QFrame()
        header_frame.setStyleSheet(
            "QFrame { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 10px; }"
        )
        h_layout = QVBoxLayout(header_frame)
        h_layout.setContentsMargins(8, 6, 8, 6)
        h_layout.setSpacing(4)

        title_lbl = QLabel(f"<b>📥 Select Chairman's Notes to Download — {meeting_label}</b>")
        title_lbl.setStyleSheet("font-size: 14px; color: #0F172A;")
        h_layout.addWidget(title_lbl)

        count_text = f"Found {len(self.analyzed_files)} candidate Chairman's Notes file(s) in server <code>/Inbox/Chair_Notes</code>"
        sub_lbl = QLabel(count_text)
        sub_lbl.setStyleSheet("font-size: 12px; color: #475569;")
        sub_lbl.setTextFormat(Qt.RichText)
        h_layout.addWidget(sub_lbl)

        tip_lbl = QLabel(
            "💡 <b>Pre-selection:</b> The newest file for each chair / track is pre-selected. "
            "Older versions remain unchecked by default. You can customize your selection below."
        )
        tip_lbl.setStyleSheet("font-size: 11px; color: #0369A1; background: #E0F2FE; border-radius: 4px; padding: 4px 8px;")
        tip_lbl.setTextFormat(Qt.RichText)
        h_layout.addWidget(tip_lbl)

        layout.addWidget(header_frame)

        # 2. Quick Selection and Search Toolbar
        tb_layout = QHBoxLayout()
        tb_layout.setSpacing(8)

        self.btn_select_recommended = QPushButton("⭐ Select Pre-picked (Latest per Chair)")
        self.btn_select_recommended.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_select_recommended.clicked.connect(self._select_recommended)
        tb_layout.addWidget(self.btn_select_recommended)

        self.btn_select_new = QPushButton("☁️ Select Un-downloaded Only")
        self.btn_select_new.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_select_new.clicked.connect(self._select_undownloaded_only)
        tb_layout.addWidget(self.btn_select_new)

        self.btn_select_all = QPushButton("☑️ Select All")
        self.btn_select_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_select_all.clicked.connect(self._select_all)
        tb_layout.addWidget(self.btn_select_all)

        self.btn_clear_all = QPushButton("⬜ Clear All")
        self.btn_clear_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_clear_all.clicked.connect(self._clear_all)
        tb_layout.addWidget(self.btn_clear_all)

        tb_layout.addStretch()

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 Filter files / chairs...")
        self.search_edit.setStyleSheet(SEARCH_INPUT_STYLE)
        self.search_edit.setFixedWidth(200)
        self.search_edit.textChanged.connect(self._filter_table)
        tb_layout.addWidget(self.search_edit)

        layout.addLayout(tb_layout)

        # 3. Candidate Files Table
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "",
            "File Name",
            "Chair / Track",
            "Recommendation",
            "Local Status",
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.setColumnWidth(0, 36)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.cellClicked.connect(self._on_cell_clicked)
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)
        layout.addWidget(self.table)

        # 4. Bottom Controls and Confirmation
        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(10)

        v_summary = QVBoxLayout()
        self.summary_label = QLabel("Selected: 0 files")
        self.summary_label.setStyleSheet("font-weight: bold; color: #1E293B;")
        v_summary.addWidget(self.summary_label)

        self.overwrite_cb = QCheckBox("Overwrite local files if already downloaded")
        self.overwrite_cb.setChecked(True)
        self.overwrite_cb.setToolTip("When checked, downloaded files will replace existing files in Agenda/ChairNotes.")
        v_summary.addWidget(self.overwrite_cb)
        bottom_layout.addLayout(v_summary)

        bottom_layout.addStretch()

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_cancel.setFixedWidth(100)
        self.btn_cancel.clicked.connect(self.reject)
        bottom_layout.addWidget(self.btn_cancel)

        self.btn_download = QPushButton("📥 Download Selected")
        self.btn_download.setStyleSheet(BUTTON_STYLE_PRIMARY)
        self.btn_download.setFixedWidth(180)
        self.btn_download.clicked.connect(self.accept)
        bottom_layout.addWidget(self.btn_download)

        layout.addLayout(bottom_layout)

    def _populate_table(self):
        self.table.setRowCount(len(self.analyzed_files))
        self._row_to_record.clear()

        for row, rec in enumerate(self.analyzed_files):
            self._row_to_record[row] = rec

            # Col 0: Checkbox
            chk_item = QTableWidgetItem()
            chk_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            # Pre-pick the latest file for each chair!
            chk_item.setCheckState(Qt.Checked if rec["is_latest_for_chair"] else Qt.Unchecked)
            self.table.setItem(row, 0, chk_item)

            # Col 1: File Name
            name_item = QTableWidgetItem(rec["filename"])
            name_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            name_item.setToolTip(f"URL: {rec['url']}")
            if rec["is_latest_for_chair"]:
                font = name_item.font()
                font.setBold(True)
                name_item.setFont(font)
            self.table.setItem(row, 1, name_item)

            # Col 2: Chair / Track
            track_item = QTableWidgetItem(rec["track_label"])
            track_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            track_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 2, track_item)

            # Col 3: Recommendation
            if rec["is_latest_for_chair"]:
                rec_text = "⭐ Latest Version"
                rec_item = QTableWidgetItem(rec_text)
                rec_item.setForeground(QColor("#047857"))  # Green
                font = rec_item.font()
                font.setBold(True)
                rec_item.setFont(font)
            else:
                rec_text = "Older Version"
                rec_item = QTableWidgetItem(rec_text)
                rec_item.setForeground(QColor("#64748B"))  # Slate grey
            rec_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            rec_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 3, rec_item)

            # Col 4: Local Status
            if rec["is_already_downloaded"]:
                size_kb = rec["local_size"] / 1024.0
                status_text = f"🟢 Downloaded ({size_kb:.0f} KB)" if size_kb > 0 else "🟢 Downloaded"
                status_item = QTableWidgetItem(status_text)
                status_item.setForeground(QColor("#059669"))
            else:
                status_item = QTableWidgetItem("☁️ Not Downloaded")
                status_item.setForeground(QColor("#2563EB"))
            status_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            status_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 4, status_item)

        self.table.itemChanged.connect(self._on_item_changed)

    def _on_cell_clicked(self, row: int, col: int):
        if col != 0:
            item = self.table.item(row, 0)
            if item:
                new_state = Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked
                item.setCheckState(new_state)

    def _on_cell_double_clicked(self, row: int, col: int):
        self._on_cell_clicked(row, col)

    def _on_item_changed(self, item: QTableWidgetItem):
        if item.column() == 0:
            self._update_selection_summary()

    def _update_selection_summary(self):
        selected_count = 0
        new_count = 0
        cached_count = 0

        for row in range(self.table.rowCount()):
            chk_item = self.table.item(row, 0)
            if chk_item and chk_item.checkState() == Qt.Checked:
                selected_count += 1
                rec = self._row_to_record.get(row, {})
                if rec.get("is_already_downloaded", False):
                    cached_count += 1
                else:
                    new_count += 1

        total = len(self.analyzed_files)
        detail = []
        if new_count > 0:
            detail.append(f"{new_count} new remote")
        if cached_count > 0:
            detail.append(f"{cached_count} already downloaded")
        detail_str = f" ({', '.join(detail)})" if detail else ""

        self.summary_label.setText(f"Selected: {selected_count} of {total} files{detail_str}")
        self.btn_download.setText(f"📥 Download Selected ({selected_count})")
        self.btn_download.setEnabled(selected_count > 0)

    def _select_recommended(self):
        self.table.blockSignals(True)
        for row, rec in self._row_to_record.items():
            chk_item = self.table.item(row, 0)
            if chk_item:
                chk_item.setCheckState(Qt.Checked if rec["is_latest_for_chair"] else Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selection_summary()

    def _select_undownloaded_only(self):
        self.table.blockSignals(True)
        for row, rec in self._row_to_record.items():
            chk_item = self.table.item(row, 0)
            if chk_item:
                chk_item.setCheckState(
                    Qt.Checked if not rec["is_already_downloaded"] else Qt.Unchecked
                )
        self.table.blockSignals(False)
        self._update_selection_summary()

    def _select_all(self):
        self.table.blockSignals(True)
        for row in range(self.table.rowCount()):
            chk_item = self.table.item(row, 0)
            if chk_item:
                chk_item.setCheckState(Qt.Checked)
        self.table.blockSignals(False)
        self._update_selection_summary()

    def _clear_all(self):
        self.table.blockSignals(True)
        for row in range(self.table.rowCount()):
            chk_item = self.table.item(row, 0)
            if chk_item:
                chk_item.setCheckState(Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selection_summary()

    def _filter_table(self, query: str):
        q = (query or "").strip().lower()
        for row in range(self.table.rowCount()):
            rec = self._row_to_record.get(row, {})
            fname = rec.get("filename", "").lower()
            track = rec.get("track_label", "").lower()
            matches = not q or (q in fname or q in track)
            self.table.setRowHidden(row, not matches)

    def get_selected_files(self) -> List[Tuple[str, str]]:
        """Returns the list of selected `(filename, url)` tuples."""
        selected: List[Tuple[str, str]] = []
        for row in range(self.table.rowCount()):
            chk_item = self.table.item(row, 0)
            if chk_item and chk_item.checkState() == Qt.Checked:
                rec = self._row_to_record.get(row)
                if rec:
                    selected.append((rec["filename"], rec["url"]))
        return selected

    def should_overwrite(self) -> bool:
        """Returns whether the user wants to overwrite files that are already cached locally."""
        return self.overwrite_cb.isChecked()
