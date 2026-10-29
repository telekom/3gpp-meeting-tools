"""Card 1: Batch PlantUML Converter."""

import logging
from pathlib import Path
from typing import Callable, List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    TABLE_STYLE_CLEAN,
    InteractiveDropLabel,
)
from modules.puml2visio.ui.cards.batch_tab_utils import (
    detect_puml_diagram_type,
    format_file_size,
)

logger = logging.getLogger(__name__)


class PumlBatchCard(QWidget):
    """Card 1: Batch PlantUML Converter with staging queue and diagram type detection."""

    files_dropped = pyqtSignal(list, str)

    def __init__(self, get_export_dir: Optional[Callable[[], Path]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._get_export_dir = get_export_dir or (lambda: Path("output"))
        self._staged_puml_files: List[Path] = []
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Options & Target Format Ribbon
        top_bar = QHBoxLayout()
        top_bar.setSpacing(10)

        top_bar.addWidget(QLabel("<b>Target Format:</b>"))
        self.puml_target_combo = QComboBox()
        self.puml_target_combo.addItems([
            "Microsoft Visio (.vsdx)",
            "Microsoft PowerPoint (.pptx)",
            "Scalable Vector Graphic (.svg)",
            "Unicode Text Art (.txt)",
        ])
        self.puml_target_combo.setStyleSheet("font-size: 12px; padding: 3px 8px;")
        top_bar.addWidget(self.puml_target_combo)

        self.chk_embed_source = QCheckBox("Embed PlantUML in Visio page")
        self.chk_embed_source.setChecked(True)
        self.chk_embed_source.setToolTip("Embeds original PlantUML source code into a hidden Visio page for 100% lossless reverse recovery.")
        top_bar.addWidget(self.chk_embed_source)

        self.chk_auto_open = QCheckBox("Auto-open output upon completion")
        self.chk_auto_open.setChecked(True)
        top_bar.addWidget(self.chk_auto_open)

        top_bar.addStretch()
        layout.addLayout(top_bar)

        # Drop Zone (Compact)
        self.drop_label_puml = InteractiveDropLabel(
            "📥 Drag & Drop PlantUML file(s) (.puml, .plantuml, .txt) here to stage for batch conversion",
            [".puml", ".plantuml", ".txt"],
        )
        self.drop_label_puml.setMaximumHeight(80)
        self.drop_label_puml.file_dropped.connect(self._on_puml_files_dropped)
        self.drop_label_puml.file_selected.connect(self._on_puml_files_dropped)
        layout.addWidget(self.drop_label_puml)

        # Buttons Toolbar
        btn_bar = QHBoxLayout()
        btn_bar.setSpacing(6)

        self.add_files_btn = QPushButton("📂 Add Files...")
        self.add_files_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.add_files_btn.clicked.connect(self._browse_puml_files)

        self.add_folder_btn = QPushButton("📁 Add Folder...")
        self.add_folder_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.add_folder_btn.clicked.connect(self._browse_puml_folder)

        self.clear_staged_btn = QPushButton("🗑️ Clear Staging")
        self.clear_staged_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_staged_btn.clicked.connect(self._clear_staged_puml)

        self.staged_count_lbl = QLabel("0 diagrams staged")
        self.staged_count_lbl.setStyleSheet("color: #64748B; font-size: 11px; margin-left: 6px;")

        btn_bar.addWidget(self.add_files_btn)
        btn_bar.addWidget(self.add_folder_btn)
        btn_bar.addWidget(self.clear_staged_btn)
        btn_bar.addWidget(self.staged_count_lbl)
        btn_bar.addStretch()

        self.convert_puml_btn = QPushButton("🚀 Convert All Staged Diagrams")
        self.convert_puml_btn.setObjectName("primaryBtn")
        self.convert_puml_btn.setEnabled(False)
        self.convert_puml_btn.clicked.connect(self._start_puml_batch_conversion)
        btn_bar.addWidget(self.convert_puml_btn)

        layout.addLayout(btn_bar)

        # Staging Table
        self.puml_table = QTableWidget(0, 5)
        self.puml_table.setHorizontalHeaderLabels(["#", "Filename", "Diagram Type", "Size", "Action"])
        self.puml_table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.puml_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.puml_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        header = self.puml_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        layout.addWidget(self.puml_table, stretch=1)

    def _on_puml_files_dropped(self, paths: list):
        for p in paths:
            path = Path(p)
            if path.is_file() and path not in self._staged_puml_files:
                self._staged_puml_files.append(path)
        self._update_puml_table()

    def _browse_puml_files(self):
        export_dir = self._get_export_dir()
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PlantUML Diagram Files",
            str(export_dir.resolve()),
            "PlantUML Diagrams (*.puml *.plantuml *.txt);;All Files (*.*)",
        )
        if files:
            self._on_puml_files_dropped(files)

    def _browse_puml_folder(self):
        export_dir = self._get_export_dir()
        folder = QFileDialog.getExistingDirectory(
            self, "Select Folder with PlantUML Diagrams", str(export_dir.resolve())
        )
        if folder:
            p_folder = Path(folder)
            found = []
            for ext in (".puml", ".plantuml", ".txt"):
                found.extend(list(p_folder.glob(f"*{ext}")))
            if found:
                self._on_puml_files_dropped(found)
            else:
                QMessageBox.information(
                    self, "No Diagrams Found", f"No .puml or .txt diagram files found in:\n{folder}"
                )

    def _clear_staged_puml(self):
        self._staged_puml_files.clear()
        self._update_puml_table()

    def _remove_staged_puml_index(self, idx: int):
        if 0 <= idx < len(self._staged_puml_files):
            self._staged_puml_files.pop(idx)
            self._update_puml_table()

    def _update_puml_table(self):
        self.puml_table.setRowCount(0)
        for i, file_path in enumerate(self._staged_puml_files, start=1):
            row = self.puml_table.rowCount()
            self.puml_table.insertRow(row)

            # 0. Index
            idx_item = QTableWidgetItem(str(i))
            idx_item.setTextAlignment(Qt.AlignCenter)

            # 1. Filename
            fn_item = QTableWidgetItem(file_path.name)
            fn_item.setToolTip(str(file_path.resolve()))

            # 2. Diagram Type
            dtype = detect_puml_diagram_type(file_path)
            dt_item = QTableWidgetItem(dtype)
            dt_item.setTextAlignment(Qt.AlignCenter)

            # 3. Size
            sz_str = format_file_size(file_path.stat().st_size) if file_path.exists() else "—"
            sz_item = QTableWidgetItem(sz_str)
            sz_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            # 4. Remove Action
            btn_remove = QPushButton("✕")
            btn_remove.setFixedSize(24, 22)
            btn_remove.setStyleSheet("font-size: 11px; padding: 0px; color: #EF4444; font-weight: bold;")
            btn_remove.setToolTip("Remove from staging")
            btn_remove.clicked.connect(lambda checked, idx=row: self._remove_staged_puml_index(idx))

            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(4, 2, 4, 2)
            action_layout.setAlignment(Qt.AlignCenter)
            action_layout.addWidget(btn_remove)

            self.puml_table.setItem(row, 0, idx_item)
            self.puml_table.setItem(row, 1, fn_item)
            self.puml_table.setItem(row, 2, dt_item)
            self.puml_table.setItem(row, 3, sz_item)
            self.puml_table.setCellWidget(row, 4, action_widget)

        count = len(self._staged_puml_files)
        self.staged_count_lbl.setText(f"{count} diagram(s) staged")
        self.convert_puml_btn.setEnabled(count > 0)

    def _start_puml_batch_conversion(self):
        if not self._staged_puml_files:
            return

        fmt_choice = self.puml_target_combo.currentText()
        if ".vsdx" in fmt_choice:
            target_fmt = "vsdx"
        elif ".pptx" in fmt_choice:
            target_fmt = "pptx"
        elif ".svg" in fmt_choice:
            target_fmt = "svg"
        else:
            target_fmt = "ascii"

        paths_to_send = list(self._staged_puml_files)
        self.files_dropped.emit(paths_to_send, target_fmt)
        self._staged_puml_files.clear()
        self._update_puml_table()

    def set_state(self, state: str, text: Optional[str] = None):
        """Update drop label state and text."""
        if hasattr(self, "drop_label_puml"):
            self.drop_label_puml.set_state(state, text)

