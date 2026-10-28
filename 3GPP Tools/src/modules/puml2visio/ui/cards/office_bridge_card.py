"""Card 2: PowerPoint ⇄ Visio Bridge."""

import logging
from pathlib import Path
from typing import Callable, List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    TABLE_STYLE_CLEAN,
    InteractiveDropLabel,
)
from modules.puml2visio.ui.cards.batch_tab_utils import format_file_size

logger = logging.getLogger(__name__)


class OfficeBridgeCard(QWidget):
    """Card 2: Office Format Bridge (PPTX ⇄ Visio) with EMF lossless vector bridging."""

    files_dropped = pyqtSignal(list, str)

    def __init__(self, get_export_dir: Optional[Callable[[], Path]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._get_export_dir = get_export_dir or (lambda: Path("output"))
        self._staged_bridge_files: List[Path] = []
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        # Direction Switcher & Banner
        dir_frame = QFrame()
        dir_frame.setObjectName("cardFrame")
        dir_frame.setStyleSheet(CARD_FRAME_STYLE + "QFrame#cardFrame { background-color: #F8FAFC; padding: 8px; }")
        dir_layout = QVBoxLayout(dir_frame)
        dir_layout.setSpacing(6)

        opts_layout = QHBoxLayout()
        opts_layout.addWidget(QLabel("<b>Conversion Direction:</b>"))

        self.rb_pptx2visio = QRadioButton("📊 PowerPoint (.pptx) ➔ Microsoft Visio (.vsdx)")
        self.rb_pptx2visio.setChecked(True)
        self.rb_pptx2visio.toggled.connect(self._on_bridge_direction_changed)
        opts_layout.addWidget(self.rb_pptx2visio)

        self.rb_visio2pptx = QRadioButton("📐 Microsoft Visio (.vsdx) ➔ PowerPoint Presentation (.pptx)")
        self.rb_visio2pptx.toggled.connect(self._on_bridge_direction_changed)
        opts_layout.addWidget(self.rb_visio2pptx)
        opts_layout.addStretch()

        dir_layout.addLayout(opts_layout)

        info_lbl = QLabel(
            "💡 <b>EMF Lossless Vector Bridging:</b> Converts multi-page presentations or Visio drawings into "
            "native editable Office vector shapes with automated ungrouping and text shrink-wrapping."
        )
        info_lbl.setWordWrap(True)
        info_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        dir_layout.addWidget(info_lbl)

        layout.addWidget(dir_frame)

        # Drop Zone for Bridge
        self.drop_label_bridge = InteractiveDropLabel(
            "📥 Drag & Drop PowerPoint presentation (.pptx) here to convert to multi-page Visio (.vsdx)",
            [".pptx"],
        )
        self.drop_label_bridge.file_dropped.connect(self._on_bridge_files_dropped)
        layout.addWidget(self.drop_label_bridge)

        # Actions Row
        bridge_btn_bar = QHBoxLayout()
        self.add_bridge_file_btn = QPushButton("📂 Browse File(s)...")
        self.add_bridge_file_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.add_bridge_file_btn.clicked.connect(self._browse_bridge_files)

        self.clear_bridge_btn = QPushButton("🗑️ Clear")
        self.clear_bridge_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_bridge_btn.clicked.connect(self._clear_staged_bridge)

        self.bridge_count_lbl = QLabel("0 presentation(s) staged")
        self.bridge_count_lbl.setStyleSheet("color: #64748B; font-size: 11px;")

        bridge_btn_bar.addWidget(self.add_bridge_file_btn)
        bridge_btn_bar.addWidget(self.clear_bridge_btn)
        bridge_btn_bar.addWidget(self.bridge_count_lbl)
        bridge_btn_bar.addStretch()

        self.convert_bridge_btn = QPushButton("🚀 Convert Office Presentation")
        self.convert_bridge_btn.setObjectName("primaryBtn")
        self.convert_bridge_btn.setEnabled(False)
        self.convert_bridge_btn.clicked.connect(self._start_bridge_conversion)
        bridge_btn_bar.addWidget(self.convert_bridge_btn)

        layout.addLayout(bridge_btn_bar)

        # Staged Bridge Table
        self.bridge_table = QTableWidget(0, 4)
        self.bridge_table.setHorizontalHeaderLabels(["#", "File Path", "Size", "Action"])
        self.bridge_table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.bridge_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.bridge_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        b_header = self.bridge_table.horizontalHeader()
        b_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        b_header.setSectionResizeMode(1, QHeaderView.Stretch)
        b_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        b_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        layout.addWidget(self.bridge_table, stretch=1)

    def _on_bridge_direction_changed(self):
        if self.rb_pptx2visio.isChecked():
            self.drop_label_bridge.accepted_extensions = [".pptx"]
            self.drop_label_bridge.setText(
                "📥 Drag & Drop PowerPoint presentation (.pptx) here to convert to multi-page Visio (.vsdx)"
            )
            self.convert_bridge_btn.setText("🚀 Convert PPTX ➔ Visio (.vsdx)")
        else:
            self.drop_label_bridge.accepted_extensions = [".vsdx"]
            self.drop_label_bridge.setText(
                "📥 Drag & Drop Microsoft Visio (.vsdx) here to convert to PowerPoint presentation (.pptx)"
            )
            self.convert_bridge_btn.setText("🚀 Convert Visio ➔ PowerPoint (.pptx)")
        self._clear_staged_bridge()

    def _on_bridge_files_dropped(self, paths: list):
        target_ext = ".pptx" if self.rb_pptx2visio.isChecked() else ".vsdx"
        for p in paths:
            path = Path(p)
            if path.is_file() and path.suffix.lower() == target_ext and path not in self._staged_bridge_files:
                self._staged_bridge_files.append(path)
        self._update_bridge_table()

    def _browse_bridge_files(self):
        export_dir = self._get_export_dir()
        if self.rb_pptx2visio.isChecked():
            filter_str = "PowerPoint Presentations (*.pptx);;All Files (*.*)"
            title = "Select PowerPoint Presentation(s)"
        else:
            filter_str = "Visio Drawings (*.vsdx);;All Files (*.*)"
            title = "Select Visio Drawing(s)"

        files, _ = QFileDialog.getOpenFileNames(self, title, str(export_dir.resolve()), filter_str)
        if files:
            self._on_bridge_files_dropped(files)

    def _clear_staged_bridge(self):
        self._staged_bridge_files.clear()
        self._update_bridge_table()

    def _remove_staged_bridge_index(self, idx: int):
        if 0 <= idx < len(self._staged_bridge_files):
            self._staged_bridge_files.pop(idx)
            self._update_bridge_table()

    def _update_bridge_table(self):
        self.bridge_table.setRowCount(0)
        for i, file_path in enumerate(self._staged_bridge_files, start=1):
            row = self.bridge_table.rowCount()
            self.bridge_table.insertRow(row)

            idx_item = QTableWidgetItem(str(i))
            idx_item.setTextAlignment(Qt.AlignCenter)

            fn_item = QTableWidgetItem(file_path.name)
            fn_item.setToolTip(str(file_path.resolve()))

            sz_str = format_file_size(file_path.stat().st_size) if file_path.exists() else "—"
            sz_item = QTableWidgetItem(sz_str)
            sz_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            btn_remove = QPushButton("✕")
            btn_remove.setFixedSize(24, 22)
            btn_remove.setStyleSheet("font-size: 11px; padding: 0px; color: #EF4444; font-weight: bold;")
            btn_remove.clicked.connect(lambda checked, idx=row: self._remove_staged_bridge_index(idx))

            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(4, 2, 4, 2)
            action_layout.setAlignment(Qt.AlignCenter)
            action_layout.addWidget(btn_remove)

            self.bridge_table.setItem(row, 0, idx_item)
            self.bridge_table.setItem(row, 1, fn_item)
            self.bridge_table.setItem(row, 2, sz_item)
            self.bridge_table.setCellWidget(row, 3, action_widget)

        count = len(self._staged_bridge_files)
        self.bridge_count_lbl.setText(f"{count} presentation(s) staged")
        self.convert_bridge_btn.setEnabled(count > 0)

    def _start_bridge_conversion(self):
        if not self._staged_bridge_files:
            return
        target_fmt = "pptx_to_visio" if self.rb_pptx2visio.isChecked() else "vsdx_to_pptx"
        paths_to_send = list(self._staged_bridge_files)
        self.files_dropped.emit(paths_to_send, target_fmt)
        self._staged_bridge_files.clear()
        self._update_bridge_table()
