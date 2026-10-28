"""Bottom drawer for recent diagram exports."""

import datetime
import logging
import os
from pathlib import Path
import subprocess
from typing import Callable, Optional

from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import CARD_FRAME_STYLE
from modules.puml2visio.ui.cards.batch_tab_utils import format_file_size

logger = logging.getLogger(__name__)


class RecentExportsStrip(QFrame):
    """Recent Exports strip displaying recently generated diagrams with 1-click launch shortcuts."""

    def __init__(self, get_export_dir: Optional[Callable[[], Path]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._get_export_dir = get_export_dir or (lambda: Path("output"))
        self._setup_ui()
        self.refresh()

    def _setup_ui(self):
        self.setObjectName("cardFrame")
        self.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 4px 8px;
            }
        """)
        strip_layout = QVBoxLayout(self)
        strip_layout.setContentsMargins(6, 4, 6, 4)
        strip_layout.setSpacing(4)

        header_row = QHBoxLayout()
        header_lbl = QLabel("<b>🕒 Recent Exports:</b>")
        header_lbl.setStyleSheet("color: #475569; font-size: 11px;")

        self.recent_count_lbl = QLabel("(0 files)")
        self.recent_count_lbl.setStyleSheet("color: #94A3B8; font-size: 11px;")

        self.refresh_recent_btn = QPushButton("🔄 Refresh")
        self.refresh_recent_btn.setStyleSheet("font-size: 10px; padding: 1px 6px;")
        self.refresh_recent_btn.clicked.connect(self.refresh)

        header_row.addWidget(header_lbl)
        header_row.addWidget(self.recent_count_lbl)
        header_row.addStretch()
        header_row.addWidget(self.refresh_recent_btn)
        strip_layout.addLayout(header_row)

        self.recent_container = QVBoxLayout()
        self.recent_container.setSpacing(2)
        strip_layout.addLayout(self.recent_container)

    def refresh(self):
        # Clear existing items in recent_container
        while self.recent_container.count():
            item = self.recent_container.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        export_dir = self._get_export_dir()
        if not export_dir.exists():
            return

        valid_exts = {".vsdx", ".pptx", ".svg", ".txt"}
        recent_files = []
        try:
            for f in export_dir.iterdir():
                if f.is_file() and f.suffix.lower() in valid_exts:
                    recent_files.append((f, f.stat().st_mtime))
        except Exception:
            pass

        recent_files.sort(key=lambda x: x[1], reverse=True)
        top_files = recent_files[:5]

        self.recent_count_lbl.setText(f"({len(top_files)} files)")

        if not top_files:
            empty_lbl = QLabel("No diagrams exported yet in output/ directory.")
            empty_lbl.setStyleSheet("color: #94A3B8; font-style: italic; font-size: 11px;")
            self.recent_container.addWidget(empty_lbl)
            return

        for f, mtime in top_files:
            dt_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(2, 1, 2, 1)
            row_layout.setSpacing(6)

            # Icon based on ext
            ext = f.suffix.lower()
            icon_str = "📄" if ext == ".vsdx" else ("📊" if ext == ".pptx" else ("🌐" if ext == ".svg" else "🔤"))

            icon_lbl = QLabel(icon_str)
            name_lbl = QLabel(f.name)
            name_lbl.setStyleSheet("font-weight: 500; color: #1E293B; font-size: 11px;")
            name_lbl.setToolTip(str(f.resolve()))

            time_lbl = QLabel(dt_str)
            time_lbl.setStyleSheet("color: #94A3B8; font-size: 10px; margin-left: 4px;")

            size_str = format_file_size(f.stat().st_size)
            size_lbl = QLabel(size_str)
            size_lbl.setStyleSheet("color: #64748B; font-size: 10px;")

            btn_open = QPushButton("👁️ Open")
            btn_open.setStyleSheet("font-size: 10px; padding: 1px 6px;")
            btn_open.clicked.connect(lambda checked, path=f: self._open_file(path))

            btn_folder = QPushButton("📂 Folder")
            btn_folder.setStyleSheet("font-size: 10px; padding: 1px 6px;")
            btn_folder.clicked.connect(lambda checked, path=f: self._show_in_folder(path))

            btn_copy = QPushButton("🔗 Copy Path")
            btn_copy.setStyleSheet("font-size: 10px; padding: 1px 6px;")
            btn_copy.clicked.connect(lambda checked, path=f: self._copy_file_path(path))

            row_layout.addWidget(icon_lbl)
            row_layout.addWidget(name_lbl)
            row_layout.addWidget(time_lbl)
            row_layout.addWidget(size_lbl)
            row_layout.addStretch()
            row_layout.addWidget(btn_open)
            row_layout.addWidget(btn_folder)
            row_layout.addWidget(btn_copy)

            self.recent_container.addWidget(row_widget)

    def _open_file(self, file_path: Path):
        try:
            os.startfile(str(file_path.resolve()))
        except Exception as e:
            QMessageBox.warning(self, "Error Opening File", f"Could not open {file_path.name}:\n{e}")

    def _show_in_folder(self, file_path: Path):
        try:
            subprocess.run(["explorer", f"/select,{str(file_path.resolve())}"])
        except Exception:
            try:
                os.startfile(str(file_path.parent.resolve()))
            except Exception:
                pass

    def _copy_file_path(self, file_path: Path):
        QApplication.clipboard().setText(str(file_path.resolve()))
        QToolTip.showText(QCursor.pos(), f"🔗 Copied path:\n{file_path.name}", self, msecShowTime=1800)
