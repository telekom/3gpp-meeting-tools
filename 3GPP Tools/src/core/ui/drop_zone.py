# --- File: src/core/ui/drop_zone.py ---
"""
ModernDropZone: High-fidelity Fluent-style Drag & Drop component for 3GPP Tools.

Features:
- Pure anti-aliased QPainter dashed-border rendering (eliminates Qt QSS dashed-corner bugs).
- Multi-state visual machine: Empty/Idle, Drag-Over (Active), Loaded (Ready Card), Busy, Error.
- Dual layout modes: Canvas Mode (spacious stacked layout) and Compact Ribbon Mode (~58px height).
- Rich file type recognition with dedicated color badges (.docx, .doc, .md, .puml, .pptx, .vsdx, .pdf).
- Automatic file size formatting, OpenXML sensitivity label detection, and path elision.
- Interactive controls: Click-to-browse, inline Change/Replace, and Clear actions.
- 100% backward-compatible API with legacy InteractiveDropLabel and ClickableDropLabel.
"""

import logging
import os
from pathlib import Path
from typing import List, Optional, Union

from PyQt5.QtCore import QPoint, QRectF, QSize, Qt, pyqtSignal
from PyQt5.QtGui import (
    QBrush,
    QColor,
    QCursor,
    QFont,
    QFontMetrics,
    QIcon,
    QPainter,
    QPainterPath,
    QPen,
)
from PyQt5.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

logger = logging.getLogger(__name__)

# Design Tokens (matching 3GPP Tools Slate/Blue design language)
COLOR_TEXT_PRIMARY = "#0F172A"       # Slate 900
COLOR_TEXT_SECONDARY = "#475569"     # Slate 600
COLOR_TEXT_MUTED = "#64748B"         # Slate 500

COLOR_IDLE_BG = "#F8FAFC"            # Slate 50
COLOR_IDLE_BORDER = "#CBD5E1"        # Slate 300
COLOR_IDLE_HOVER_BORDER = "#94A3B8"  # Slate 400

COLOR_DRAG_BG = "#EFF6FF"            # Blue 50
COLOR_DRAG_BORDER = "#2563EB"        # Blue 600
COLOR_DRAG_ACCENT = "#1D4ED8"        # Blue 700

COLOR_LOADED_BG = "#FFFFFF"          # Pure White
COLOR_LOADED_BORDER = "#E2E8F0"      # Slate 200

COLOR_BUSY_BG = "#FFFBEB"            # Amber 50
COLOR_BUSY_BORDER = "#F59E0B"        # Amber 500

COLOR_ERROR_BG = "#FEF2F2"           # Red 50
COLOR_ERROR_BORDER = "#EF4444"       # Red 500

# File Type Archetypes (Badge Color, Background, Icon)
FILE_TYPE_METADATA = {
    ".docx": {"icon": "📘", "name": "Word Document", "bg": "#EFF6FF", "fg": "#1D4ED8", "border": "#BFDBFE"},
    ".doc":  {"icon": "📘", "name": "Legacy Word Doc", "bg": "#EFF6FF", "fg": "#1D4ED8", "border": "#BFDBFE"},
    ".docm": {"icon": "📘", "name": "Word Macro-Enabled", "bg": "#EFF6FF", "fg": "#1D4ED8", "border": "#BFDBFE"},
    ".md":   {"icon": "📝", "name": "Markdown Document", "bg": "#F5F3FF", "fg": "#6D28D9", "border": "#DDD6FE"},
    ".markdown": {"icon": "📝", "name": "Markdown Document", "bg": "#F5F3FF", "fg": "#6D28D9", "border": "#DDD6FE"},
    ".puml": {"icon": "📐", "name": "PlantUML Diagram", "bg": "#ECFDF5", "fg": "#047857", "border": "#A7F3D0"},
    ".plantuml": {"icon": "📐", "name": "PlantUML Diagram", "bg": "#ECFDF5", "fg": "#047857", "border": "#A7F3D0"},
    ".pptx": {"icon": "📙", "name": "PowerPoint Deck", "bg": "#FFF7ED", "fg": "#C2410C", "border": "#FED7AA"},
    ".ppt":  {"icon": "📙", "name": "Legacy PowerPoint", "bg": "#FFF7ED", "fg": "#C2410C", "border": "#FED7AA"},
    ".vsdx": {"icon": "🔷", "name": "Visio Drawing", "bg": "#F0F9FF", "fg": "#0369A1", "border": "#BAE6FD"},
    ".vsd":  {"icon": "🔷", "name": "Legacy Visio Drawing", "bg": "#F0F9FF", "fg": "#0369A1", "border": "#BAE6FD"},
    ".pdf":  {"icon": "📕", "name": "PDF Document", "bg": "#FFF1F2", "fg": "#BE123C", "border": "#FECDD3"},
    ".txt":  {"icon": "📄", "name": "Text Document", "bg": "#F8FAFC", "fg": "#334155", "border": "#CBD5E1"},
    ".zip":  {"icon": "🗜️", "name": "Zip Archive", "bg": "#F3F4F6", "fg": "#374151", "border": "#D1D5DB"},
}

DEFAULT_FILE_TYPE = {"icon": "📄", "name": "Document", "bg": "#F8FAFC", "fg": "#475569", "border": "#CBD5E1"}


def format_file_size(size_bytes: int) -> str:
    """Format bytes into human-readable B, KB, MB, or GB."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    kb = size_bytes / 1024.0
    if kb < 1024:
        return f"{kb:.1f} KB"
    mb = kb / 1024.0
    if mb < 1024:
        return f"{mb:.2f} MB"
    gb = mb / 1024.0
    return f"{gb:.2f} GB"


class ModernDropZone(QFrame):
    """
    A modern, Fluent-styled drag & drop target component for PyQt5.
    
    Renders anti-aliased dashed borders via QPainter to prevent corner rasterization bugs,
    and transitions into a solid Active Document Card when a file is selected.
    """

    file_dropped = pyqtSignal(list)
    file_selected = pyqtSignal(list)
    file_cleared = pyqtSignal()
    clicked = pyqtSignal()

    def __init__(
        self,
        accepted_extensions: Optional[List[str]] = None,
        title: str = "Drag & drop file(s) here",
        subtitle: Optional[str] = "or click to browse from computer",
        compact: bool = False,
        browse_on_click: bool = True,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.accepted_extensions = [ext.lower() for ext in (accepted_extensions or [])]
        self.primary_title = title
        self.subtitle = subtitle
        self.compact = compact
        self.browse_on_click = browse_on_click

        # State tracking
        self._current_state = "idle"  # "idle", "drag_over", "loaded", "busy", "error"
        self._selected_paths: List[Path] = []
        self._custom_status_message: Optional[str] = None
        self._is_hovered = False

        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor if browse_on_click else Qt.ArrowCursor)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self._init_ui()
        self._apply_state_style()

    # -------------------------------------------------------------------------
    # UI Setup & State Pages
    # -------------------------------------------------------------------------
    def _init_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self.stack = QStackedWidget(self)
        self.stack.setStyleSheet("background: transparent; border: none;")

        # Page 0: Empty / Drop Prompt Page
        self.empty_page = self._create_empty_page()
        self.stack.addWidget(self.empty_page)

        # Page 1: File Loaded / Ready Card Page
        self.loaded_page = self._create_loaded_page()
        self.stack.addWidget(self.loaded_page)

        root_layout.addWidget(self.stack)

        if self.compact:
            self.setFixedHeight(60)
        else:
            self.setFixedHeight(105)

    def _create_empty_page(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet("background: transparent;")

        if self.compact:
            # Compact Ribbon Layout (Horizontal)
            layout = QHBoxLayout(page)
            layout.setContentsMargins(14, 8, 14, 8)
            layout.setSpacing(12)

            self.empty_icon_lbl = QLabel("📥")
            self.empty_icon_lbl.setStyleSheet(
                "font-size: 20px; background-color: #EEF2F6; border-radius: 18px; padding: 4px;"
            )
            self.empty_icon_lbl.setFixedSize(36, 36)
            self.empty_icon_lbl.setAlignment(Qt.AlignCenter)
            layout.addWidget(self.empty_icon_lbl)

            text_col = QVBoxLayout()
            text_col.setSpacing(2)
            text_col.setAlignment(Qt.AlignVCenter)

            self.lbl_empty_title = QLabel(self.primary_title)
            self.lbl_empty_title.setStyleSheet(f"font-size: 13px; font-weight: 600; color: {COLOR_TEXT_PRIMARY};")
            text_col.addWidget(self.lbl_empty_title)

            exts_hint = f"Supports {', '.join(self.accepted_extensions)}" if self.accepted_extensions else ""
            sub_text = f"{exts_hint} • {self.subtitle}" if (exts_hint and self.subtitle) else (exts_hint or self.subtitle or "")
            self.lbl_empty_sub = QLabel(sub_text)
            self.lbl_empty_sub.setStyleSheet(f"font-size: 11px; color: {COLOR_TEXT_MUTED};")
            text_col.addWidget(self.lbl_empty_sub)

            layout.addLayout(text_col, 1)

            self.btn_empty_browse = QPushButton("📂 Browse...")
            self.btn_empty_browse.setStyleSheet("""
                QPushButton {
                    background-color: #F8FAFC;
                    color: #1E293B;
                    border: 1px solid #CBD5E1;
                    border-radius: 5px;
                    padding: 5px 12px;
                    font-size: 12px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #EFF6FF;
                    border-color: #3B82F6;
                    color: #1D4ED8;
                }
            """)
            self.btn_empty_browse.setCursor(Qt.PointingHandCursor)
            self.btn_empty_browse.clicked.connect(self._on_browse_clicked)
            layout.addWidget(self.btn_empty_browse)

        else:
            # Canvas Stacked Layout (Spacious)
            layout = QVBoxLayout(page)
            layout.setContentsMargins(16, 12, 16, 12)
            layout.setSpacing(4)
            layout.setAlignment(Qt.AlignCenter)

            self.empty_icon_lbl = QLabel("📥")
            self.empty_icon_lbl.setStyleSheet(
                "font-size: 22px; background-color: #EEF2F6; border-radius: 18px; padding: 2px;"
            )
            self.empty_icon_lbl.setFixedSize(36, 36)
            self.empty_icon_lbl.setAlignment(Qt.AlignCenter)
            layout.addWidget(self.empty_icon_lbl, 0, Qt.AlignCenter)

            self.lbl_empty_title = QLabel(self.primary_title)
            self.lbl_empty_title.setStyleSheet(f"font-size: 13px; font-weight: 600; color: {COLOR_TEXT_PRIMARY};")
            self.lbl_empty_title.setAlignment(Qt.AlignCenter)
            layout.addWidget(self.lbl_empty_title)

            self.lbl_empty_sub = QLabel(self.subtitle or "")
            self.lbl_empty_sub.setStyleSheet(f"font-size: 11px; color: {COLOR_TEXT_MUTED};")
            self.lbl_empty_sub.setAlignment(Qt.AlignCenter)
            layout.addWidget(self.lbl_empty_sub)

            # Optional format badges pill row
            if self.accepted_extensions:
                badges_row = QHBoxLayout()
                badges_row.setSpacing(4)
                badges_row.setAlignment(Qt.AlignCenter)
                for ext in self.accepted_extensions[:5]:
                    badge = QLabel(ext)
                    badge.setStyleSheet("""
                        QLabel {
                            background-color: #F1F5F9;
                            color: #475569;
                            font-size: 10px;
                            font-weight: 600;
                            border: 1px solid #E2E8F0;
                            border-radius: 4px;
                            padding: 1px 5px;
                        }
                    """)
                    badges_row.addWidget(badge)
                layout.addLayout(badges_row)

        return page

    def _create_loaded_page(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet("background: transparent;")

        layout = QHBoxLayout(page)
        layout.setContentsMargins(14, 8, 14, 8)
        layout.setSpacing(12)

        # File Archetype Icon Badge
        self.loaded_icon_lbl = QLabel("📄")
        self.loaded_icon_lbl.setFixedSize(40, 40)
        self.loaded_icon_lbl.setAlignment(Qt.AlignCenter)
        self.loaded_icon_lbl.setStyleSheet("font-size: 20px; border-radius: 8px;")
        layout.addWidget(self.loaded_icon_lbl)

        # File Details (Title + Metadata row)
        details_col = QVBoxLayout()
        details_col.setSpacing(3)
        details_col.setAlignment(Qt.AlignVCenter)

        self.lbl_file_name = QLabel()
        self.lbl_file_name.setStyleSheet(f"font-size: 13px; font-weight: 600; color: {COLOR_TEXT_PRIMARY};")
        details_col.addWidget(self.lbl_file_name)

        meta_row = QHBoxLayout()
        meta_row.setSpacing(6)
        meta_row.setAlignment(Qt.AlignLeft)

        self.lbl_file_size = QLabel()
        self.lbl_file_size.setStyleSheet("""
            QLabel {
                background-color: #F1F5F9;
                color: #334155;
                font-size: 10px;
                font-weight: 600;
                border: 1px solid #E2E8F0;
                border-radius: 4px;
                padding: 1px 6px;
            }
        """)
        meta_row.addWidget(self.lbl_file_size)

        self.lbl_file_desc = QLabel()
        self.lbl_file_desc.setStyleSheet(f"font-size: 11px; color: {COLOR_TEXT_MUTED};")
        meta_row.addWidget(self.lbl_file_desc)

        self.lbl_sensitivity = QLabel()
        self.lbl_sensitivity.setStyleSheet("""
            QLabel {
                background-color: #ECFDF5;
                color: #065F46;
                font-size: 10px;
                font-weight: 600;
                border: 1px solid #A7F3D0;
                border-radius: 4px;
                padding: 1px 6px;
            }
        """)
        self.lbl_sensitivity.hide()
        meta_row.addWidget(self.lbl_sensitivity)

        details_col.addLayout(meta_row)
        layout.addLayout(details_col, 1)

        # Action Buttons: Change & Clear
        actions_row = QHBoxLayout()
        actions_row.setSpacing(6)
        actions_row.setAlignment(Qt.AlignVCenter)

        self.btn_change = QPushButton("🔄 Replace")
        self.btn_change.setStyleSheet("""
            QPushButton {
                background-color: #F8FAFC;
                color: #1E293B;
                border: 1px solid #CBD5E1;
                border-radius: 5px;
                padding: 5px 12px;
                font-size: 12px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #EFF6FF;
                border-color: #3B82F6;
                color: #1D4ED8;
            }
        """)
        self.btn_change.setCursor(Qt.PointingHandCursor)
        self.btn_change.clicked.connect(self._on_browse_clicked)
        actions_row.addWidget(self.btn_change)

        self.btn_clear = QPushButton("✕")
        self.btn_clear.setToolTip("Clear selected document")
        self.btn_clear.setFixedSize(28, 28)
        self.btn_clear.setStyleSheet("""
            QPushButton {
                background-color: #F8FAFC;
                color: #64748B;
                border: 1px solid #CBD5E1;
                border-radius: 14px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #FEE2E2;
                border-color: #EF4444;
                color: #B91C1C;
            }
        """)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.clicked.connect(self.clear)
        actions_row.addWidget(self.btn_clear)

        layout.addLayout(actions_row)
        return page

    # -------------------------------------------------------------------------
    # Anti-Aliased Border & Background Painter
    # -------------------------------------------------------------------------
    def paintEvent(self, event):
        """
        Custom paint event ensuring high-DPI smooth anti-aliased rounded dashed borders,
        completely resolving the Qt QSS dashed-corner broken bracket artifact.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = QRectF(self.rect()).adjusted(1.0, 1.0, -1.0, -1.0)
        radius = 8.0

        if self._current_state == "drag_over":
            bg_color = QColor(COLOR_DRAG_BG)
            border_color = QColor(COLOR_DRAG_BORDER)
            pen = QPen(border_color, 2.0, Qt.DashLine)
            pen.setDashPattern([6, 4])
            pen.setCapStyle(Qt.RoundCap)
            pen.setJoinStyle(Qt.RoundJoin)
        elif self._current_state == "loaded":
            bg_color = QColor(COLOR_LOADED_BG)
            border_color = QColor(COLOR_LOADED_BORDER)
            pen = QPen(border_color, 1.0, Qt.SolidLine)
        elif self._current_state == "busy":
            bg_color = QColor(COLOR_BUSY_BG)
            border_color = QColor(COLOR_BUSY_BORDER)
            pen = QPen(border_color, 1.5, Qt.SolidLine)
        elif self._current_state == "error":
            bg_color = QColor(COLOR_ERROR_BG)
            border_color = QColor(COLOR_ERROR_BORDER)
            pen = QPen(border_color, 1.5, Qt.SolidLine)
        else:  # idle / empty
            bg_color = QColor(COLOR_IDLE_BG)
            border_color = QColor(COLOR_IDLE_HOVER_BORDER if self._is_hovered else COLOR_IDLE_BORDER)
            pen = QPen(border_color, 1.5, Qt.DashLine)
            pen.setDashPattern([6, 4])
            pen.setCapStyle(Qt.RoundCap)
            pen.setJoinStyle(Qt.RoundJoin)

        painter.setPen(pen)
        painter.setBrush(QBrush(bg_color))
        painter.drawRoundedRect(rect, radius, radius)

    def enterEvent(self, event):
        self._is_hovered = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._is_hovered = False
        self.update()
        super().leaveEvent(event)

    # -------------------------------------------------------------------------
    # Drag and Drop Events
    # -------------------------------------------------------------------------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            valid = any(self._is_accepted(url.toLocalFile()) for url in urls)
            if valid:
                self._current_state = "drag_over"
                self.update()
                event.acceptProposedAction()
                return
        event.ignore()

    def dragLeaveEvent(self, event):
        self._current_state = "loaded" if self._selected_paths else "idle"
        self.update()
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        self._current_state = "idle"
        valid_paths: List[Path] = []
        for url in event.mimeData().urls():
            local_path = Path(url.toLocalFile())
            if local_path.is_file() and self._is_accepted(str(local_path)):
                valid_paths.append(local_path)
            elif local_path.is_dir():
                valid_paths.append(local_path)

        if valid_paths:
            self.set_loaded_file(valid_paths)
            self.file_dropped.emit([str(p) for p in valid_paths])
            self.file_selected.emit([str(p) for p in valid_paths])
            event.acceptProposedAction()
        else:
            self.update()
            event.ignore()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
            if self._current_state != "loaded" and self.browse_on_click:
                self._on_browse_clicked()
        super().mousePressEvent(event)

    # -------------------------------------------------------------------------
    # State & File Loading Management
    # -------------------------------------------------------------------------
    def _is_accepted(self, file_path_str: str) -> bool:
        if not self.accepted_extensions:
            return True
        path_lower = file_path_str.lower()
        return any(path_lower.endswith(ext) for ext in self.accepted_extensions)

    def set_loaded_file(
        self,
        file_path_or_paths: Union[str, Path, List[Union[str, Path]]],
        metadata_override: Optional[str] = None,
    ):
        """Transition into the solid Active Document Card state."""
        if isinstance(file_path_or_paths, (str, Path)):
            paths = [Path(file_path_or_paths)]
        else:
            paths = [Path(p) for p in file_path_or_paths]

        if not paths:
            self.clear()
            return

        self._selected_paths = paths
        self._current_state = "loaded"

        if len(paths) == 1:
            p = paths[0]
            suffix = p.suffix.lower()
            archetype = FILE_TYPE_METADATA.get(suffix, DEFAULT_FILE_TYPE)

            # Icon badge styling
            self.loaded_icon_lbl.setText(archetype["icon"])
            self.loaded_icon_lbl.setStyleSheet(f"""
                QLabel {{
                    background-color: {archetype['bg']};
                    color: {archetype['fg']};
                    border: 1px solid {archetype['border']};
                    border-radius: 8px;
                    font-size: 20px;
                }}
            """)

            # Title
            self.lbl_file_name.setText(p.name)
            self.lbl_file_name.setToolTip(str(p.resolve()))

            # File size
            if p.exists() and p.is_file():
                size_str = format_file_size(p.stat().st_size)
            else:
                size_str = "File"
            self.lbl_file_size.setText(size_str)
            self.lbl_file_size.show()

            # Description / Metadata
            self.lbl_file_desc.setText(metadata_override or archetype["name"])

            # Sensitivity label detection (Word documents)
            label_name = self._detect_sensitivity(p)
            if label_name:
                self.lbl_sensitivity.setText(f"🛡️ {label_name}")
                self.lbl_sensitivity.show()
            else:
                self.lbl_sensitivity.hide()

        else:
            # Multi-file staging
            self.loaded_icon_lbl.setText("📚")
            self.loaded_icon_lbl.setStyleSheet("""
                QLabel {
                    background-color: #EFF6FF;
                    color: #1D4ED8;
                    border: 1px solid #BFDBFE;
                    border-radius: 8px;
                    font-size: 20px;
                }
            """)
            self.lbl_file_name.setText(f"{len(paths)} files selected")
            self.lbl_file_name.setToolTip("\n".join(str(p) for p in paths))

            total_bytes = sum(p.stat().st_size for p in paths if p.exists() and p.is_file())
            self.lbl_file_size.setText(format_file_size(total_bytes))
            self.lbl_file_size.show()
            self.lbl_file_desc.setText("Batch files queued")
            self.lbl_sensitivity.hide()

        self.stack.setCurrentIndex(1)
        self.update()

    def clear(self):
        """Reset to the empty / drop prompt state."""
        if not self._selected_paths and self._current_state == "idle":
            return
        self._selected_paths = []
        self._current_state = "idle"
        self._custom_status_message = None
        self.stack.setCurrentIndex(0)
        self.update()
        self.file_cleared.emit()

    def clear_selection(self):
        """Backward-compatibility alias for clear()."""
        self.clear()

    def get_selected_files(self) -> List[Path]:
        """Return the list of currently selected Path objects."""
        return list(self._selected_paths)

    def get_selected_file(self) -> Optional[Path]:
        """Return the first selected Path, or None if empty."""
        return self._selected_paths[0] if self._selected_paths else None

    # -------------------------------------------------------------------------
    # Backward Compatibility Methods for InteractiveDropLabel Callers
    # -------------------------------------------------------------------------
    def set_state(self, state: str, text: Optional[str] = None):
        """Backward-compatible state setter matching InteractiveDropLabel API."""
        state_lower = str(state).lower()
        if text:
            self._custom_status_message = text

        if state_lower in ("ready", "loaded", "idle", "empty"):
            if text and ("\n" in text or "ready:" in text.lower()):
                # Parse lines like "Ready:\nfilename.docx (30.5 KB)"
                lines = [line.strip() for line in text.split("\n") if line.strip()]
                target_line = lines[-1] if lines else text
                # Extract filename if it matches
                clean_name = target_line.split(" (")[0].replace("Ready:", "").strip()
                if clean_name and (Path(clean_name).exists() or "." in clean_name):
                    self.set_loaded_file(clean_name, metadata_override=target_line)
                    return

            if state_lower in ("idle", "empty"):
                self._current_state = "idle"
                self.stack.setCurrentIndex(0)
                if text:
                    self.lbl_empty_title.setText(text)
            else:
                self._current_state = "loaded" if self._selected_paths else "idle"
        elif state_lower == "busy":
            self._current_state = "busy"
            if text and hasattr(self, "lbl_empty_title"):
                self.lbl_empty_title.setText(text)
        elif state_lower == "error":
            self._current_state = "error"
            if text and hasattr(self, "lbl_empty_title"):
                self.lbl_empty_title.setText(text)

        self.update()

    def setText(self, text: str):
        """Backward-compatibility alias to set the primary empty-state title."""
        self.primary_title = text
        self._custom_status_message = text
        if hasattr(self, "lbl_empty_title"):
            self.lbl_empty_title.setText(text)

    def text(self) -> str:
        """Backward-compatibility getter."""
        if self._custom_status_message is not None:
            return self._custom_status_message
        if self._current_state == "loaded" and self._selected_paths:
            return str(self._selected_paths[0].name)
        return self.primary_title

    # -------------------------------------------------------------------------
    # Internal Helpers
    # -------------------------------------------------------------------------
    def _apply_state_style(self):
        self.update()

    def _on_browse_clicked(self):
        """Launch local system file dialog based on accepted extensions."""
        ext_filter = "All Files (*.*)"
        if self.accepted_extensions:
            filters = [f"*{ext}" for ext in self.accepted_extensions]
            ext_filter = f"Supported Files ({' '.join(filters)});;All Files (*.*)"

        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Document(s)",
            "",
            ext_filter,
        )
        if files:
            paths = [Path(f) for f in files]
            self.set_loaded_file(paths)
            self.file_selected.emit(files)
            self.file_dropped.emit(files)

    def _detect_sensitivity(self, path: Path) -> Optional[str]:
        """Safely extract MIP sensitivity label without throwing errors."""
        try:
            if path.suffix.lower() in (".docx", ".docm", ".dotx") and path.exists():
                from modules.word_tools.core.sensitivity_label import get_document_sensitivity_label
                label_info = get_document_sensitivity_label(path)
                if label_info and label_info.get("name"):
                    return label_info["name"]
        except Exception as e:
            logger.debug("Sensitivity check skipped for %s: %s", path, e)
        return None


# -----------------------------------------------------------------------------
# Backward-Compatible Drop-In Subclasses
# -----------------------------------------------------------------------------
class InteractiveDropLabel(ModernDropZone):
    """
    Direct drop-in replacement for legacy InteractiveDropLabel.
    Preserves constructor signature (text, accepted_extensions).
    """

    def __init__(self, text: str, accepted_extensions: Optional[List[str]] = None, parent: Optional[QWidget] = None):
        super().__init__(
            accepted_extensions=accepted_extensions,
            title=text,
            compact=True,  # Default to compact ribbon mode for toolbars
            parent=parent,
        )


class ClickableDropLabel(ModernDropZone):
    """
    Direct drop-in replacement for ClickableDropLabel in DocumentSelectorPane.
    """

    def __init__(self, text: str, accepted_extensions: Optional[List[str]] = None, parent: Optional[QWidget] = None):
        super().__init__(
            accepted_extensions=accepted_extensions,
            title=text,
            compact=False,  # Canvas mode for document selector pane
            browse_on_click=True,
            parent=parent,
        )
