"""ModernDropZone widget and backward-compatible drop label subclasses."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import List, Optional, Union

from PyQt5.QtCore import QRectF, Qt, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QPainter, QPen
from PyQt5.QtWidgets import QFileDialog, QFrame, QSizePolicy, QWidget

from .metadata import (
    COLOR_BUSY_BG,
    COLOR_BUSY_BORDER,
    COLOR_DRAG_BG,
    COLOR_DRAG_BORDER,
    COLOR_ERROR_BG,
    COLOR_ERROR_BORDER,
    COLOR_IDLE_BG,
    COLOR_IDLE_BORDER,
    COLOR_IDLE_HOVER_BORDER,
    COLOR_LOADED_BG,
    COLOR_LOADED_BORDER,
    DEFAULT_FILE_TYPE,
    FILE_TYPE_METADATA,
    format_file_size,
)
from .pages import _DropZonePagesMixin
from .styles import MULTI_FILE_ICON_BADGE_STYLE, build_file_icon_badge_style

logger = logging.getLogger(__name__)


class ModernDropZone(_DropZonePagesMixin, QFrame):
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
        if compact:
            self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        else:
            self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self._init_ui()
        self._apply_state_style()

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
            self.loaded_icon_lbl.setStyleSheet(
                build_file_icon_badge_style(
                    archetype["bg"],
                    archetype["fg"],
                    archetype["border"],
                )
            )

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
            self.loaded_icon_lbl.setStyleSheet(MULTI_FILE_ICON_BADGE_STYLE)
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
