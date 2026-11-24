"""PlantUML code editor widget with line-number gutter, highlighting, and drag-and-drop."""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Union

from PyQt5.QtCore import QRect, QSize, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QPainter, QTextCharFormat, QTextCursor, QTextFormat
from PyQt5.QtWidgets import QPlainTextEdit, QTextEdit, QWidget

from core.ui.theme import C
from modules.puml2visio.ui.editor.editor_mixins import CodeEditorEditingMixin
from modules.puml2visio.ui.syntax_highlighter import PlantUMLHighlighter

PUML_EXTENSIONS = (".puml", ".plantuml", ".pu", ".iuml", ".wsd", ".txt")
VISIO_EXTENSIONS = (".vsdx", ".vsdm")
SUPPORTED_EDITOR_EXTENSIONS = PUML_EXTENSIONS + VISIO_EXTENSIONS


class LineNumberArea(QWidget):
    """Gutter widget rendering 1-based line numbers for CodeDropTextEdit."""

    def __init__(self, editor: "CodeDropTextEdit"):
        super().__init__(editor)
        self.code_editor = editor

    def sizeHint(self):
        return QSize(self.code_editor.line_number_area_width(), 0)

    def paintEvent(self, event):
        self.code_editor.lineNumberAreaPaintEvent(event)


class CodeDropTextEdit(CodeEditorEditingMixin, QPlainTextEdit):
    """Enhanced PlantUML code editor with syntax highlighting, smart indentation,
    comment toggling, snippet insertion, and error-line pinpointing.
    """

    file_dropped = pyqtSignal(str)
    find_requested = pyqtSignal()

    PUML_EXTENSIONS = PUML_EXTENSIONS
    VISIO_EXTENSIONS = VISIO_EXTENSIONS
    SUPPORTED_EXTENSIONS = SUPPORTED_EDITOR_EXTENSIONS

    @classmethod
    def is_supported_file(cls, file_path: str) -> bool:
        if not file_path:
            return False
        return Path(file_path).suffix.lower() in cls.SUPPORTED_EXTENSIONS

    @classmethod
    def is_puml_file(cls, file_path: str) -> bool:
        if not file_path:
            return False
        return Path(file_path).suffix.lower() in cls.PUML_EXTENSIONS

    @classmethod
    def is_visio_file(cls, file_path: str) -> bool:
        if not file_path:
            return False
        return Path(file_path).suffix.lower() in cls.VISIO_EXTENSIONS

    def __init__(self, parent=None):
        super().__init__(parent)
        self.line_number_area = LineNumberArea(self)
        self.error_line: Optional[int] = None
        self.error_tooltip: str = ""

        self.highlighter = PlantUMLHighlighter(self.document())

        self.blockCountChanged.connect(self.update_line_number_area_width)
        self.updateRequest.connect(self.update_line_number_area)
        self.cursorPositionChanged.connect(self.highlight_current_line)

        self.update_line_number_area_width(0)
        self.highlight_current_line()

        self.default_style = f"""
            QPlainTextEdit {{
                font-family: Consolas, 'Cascadia Code', 'Courier New', monospace; 
                font-size: 13px; 
                border: 1px solid {C.BORDER_STRONG}; 
                border-radius: 6px; 
                background-color: {C.BG_CARD};
                color: {C.TEXT_STRONG};
                line-height: 1.4;
            }}
            QPlainTextEdit:focus {{
                border: 2px solid {C.ACCENT_INDIGO_BLUE};
            }}
        """
        self.hover_style = self.default_style.replace(
            f"border: 1px solid {C.BORDER_STRONG};",
            f"border: 2px dashed {C.ACCENT_INDIGO_BLUE}; background-color: {C.ACCENT_SOFT};",
        )
        self.setStyleSheet(self.default_style)
        self.setLineWrapMode(QPlainTextEdit.NoWrap)

        self.setTabStopDistance(
            self.fontMetrics().horizontalAdvance(" ") * 4
            if hasattr(self.fontMetrics(), "horizontalAdvance")
            else self.fontMetrics().width(" ") * 4
        )

    def line_number_area_width(self) -> int:
        digits = 1
        max_val = max(1, self.blockCount())
        while max_val >= 10:
            max_val /= 10
            digits += 1

        fm = self.fontMetrics()
        width_char = fm.horizontalAdvance("9") if hasattr(fm, "horizontalAdvance") else fm.width("9")
        return 14 + width_char * digits

    def update_line_number_area_width(self, _):
        self.setViewportMargins(self.line_number_area_width(), 0, 0, 0)

    def update_line_number_area(self, rect, dy):
        if dy:
            self.line_number_area.scroll(0, dy)
        else:
            self.line_number_area.update(0, rect.y(), self.line_number_area.width(), rect.height())
        if rect.contains(self.viewport().rect()):
            self.update_line_number_area_width(0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        cr = self.contentsRect()
        self.line_number_area.setGeometry(QRect(cr.left(), cr.top(), self.line_number_area_width(), cr.height()))

    def lineNumberAreaPaintEvent(self, event):
        painter = QPainter(self.line_number_area)
        painter.fillRect(event.rect(), QColor(C.BG_SURFACE))

        block = self.firstVisibleBlock()
        block_number = block.blockNumber()
        top = round(self.blockBoundingGeometry(block).translated(self.contentOffset()).top())
        bottom = top + round(self.blockBoundingRect(block).height())

        fm = self.fontMetrics()
        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                line_idx = block_number + 1
                number_str = str(line_idx)

                if self.error_line is not None and self.error_line == line_idx:
                    painter.fillRect(0, top, self.line_number_area.width(), fm.height(), QColor(C.DANGER_SOFT_HOVER))
                    painter.setPen(QColor(C.DANGER_ALT))
                else:
                    painter.setPen(QColor(C.TEXT_FAINT))

                painter.drawText(
                    0,
                    top,
                    self.line_number_area.width() - 6,
                    fm.height(),
                    Qt.AlignRight | Qt.AlignVCenter,
                    number_str,
                )

            block = block.next()
            top = bottom
            bottom = top + round(self.blockBoundingRect(block).height())
            block_number += 1

    def highlight_current_line(self):
        extra_selections = []

        if self.error_line is not None and self.error_line > 0:
            block = self.document().findBlockByNumber(self.error_line - 1)
            if block.isValid():
                error_sel = QTextEdit.ExtraSelection()
                error_format = QTextCharFormat()
                error_format.setBackground(QColor(C.DANGER_SOFT_HOVER))
                error_format.setProperty(QTextFormat.FullWidthSelection, True)
                error_format.setUnderlineColor(QColor(C.DANGER_ALT))
                error_format.setUnderlineStyle(QTextCharFormat.WaveUnderline)
                error_sel.format = error_format
                error_sel.cursor = QTextCursor(block)
                error_sel.cursor.clearSelection()
                extra_selections.append(error_sel)

        if not self.isReadOnly():
            selection = QTextEdit.ExtraSelection()
            line_color = QColor(C.BG_CANVAS)
            selection.format.setBackground(line_color)
            selection.format.setProperty(QTextFormat.FullWidthSelection, True)
            selection.cursor = self.textCursor()
            selection.cursor.clearSelection()
            extra_selections.append(selection)

        self.setExtraSelections(extra_selections)

    def set_error_line(self, line_no: Optional[int], error_msg: str = ""):
        """Sets or clears the highlighted error line reported by the PlantUML engine."""
        self.error_line = line_no
        self.error_tooltip = error_msg
        self.line_number_area.update()
        self.highlight_current_line()

    def load_file(self, file_path: Union[str, Path]) -> bool:
        """Loads a PlantUML or text file from disk into the editor with robust encoding detection."""
        path = Path(file_path)
        if not path.is_file():
            return False
        content = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                content = path.read_text(encoding=enc)
                break
            except (UnicodeDecodeError, LookupError):
                continue
        if content is None:
            content = path.read_text(encoding="utf-8", errors="replace")

        cursor = self.textCursor()
        cursor.beginEditBlock()
        cursor.select(QTextCursor.Document)
        cursor.insertText(content)
        cursor.endEditBlock()
        cursor.setPosition(0)
        self.setTextCursor(cursor)
        self.set_error_line(None)
        return True

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(self.is_supported_file(url.toLocalFile()) for url in urls):
                self.setStyleSheet(self.hover_style)
                event.acceptProposedAction()
                return
            event.ignore()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(self.is_supported_file(url.toLocalFile()) for url in urls):
                event.acceptProposedAction()
                return
            event.ignore()
            return
        super().dragMoveEvent(event)

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        self.setStyleSheet(self.default_style)
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if self.is_supported_file(file_path):
                    self.file_dropped.emit(file_path)
                    event.acceptProposedAction()
                    return
            event.ignore()
            return
        super().dropEvent(event)
