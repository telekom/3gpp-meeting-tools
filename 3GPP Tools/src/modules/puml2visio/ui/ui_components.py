# --- File: src/modules/puml2visio/ui/ui_components.py ---
"""Enhanced code editor components for PlantUML Mini-IDE:

- LineNumberArea: Gutter rendering line numbers
- CodeDropTextEdit: PlantUML editor with syntax highlighting, smart indentation,
  comment toggling, snippet insertion, error line highlighting, and .vsdx drag-drop
- FindReplaceBar: Inline non-intrusive find and replace toolbar
"""

from typing import Optional
import re

from PyQt5.QtCore import QPoint, QRect, QRegularExpression, QSize, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QPainter, QTextCharFormat, QTextCursor, QTextDocument, QTextFormat
from PyQt5.QtWidgets import (
    QAction,
    QCheckBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.puml2visio.templates.plantuml_snippets import (
    PLANTUML_SNIPPETS,
    PlantUMLSnippet,
    resolve_snippet_text,
)
from modules.puml2visio.ui.syntax_highlighter import PlantUMLHighlighter


class LineNumberArea(QWidget):
    """Gutter widget rendering 1-based line numbers for CodeDropTextEdit."""

    def __init__(self, editor: "CodeDropTextEdit"):
        super().__init__(editor)
        self.code_editor = editor

    def sizeHint(self):
        return QSize(self.code_editor.line_number_area_width(), 0)

    def paintEvent(self, event):
        self.code_editor.lineNumberAreaPaintEvent(event)


class CodeDropTextEdit(QPlainTextEdit):
    """Enhanced PlantUML code editor with syntax highlighting, smart indentation,

    comment toggling, snippet insertion, and error-line pinpointing.
    """
    file_dropped = pyqtSignal(str)
    find_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.line_number_area = LineNumberArea(self)
        self.error_line: Optional[int] = None
        self.error_tooltip: str = ""

        # Syntax Highlighting
        self.highlighter = PlantUMLHighlighter(self.document())

        # Signals
        self.blockCountChanged.connect(self.update_line_number_area_width)
        self.updateRequest.connect(self.update_line_number_area)
        self.cursorPositionChanged.connect(self.highlight_current_line)

        self.update_line_number_area_width(0)
        self.highlight_current_line()

        # Modern Visual Styling
        self.default_style = """
            QPlainTextEdit {
                font-family: Consolas, 'Cascadia Code', 'Courier New', monospace; 
                font-size: 13px; 
                border: 1px solid #CBD5E1; 
                border-radius: 6px; 
                background-color: #FFFFFF;
                color: #1E293B;
                line-height: 1.4;
            }
            QPlainTextEdit:focus {
                border: 2px solid #395396;
            }
        """
        self.hover_style = self.default_style.replace(
            "border: 1px solid #CBD5E1;",
            "border: 2px dashed #395396; background-color: #EBF3FC;"
        )
        self.setStyleSheet(self.default_style)
        self.setLineWrapMode(QPlainTextEdit.NoWrap)

        # Tab key inserts 2 spaces by default
        self.setTabStopDistance(self.fontMetrics().horizontalAdvance(' ') * 4 if hasattr(self.fontMetrics(), 'horizontalAdvance') else self.fontMetrics().width(' ') * 4)

    # -------------------------------------------------------------------------
    # Line Number Gutter Logic
    # -------------------------------------------------------------------------
    def line_number_area_width(self) -> int:
        digits = 1
        max_val = max(1, self.blockCount())
        while max_val >= 10:
            max_val /= 10
            digits += 1

        fm = self.fontMetrics()
        width_char = fm.horizontalAdvance('9') if hasattr(fm, 'horizontalAdvance') else fm.width('9')
        space = 14 + width_char * digits
        return space

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
        painter.fillRect(event.rect(), QColor("#F8FAFC"))

        block = self.firstVisibleBlock()
        block_number = block.blockNumber()
        top = round(self.blockBoundingGeometry(block).translated(self.contentOffset()).top())
        bottom = top + round(self.blockBoundingRect(block).height())

        fm = self.fontMetrics()
        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                line_idx = block_number + 1
                number_str = str(line_idx)

                # Distinguish error line in gutter
                if self.error_line is not None and self.error_line == line_idx:
                    painter.fillRect(0, top, self.line_number_area.width(), fm.height(), QColor("#FEE2E2"))
                    painter.setPen(QColor("#DC2626"))
                else:
                    painter.setPen(QColor("#94A3B8"))

                painter.drawText(
                    0, top, self.line_number_area.width() - 6, fm.height(),
                    Qt.AlignRight | Qt.AlignVCenter, number_str
                )

            block = block.next()
            top = bottom
            bottom = top + round(self.blockBoundingRect(block).height())
            block_number += 1

    # -------------------------------------------------------------------------
    # Line Highlighting (Current Line + Error Line)
    # -------------------------------------------------------------------------
    def highlight_current_line(self):
        extra_selections = []

        # 1. Error Line Highlighting (Soft red full-width selection)
        if self.error_line is not None and self.error_line > 0:
            block = self.document().findBlockByNumber(self.error_line - 1)
            if block.isValid():
                error_sel = QTextEdit.ExtraSelection()
                error_format = QTextCharFormat()
                error_format.setBackground(QColor("#FEE2E2"))
                error_format.setProperty(QTextFormat.FullWidthSelection, True)
                error_format.setUnderlineColor(QColor("#DC2626"))
                error_format.setUnderlineStyle(QTextCharFormat.WaveUnderline)
                error_sel.format = error_format
                error_sel.cursor = QTextCursor(block)
                error_sel.cursor.clearSelection()
                extra_selections.append(error_sel)

        # 2. Active Cursor Line Highlighting (Subtle light blue)
        if not self.isReadOnly():
            selection = QTextEdit.ExtraSelection()
            line_color = QColor("#F1F5F9")
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

    # -------------------------------------------------------------------------
    # Snippet & Syntax Element Insertion
    # -------------------------------------------------------------------------
    def insert_snippet(self, snippet: PlantUMLSnippet):
        """Inserts a syntax snippet at cursor, intelligently wrapping selections

        and resolving participant tokens.
        """
        cursor = self.textCursor()
        cursor.beginEditBlock()

        # Qt uses \u2029 for paragraph separators in selectedText()
        selected_text = cursor.selectedText().replace("\u2029", "\n")

        # Determine indentation of the current line
        current_line = cursor.block().text()
        indent_len = len(current_line) - len(current_line.lstrip())
        base_indent = " " * indent_len

        resolved_text, start_offset, end_offset = resolve_snippet_text(
            puml_text=self.toPlainText(),
            snippet=snippet,
            selected_text=selected_text,
            base_indent=base_indent
        )

        insert_pos = cursor.selectionStart()
        cursor.insertText(resolved_text)

        # Reposition cursor to designated placeholder or target
        cursor.setPosition(insert_pos + start_offset)
        if end_offset > start_offset:
            cursor.setPosition(insert_pos + end_offset, QTextCursor.KeepAnchor)

        cursor.endEditBlock()
        self.setTextCursor(cursor)
        self.setFocus()

        # Clear any obsolete error highlight on edit
        if self.error_line is not None:
            self.set_error_line(None)

    # -------------------------------------------------------------------------
    # Keyboard & Editing Shortcuts (Indentation, Auto-indent, Comment toggle)
    # -------------------------------------------------------------------------
    def toggle_comment(self):
        """Toggles PlantUML single-quote comments (' ) on selected lines."""
        cursor = self.textCursor()
        cursor.beginEditBlock()

        had_selection = cursor.hasSelection()
        start = cursor.selectionStart()
        end = cursor.selectionEnd()

        cursor.setPosition(start)
        first_block = cursor.block()
        cursor.setPosition(end)
        last_block = cursor.block()

        # If selection ends right at the beginning of a block, exclude that block
        if end > start and cursor.positionInBlock() == 0:
            last_block = last_block.previous()

        first_num = first_block.blockNumber()
        last_num = last_block.blockNumber()

        # Check if all blocks are already commented
        blocks = []
        curr = first_block
        all_commented = True
        while curr.isValid():
            blocks.append(curr)
            stripped = curr.text().strip()
            if stripped and not stripped.startswith("'"):
                all_commented = False
            if curr == last_block:
                break
            curr = curr.next()

        for b in blocks:
            b_cursor = QTextCursor(b)
            line_str = b.text()
            if not line_str.strip():
                continue
            if all_commented:
                # Remove leading comment
                idx = line_str.find("'")
                if idx != -1:
                    b_cursor.setPosition(b.position() + idx)
                    remove_count = 2 if line_str[idx:idx + 2] == "' " else 1
                    b_cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, remove_count)
                    b_cursor.removeSelectedText()
            else:
                # Add leading comment after indentation
                indent = len(line_str) - len(line_str.lstrip())
                b_cursor.setPosition(b.position() + indent)
                b_cursor.insertText("' ")

        cursor.endEditBlock()

        if had_selection:
            new_first = self.document().findBlockByNumber(first_num)
            new_last = self.document().findBlockByNumber(last_num)
            if new_first.isValid() and new_last.isValid():
                new_cursor = QTextCursor(new_first)
                new_cursor.setPosition(new_last.position() + new_last.length() - 1, QTextCursor.KeepAnchor)
                self.setTextCursor(new_cursor)
        else:
            self.setTextCursor(cursor)

    def keyPressEvent(self, event):
        # 1. Comment Toggle (Ctrl+/ or Ctrl+')
        if event.modifiers() == Qt.ControlModifier and event.key() in (Qt.Key_Slash, Qt.Key_Apostrophe):
            self.toggle_comment()
            event.accept()
            return

        # 2. Find Bar (Ctrl+F)
        if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_F:
            self.find_requested.emit()
            event.accept()
            return

        # 3. Tab / Shift+Tab Block Indentation
        if event.key() == Qt.Key_Tab or event.key() == Qt.Key_Backtab:
            cursor = self.textCursor()
            has_selection = cursor.hasSelection()
            is_backtab = event.key() == Qt.Key_Backtab or (event.modifiers() & Qt.ShiftModifier)

            if has_selection:
                cursor.beginEditBlock()
                start = cursor.selectionStart()
                end = cursor.selectionEnd()

                cursor.setPosition(start)
                first_block = cursor.block()
                cursor.setPosition(end)
                last_block = cursor.block()
                if end > start and cursor.positionInBlock() == 0:
                    last_block = last_block.previous()

                curr = first_block
                while curr.isValid():
                    b_cursor = QTextCursor(curr)
                    text = curr.text()
                    if is_backtab:
                        # Unindent by up to 2 spaces
                        if text.startswith("  "):
                            b_cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 2)
                            b_cursor.removeSelectedText()
                        elif text.startswith(" "):
                            b_cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
                            b_cursor.removeSelectedText()
                    else:
                        # Indent by 2 spaces
                        b_cursor.insertText("  ")

                    if curr == last_block:
                        break
                    curr = curr.next()

                cursor.endEditBlock()
                event.accept()
                return
            elif is_backtab:
                # Unindent single line if no selection
                cursor.beginEditBlock()
                b_cursor = QTextCursor(cursor.block())
                text = cursor.block().text()
                if text.startswith("  "):
                    b_cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 2)
                    b_cursor.removeSelectedText()
                elif text.startswith(" "):
                    b_cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
                    b_cursor.removeSelectedText()
                cursor.endEditBlock()
                event.accept()
                return
            else:
                # Insert 2 spaces for soft tabs
                cursor.insertText("  ")
                event.accept()
                return

        # 4. Smart Enter Auto-Indentation
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            cursor = self.textCursor()
            line = cursor.block().text()
            indent_match = re.match(r"^(\s*)", line)
            indent = indent_match.group(1) if indent_match else ""

            # Check if current line opens an indented block
            stripped = line.strip()
            block_openers = (
                "alt", "else", "opt", "loop", "par", "critical", "break",
                "group", "box", "package", "node", "partition", "fork", "split"
            )
            is_opener = any(stripped.startswith(op) for op in block_openers) or stripped.startswith("note ") and not ":" in stripped

            super().keyPressEvent(event)

            # Insert matching indentation (+ extra 2 spaces if opening a block)
            if indent or is_opener:
                extra = "  " if is_opener else ""
                self.textCursor().insertText(indent + extra)
            event.accept()
            return

        super().keyPressEvent(event)

    # -------------------------------------------------------------------------
    # Context Menu Integration
    # -------------------------------------------------------------------------
    def contextMenuEvent(self, event):
        menu = self.createStandardContextMenu()

        menu.addSeparator()

        comment_act = QAction("💬 Toggle Comment (Ctrl+/)", self)
        comment_act.triggered.connect(self.toggle_comment)
        menu.addAction(comment_act)

        # Nested Quick-Insert Snippets Submenu
        insert_menu = QMenu("🧩 Insert Element", self)
        categories = {}
        for snip in PLANTUML_SNIPPETS:
            categories.setdefault(snip.category, []).append(snip)

        for cat_name, snips in categories.items():
            # Escape '&' as '&&' so Qt does not treat it as a keyboard mnemonic
            cat_menu = insert_menu.addMenu(cat_name.replace("&", "&&"))
            for snip in snips:
                act = QAction(snip.label.replace("&", "&&"), self)
                act.setToolTip(snip.description)
                act.triggered.connect(lambda checked, s=snip: self.insert_snippet(s))
                cat_menu.addAction(act)

        menu.addMenu(insert_menu)
        menu.exec_(event.globalPos())

    # -------------------------------------------------------------------------
    # Drag & Drop for .vsdx file extraction
    # -------------------------------------------------------------------------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith('.vsdx') for url in urls):
                self.setStyleSheet(self.hover_style)
                event.acceptProposedAction()
                return
        super().dragEnterEvent(event)

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        self.setStyleSheet(self.default_style)
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith('.vsdx'):
                    self.file_dropped.emit(file_path)
                    event.acceptProposedAction()
                    return
        super().dropEvent(event)


class FindReplaceBar(QWidget):
    """Inline non-intrusive Find & Replace toolbar for CodeDropTextEdit."""

    def __init__(self, editor: CodeDropTextEdit, parent=None):
        super().__init__(parent)
        self.editor = editor
        self._setup_ui()
        self.hide()

    def _setup_ui(self):
        layout = QHBoxLayout()
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(6)

        # Find Input
        self.find_lbl = QLabel("🔍 Find:")
        self.find_lbl.setStyleSheet("font-weight: bold; color: #475569; font-size: 11px;")
        self.find_input = QLineEdit()
        self.find_input.setPlaceholderText("Search...")
        self.find_input.setMinimumWidth(160)
        self.find_input.textChanged.connect(self._find_next)
        self.find_input.returnPressed.connect(self._find_next)

        self.match_count_lbl = QLabel("")
        self.match_count_lbl.setStyleSheet("color: #64748B; font-size: 11px; min-width: 45px;")

        self.prev_btn = QPushButton("▲")
        self.prev_btn.setToolTip("Find Previous (Shift+F3)")
        self.prev_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.prev_btn.setFixedWidth(28)
        self.prev_btn.clicked.connect(self._find_prev)

        self.next_btn = QPushButton("▼")
        self.next_btn.setToolTip("Find Next (F3)")
        self.next_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.next_btn.setFixedWidth(28)
        self.next_btn.clicked.connect(self._find_next)

        # Replace Input
        self.replace_lbl = QLabel("Replace:")
        self.replace_lbl.setStyleSheet("font-weight: bold; color: #475569; font-size: 11px;")
        self.replace_input = QLineEdit()
        self.replace_input.setPlaceholderText("Replace with...")
        self.replace_input.setMinimumWidth(140)

        self.replace_btn = QPushButton("Replace")
        self.replace_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.replace_btn.clicked.connect(self._replace_one)

        self.replace_all_btn = QPushButton("Replace All")
        self.replace_all_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.replace_all_btn.clicked.connect(self._replace_all)

        # Case Sensitive Checkbox
        self.case_check = QCheckBox("Aa")
        self.case_check.setToolTip("Match Case")
        self.case_check.setStyleSheet("font-weight: bold; color: #475569;")
        self.case_check.toggled.connect(self._find_next)

        # Close button
        self.close_btn = QPushButton("✕")
        self.close_btn.setToolTip("Close Find Bar (Esc)")
        self.close_btn.setStyleSheet("QPushButton { border: none; font-weight: bold; color: #64748B; padding: 2px 6px; } QPushButton:hover { color: #DC2626; }")
        self.close_btn.clicked.connect(self.hide_and_focus_editor)

        layout.addWidget(self.find_lbl)
        layout.addWidget(self.find_input)
        layout.addWidget(self.match_count_lbl)
        layout.addWidget(self.prev_btn)
        layout.addWidget(self.next_btn)
        layout.addSpacing(10)
        layout.addWidget(self.replace_lbl)
        layout.addWidget(self.replace_input)
        layout.addWidget(self.replace_btn)
        layout.addWidget(self.replace_all_btn)
        layout.addWidget(self.case_check)
        layout.addStretch()
        layout.addWidget(self.close_btn)

        self.setLayout(layout)
        self.setStyleSheet("""
            QWidget {
                background-color: #F8FAFC; 
                border-bottom: 1px solid #E2E8F0; 
                border-radius: 4px;
            }
            QLineEdit {
                border: 1px solid #CBD5E1; 
                border-radius: 4px; 
                padding: 3px 6px; 
                background: white;
            }
            QLineEdit:focus {
                border-color: #395396;
            }
        """)

    def show_and_focus(self):
        """Shows the find bar and focuses the search input."""
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            sel = cursor.selectedText().replace("\u2029", " ")
            if "\n" not in sel and len(sel) < 80:
                self.find_input.setText(sel)
        self.show()
        self.find_input.selectAll()
        self.find_input.setFocus()
        self._update_match_count()

    def hide_and_focus_editor(self):
        self.hide()
        self.editor.setFocus()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.hide_and_focus_editor()
            event.accept()
            return
        super().keyPressEvent(event)

    def _find_next(self):
        query = self.find_input.text()
        if not query:
            self.match_count_lbl.setText("")
            return
        flags = QTextDocument.FindFlags()
        if self.case_check.isChecked():
            flags |= QTextDocument.FindCaseSensitively

        found = self.editor.find(query, flags)
        if not found:
            # Wrap around to document beginning
            cursor = self.editor.textCursor()
            cursor.movePosition(QTextCursor.Start)
            self.editor.setTextCursor(cursor)
            found = self.editor.find(query, flags)

        self._update_match_count()

    def _find_prev(self):
        query = self.find_input.text()
        if not query:
            self.match_count_lbl.setText("")
            return
        flags = QTextDocument.FindBackward
        if self.case_check.isChecked():
            flags |= QTextDocument.FindCaseSensitively

        found = self.editor.find(query, flags)
        if not found:
            # Wrap around to document end
            cursor = self.editor.textCursor()
            cursor.movePosition(QTextCursor.End)
            self.editor.setTextCursor(cursor)
            found = self.editor.find(query, flags)

        self._update_match_count()

    def _update_match_count(self):
        query = self.find_input.text()
        if not query:
            self.match_count_lbl.setText("")
            return
        flags = re.IGNORECASE if not self.case_check.isChecked() else 0
        matches = len(re.findall(re.escape(query), self.editor.toPlainText(), flags=flags))
        self.match_count_lbl.setText(f"{matches} found" if matches else "No match")

    def _replace_one(self):
        cursor = self.editor.textCursor()
        query = self.find_input.text()
        replacement = self.replace_input.text()
        if not query:
            return

        flags = 0 if self.case_check.isChecked() else re.IGNORECASE
        if cursor.hasSelection():
            sel = cursor.selectedText().replace("\u2029", "\n")
            if (sel == query if self.case_check.isChecked() else sel.lower() == query.lower()):
                cursor.insertText(replacement)
                self._find_next()
                return

        self._find_next()
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            cursor.insertText(replacement)

    def _replace_all(self):
        query = self.find_input.text()
        replacement = self.replace_input.text()
        if not query:
            return

        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        text = self.editor.toPlainText()
        flags = 0 if self.case_check.isChecked() else re.IGNORECASE
        new_text, count = re.subn(re.escape(query), replacement, text, flags=flags)
        if count > 0:
            cursor.select(QTextCursor.Document)
            cursor.insertText(new_text)
            self.match_count_lbl.setText(f"Replaced {count}")
        else:
            self.match_count_lbl.setText("0 replaced")
        cursor.endEditBlock()
