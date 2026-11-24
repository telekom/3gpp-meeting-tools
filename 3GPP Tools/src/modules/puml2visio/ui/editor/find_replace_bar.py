"""Inline non-intrusive Find & Replace toolbar for CodeDropTextEdit."""

from __future__ import annotations

import re

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QTextCursor, QTextDocument
from PyQt5.QtWidgets import (
    QCheckBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.puml2visio.ui.editor.code_editor import CodeDropTextEdit


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

        self.find_lbl = QLabel("🔍 Find:")
        self.find_lbl.setStyleSheet(f"font-weight: bold; color: {C.TEXT_SECONDARY}; font-size: 11px;")
        self.find_input = QLineEdit()
        self.find_input.setPlaceholderText("Search...")
        self.find_input.setMinimumWidth(160)
        self.find_input.textChanged.connect(self._find_next)
        self.find_input.returnPressed.connect(self._find_next)

        self.match_count_lbl = QLabel("")
        self.match_count_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; min-width: 45px;")

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

        self.replace_lbl = QLabel("Replace:")
        self.replace_lbl.setStyleSheet(f"font-weight: bold; color: {C.TEXT_SECONDARY}; font-size: 11px;")
        self.replace_input = QLineEdit()
        self.replace_input.setPlaceholderText("Replace with...")
        self.replace_input.setMinimumWidth(140)

        self.replace_btn = QPushButton("Replace")
        self.replace_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.replace_btn.clicked.connect(self._replace_one)

        self.replace_all_btn = QPushButton("Replace All")
        self.replace_all_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.replace_all_btn.clicked.connect(self._replace_all)

        self.case_check = QCheckBox("Aa")
        self.case_check.setToolTip("Match Case")
        self.case_check.setStyleSheet(f"font-weight: bold; color: {C.TEXT_SECONDARY};")
        self.case_check.toggled.connect(self._find_next)

        self.close_btn = QPushButton("✕")
        self.close_btn.setToolTip("Close Find Bar (Esc)")
        self.close_btn.setStyleSheet(
            f"QPushButton {{ border: none; font-weight: bold; color: {C.TEXT_MUTED}; padding: 2px 6px; }} "
            f"QPushButton:hover {{ color: {C.DANGER_ALT}; }}"
        )
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
        self.setStyleSheet(
            f"""
            QWidget {{
                background-color: {C.BG_SURFACE}; 
                border-bottom: 1px solid {C.BORDER_SUBTLE}; 
                border-radius: 4px;
            }}
            QLineEdit {{
                border: 1px solid {C.BORDER_STRONG}; 
                border-radius: 4px; 
                padding: 3px 6px; 
                background: white;
            }}
            QLineEdit:focus {{
                border-color: {C.ACCENT_INDIGO_BLUE};
            }}
        """
        )

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
            cursor = self.editor.textCursor()
            cursor.movePosition(QTextCursor.Start)
            self.editor.setTextCursor(cursor)
            self.editor.find(query, flags)

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
            cursor = self.editor.textCursor()
            cursor.movePosition(QTextCursor.End)
            self.editor.setTextCursor(cursor)
            self.editor.find(query, flags)

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

        if cursor.hasSelection():
            sel = cursor.selectedText().replace("\u2029", "\n")
            if sel == query if self.case_check.isChecked() else sel.lower() == query.lower():
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
