"""Editing shortcuts, snippet insertion, and context-menu mixin for CodeDropTextEdit."""

from __future__ import annotations

import re

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QTextCursor
from PyQt5.QtWidgets import QAction, QMenu

from modules.puml2visio.templates.plantuml_snippets import (
    PLANTUML_SNIPPETS,
    PlantUMLSnippet,
    resolve_snippet_text,
)


class CodeEditorEditingMixin:
    """Snippet insertion, comment toggling, smart indentation, and context menu."""

    def insert_snippet(self, snippet: PlantUMLSnippet):
        """Inserts a syntax snippet at cursor, intelligently wrapping selections
        and resolving participant tokens.
        """
        cursor = self.textCursor()
        cursor.beginEditBlock()

        selected_text = cursor.selectedText().replace("\u2029", "\n")

        current_line = cursor.block().text()
        indent_len = len(current_line) - len(current_line.lstrip())
        base_indent = " " * indent_len

        resolved_text, start_offset, end_offset = resolve_snippet_text(
            puml_text=self.toPlainText(),
            snippet=snippet,
            selected_text=selected_text,
            base_indent=base_indent,
        )

        insert_pos = cursor.selectionStart()
        cursor.insertText(resolved_text)

        cursor.setPosition(insert_pos + start_offset)
        if end_offset > start_offset:
            cursor.setPosition(insert_pos + end_offset, QTextCursor.KeepAnchor)

        cursor.endEditBlock()
        self.setTextCursor(cursor)
        self.setFocus()

        if self.error_line is not None:
            self.set_error_line(None)

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

        if end > start and cursor.positionInBlock() == 0:
            last_block = last_block.previous()

        first_num = first_block.blockNumber()
        last_num = last_block.blockNumber()

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
                idx = line_str.find("'")
                if idx != -1:
                    b_cursor.setPosition(b.position() + idx)
                    remove_count = 2 if line_str[idx : idx + 2] == "' " else 1
                    b_cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, remove_count)
                    b_cursor.removeSelectedText()
            else:
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
                        if text.startswith("  "):
                            b_cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 2)
                            b_cursor.removeSelectedText()
                        elif text.startswith(" "):
                            b_cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
                            b_cursor.removeSelectedText()
                    else:
                        b_cursor.insertText("  ")

                    if curr == last_block:
                        break
                    curr = curr.next()

                cursor.endEditBlock()
                event.accept()
                return
            elif is_backtab:
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
                cursor.insertText("  ")
                event.accept()
                return

        # 4. Smart Enter Auto-Indentation
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            cursor = self.textCursor()
            line = cursor.block().text()
            indent_match = re.match(r"^(\s*)", line)
            indent = indent_match.group(1) if indent_match else ""

            stripped = line.strip()
            block_openers = (
                "alt",
                "else",
                "opt",
                "loop",
                "par",
                "critical",
                "break",
                "group",
                "box",
                "package",
                "node",
                "partition",
                "fork",
                "split",
            )
            is_opener = any(stripped.startswith(op) for op in block_openers) or (
                stripped.startswith("note ") and ":" not in stripped
            )

            super().keyPressEvent(event)

            if indent or is_opener:
                extra = "  " if is_opener else ""
                self.textCursor().insertText(indent + extra)
            event.accept()
            return

        super().keyPressEvent(event)

    def contextMenuEvent(self, event):
        menu = self.createStandardContextMenu()
        menu.addSeparator()

        comment_act = QAction("💬 Toggle Comment (Ctrl+/)", self)
        comment_act.triggered.connect(self.toggle_comment)
        menu.addAction(comment_act)

        insert_menu = QMenu("🧩 Insert Element", self)
        categories = {}
        for snip in PLANTUML_SNIPPETS:
            categories.setdefault(snip.category, []).append(snip)

        for cat_name, snips in categories.items():
            cat_menu = insert_menu.addMenu(cat_name.replace("&", "&&"))
            for snip in snips:
                act = QAction(snip.label.replace("&", "&&"), self)
                act.setToolTip(snip.description)
                act.triggered.connect(lambda checked, s=snip: self.insert_snippet(s))
                cat_menu.addAction(act)

        insert_menu.addSeparator()
        add_preset_act = QAction("➕ Add Custom Preset...", self)
        add_preset_act.triggered.connect(self._open_add_preset_dialog)
        insert_menu.addAction(add_preset_act)

        reload_act = QAction("🔄 Reload Presets from Disk", self)
        reload_act.triggered.connect(self._reload_presets)
        insert_menu.addAction(reload_act)

        menu.addMenu(insert_menu)
        menu.exec_(event.globalPos())

    def _open_add_preset_dialog(self):
        from modules.puml2visio.ui.preset_dialog import UserPresetDialog

        dialog = UserPresetDialog(parent=self)
        dialog.exec_()

    def _reload_presets(self):
        from modules.puml2visio.templates.plantuml_snippets import reload_snippets

        reload_snippets()
