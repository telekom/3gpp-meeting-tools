# --- File: src/modules/specifications/ui/spec_annotations_dialog.py ---
"""
Dialogs for managing personal specification tags and comments/notes.
Provides EditSpecAnnotationsDialog for individual specification annotations
and BatchTagDialog for batch-tagging multiple specifications.
"""
import logging
from typing import List, Optional, Set

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QCompleter,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BADGE_STYLE_INFO,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SEARCH_INPUT_STYLE,
)

logger = logging.getLogger(__name__)

PRIMARY_BUTTON_STYLE = """
QPushButton {
    background-color: #1E5C99;
    color: #FFFFFF;
    border: 1px solid #164877;
    border-radius: 4px;
    padding: 6px 16px;
    font-size: 12px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #184D80;
}
QPushButton:pressed {
    background-color: #123A61;
}
"""

TAG_CHIP_STYLE = """
QWidget#tagChip {
    background-color: #EBF3FC;
    border: 1px solid #BDD7EE;
    border-radius: 12px;
}
QLabel#tagText {
    color: #1E5C99;
    font-size: 11px;
    font-weight: 600;
}
QPushButton#tagRemoveBtn {
    background: transparent;
    border: none;
    color: #8FAADC;
    font-size: 11px;
    font-weight: bold;
    padding: 0 4px;
    margin: 0;
}
QPushButton#tagRemoveBtn:hover {
    color: #C00000;
}
"""

SUGGESTION_CHIP_STYLE = """
QPushButton {
    background-color: #F1F5F9;
    color: #475569;
    border: 1px solid #CBD5E1;
    border-radius: 10px;
    padding: 2px 8px;
    font-size: 11px;
}
QPushButton:hover {
    background-color: #E2E8F0;
    color: #1E293B;
    border-color: #94A3B8;
}
"""


class TagChip(QWidget):
    """Visual pill displaying a tag with a clickable removal button."""
    remove_requested = pyqtSignal(str)

    def __init__(self, tag_text: str, parent=None):
        super().__init__(parent)
        self.tag_text = tag_text
        self.setObjectName("tagChip")
        self.setStyleSheet(TAG_CHIP_STYLE)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 2, 6, 2)
        layout.setSpacing(4)

        label = QLabel(f"🏷️ {tag_text}")
        label.setObjectName("tagText")
        layout.addWidget(label)

        btn = QPushButton("✕")
        btn.setObjectName("tagRemoveBtn")
        btn.setCursor(Qt.PointingHandCursor)
        btn.setToolTip(f"Remove '{tag_text}'")
        btn.clicked.connect(lambda: self.remove_requested.emit(self.tag_text))
        layout.addWidget(btn)


class EditSpecAnnotationsDialog(QDialog):
    """
    Modal dialog allowing delegates to view, add, and modify custom tags
    and multi-line technical notes / comments for a specification.
    """

    def __init__(self, spec_number: str, spec_title: str, db, parent=None):
        super().__init__(parent)
        self.spec_number = spec_number.strip()
        self.spec_title = (spec_title or "").strip()
        self.db = db
        self.current_tags: Set[str] = set()

        self.setWindowTitle(f"Edit Tags & Notes: TS {self.spec_number}")
        self.resize(540, 500)
        self.setMinimumWidth(480)
        self.setMinimumHeight(420)
        self.setStyleSheet(f"""
            QDialog {{ background-color: #F8F9FA; }}
            {CARD_FRAME_STYLE}
        """)

        self._init_ui()
        self._load_data()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        # 1. Header Card
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(4)

        top_row = QHBoxLayout()
        spec_badge = QLabel(f"<b>TS {self.spec_number}</b>")
        spec_badge.setStyleSheet(BADGE_STYLE_INFO)
        top_row.addWidget(spec_badge)
        top_row.addStretch()
        header_layout.addLayout(top_row)

        if self.spec_title:
            title_lbl = QLabel(self.spec_title)
            title_lbl.setWordWrap(True)
            title_lbl.setStyleSheet("color: #4A5568; font-size: 12px;")
            header_layout.addWidget(title_lbl)

        main_layout.addWidget(header_card)

        # 2. Tags Section Card
        tags_card = QFrame()
        tags_card.setObjectName("cardFrame")
        tags_layout = QVBoxLayout(tags_card)
        tags_layout.setContentsMargins(12, 12, 12, 12)
        tags_layout.setSpacing(8)

        tags_header = QLabel("<b>Tags</b> (e.g., <i>data framework, analytics, charging</i>)")
        tags_header.setStyleSheet("color: #2D3748; font-size: 12px;")
        tags_layout.addWidget(tags_header)

        # Tag Input Row
        input_row = QHBoxLayout()
        input_row.setSpacing(6)

        self.tag_input = QLineEdit()
        self.tag_input.setPlaceholderText("Type a tag and press Enter or click Add...")
        self.tag_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.tag_input.returnPressed.connect(self._add_current_tag)

        # Autocomplete using unique tags in DB
        existing_tags = self.db.get_all_unique_tags() if self.db else []
        completer = QCompleter(existing_tags, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.tag_input.setCompleter(completer)
        input_row.addWidget(self.tag_input, stretch=1)

        add_tag_btn = QPushButton("➕ Add")
        add_tag_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        add_tag_btn.setCursor(Qt.PointingHandCursor)
        add_tag_btn.clicked.connect(self._add_current_tag)
        input_row.addWidget(add_tag_btn)
        tags_layout.addLayout(input_row)

        # Active Tag Pills Container
        self.pills_container = QWidget()
        self.pills_layout = QHBoxLayout(self.pills_container)
        self.pills_layout.setContentsMargins(0, 4, 0, 4)
        self.pills_layout.setSpacing(6)
        self.pills_layout.setAlignment(Qt.AlignLeft)

        pills_scroll = QScrollArea()
        pills_scroll.setWidgetResizable(True)
        pills_scroll.setWidget(self.pills_container)
        pills_scroll.setFixedHeight(46)
        pills_scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        tags_layout.addWidget(pills_scroll)

        # Suggested Existing Tags Row
        self.suggestions_layout = QHBoxLayout()
        self.suggestions_layout.setContentsMargins(0, 0, 0, 0)
        self.suggestions_layout.setSpacing(4)
        self.suggestions_layout.setAlignment(Qt.AlignLeft)

        self.suggestions_widget = QWidget()
        self.suggestions_widget.setLayout(self.suggestions_layout)
        tags_layout.addWidget(self.suggestions_widget)

        main_layout.addWidget(tags_card)

        # 3. Comments / Notes Section Card
        notes_card = QFrame()
        notes_card.setObjectName("cardFrame")
        notes_layout = QVBoxLayout(notes_card)
        notes_layout.setContentsMargins(12, 12, 12, 12)
        notes_layout.setSpacing(6)

        notes_header = QLabel("<b>Personal Notes & Comments</b> (Searchable across application)")
        notes_header.setStyleSheet("color: #2D3748; font-size: 12px;")
        notes_layout.addWidget(notes_header)

        self.comments_edit = QPlainTextEdit()
        self.comments_edit.setPlaceholderText(
            "Write personal notes, technical context, meeting relevance, or delegation stance...\n"
            "These comments are fully searchable in the Specifications search bar."
        )
        self.comments_edit.setStyleSheet("""
            QPlainTextEdit {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #FFFFFF;
                font-size: 12px;
                padding: 6px;
                line-height: 1.4;
            }
            QPlainTextEdit:focus {
                border: 1px solid #1E5C99;
            }
        """)
        notes_layout.addWidget(self.comments_edit, stretch=1)

        self.updated_lbl = QLabel("")
        self.updated_lbl.setStyleSheet("color: #718096; font-size: 11px; font-style: italic;")
        notes_layout.addWidget(self.updated_lbl)

        main_layout.addWidget(notes_card, stretch=1)

        # 4. Action Buttons Footer
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        save_btn = QPushButton("💾 Save Annotations")
        save_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        save_btn.setCursor(Qt.PointingHandCursor)
        save_btn.clicked.connect(self._save_and_close)
        btn_layout.addWidget(save_btn)

        main_layout.addLayout(btn_layout)

    def _load_data(self):
        """Loads saved tags and comments from the database."""
        if not self.db:
            return

        data = self.db.get_spec_annotations(self.spec_number)
        self.current_tags = set(data.get("tags", []))
        self.comments_edit.setPlainText(data.get("comments", ""))

        updated = data.get("updated_at")
        if updated:
            self.updated_lbl.setText(f"Last updated: {updated}")
        else:
            self.updated_lbl.setText("No previous notes saved.")

        self._render_tag_pills()
        self._render_suggestions()

    def _add_current_tag(self):
        raw = self.tag_input.text().strip()
        if not raw:
            return

        # Support comma-separated entry (e.g. typing "tag1, tag2")
        added = False
        for piece in raw.split(","):
            clean = piece.strip()
            if clean and clean not in self.current_tags:
                self.current_tags.add(clean)
                added = True

        self.tag_input.clear()
        if added:
            self._render_tag_pills()
            self._render_suggestions()

    def _remove_tag(self, tag: str):
        if tag in self.current_tags:
            self.current_tags.remove(tag)
            self._render_tag_pills()
            self._render_suggestions()

    def _render_tag_pills(self):
        # Clear existing pill widgets
        while self.pills_layout.count():
            item = self.pills_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        if not self.current_tags:
            no_tags = QLabel("<i>No tags assigned yet.</i>")
            no_tags.setStyleSheet("color: #A0AEC0; font-size: 11px;")
            self.pills_layout.addWidget(no_tags)
        else:
            for t in sorted(list(self.current_tags), key=lambda x: x.lower()):
                chip = TagChip(t, self.pills_container)
                chip.remove_requested.connect(self._remove_tag)
                self.pills_layout.addWidget(chip)

        self.pills_layout.addStretch()

    def _render_suggestions(self):
        # Clear existing suggestions
        while self.suggestions_layout.count():
            item = self.suggestions_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        if not self.db:
            return

        all_unique = self.db.get_all_unique_tags()
        available = [t for t in all_unique if t not in self.current_tags][:8]

        if available:
            lbl = QLabel("Suggested:")
            lbl.setStyleSheet("color: #718096; font-size: 11px;")
            self.suggestions_layout.addWidget(lbl)

            for t in available:
                s_btn = QPushButton(f"+ {t}")
                s_btn.setStyleSheet(SUGGESTION_CHIP_STYLE)
                s_btn.setCursor(Qt.PointingHandCursor)
                s_btn.setToolTip(f"Add '{t}'")
                s_btn.clicked.connect(lambda _, tag=t: self._add_suggested_tag(tag))
                self.suggestions_layout.addWidget(s_btn)

            self.suggestions_layout.addStretch()
            self.suggestions_widget.setVisible(True)
        else:
            self.suggestions_widget.setVisible(False)

    def _add_suggested_tag(self, tag: str):
        self.current_tags.add(tag)
        self._render_tag_pills()
        self._render_suggestions()

    def _save_and_close(self):
        if not self.db:
            self.reject()
            return

        comments = self.comments_edit.toPlainText().strip()
        tags_list = sorted(list(self.current_tags))

        success = self.db.set_spec_annotations(self.spec_number, tags_list, comments)
        if success:
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "Failed to save annotations to the database.")


class BatchTagDialog(QDialog):
    """
    Dialog for batch adding a tag to multiple selected specifications at once.
    """

    def __init__(self, spec_numbers: List[str], db, parent=None):
        super().__init__(parent)
        self.spec_numbers = spec_numbers
        self.db = db

        self.setWindowTitle(f"Tag Selected Specifications ({len(spec_numbers)})")
        self.resize(460, 260)
        self.setStyleSheet(f"""
            QDialog {{ background-color: #F8F9FA; }}
            {CARD_FRAME_STYLE}
        """)

        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Card
        card = QFrame()
        card.setObjectName("cardFrame")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 14, 14, 14)
        card_layout.setSpacing(8)

        count = len(self.spec_numbers)
        specs_preview = ", ".join(self.spec_numbers[:5])
        if count > 5:
            specs_preview += f" ... (+{count - 5} more)"

        info_lbl = QLabel(f"Apply a tag to <b>{count} specifications</b>:<br><span style='color: #4A5568;'>{specs_preview}</span>")
        info_lbl.setWordWrap(True)
        info_lbl.setStyleSheet("font-size: 12px;")
        card_layout.addWidget(info_lbl)

        tag_lbl = QLabel("<b>Tag to apply:</b>")
        tag_lbl.setStyleSheet("font-size: 12px; margin-top: 4px;")
        card_layout.addWidget(tag_lbl)

        self.tag_input = QLineEdit()
        self.tag_input.setPlaceholderText("e.g. data framework, security, architecture...")
        self.tag_input.setStyleSheet(SEARCH_INPUT_STYLE)

        existing_tags = self.db.get_all_unique_tags() if self.db else []
        completer = QCompleter(existing_tags, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.tag_input.setCompleter(completer)
        card_layout.addWidget(self.tag_input)

        if existing_tags:
            sugg_layout = QHBoxLayout()
            sugg_layout.setSpacing(4)
            sugg_layout.addWidget(QLabel("<span style='color: #718096; font-size: 11px;'>Recent:</span>"))
            for t in existing_tags[:5]:
                b = QPushButton(t)
                b.setStyleSheet(SUGGESTION_CHIP_STYLE)
                b.setCursor(Qt.PointingHandCursor)
                b.clicked.connect(lambda _, tag=t: self.tag_input.setText(tag))
                sugg_layout.addWidget(b)
            sugg_layout.addStretch()
            card_layout.addLayout(sugg_layout)

        layout.addWidget(card)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        apply_btn = QPushButton("🏷️ Apply Tag")
        apply_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        apply_btn.setCursor(Qt.PointingHandCursor)
        apply_btn.clicked.connect(self._apply_tag)
        btn_layout.addWidget(apply_btn)

        layout.addLayout(btn_layout)

    def _apply_tag(self):
        tag = self.tag_input.text().strip()
        if not tag:
            QMessageBox.warning(self, "Missing Tag", "Please enter a tag name to apply.")
            return

        if not self.db:
            self.reject()
            return

        count = self.db.add_tag_to_specs(self.spec_numbers, tag)
        logger.info(f"Batch applied tag '{tag}' to {count} specifications.")
        self.accept()
