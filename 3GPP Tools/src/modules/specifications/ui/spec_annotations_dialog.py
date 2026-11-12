# --- File: src/modules/specifications/ui/spec_annotations_dialog.py ---
"""
Dialogs for managing personal specification tags and comments/notes.
Provides EditSpecAnnotationsDialog for individual specification annotations
and BatchTagDialog for batch-tagging multiple specifications.
"""
import logging
from typing import List, Set

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
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.ui_components import (
    BADGE_STYLE_INFO,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SEARCH_INPUT_STYLE,
)
from core.ui.theme import C

logger = logging.getLogger(__name__)

PRIMARY_BUTTON_STYLE = f"""
QPushButton {{
    background-color: {C.ACCENT_PRIMARY};
    color: {C.BG_CARD};
    border: 1px solid #164877;
    border-radius: 4px;
    padding: 6px 16px;
    font-size: 12px;
    font-weight: bold;
}}
QPushButton:hover {{
    background-color: #184D80;
}}
QPushButton:pressed {{
    background-color: #123A61;
}}
"""

TAG_CHIP_STYLE = f"""
QFrame#tagChip {{
    background-color: {C.ACCENT_SOFT};
    border: 1px solid {C.INFO_BORDER_MEDIUM};
    border-radius: 12px;
}}
QFrame#tagChip:hover {{
    background-color: {C.INFO_SOFT_STRONG};
    border-color: #60A5FA;
}}
QLabel#tagText {{
    color: {C.INFO_DARK};
    font-size: 11px;
    font-weight: 600;
    background: transparent;
    border: none;
}}
QPushButton#tagRemoveBtn {{
    background: transparent;
    border: none;
    color: {C.TEXT_MUTED};
    font-size: 11px;
    font-weight: bold;
    padding: 0 4px;
    margin: 0;
}}
QPushButton#tagRemoveBtn:hover {{
    color: {C.DANGER_ALT};
}}
"""

SUGGESTION_CHIP_STYLE = f"""
QPushButton {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_BODY};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: 11px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 500;
}}
QPushButton:hover {{
    background-color: {C.ACCENT_SOFT};
    color: {C.ACCENT_PRIMARY};
    border-color: {C.INFO_BORDER_MEDIUM};
}}
QPushButton:pressed {{
    background-color: {C.INFO_SOFT_STRONG};
}}
"""


class TagChip(QFrame):
    """Visual pill displaying an assigned tag with a clickable removal button."""
    remove_requested = pyqtSignal(str)

    def __init__(self, tag_text: str, parent=None):
        super().__init__(parent)
        self.tag_text = tag_text
        self.setObjectName("tagChip")
        self.setStyleSheet(TAG_CHIP_STYLE)
        self.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(9, 3, 6, 3)
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
        self.resize(600, 540)
        self.setMinimumWidth(540)
        self.setMinimumHeight(440)
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

        tags_header_row = QHBoxLayout()
        tags_header = QLabel("<b>Tags</b> (e.g., <i>data framework, analytics, charging</i>)")
        tags_header.setStyleSheet("color: #2D3748; font-size: 12px;")
        tags_header_row.addWidget(tags_header)
        tags_header_row.addStretch()

        self.tag_count_lbl = QLabel("")
        self.tag_count_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; font-weight: 600;")
        tags_header_row.addWidget(self.tag_count_lbl)
        tags_layout.addLayout(tags_header_row)

        # Tag Input Row
        input_row = QHBoxLayout()
        input_row.setSpacing(6)

        self.tag_input = QLineEdit()
        self.tag_input.setPlaceholderText("Type one or more comma-separated tags and press Enter...")
        self.tag_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.tag_input.returnPressed.connect(self._add_current_tag)
        self.tag_input.textChanged.connect(lambda _: self._render_suggestions())

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

        # Active Tag Pills Container (auto-wrapping FlowLayout)
        self.pills_container = QFrame()
        self.pills_container.setObjectName("assignedTagsBox")
        self.pills_layout = FlowLayout(self.pills_container, margin=8, h_spacing=6, v_spacing=6)
        tags_layout.addWidget(self.pills_container)

        # Suggested Existing Tags Section (Header + auto-wrapping FlowLayout)
        self.suggestions_widget = QWidget()
        sugg_vbox = QVBoxLayout(self.suggestions_widget)
        sugg_vbox.setContentsMargins(0, 2, 0, 0)
        sugg_vbox.setSpacing(5)

        self.suggestions_header_lbl = QLabel(f"💡 <b>Suggested Tags</b> <span style='color: {C.TEXT_FAINT};'>(click to add)</span>:")
        self.suggestions_header_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        sugg_vbox.addWidget(self.suggestions_header_lbl)

        self.suggestions_chips_container = QWidget()
        self.suggestions_layout = FlowLayout(
            self.suggestions_chips_container, margin=0, h_spacing=6, v_spacing=6
        )
        sugg_vbox.addWidget(self.suggestions_chips_container)
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
        self.comments_edit.setStyleSheet(f"""
            QPlainTextEdit {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background-color: {C.BG_CARD};
                font-size: 12px;
                padding: 6px;
                line-height: 1.4;
            }}
            QPlainTextEdit:focus {{
                border: 1px solid {C.ACCENT_PRIMARY};
            }}
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

        self.tag_input.blockSignals(True)
        self.tag_input.clear()
        self.tag_input.blockSignals(False)

        if added:
            self._render_tag_pills()
        self._render_suggestions()

    def _remove_tag(self, tag: str):
        if tag in self.current_tags:
            self.current_tags.remove(tag)
            self._render_tag_pills()
            self._render_suggestions()

    def _render_tag_pills(self):
        self.pills_layout.clear()

        count = len(self.current_tags)
        self.tag_count_lbl.setText(f"{count} tag{'s' if count != 1 else ''} assigned" if count else "")

        if not self.current_tags:
            self.pills_container.setStyleSheet(f"""
                QFrame#assignedTagsBox {{
                    background-color: {C.BG_SURFACE};
                    border: 1px dashed {C.BORDER_STRONG};
                    border-radius: 6px;
                }}
            """)
            no_tags = QLabel("<i>🏷️ No tags assigned yet — type above or click a suggestion below.</i>")
            no_tags.setStyleSheet(f"color: {C.TEXT_FAINT}; font-size: 11px; border: none; background: transparent;")
            self.pills_layout.addWidget(no_tags)
        else:
            self.pills_container.setStyleSheet(f"""
                QFrame#assignedTagsBox {{
                    background-color: {C.BG_SURFACE};
                    border: 1px solid {C.BORDER_SUBTLE};
                    border-radius: 6px;
                }}
            """)
            for t in sorted(list(self.current_tags), key=lambda x: x.lower()):
                chip = TagChip(t, self.pills_container)
                chip.remove_requested.connect(self._remove_tag)
                self.pills_layout.addWidget(chip)

        self.pills_container.updateGeometry()

    def _render_suggestions(self):
        self.suggestions_layout.clear()

        if not self.db:
            self.suggestions_widget.setVisible(False)
            return

        all_unique = self.db.get_all_unique_tags()
        available = [t for t in all_unique if t not in self.current_tags]

        filter_text = self.tag_input.text().strip().lower() if hasattr(self, "tag_input") else ""
        if filter_text:
            matching = [t for t in available if filter_text in t.lower()]
            others = [t for t in available if filter_text not in t.lower()]
            available = matching + others

        shown = available[:16]

        if shown:
            for t in shown:
                s_btn = QPushButton(f"+ {t}")
                s_btn.setStyleSheet(SUGGESTION_CHIP_STYLE)
                s_btn.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
                s_btn.setCursor(Qt.PointingHandCursor)
                s_btn.setToolTip(f"Click to assign tag '{t}'")
                s_btn.clicked.connect(lambda _, tag=t: self._add_suggested_tag(tag))
                self.suggestions_layout.addWidget(s_btn)

            self.suggestions_widget.setVisible(True)
            self.suggestions_chips_container.updateGeometry()
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
        self.resize(500, 290)
        self.setMinimumWidth(460)
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

        info_lbl = QLabel(
            f"Apply a tag to <b>{count} specifications</b>:<br>"
            f"<span style='color: #4A5568;'>{specs_preview}</span>"
        )
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
            sugg_lbl = QLabel(f"💡 <b>Existing Tags</b> <span style='color: {C.TEXT_FAINT};'>(click to select)</span>:")
            sugg_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px; margin-top: 2px;")
            card_layout.addWidget(sugg_lbl)

            sugg_container = QWidget()
            sugg_flow = FlowLayout(sugg_container, margin=0, h_spacing=6, v_spacing=6)
            for t in existing_tags[:10]:
                b = QPushButton(f"+ {t}")
                b.setStyleSheet(SUGGESTION_CHIP_STYLE)
                b.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
                b.setCursor(Qt.PointingHandCursor)
                b.setToolTip(f"Select '{t}'")
                b.clicked.connect(lambda _, tag=t: self.tag_input.setText(tag))
                sugg_flow.addWidget(b)
            card_layout.addWidget(sugg_container)

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
