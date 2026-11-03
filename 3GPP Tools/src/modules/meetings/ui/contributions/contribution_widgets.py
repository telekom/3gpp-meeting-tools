# --- File: src/modules/meetings/ui/contribution_widgets.py ---
import re
import logging
from PyQt5.QtCore import Qt, pyqtSignal, QStringListModel, QTimer
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QRadioButton, QButtonGroup, QListWidget, QListWidgetItem, QFrame,
    QCompleter
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEARCH_INPUT_STYLE
)
from core.utils.company_sanitizer import CompanySanitizer
from modules.work_items.core.wi_database import WorkItemsDatabase

logger = logging.getLogger(__name__)

class KPICard(QFrame):
    """Lightweight summary metric card."""

    def __init__(self, title: str, default_val: str = "-", subtitle: str = "", parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
            QLabel#kpiTitle {
                color: #64748B;
                font-size: 10px;
                font-weight: bold;
                border: none;
            }
            QLabel#kpiValue {
                color: #1E5C99;
                font-size: 16px;
                font-weight: bold;
                border: none;
            }
            QLabel#kpiSub {
                color: #94A3B8;
                font-size: 10px;
                border: none;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 6, 10, 6)
        layout.setSpacing(1)

        self.lbl_title = QLabel(title.upper())
        self.lbl_title.setObjectName("kpiTitle")
        self.lbl_value = QLabel(default_val)
        self.lbl_value.setObjectName("kpiValue")
        self.lbl_sub = QLabel(subtitle)
        self.lbl_sub.setObjectName("kpiSub")

        layout.addWidget(self.lbl_title)
        layout.addWidget(self.lbl_value)
        layout.addWidget(self.lbl_sub)

    def set_value(self, val: str, sub: str = ""):
        self.lbl_value.setText(val)
        if sub:
            self.lbl_sub.setText(sub)


class FilterTokenChip(QFrame):
    removed = pyqtSignal(str)

    def __init__(self, token_text: str, parent=None):
        super().__init__(parent)
        self.token_text = token_text
        self.setStyleSheet("""
            QFrame {
                background-color: #EBF3FC;
                border: 1px solid #BFDBFE;
                border-radius: 12px;
                padding: 1px 6px;
            }
            QLabel {
                color: #1E5C99;
                font-size: 11px;
                font-weight: bold;
                border: none;
            }
            QPushButton {
                background: transparent;
                border: none;
                color: #64748B;
                font-weight: bold;
                font-size: 11px;
                padding: 0px 2px;
            }
            QPushButton:hover {
                color: #DC2626;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 2, 4, 2)
        layout.setSpacing(4)

        lbl = QLabel(token_text)
        btn_close = QPushButton("✕")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.clicked.connect(lambda: self.removed.emit(self.token_text))

        layout.addWidget(lbl)
        layout.addWidget(btn_close)


class WITokenInputWidget(QWidget):
    tokens_changed = pyqtSignal()

    def __init__(self, wi_db: WorkItemsDatabase, parent=None):
        super().__init__(parent)
        self.wi_db = wi_db
        self.selected_tokens = set()

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(4)

        self.chips_widget = QWidget()
        self.chips_layout = QHBoxLayout(self.chips_widget)
        self.chips_layout.setContentsMargins(0, 0, 0, 0)
        self.chips_layout.setSpacing(4)
        self.chips_layout.setAlignment(Qt.AlignLeft)
        main_layout.addWidget(self.chips_widget)

        input_layout = QHBoxLayout()
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("Search WI acronym or code (e.g., FS_6G_ARC)...")
        self.input_edit.setStyleSheet(SEARCH_INPUT_STYLE)
        self.input_edit.returnPressed.connect(self._add_current_input)

        self._setup_completer()

        btn_add = QPushButton("Add")
        btn_add.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_add.setFixedWidth(50)
        btn_add.clicked.connect(self._add_current_input)

        input_layout.addWidget(self.input_edit)
        input_layout.addWidget(btn_add)
        main_layout.addLayout(input_layout)

    def _setup_completer(self):
        # Query the WI DB as the user types instead of taking an arbitrary
        # first-N snapshot. This keeps older WIs (e.g. eNA) discoverable.
        self._suggestion_model = QStringListModel(self)
        self._completer = QCompleter(self._suggestion_model, self)
        self._completer.setCaseSensitivity(Qt.CaseInsensitive)
        self._completer.setFilterMode(Qt.MatchContains)
        self._completer.setCompletionMode(QCompleter.PopupCompletion)
        self.input_edit.setCompleter(self._completer)

        self._suggest_timer = QTimer(self)
        self._suggest_timer.setSingleShot(True)
        self._suggest_timer.setInterval(180)
        self._suggest_timer.timeout.connect(self._refresh_suggestions)
        self.input_edit.textEdited.connect(lambda _: self._suggest_timer.start())

    def _refresh_suggestions(self):
        text = self.input_edit.text().strip()
        if not text:
            self._suggestion_model.setStringList([])
            return
        try:
            rows = self.wi_db.search_work_item_suggestions(text, limit=40)
            self._suggestion_model.setStringList([r["acronym"] for r in rows])
            self._completer.complete()
        except Exception as e:
            logging.warning(f"Could not refresh WI suggestions: {e}")

    def _add_current_input(self):
        text = self.input_edit.text().strip()
        if text and text not in self.selected_tokens:
            self.selected_tokens.add(text)
            chip = FilterTokenChip(text, self)
            chip.removed.connect(self._remove_token)
            self.chips_layout.addWidget(chip)
            self.input_edit.clear()
            self.tokens_changed.emit()

    def _remove_token(self, token: str):
        if token in self.selected_tokens:
            self.selected_tokens.remove(token)
            for i in range(self.chips_layout.count()):
                widget = self.chips_layout.itemAt(i).widget()
                if isinstance(widget, FilterTokenChip) and widget.token_text == token:
                    widget.deleteLater()
                    break
            self.tokens_changed.emit()

    def get_tokens(self) -> list:
        return list(self.selected_tokens)


class AITokenInputWidget(QWidget):
    """Token input for exact Agenda Item filtering across meetings."""
    tokens_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_tokens = set()

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(4)

        self.chips_widget = QWidget()
        self.chips_layout = QHBoxLayout(self.chips_widget)
        self.chips_layout.setContentsMargins(0, 0, 0, 0)
        self.chips_layout.setSpacing(4)
        self.chips_layout.setAlignment(Qt.AlignLeft)
        main_layout.addWidget(self.chips_widget)

        input_layout = QHBoxLayout()
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("Add Agenda Item (e.g., 6.2.1)...")
        self.input_edit.setStyleSheet(SEARCH_INPUT_STYLE)
        self.input_edit.setToolTip(
            "Filter by exact Agenda Item number. Add multiple AIs to match any of them."
        )
        self.input_edit.returnPressed.connect(self._add_current_input)

        btn_add = QPushButton("Add")
        btn_add.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_add.setFixedWidth(50)
        btn_add.clicked.connect(self._add_current_input)

        input_layout.addWidget(self.input_edit)
        input_layout.addWidget(btn_add)
        main_layout.addLayout(input_layout)

    @staticmethod
    def _normalize_ai(value: str) -> str:
        text = str(value or "").strip()
        text = re.sub(r"^AI\s*[:#-]?\s*", "", text, flags=re.IGNORECASE)
        return re.sub(r"\s+", "", text)

    def _add_current_input(self):
        text = self._normalize_ai(self.input_edit.text())
        if text and text not in self.selected_tokens:
            self.selected_tokens.add(text)
            chip = FilterTokenChip(text, self)
            chip.removed.connect(self._remove_token)
            self.chips_layout.addWidget(chip)
            self.input_edit.clear()
            self.tokens_changed.emit()

    def _remove_token(self, token: str):
        if token in self.selected_tokens:
            self.selected_tokens.remove(token)
            for i in range(self.chips_layout.count()):
                widget = self.chips_layout.itemAt(i).widget()
                if isinstance(widget, FilterTokenChip) and widget.token_text == token:
                    widget.deleteLater()
                    break
            self.tokens_changed.emit()

    def get_tokens(self) -> list:
        return sorted(self.selected_tokens)


class CompanySelectorWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter companies list...")
        self.search_input.textChanged.connect(self._filter_list)
        layout.addWidget(self.search_input)

        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet("QListWidget::item { padding: 3px; }")
        layout.addWidget(self.list_widget)

        actions_layout = QHBoxLayout()
        btn_select_all = QPushButton("Select All")
        btn_select_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_select_all.clicked.connect(lambda: self._set_all_checked(True))

        btn_clear_all = QPushButton("Clear All")
        btn_clear_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_clear_all.clicked.connect(lambda: self._set_all_checked(False))

        actions_layout.addWidget(btn_select_all)
        actions_layout.addWidget(btn_clear_all)
        layout.addLayout(actions_layout)

        match_frame = QFrame()
        match_frame.setStyleSheet("QFrame { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 4px; padding: 4px; }")
        match_layout = QVBoxLayout(match_frame)
        match_layout.setContentsMargins(4, 4, 4, 4)
        match_layout.setSpacing(2)

        match_layout.addWidget(QLabel("<b>Author Match Rule:</b>"))
        self.rb_any = QRadioButton("Any co-author (joint contribution)")
        self.rb_primary = QRadioButton("Primary source only (first listed)")
        self.rb_any.setChecked(True)

        self.btn_group = QButtonGroup(self)
        self.btn_group.addButton(self.rb_any)
        self.btn_group.addButton(self.rb_primary)

        match_layout.addWidget(self.rb_any)
        match_layout.addWidget(self.rb_primary)
        layout.addWidget(match_frame)

        self._populate_companies()

    def _populate_companies(self):
        companies = sorted(list(CompanySanitizer.SIGNATURE_SYNONYMS_REGEX.keys()), key=lambda s: s.lower())
        for comp in companies:
            item = QListWidgetItem(comp, self.list_widget)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)

    def _filter_list(self, text: str):
        query = text.strip().lower()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setHidden(query not in item.text().lower())

    def _set_all_checked(self, checked: bool):
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if not item.isHidden():
                item.setCheckState(state)

    def get_selected_companies(self) -> set:
        return {
            self.list_widget.item(i).text()
            for i in range(self.list_widget.count())
            if self.list_widget.item(i).checkState() == Qt.Checked
        }

    def is_primary_only(self) -> bool:
        return self.rb_primary.isChecked()
