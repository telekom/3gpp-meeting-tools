# --- File: src/modules/meetings/ui/topic_pills.py ---
"""
Pill & Tag UI Components for Cross-WG Topic Management.

Provides:
- WgTogglePill: Checkable pill button color-coded by Working Group archetype (SA blue, RAN purple, CT teal).
- TagCloudWidget: Interactive tag cloud editor with removable pill chips and inline entry for WI tokens and keywords.
- AgendaItemPillsWidget: Structured per-WG Agenda Item pattern manager with removable pills and inline adding.
"""

import re
from typing import Callable, Dict, List, Optional, Set

from PyQt5.QtCore import QEvent, QObject, QStringListModel, QTimer, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QKeyEvent
from PyQt5.QtWidgets import (
    QAbstractItemView, QAction, QCompleter, QFrame, QHBoxLayout, QInputDialog, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QMenu, QPushButton, QScrollArea,
    QSizePolicy, QVBoxLayout, QWidget, QWidgetAction
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from core.ui.theme import C


def get_wg_colors(wg: str) -> Dict[str, str]:
    """Returns archetype color schemes based on Working Group family."""
    wg_u = str(wg or "").strip().upper()
    if wg_u.startswith("SA"):
        return {
            "bg_active": C.INFO_SOFT_STRONG,
            "fg_active": C.INFO_DARK,
            "border_active": C.INFO_BORDER_MEDIUM,
            "bg_inactive": C.BG_SURFACE,
            "fg_inactive": C.TEXT_MUTED,
            "border_inactive": C.BORDER_STRONG,
            "hover": C.INFO_SOFT,
        }
    elif wg_u.startswith("RAN"):
        return {
            "bg_active": C.VIOLET_SOFT,
            "fg_active": C.VIOLET_DARK,
            "border_active": C.VIOLET_BORDER_STRONG,
            "bg_inactive": C.BG_SURFACE,
            "fg_inactive": C.TEXT_MUTED,
            "border_inactive": C.BORDER_STRONG,
            "hover": C.VIOLET_SURFACE,
        }
    else:  # CT or others
        return {
            "bg_active": C.TEAL_SOFT,
            "fg_active": C.TEAL_DARK,
            "border_active": C.TEAL_BORDER,
            "bg_inactive": C.BG_SURFACE,
            "fg_inactive": C.TEXT_MUTED,
            "border_inactive": C.BORDER_STRONG,
            "hover": C.TEAL_SURFACE,
        }


class WgTogglePill(QPushButton):
    """
    Checkable pill button for selecting Target Working Groups.
    Displays '✓ WG' when checked and '+ WG' when unchecked with family-specific colors.
    """
    def __init__(self, wg: str, checked: bool = False, parent=None):
        super().__init__(parent)
        self.wg = str(wg).strip().upper()
        self.colors = get_wg_colors(self.wg)
        self.setCheckable(True)
        self.setChecked(checked)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.setCursor(Qt.PointingHandCursor)
        self._update_appearance()
        self.toggled.connect(self._on_toggled)

    def _on_toggled(self, checked: bool):
        self._update_appearance()

    def setChecked(self, checked: bool):
        super().setChecked(checked)
        self._update_appearance()

    def _update_appearance(self):
        c = self.colors
        if self.isChecked():
            self.setText(f"✓ {self.wg}")
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {c['bg_active']};
                    color: {c['fg_active']};
                    border: 1px solid {c['border_active']};
                    border-radius: 12px;
                    padding: 4px 10px;
                    font-size: 11px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    border-color: {c['fg_active']};
                }}
            """)
        else:
            self.setText(f"+ {self.wg}")
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {c['bg_inactive']};
                    color: {c['fg_inactive']};
                    border: 1px solid {c['border_inactive']};
                    border-radius: 12px;
                    padding: 4px 10px;
                    font-size: 11px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {c['hover']};
                    color: {C.TEXT_STRONG};
                    border-color: {C.TEXT_FAINT};
                }}
            """)


class TagPillItem(QFrame):
    """Individual tag pill with a remove button."""
    removed = pyqtSignal(str)

    def __init__(self, text: str, parent=None, colors: Optional[Dict[str, str]] = None):
        super().__init__(parent)
        self.tag_text = text
        bg = (colors or {}).get("bg", C.BG_CANVAS)
        border = (colors or {}).get("border", C.BORDER_STRONG)
        fg = (colors or {}).get("fg", C.TEXT_BODY)
        self.setStyleSheet(f"""
            QFrame {{
                background-color: {bg};
                border: 1px solid {border};
                border-radius: 11px;
                padding: 1px 4px;
            }}
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 2, 4, 2)
        layout.setSpacing(4)

        lbl = QLabel(text)
        lbl.setStyleSheet(f"color: {fg}; font-size: 11px; font-weight: 500; background: transparent; border: none;")
        layout.addWidget(lbl)

        btn_close = QPushButton("✕")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setFixedSize(14, 14)
        btn_close.setStyleSheet(f"""
            QPushButton {{
                color: {C.TEXT_FAINT};
                background: transparent;
                border: none;
                font-size: 10px;
                font-weight: bold;
                padding: 0;
            }}
            QPushButton:hover {{
                color: {C.DANGER_ACCENT};
            }}
        """)
        btn_close.clicked.connect(lambda: self.removed.emit(self.tag_text))
        layout.addWidget(btn_close)


class _TagLineEdit(QLineEdit):
    """Custom QLineEdit that captures Backspace to remove last tag and handles comma separation."""
    backspace_pressed = pyqtSignal()

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key_Backspace and not self.text():
            self.backspace_pressed.emit()
            return
        super().keyPressEvent(event)


class TagCloudWidget(QFrame):
    """
    Modern Tag Cloud input displaying tokens as removable pills with an inline editor.
    Supports comma separation, Enter to add, paste with comma parsing, and backspace removal.
    """
    tagsChanged = pyqtSignal()

    def __init__(
        self,
        placeholder: str = "Type term and press Enter or comma...",
        parent=None,
        pill_colors: Optional[Dict[str, str]] = None,
        strict_dropdown: bool = False,
        dropdown_label: str = "+ Select ▾",
        empty_dropdown_label: str = "None available",
    ):
        super().__init__(parent)
        self.pill_colors = pill_colors
        self.strict_dropdown = strict_dropdown
        self._dropdown_label = dropdown_label
        self._empty_dropdown_label = empty_dropdown_label
        self._suggestions: List[str] = []
        self._completer: Optional[QCompleter] = None
        self._value_extractor: Optional[Callable[[str], str]] = None
        self.setObjectName("TagCloudFrame")
        self.setStyleSheet(f"""
            QFrame#TagCloudFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                padding: 3px;
            }}
            QFrame#TagCloudFrame:focus-within {{
                border-color: {C.INFO_ACCENT};
            }}
        """)
        self._tags: List[str] = []

        self.flow_layout = FlowLayout(self, margin=2, h_spacing=6, v_spacing=4)

        self.input_edit = _TagLineEdit()
        self.input_edit.setObjectName("TagCloudInput")
        self.input_edit.setPlaceholderText(placeholder)
        self.input_edit.setMinimumWidth(190)
        self.input_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.input_edit.setStyleSheet(
            "QLineEdit#TagCloudInput { border: none; background: transparent; font-size: 11px; padding: 2px 4px; }"
        )
        self.input_edit.returnPressed.connect(self._on_enter_pressed)
        self.input_edit.textChanged.connect(self._on_text_changed)
        self.input_edit.backspace_pressed.connect(self._on_backspace_pressed)

        self.pick_btn: Optional[QPushButton] = None
        self.pick_menu: Optional[QMenu] = None
        self.pick_search_edit: Optional[QLineEdit] = None
        self.pick_list: Optional[QListWidget] = None
        if self.strict_dropdown:
            self.input_edit.setReadOnly(True)
            self.input_edit.setVisible(False)
            self.pick_btn = QPushButton(self._empty_dropdown_label)
            self.pick_btn.setCursor(Qt.PointingHandCursor)
            self.pick_btn.setEnabled(False)
            self.pick_btn.setStyleSheet(
                f"QPushButton {{ background: {C.BG_SURFACE}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_STRONG}; "
                "border-radius: 10px; padding: 2px 8px; font-size: 11px; }"
                f"QPushButton:hover:enabled {{ background: {C.INFO_SOFT}; border-color: {C.INFO_ACCENT}; }}"
                f"QPushButton:disabled {{ color: {C.TEXT_MUTED}; border-style: dashed; }}"
            )
            self.pick_menu = QMenu(self.pick_btn)
            self.pick_menu.setStyleSheet(
                f"QMenu {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; "
                "border-radius: 6px; padding: 0px; }}"
            )
            self.pick_menu.aboutToShow.connect(self._populate_pick_menu)
            self.pick_btn.setMenu(self.pick_menu)
            self.flow_layout.addWidget(self.pick_btn)
        else:
            self.flow_layout.addWidget(self.input_edit)

    def set_empty_dropdown_label(self, label: str):
        """Updates the disabled label shown when strict_dropdown has no selectable suggestions."""
        self._empty_dropdown_label = label
        self._update_pick_button_state()

    def _update_pick_button_state(self):
        if not self.pick_btn:
            return
        has_items = bool(self._suggestions)
        self.pick_btn.setEnabled(has_items)
        self.pick_btn.setText(self._dropdown_label if has_items else self._empty_dropdown_label)

    def _populate_pick_menu(self):
        if not self.pick_menu:
            return
        menu = self.pick_menu
        menu.clear()
        self.pick_search_edit = None
        self.pick_list = None
        if not self._suggestions:
            empty_act = menu.addAction(self._empty_dropdown_label)
            empty_act.setEnabled(False)
            return

        # Compute readable single-column popup width clamped cleanly to viewport ergonomics
        max_len = max((len(str(lbl)) for lbl in self._suggestions), default=20)
        popup_w = max(scale_px(300), min(scale_px(540), scale_px(120 + max_len * 7)))

        container = QWidget(menu)
        container.setFixedWidth(popup_w)
        c_layout = QVBoxLayout(container)
        c_layout.setContentsMargins(6, 6, 6, 6)
        c_layout.setSpacing(5)

        search_edit: Optional[QLineEdit] = None
        if len(self._suggestions) > 5:
            search_edit = QLineEdit(container)
            search_edit.setClearButtonEnabled(True)
            search_edit.setPlaceholderText("🔍 Search by # or keyword (Space-separated terms)...")
            search_edit.setStyleSheet(
                f"QLineEdit {{ background: {C.BG_CARD}; color: {C.TEXT_BODY}; "
                f"border: 1px solid {C.BORDER_STRONG}; border-radius: 5px; padding: 4px 8px; font-size: 12px; }} "
                f"QLineEdit:focus {{ border-color: {C.INFO_ACCENT}; }}"
            )
            c_layout.addWidget(search_edit)
            self.pick_search_edit = search_edit

        list_widget = QListWidget(container)
        list_widget.setSelectionMode(QAbstractItemView.NoSelection)
        list_widget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        list_widget.setWordWrap(True)
        list_widget.setUniformItemSizes(False)
        list_widget.setFocusPolicy(Qt.StrongFocus)
        list_widget.setStyleSheet(
            f"QListWidget {{ background: {C.BG_CARD}; color: {C.TEXT_BODY}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 5px; outline: none; font-size: 12px; }} "
            f"QListWidget::item {{ padding: 4px 6px; border-radius: 3px; }} "
            f"QListWidget::item:hover {{ background-color: {C.BG_HOVER}; }}"
        )
        self.pick_list = list_widget

        active_lower = {t.lower() for t in self._tags}
        normal_font = QFont(list_widget.font())
        bold_font = QFont(normal_font)
        bold_font.setBold(True)
        active_bg = QColor(C.INFO_SOFT)
        active_fg = QColor(C.INFO_DARK)
        default_bg = QColor(C.BG_CARD)
        default_fg = QColor(C.TEXT_BODY)

        def _style_item(it: QListWidgetItem, checked: bool):
            it.setData(Qt.CheckStateRole, Qt.Checked if checked else Qt.Unchecked)
            it.setFont(bold_font if checked else normal_font)
            it.setBackground(active_bg if checked else default_bg)
            it.setForeground(active_fg if checked else default_fg)

        for label in self._suggestions:
            val = self._value_extractor(label) if self._value_extractor else label
            it = QListWidgetItem(label)
            it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            it.setData(Qt.UserRole, val)
            it.setToolTip(label)
            _style_item(it, val.lower() in active_lower)
            list_widget.addItem(it)

        est_rows_h = sum(max(scale_px(26), list_widget.sizeHintForRow(i) + 4) for i in range(min(14, list_widget.count())))
        list_h = max(scale_px(44), min(scale_px(380), est_rows_h + scale_px(6)))
        list_widget.setFixedHeight(list_h)
        c_layout.addWidget(list_widget)

        footer = QHBoxLayout()
        footer.setContentsMargins(2, 0, 2, 0)
        footer.setSpacing(6)
        status_lbl = QLabel(container)
        status_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        footer.addWidget(status_lbl, 1)

        btn_done = QPushButton("Done ✓", container)
        btn_done.setCursor(Qt.PointingHandCursor)
        btn_done.setStyleSheet(
            f"QPushButton {{ background: {C.INFO_SOFT}; color: {C.INFO_DARK}; "
            f"border: 1px solid {C.INFO_BORDER_MEDIUM}; border-radius: 4px; "
            "padding: 2px 10px; font-size: 11px; font-weight: bold; } "
            f"QPushButton:hover {{ background: {C.INFO_SOFT_STRONG}; }}"
        )
        btn_done.clicked.connect(menu.close)
        footer.addWidget(btn_done)
        c_layout.addLayout(footer)

        def _update_status():
            total = list_widget.count()
            shown = sum(1 for i in range(total) if not list_widget.item(i).isHidden())
            sel_cnt = sum(1 for i in range(total) if list_widget.item(i).data(Qt.CheckStateRole) == Qt.Checked)
            if shown < total:
                status_lbl.setText(f"Showing {shown} of {total} · {sel_cnt} selected")
            else:
                status_lbl.setText(f"{total} item(s) · {sel_cnt} selected")

        def _toggle_item(it: Optional[QListWidgetItem]):
            if it is None:
                return
            val = str(it.data(Qt.UserRole) or it.text())
            now_checked = it.data(Qt.CheckStateRole) != Qt.Checked
            _style_item(it, now_checked)
            if now_checked:
                self.add_tag(val)
            else:
                self.remove_tag(val)
            _update_status()

        list_widget.itemClicked.connect(_toggle_item)

        if search_edit is not None:
            def _filter_list(txt: str):
                tokens = [tok for tok in txt.strip().lower().split() if tok]
                for i in range(list_widget.count()):
                    it = list_widget.item(i)
                    lbl_l = it.text().lower()
                    it.setHidden(bool(tokens) and not all(tok in lbl_l for tok in tokens))
                _update_status()

            def _on_search_return():
                visible_items = [list_widget.item(i) for i in range(list_widget.count()) if not list_widget.item(i).isHidden()]
                if len(visible_items) == 1:
                    _toggle_item(visible_items[0])

            search_edit.textChanged.connect(_filter_list)
            search_edit.returnPressed.connect(_on_search_return)

        _update_status()
        wa = QWidgetAction(menu)
        wa.setDefaultWidget(container)
        menu.addAction(wa)
        if search_edit is not None:
            QTimer.singleShot(0, search_edit.setFocus)

    def set_suggestions(self, suggestions: List[str], value_extractor: Optional[Callable[[str], str]] = None):
        """Attaches an autocomplete QCompleter or updates strict dropdown options."""
        self._value_extractor = value_extractor
        self._suggestions = list(suggestions or [])
        if self.strict_dropdown:
            self._update_pick_button_state()
            return
        model = QStringListModel(self._suggestions, self)
        if self._completer is None:
            self._completer = QCompleter(model, self)
            self._completer.setCaseSensitivity(Qt.CaseInsensitive)
            self._completer.setFilterMode(Qt.MatchContains)
            self._completer.setCompletionMode(QCompleter.PopupCompletion)
            self._completer.activated[str].connect(self._on_completer_activated)
            self.input_edit.setCompleter(self._completer)
        else:
            self._completer.setModel(model)

    def _clear_input_safely(self):
        try:
            if self.input_edit is not None:
                self.input_edit.clear()
        except RuntimeError:
            pass

    def _on_completer_activated(self, selected: str):
        val = self._value_extractor(selected) if self._value_extractor else selected
        if val:
            self.add_tag(val)
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(0, self._clear_input_safely)

    def _on_enter_pressed(self):
        text = self.input_edit.text()
        if self._value_extractor and text.strip():
            text = self._value_extractor(text.strip())
        self._ingest_text(text)
        self.input_edit.clear()

    def _on_text_changed(self, text: str):
        if "," in text or ";" in text:
            self._ingest_text(text)
            self.input_edit.clear()

    def _on_backspace_pressed(self):
        if self._tags:
            self.remove_tag(self._tags[-1])

    def _ingest_text(self, text: str):
        parts = re.split(r'[,;]+', text)
        changed = False
        for p in parts:
            p_clean = p.strip()
            if p_clean and p_clean.lower() not in [t.lower() for t in self._tags]:
                self._tags.append(p_clean)
                changed = True
        if changed:
            self._rebuild_pills()
            self.tagsChanged.emit()

    def add_tag(self, tag: str):
        tag_clean = tag.strip()
        if tag_clean and tag_clean.lower() not in [t.lower() for t in self._tags]:
            self._tags.append(tag_clean)
            self._rebuild_pills()
            self.tagsChanged.emit()

    def remove_tag(self, tag: str):
        target = tag.strip().lower()
        self._tags = [t for t in self._tags if t.lower() != target]
        self._rebuild_pills()
        self.tagsChanged.emit()

    def _rebuild_pills(self):
        # Detach persistent input/dropdown controls first so FlowLayout.clear() never deletes them
        self.flow_layout.removeWidget(self.input_edit)
        if self.pick_btn is not None:
            self.flow_layout.removeWidget(self.pick_btn)
        self.flow_layout.clear()

        # Add tag pill widgets
        for t in self._tags:
            pill = TagPillItem(t, self, colors=self.pill_colors)
            pill.removed.connect(self.remove_tag)
            self.flow_layout.addWidget(pill)

        if self.pick_btn is not None:
            self.flow_layout.addWidget(self.pick_btn)
            self._update_pick_button_state()
        else:
            self.flow_layout.addWidget(self.input_edit)
        self.updateGeometry()

    def tags(self) -> List[str]:
        return list(self._tags)

    def text(self) -> str:
        """Returns comma-separated string for backward compatibility."""
        return ", ".join(self._tags)

    def setText(self, text: str):
        """Sets tags from comma-separated or semicolon-separated string."""
        self._tags = []
        if text:
            parts = re.split(r'[,;]+', str(text))
            for p in parts:
                p_clean = p.strip()
                if p_clean and p_clean.lower() not in [t.lower() for t in self._tags]:
                    self._tags.append(p_clean)
        self._rebuild_pills()
        self.tagsChanged.emit()


class AgendaItemPill(QFrame):
    """Pill chip representing a specific Agenda Item pattern within a Working Group."""
    removed = pyqtSignal(str, str)  # wg, pattern

    def __init__(self, wg: str, pattern: str, parent=None):
        super().__init__(parent)
        self.wg = wg.strip().upper()
        self.pattern = pattern.strip()
        colors = get_wg_colors(self.wg)

        self.setStyleSheet(f"""
            QFrame {{
                background-color: {colors['bg_active']};
                border: 1px solid {colors['border_active']};
                border-radius: 10px;
                padding: 1px 4px;
            }}
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 2, 4, 2)
        layout.setSpacing(4)

        lbl = QLabel(f"<b>{self.pattern}</b>")
        lbl.setStyleSheet(f"color: {colors['fg_active']}; font-size: 11px; background: transparent; border: none;")
        layout.addWidget(lbl)

        btn_close = QPushButton("✕")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setFixedSize(14, 14)
        btn_close.setStyleSheet(f"""
            QPushButton {{
                color: {colors['fg_active']};
                background: transparent;
                border: none;
                font-size: 10px;
                font-weight: bold;
                padding: 0;
            }}
            QPushButton:hover {{
                color: {C.DANGER_ACCENT};
            }}
        """)
        btn_close.clicked.connect(lambda: self.removed.emit(self.wg, self.pattern))
        layout.addWidget(btn_close)


class AgendaItemPillsWidget(QFrame):
    """
    Per-WG grouped Agenda Item pattern manager.
    Displays a row per Target Working Group with removable pattern pills and '+ Add AI' button.
    """
    rulesChanged = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("AiPillsFrame")
        self.setStyleSheet(f"""
            QFrame#AiPillsFrame {{
                background-color: {C.BG_ALT_LIGHT};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 6px;
                padding: 6px;
            }}
        """)
        self._rules: Dict[str, List[str]] = {}
        self._active_wgs: List[str] = ["SA2", "RAN2", "CT1"]

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(4, 4, 4, 4)
        self.main_layout.setSpacing(6)

        self.lbl_empty = QLabel("No Agenda Items defined. Use '🔍 Auto-Suggest AIs...' or click '+ Add AI' below.")
        self.lbl_empty.setWordWrap(True)
        self.lbl_empty.setStyleSheet(
            f"color: {C.TEXT_FAINT}; font-size: 11px; font-style: italic; padding: 4px; background: transparent; border: none;"
        )
        self.main_layout.addWidget(self.lbl_empty)

        self._rows_container = QWidget()
        self._rows_layout = QVBoxLayout(self._rows_container)
        self._rows_layout.setContentsMargins(0, 0, 0, 0)
        self._rows_layout.setSpacing(6)
        self.main_layout.addWidget(self._rows_container)

    def set_active_wgs(self, wgs: List[str]):
        """Updates which Working Groups are visible/managed based on Target WGs selection."""
        self._active_wgs = [w.strip().upper() for w in wgs if w.strip()]
        self._rebuild_rows()

    def add_pattern(self, wg: str, pattern: str):
        wg = wg.strip().upper()
        pat = pattern.strip()
        if not wg or not pat:
            return
        if wg not in self._rules:
            self._rules[wg] = []
        if pat not in self._rules[wg]:
            self._rules[wg].append(pat)
            self._rebuild_rows()
            self.rulesChanged.emit()

    def remove_pattern(self, wg: str, pattern: str):
        wg = wg.strip().upper()
        pat = pattern.strip()
        if wg in self._rules and pat in self._rules[wg]:
            self._rules[wg].remove(pat)
            if not self._rules[wg]:
                del self._rules[wg]
            self._rebuild_rows()
            self.rulesChanged.emit()

    def _on_add_clicked(self, wg: str):
        text, ok = QInputDialog.getText(
            self,
            f"Add Agenda Item for {wg}",
            f"Enter Agenda Item pattern for {wg} (e.g. 5.14*, 9.3*):",
        )
        if ok and text.strip():
            for p in text.split(","):
                p_c = p.strip()
                if p_c:
                    self.add_pattern(wg, p_c)

    def _rebuild_rows(self):
        # Clear existing rows immediately so back-to-back rebuilds never overlap visually
        while self._rows_layout.count():
            item = self._rows_layout.takeAt(0)
            if item and item.widget():
                w = item.widget()
                w.hide()
                w.setParent(None)
                w.deleteLater()

        has_any_pills = False
        # Display rows for all active WGs that either have patterns or are in active_wgs
        display_wgs = sorted(list(set(self._active_wgs) | set(self._rules.keys())))

        for wg in display_wgs:
            pats = self._rules.get(wg, [])
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(2, 2, 2, 2)
            row_layout.setSpacing(6)

            # WG Badge
            colors = get_wg_colors(wg)
            lbl_wg = QLabel(f"<b>{wg}:</b>")
            lbl_wg.setFixedWidth(50)
            lbl_wg.setStyleSheet(
                f"color: {colors['fg_active']}; font-size: 11px; padding-top: 3px; background: transparent; border: none;"
            )
            row_layout.addWidget(lbl_wg, 0, Qt.AlignTop)

            # Pills in a FlowLayout
            flow_container = QWidget()
            flow_layout = FlowLayout(flow_container, margin=0, h_spacing=4, v_spacing=4)

            for pat in pats:
                has_any_pills = True
                pill = AgendaItemPill(wg, pat, self)
                pill.removed.connect(self.remove_pattern)
                flow_layout.addWidget(pill)

            btn_add = QPushButton("+ Add AI")
            btn_add.setCursor(Qt.PointingHandCursor)
            btn_add.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.BG_CARD};
                    color: {C.TEXT_SECONDARY};
                    border: 1px dashed {C.BORDER_STRONG};
                    border-radius: 10px;
                    padding: 2px 8px;
                    font-size: 10px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {C.BG_CANVAS};
                    color: {C.TEXT_PRIMARY};
                    border-color: {C.TEXT_FAINT};
                }}
            """)
            btn_add.clicked.connect(lambda _, w=wg: self._on_add_clicked(w))
            flow_layout.addWidget(btn_add)

            row_layout.addWidget(flow_container, 1)
            self._rows_layout.addWidget(row_widget)

        self.lbl_empty.setVisible(not has_any_pills)
        self.updateGeometry()

    def text(self) -> str:
        """Returns standard format: 'SA2: 5.14*, 5.12*; RAN2: 8.7*; CT1: 6.18*' for backward compatibility."""
        out = []
        for wg in sorted(self._rules.keys()):
            pats = self._rules[wg]
            if pats:
                out.append(f"{wg}: {', '.join(pats)}")
        return "; ".join(out)

    def setText(self, rule_str: str):
        """Parses standard format into internal map and updates pills."""
        self._rules = {}
        if rule_str:
            for chunk in str(rule_str).split(";"):
                chunk = chunk.strip()
                if ":" in chunk:
                    wg_k, pats = chunk.split(":", 1)
                    wg_k = wg_k.strip().upper()
                    p_list = [p.strip() for p in pats.split(",") if p.strip()]
                    if p_list:
                        self._rules[wg_k] = p_list
        self._rebuild_rows()
        self.rulesChanged.emit()

    def patterns_map(self) -> Dict[str, List[str]]:
        return {k: list(v) for k, v in self._rules.items()}


class FilterPillButton(QPushButton):
    """
    Clickable, checkable filter pill with title, count badge, and custom color scheme.
    Used for 1-click filtering by Working Group and Decision Category.
    """
    def __init__(self, key: str, label: str, count: Optional[int] = None, colors: Optional[Dict[str, str]] = None, parent=None):
        super().__init__(parent)
        self.key = key
        self.base_label = label
        self._count = count
        self.colors = colors or {
            "bg_active": C.ACCENT_ALT,
            "fg_active": C.BG_CARD,
            "border_active": C.INFO,
            "bg_inactive": C.BG_CARD,
            "fg_inactive": C.TEXT_SECONDARY,
            "border_inactive": C.BORDER_STRONG,
            "hover": C.BG_SURFACE,
        }
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self._update_text_and_style()
        self.toggled.connect(self._on_toggled)

    def set_count(self, count: Optional[int]):
        self._count = count
        self._update_text_and_style()

    def _on_toggled(self, checked: bool):
        self._update_text_and_style()

    def _update_text_and_style(self):
        c = self.colors
        cnt_str = f" ({self._count})" if self._count is not None and self._count >= 0 else ""
        self.setText(f"{self.base_label}{cnt_str}")
        if self.isChecked():
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {c['bg_active']};
                    color: {c['fg_active']};
                    border: 1px solid {c['border_active']};
                    border-radius: 11px;
                    padding: 3px 10px;
                    font-size: 11px;
                    font-weight: bold;
                }}
            """)
        else:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {c['bg_inactive']};
                    color: {c['fg_inactive']};
                    border: 1px solid {c['border_inactive']};
                    border-radius: 11px;
                    padding: 3px 10px;
                    font-size: 11px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {c['hover']};
                    border-color: {C.TEXT_FAINT};
                    color: {C.TEXT_PRIMARY};
                }}
            """)

