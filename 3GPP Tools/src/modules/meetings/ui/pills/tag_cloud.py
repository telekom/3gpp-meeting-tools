"""Interactive tag cloud pill editor with autocomplete and strict dropdown picker."""

from __future__ import annotations

import html
import re
from typing import Callable, Dict, List, Optional

from PyQt5.QtCore import QSize, QStringListModel, QTimer, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QKeyEvent
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCompleter,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
    QWidgetAction,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from core.ui.theme import C


class TagPillItem(QFrame):
    """Individual tag pill with a remove button."""

    removed = pyqtSignal(str)

    def __init__(
        self,
        text: str,
        parent=None,
        colors: Optional[Dict[str, str]] = None,
        compact: bool = False,
    ):
        super().__init__(parent)
        self.tag_text = text
        bg = (colors or {}).get("bg", C.BG_CANVAS)
        border = (colors or {}).get("border", C.BORDER_STRONG)
        fg = (colors or {}).get("fg", C.TEXT_BODY)
        radius = 9 if compact else 11
        pad = "0px 3px" if compact else "1px 4px"
        self.setStyleSheet(
            f"QFrame {{ background-color: {bg}; border: 1px solid {border}; "
            f"border-radius: {radius}px; padding: {pad}; }}"
        )
        layout = QHBoxLayout(self)
        if compact:
            layout.setContentsMargins(3, 1, 3, 1)
            layout.setSpacing(2)
        else:
            layout.setContentsMargins(4, 2, 4, 2)
            layout.setSpacing(4)

        lbl = QLabel(text)
        lbl.setStyleSheet(f"color: {fg}; font-size: 11px; font-weight: 500; background: transparent; border: none;")
        layout.addWidget(lbl)

        btn_close = QPushButton("✕")
        btn_close.setCursor(Qt.PointingHandCursor)
        close_px = scale_px(12) if compact else 14
        btn_close.setFixedSize(close_px, close_px)
        btn_close.setStyleSheet(
            f"QPushButton {{ color: {C.TEXT_FAINT}; background: transparent; border: none; "
            f"font-size: 10px; font-weight: bold; padding: 0; }}"
            f"QPushButton:hover {{ color: {C.DANGER_ACCENT}; }}"
        )
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
        compact: bool = False,
        prefix_label: Optional[str] = None,
        max_visible_pills: Optional[int] = None,
    ):
        super().__init__(parent)
        self.pill_colors = pill_colors
        self.strict_dropdown = strict_dropdown
        self._dropdown_label = dropdown_label
        self._empty_dropdown_label = empty_dropdown_label
        self._compact = bool(compact)
        self._prefix_text = str(prefix_label or "").strip()
        self._max_visible_pills = int(max_visible_pills) if max_visible_pills and max_visible_pills > 0 else None
        self._overflow_expanded = False
        self._overflow_btn: Optional[QPushButton] = None
        self.pick_clear_btn: Optional[QPushButton] = None
        self._suggestions: List[str] = []
        self._completer: Optional[QCompleter] = None
        self._value_extractor: Optional[Callable[[str], str]] = None
        self.setObjectName("TagCloudFrame")
        if self._compact:
            sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
            sp.setHeightForWidth(True)
            self.setSizePolicy(sp)
            self.setStyleSheet(
                "QFrame#TagCloudFrame { background-color: transparent; border: none; padding: 0px; }"
            )
            self.flow_layout = FlowLayout(self, margin=0, h_spacing=3, v_spacing=2)
        else:
            self.setStyleSheet(
                f"""
                QFrame#TagCloudFrame {{
                    background-color: {C.BG_CARD};
                    border: 1px solid {C.BORDER_STRONG};
                    border-radius: 6px;
                    padding: 3px;
                }}
                QFrame#TagCloudFrame:focus-within {{
                    border-color: {C.INFO_ACCENT};
                }}
            """
            )
            self.flow_layout = FlowLayout(self, margin=2, h_spacing=6, v_spacing=4)
        self._tags: List[str] = []

        self._prefix_lbl: Optional[QLabel] = None
        if self._prefix_text:
            self._prefix_lbl = QLabel(f"<b>{html.escape(self._prefix_text)}</b>", self)
            self._prefix_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            self._prefix_lbl.setMinimumHeight(scale_px(20))
            self._prefix_lbl.setStyleSheet(
                f"QLabel {{ color: {C.TEXT_SECONDARY}; font-size: 11px; "
                "background: transparent; border: none; padding-right: 1px; }"
            )
            self.flow_layout.addWidget(self._prefix_lbl)

        self.input_edit = _TagLineEdit()
        self.input_edit.setObjectName("TagCloudInput")
        self.input_edit.setPlaceholderText(placeholder)
        if self._compact:
            self.input_edit.setMinimumWidth(scale_px(95))
            self.input_edit.setMaximumWidth(scale_px(150))
            self.input_edit.setStyleSheet(
                f"QLineEdit#TagCloudInput {{ background: {C.BG_CARD}; color: {C.TEXT_BODY}; "
                f"border: 1px dashed {C.BORDER_STRONG}; border-radius: 9px; font-size: 11px; padding: 1px 6px; }} "
                f"QLineEdit#TagCloudInput:focus {{ border-style: solid; border-color: {C.INFO_ACCENT}; }}"
            )
        else:
            self.input_edit.setMinimumWidth(190)
            self.input_edit.setStyleSheet(
                "QLineEdit#TagCloudInput { border: none; background: transparent; font-size: 11px; padding: 2px 4px; }"
            )
        self.input_edit.setSizePolicy(QSizePolicy.Expanding if not self._compact else QSizePolicy.Preferred, QSizePolicy.Fixed)
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
            btn_radius = 9 if self._compact else 10
            btn_pad = "1px 7px" if self._compact else "2px 8px"
            self.pick_btn.setStyleSheet(
                f"QPushButton {{ background: {C.BG_SURFACE}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_STRONG}; "
                f"border-radius: {btn_radius}px; padding: {btn_pad}; font-size: 11px; }}"
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

        btn_clear = QPushButton("Clear", container)
        btn_clear.setCursor(Qt.PointingHandCursor)
        btn_clear.setToolTip("Clear all tags in this category")
        btn_clear.setStyleSheet(
            f"QPushButton {{ background: {C.BG_SURFACE}; color: {C.DANGER}; "
            f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; "
            "padding: 2px 8px; font-size: 11px; } "
            f"QPushButton:hover:enabled {{ background: {C.DANGER_SOFT}; border-color: {C.DANGER_BORDER}; }} "
            f"QPushButton:disabled {{ color: {C.TEXT_MUTED}; }}"
        )
        self.pick_clear_btn = btn_clear
        footer.addWidget(btn_clear)

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
            btn_clear.setEnabled(bool(self._tags))
            if shown < total:
                status_lbl.setText(f"Showing {shown} of {total} · {sel_cnt} selected")
            else:
                status_lbl.setText(f"{total} item(s) · {sel_cnt} selected")

        def _on_clear_clicked():
            self.clear_tags()
            for i in range(list_widget.count()):
                _style_item(list_widget.item(i), False)
            _update_status()

        btn_clear.clicked.connect(_on_clear_clicked)

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
        parts = re.split(r"[,;]+", text)
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

    def set_overflow_expanded(self, expanded: bool):
        """Expands or collapses overflow pills when max_visible_pills is configured."""
        self._overflow_expanded = bool(expanded)
        self._rebuild_pills()

    def clear_tags(self):
        """Removes all currently selected tags and emits tagsChanged if non-empty."""
        if not self._tags:
            return
        self._tags = []
        self._overflow_expanded = False
        self._rebuild_pills()
        self.tagsChanged.emit()

    def _rebuild_pills(self):
        # Detach persistent prefix/input/dropdown controls first so FlowLayout.clear() never deletes them
        if self._prefix_lbl is not None:
            self.flow_layout.removeWidget(self._prefix_lbl)
        self.flow_layout.removeWidget(self.input_edit)
        if self.pick_btn is not None:
            self.flow_layout.removeWidget(self.pick_btn)
        self._overflow_btn = None
        self.flow_layout.clear()

        if self._prefix_lbl is not None:
            self.flow_layout.addWidget(self._prefix_lbl)

        limit = self._max_visible_pills
        if limit and not self._overflow_expanded and len(self._tags) > limit:
            visible_tags = self._tags[:limit]
            hidden_tags = self._tags[limit:]
        else:
            visible_tags = self._tags
            hidden_tags = []

        for t in visible_tags:
            pill = TagPillItem(t, self, colors=self.pill_colors, compact=self._compact)
            pill.removed.connect(self.remove_tag)
            self.flow_layout.addWidget(pill)

        if limit and len(self._tags) > limit:
            ov_txt = f"+{len(hidden_tags)} more" if hidden_tags else "Show less"
            ov_btn = QPushButton(ov_txt, self)
            ov_btn.setCursor(Qt.PointingHandCursor)
            ov_btn.setToolTip(
                f"Show all {len(self._tags)} tags: {', '.join(hidden_tags)}"
                if hidden_tags
                else "Collapse extra tags"
            )
            ov_btn.setStyleSheet(
                f"QPushButton {{ background: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; "
                f"border: 1px solid {C.BORDER_STRONG}; border-radius: 9px; "
                "padding: 1px 6px; font-size: 11px; font-weight: 500; }"
                f"QPushButton:hover {{ background: {C.INFO_SOFT}; border-color: {C.INFO_ACCENT}; color: {C.INFO_DARK}; }}"
            )
            ov_btn.clicked.connect(lambda _=False, exp=bool(hidden_tags): self.set_overflow_expanded(exp))
            self._overflow_btn = ov_btn
            self.flow_layout.addWidget(ov_btn)

        if self.pick_btn is not None:
            self.flow_layout.addWidget(self.pick_btn)
            self._update_pick_button_state()
        else:
            self.flow_layout.addWidget(self.input_edit)
        self.flow_layout.invalidate()
        self.updateGeometry()
        parent = self.parentWidget()
        if parent is not None and parent.layout() is not None:
            parent.layout().invalidate()
            parent.updateGeometry()

    def hasHeightForWidth(self) -> bool:
        return True if self._compact else super().hasHeightForWidth()

    def heightForWidth(self, width: int) -> int:
        if self._compact:
            return self.flow_layout.heightForWidth(width)
        return super().heightForWidth(width)

    def sizeHint(self) -> QSize:
        if not self._compact:
            return super().sizeHint()
        margins = self.contentsMargins()
        l_m, t_m, r_m, b_m = self.flow_layout.getContentsMargins()
        h_space = max(0, self.flow_layout.horizontalSpacing())
        total_w = margins.left() + margins.right() + l_m + r_m
        max_h = 0
        vis_cnt = 0
        for i in range(self.flow_layout.count()):
            it = self.flow_layout.itemAt(i)
            wid = it.widget() if it else None
            if wid is not None and wid.isHidden():
                continue
            sz = it.sizeHint()
            total_w += sz.width()
            max_h = max(max_h, sz.height())
            vis_cnt += 1
        if vis_cnt > 1:
            total_w += h_space * (vis_cnt - 1)
        total_h = max(scale_px(22), max_h + margins.top() + margins.bottom() + t_m + b_m)
        return QSize(total_w, total_h)

    def tags(self) -> List[str]:
        return list(self._tags)

    def text(self) -> str:
        """Returns comma-separated string for backward compatibility."""
        return ", ".join(self._tags)

    def setText(self, text: str):
        """Sets tags from comma-separated or semicolon-separated string."""
        self._tags = []
        self._overflow_expanded = False
        if text:
            parts = re.split(r"[,;]+", str(text))
            for p in parts:
                p_clean = p.strip()
                if p_clean and p_clean.lower() not in [t.lower() for t in self._tags]:
                    self._tags.append(p_clean)
        self._rebuild_pills()
        self.tagsChanged.emit()
