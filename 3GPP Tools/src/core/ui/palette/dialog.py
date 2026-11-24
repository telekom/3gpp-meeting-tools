"""Modal Ctrl+K Command Palette dialog window and keyboard event filter."""

from __future__ import annotations

import logging
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from PyQt5.QtCore import QEvent, Qt, pyqtSignal
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.icons import lucide_icon
from core.ui.palette.delegate import _PaletteRowDelegate
from core.ui.palette.models import (
    CommandItem,
    _KIND_HEADER,
    _KIND_ITEM,
    _ROLE_CMD,
    _ROLE_KIND,
    _score_command,
)
from core.ui.theme import C, R, T

logger = logging.getLogger(__name__)


class CommandPaletteDialog(QDialog):
    """Modal Ctrl+K command palette with instant filter + arrow-key execution."""

    item_executed = pyqtSignal(object)  # CommandItem

    def __init__(
        self,
        static_items: Sequence[CommandItem],
        dynamic_searcher: Optional[Callable[[str], Sequence[CommandItem]]] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("commandPaletteDialog")
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setModal(True)

        self._static_items = list(static_items)
        self._dynamic_searcher = dynamic_searcher

        self._setup_ui()
        self._apply_styles()
        self.populate_results("")
        self._center_on_parent()

    def _setup_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 20)

        self.card = QFrame(self)
        self.card.setObjectName("cmdPaletteCard")
        self.card.setFixedWidth(580)
        shadow = QGraphicsDropShadowEffect(self.card)
        shadow.setBlurRadius(28)
        shadow.setOffset(0, 8)
        shadow.setColor(QColor(0, 0, 0, 55))
        self.card.setGraphicsEffect(shadow)

        card_layout = QVBoxLayout(self.card)
        card_layout.setContentsMargins(0, 0, 0, 0)
        card_layout.setSpacing(0)

        # --- Top Search Bar Row (50px) ---
        search_bar = QFrame(self.card)
        search_bar.setObjectName("cmdPaletteSearchRow")
        search_bar.setFixedHeight(50)
        sb_layout = QHBoxLayout(search_bar)
        sb_layout.setContentsMargins(14, 0, 14, 0)
        sb_layout.setSpacing(10)

        icon_lbl = QLabel(search_bar)
        icon_lbl.setPixmap(lucide_icon("search", color=C.ACCENT_PRIMARY, size=17).pixmap(17, 17))
        sb_layout.addWidget(icon_lbl)

        self.search_input = QLineEdit(search_bar)
        self.search_input.setObjectName("cmdPaletteInput")
        self.search_input.setPlaceholderText(
            "Jump to tool, spec (23.501), meeting, TDoc (S2-…), WI, or action…"
        )
        self.search_input.setClearButtonEnabled(True)
        self.search_input.installEventFilter(self)
        self.search_input.textChanged.connect(self.populate_results)
        sb_layout.addWidget(self.search_input, 1)

        esc_badge = QLabel("Esc", search_bar)
        esc_badge.setObjectName("cmdPaletteEscBadge")
        sb_layout.addWidget(esc_badge)
        card_layout.addWidget(search_bar)

        # --- Results List ---
        self.results_list = QListWidget(self.card)
        self.results_list.setObjectName("cmdPaletteList")
        self.results_list.setItemDelegate(_PaletteRowDelegate(self.results_list))
        self.results_list.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        self.results_list.setMaximumHeight(380)
        self.results_list.setMinimumHeight(180)
        self.results_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.results_list.itemActivated.connect(self._on_item_activated)
        self.results_list.itemClicked.connect(self._on_item_activated)
        card_layout.addWidget(self.results_list, 1)

        # --- Footer Hint Strip ---
        footer = QFrame(self.card)
        footer.setObjectName("cmdPaletteFooter")
        footer.setFixedHeight(32)
        f_layout = QHBoxLayout(footer)
        f_layout.setContentsMargins(14, 0, 14, 0)
        f_layout.setSpacing(12)

        hint_lbl = QLabel("↑↓ Navigate   ·   ↵ Open   ·   Esc Dismiss", footer)
        hint_lbl.setObjectName("cmdPaletteHintLabel")
        f_layout.addWidget(hint_lbl)
        f_layout.addStretch(1)

        self.count_label = QLabel("", footer)
        self.count_label.setObjectName("cmdPaletteHintLabel")
        f_layout.addWidget(self.count_label)
        card_layout.addWidget(footer)

        outer.addWidget(self.card)

    def _apply_styles(self) -> None:
        self.setStyleSheet(f"""
            QFrame#cmdPaletteCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.XL}px;
            }}
            QFrame#cmdPaletteSearchRow {{
                background-color: {C.BG_CARD};
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                border-top-left-radius: {R.XL}px;
                border-top-right-radius: {R.XL}px;
            }}
            QLineEdit#cmdPaletteInput {{
                background-color: transparent;
                border: none;
                color: {C.TEXT_PRIMARY};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_SECTION_TITLE}px;
                padding: 4px 0px;
            }}
            QLabel#cmdPaletteEscBadge {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_MUTED};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.SM}px;
                padding: 2px 6px;
                font-family: {T.FONT_FAMILY_MONO};
                font-size: 11px;
                font-weight: 600;
            }}
            QListWidget#cmdPaletteList {{
                background-color: {C.BG_CARD};
                border: none;
                outline: none;
                padding: 4px 0px;
            }}
            QFrame#cmdPaletteFooter {{
                background-color: {C.BG_SURFACE};
                border-top: 1px solid {C.BORDER_SUBTLE};
                border-bottom-left-radius: {R.XL}px;
                border-bottom-right-radius: {R.XL}px;
            }}
            QLabel#cmdPaletteHintLabel {{
                color: {C.TEXT_MUTED};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
            }}
        """)

    def _center_on_parent(self) -> None:
        parent = self.parentWidget()
        self.adjustSize()
        if parent is not None and parent.isVisible():
            pg = parent.frameGeometry()
            x = pg.center().x() - self.width() // 2
            y = pg.top() + max(48, int(pg.height() * 0.14))
            self.move(x, y)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._center_on_parent()
        self.search_input.setFocus()
        self.search_input.selectAll()

    def populate_results(self, query: str) -> None:
        q = (query or "").strip()
        self.results_list.clear()

        matched_static: List[Tuple[Tuple[int, str], CommandItem]] = []
        for item in self._static_items:
            rank = _score_command(item, q)
            if rank is not None:
                matched_static.append((rank, item))

        if q:
            matched_static.sort(key=lambda pair: pair[0])

        grouped: Dict[str, List[CommandItem]] = {}
        for _rank, item in matched_static:
            grouped.setdefault(item.category, []).append(item)

        if self._dynamic_searcher is not None and len(q) >= 2:
            try:
                dyn = self._dynamic_searcher(q) or []
            except Exception as exc:
                logger.debug("Command palette dynamic provider error: %s", exc)
                dyn = []
            for item in dyn:
                grouped.setdefault(item.category, []).append(item)

        # Render groups in clean predictable order
        preferred_order = [
            "Tools",
            "Side Panels & Actions",
            "Specifications",
            "Meetings & TDocs",
            "Work Items",
            "Protocols",
            "Word Workbenches",
        ]
        ordered_cats = [c for c in preferred_order if c in grouped] + [
            c for c in grouped if c not in preferred_order
        ]

        total_items = 0
        first_selectable_row = -1
        for cat in ordered_cats:
            cat_items = grouped[cat]
            if not q and cat == "Word Workbenches":
                cat_items = cat_items[:4]
            if not cat_items:
                continue

            hdr = QListWidgetItem(cat)
            hdr.setFlags(Qt.NoItemFlags)
            hdr.setData(_ROLE_KIND, _KIND_HEADER)
            self.results_list.addItem(hdr)

            for item in cat_items:
                row_item = QListWidgetItem(item.title)
                row_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                row_item.setData(_ROLE_KIND, _KIND_ITEM)
                row_item.setData(_ROLE_CMD, item)
                self.results_list.addItem(row_item)
                if first_selectable_row < 0:
                    first_selectable_row = self.results_list.count() - 1
                total_items += 1

        if first_selectable_row >= 0:
            self.results_list.setCurrentRow(first_selectable_row)
            self.count_label.setText(f"{total_items} command{'s' if total_items != 1 else ''}")
        else:
            empty_row = QListWidgetItem("No matching commands, specs, meetings, or work items")
            empty_row.setFlags(Qt.NoItemFlags)
            empty_row.setData(_ROLE_KIND, _KIND_HEADER)
            self.results_list.addItem(empty_row)
            self.count_label.setText("0 matches")

    def _move_selection(self, delta: int) -> None:
        count = self.results_list.count()
        if count == 0:
            return
        cur = self.results_list.currentRow()
        idx = cur
        for _ in range(count):
            idx = (idx + delta) % count
            it = self.results_list.item(idx)
            if it is not None and bool(it.flags() & Qt.ItemIsSelectable):
                self.results_list.setCurrentRow(idx)
                return

    def _execute_current(self) -> bool:
        it = self.results_list.currentItem()
        if it is None:
            return False
        return self._on_item_activated(it)

    def _on_item_activated(self, item: QListWidgetItem) -> bool:
        if item is None or item.data(_ROLE_KIND) != _KIND_ITEM:
            return False
        cmd: Optional[CommandItem] = item.data(_ROLE_CMD)
        if cmd is None:
            return False
        self.accept()
        self.item_executed.emit(cmd)
        if callable(cmd.callback):
            cmd.callback()
        return True

    def eventFilter(self, watched, event) -> bool:
        if watched is self.search_input and event.type() == QEvent.KeyPress:
            key = event.key()
            if key == Qt.Key_Down:
                self._move_selection(1)
                return True
            if key == Qt.Key_Up:
                self._move_selection(-1)
                return True
            if key in (Qt.Key_Return, Qt.Key_Enter):
                self._execute_current()
                return True
            if key == Qt.Key_Escape:
                self.reject()
                return True
        return super().eventFilter(watched, event)
