# --- File: src/core/ui/right_side_panel.py ---
"""
Collapsible right-hand utility dock for Phase 3 cross-tool workflows.

Hosts two switchable slots side-by-side alongside any workspace tab:
- ``assistant``: Compact docked 3GPP Research Assistant (Phase 3.2)
- ``spec_search``: Compact Quick Spec Search side panel (Phase 3.3)
"""

from __future__ import annotations

from typing import Dict, Optional

from PyQt5.QtCore import QSize, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.icons import lucide_icon
from core.ui.theme import C, R, S, T

MODE_ASSISTANT = "assistant"
MODE_SPEC_SEARCH = "spec_search"
VALID_MODES = (MODE_ASSISTANT, MODE_SPEC_SEARCH)

DEFAULT_PANEL_WIDTH = 390
MIN_PANEL_WIDTH = 320
MAX_PANEL_WIDTH = 560


def _build_placeholder_slot(title: str, subtitle: str) -> QWidget:
    """Create a token-styled placeholder slot widget prior to slot injection."""
    holder = QWidget()
    holder.setObjectName("rightPanelPlaceholder")
    layout = QVBoxLayout(holder)
    layout.setContentsMargins(S.LG, S.LG, S.LG, S.LG)
    layout.setSpacing(S.SM)
    layout.addStretch(1)

    lbl_title = QLabel(title)
    lbl_title.setAlignment(Qt.AlignCenter)
    lbl_title.setStyleSheet(
        f"color: {C.TEXT_STRONG}; font-family: {T.FONT_FAMILY_UI}; "
        f"font-size: {T.SIZE_SECTION_TITLE}px; font-weight: 600;"
    )
    layout.addWidget(lbl_title)

    lbl_sub = QLabel(subtitle)
    lbl_sub.setWordWrap(True)
    lbl_sub.setAlignment(Qt.AlignCenter)
    lbl_sub.setStyleSheet(
        f"color: {C.TEXT_MUTED}; font-family: {T.FONT_FAMILY_UI}; "
        f"font-size: {T.SIZE_BODY}px;"
    )
    layout.addWidget(lbl_sub)
    layout.addStretch(1)
    return holder


class RightSidePanel(QFrame):
    """Collapsible, drag-resizable right-hand dock hosting Assistant & Quick Spec Search."""

    toggled = pyqtSignal(bool)
    mode_changed = pyqtSignal(str)
    promote_requested = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("rightSidePanel")
        self._is_open: bool = False
        self._mode: str = MODE_ASSISTANT
        self._preferred_width: int = DEFAULT_PANEL_WIDTH
        self._slots: Dict[str, QWidget] = {}

        self._setup_ui()
        self._apply_styles()
        self.set_mode(MODE_ASSISTANT, emit=False)
        self.set_open(False, emit=False)

    def _setup_ui(self) -> None:
        self.setMinimumWidth(0)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # --- Compact Segmented Header ---
        self.header = QFrame(self)
        self.header.setObjectName("rightPanelHeader")
        self.header.setFixedHeight(42)
        h_layout = QHBoxLayout(self.header)
        h_layout.setContentsMargins(S.SM, S.XS, S.SM, S.XS)
        h_layout.setSpacing(S.XS)

        self.seg_frame = QFrame(self.header)
        self.seg_frame.setObjectName("rightPanelSeg")
        seg_layout = QHBoxLayout(self.seg_frame)
        seg_layout.setContentsMargins(2, 2, 2, 2)
        seg_layout.setSpacing(2)

        self._mode_group = QButtonGroup(self)
        self._mode_group.setExclusive(True)

        self.mode_assistant_btn = QPushButton("Assistant", self.seg_frame)
        self.mode_assistant_btn.setObjectName("rightPanelSegBtn")
        self.mode_assistant_btn.setCheckable(True)
        self.mode_assistant_btn.setCursor(Qt.PointingHandCursor)
        self.mode_assistant_btn.setIcon(
            lucide_icon("sparkles", color=C.TEXT_MUTED, active_color=C.ACCENT_PRIMARY, size=14)
        )
        self.mode_assistant_btn.setIconSize(QSize(14, 14))
        self.mode_assistant_btn.setToolTip("Switch right panel to Docked Assistant (Ctrl+Shift+A)")
        self.mode_assistant_btn.clicked.connect(lambda: self.set_mode(MODE_ASSISTANT))

        self.mode_search_btn = QPushButton("Quick Search", self.seg_frame)
        self.mode_search_btn.setObjectName("rightPanelSegBtn")
        self.mode_search_btn.setCheckable(True)
        self.mode_search_btn.setCursor(Qt.PointingHandCursor)
        self.mode_search_btn.setIcon(
            lucide_icon("search", color=C.TEXT_MUTED, active_color=C.ACCENT_PRIMARY, size=14)
        )
        self.mode_search_btn.setIconSize(QSize(14, 14))
        self.mode_search_btn.setToolTip("Switch right panel to Quick Spec Search (Ctrl+Shift+F)")
        self.mode_search_btn.clicked.connect(lambda: self.set_mode(MODE_SPEC_SEARCH))

        self._mode_group.addButton(self.mode_assistant_btn)
        self._mode_group.addButton(self.mode_search_btn)
        seg_layout.addWidget(self.mode_assistant_btn)
        seg_layout.addWidget(self.mode_search_btn)
        h_layout.addWidget(self.seg_frame)

        h_layout.addStretch(1)

        self.promote_btn = QPushButton(self.header)
        self.promote_btn.setObjectName("rightPanelIconBtn")
        self.promote_btn.setCursor(Qt.PointingHandCursor)
        self.promote_btn.setIcon(
            lucide_icon("maximize-2", color=C.TEXT_MUTED, active_color=C.TEXT_PRIMARY, size=14)
        )
        self.promote_btn.setIconSize(QSize(14, 14))
        self.promote_btn.setToolTip("Open active tool in full-screen workspace tab")
        self.promote_btn.clicked.connect(lambda: self.promote_requested.emit(self._mode))
        h_layout.addWidget(self.promote_btn)

        self.close_btn = QPushButton(self.header)
        self.close_btn.setObjectName("rightPanelIconBtn")
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.setIcon(
            lucide_icon("x", color=C.TEXT_MUTED, active_color=C.TEXT_PRIMARY, size=14)
        )
        self.close_btn.setIconSize(QSize(14, 14))
        self.close_btn.setToolTip("Close side panel")
        self.close_btn.clicked.connect(self.close_panel)
        h_layout.addWidget(self.close_btn)

        root.addWidget(self.header)

        # --- Stacked Content Area ---
        self.stack = QStackedWidget(self)
        self.stack.setObjectName("rightPanelStack")

        default_assistant = _build_placeholder_slot(
            "3GPP Research Assistant",
            "Docked local assistant workspace ready for Phase 3.2 wiring.",
        )
        default_search = _build_placeholder_slot(
            "Quick Spec Search",
            "Compact FTS5 clause lookup ready for Phase 3.3 wiring.",
        )
        self._slots[MODE_ASSISTANT] = default_assistant
        self._slots[MODE_SPEC_SEARCH] = default_search
        self.stack.addWidget(default_assistant)
        self.stack.addWidget(default_search)
        root.addWidget(self.stack, 1)

    def _apply_styles(self) -> None:
        self.setStyleSheet(f"""
            QFrame#rightSidePanel {{
                background-color: {C.BG_CANVAS};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.LG}px;
            }}
            QFrame#rightPanelHeader {{
                background-color: {C.BG_CARD};
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                border-top-left-radius: {R.LG}px;
                border-top-right-radius: {R.LG}px;
            }}
            QFrame#rightPanelSeg {{
                background-color: {C.BG_CANVAS};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.MD}px;
            }}
            QPushButton#rightPanelSegBtn {{
                background-color: transparent;
                color: {C.TEXT_MUTED};
                border: none;
                border-radius: {R.SM}px;
                padding: 4px 10px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_TABLE}px;
                font-weight: 600;
            }}
            QPushButton#rightPanelSegBtn:hover {{
                color: {C.TEXT_PRIMARY};
            }}
            QPushButton#rightPanelSegBtn:checked {{
                background-color: {C.BG_CARD};
                color: {C.ACCENT_PRIMARY};
            }}
            QPushButton#rightPanelIconBtn {{
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: {R.MD}px;
                min-width: 26px;
                min-height: 26px;
                padding: 2px;
            }}
            QPushButton#rightPanelIconBtn:hover {{
                background-color: {C.BG_CANVAS};
                border-color: {C.BORDER_SUBTLE};
            }}
            QStackedWidget#rightPanelStack {{
                background-color: transparent;
            }}
        """)

    # --- Public State & Slot API ---

    @property
    def is_open(self) -> bool:
        return self._is_open

    @property
    def mode(self) -> str:
        return self._mode

    @property
    def preferred_width(self) -> int:
        return self._preferred_width

    @preferred_width.setter
    def preferred_width(self, value: int) -> None:
        self._preferred_width = self.clamp_width(value)

    @staticmethod
    def clamp_width(value: int) -> int:
        return max(MIN_PANEL_WIDTH, min(MAX_PANEL_WIDTH, int(value)))

    def slot_widget(self, mode: str) -> Optional[QWidget]:
        """Return the widget currently installed for ``mode``."""
        return self._slots.get(mode)

    def set_slot_widget(self, mode: str, widget: QWidget) -> None:
        """Replace the content widget for ``mode`` (``assistant`` or ``spec_search``)."""
        if mode not in VALID_MODES:
            raise ValueError(f"Unsupported right panel mode: {mode!r}")
        old = self._slots.get(mode)
        if old is widget and self.stack.indexOf(widget) >= 0:
            if self._mode == mode:
                self.stack.setCurrentWidget(widget)
            return
        idx = self.stack.indexOf(old) if old is not None else -1
        if idx >= 0 and old is not widget:
            self.stack.removeWidget(old)
            old.setParent(None)
            self.stack.insertWidget(idx, widget)
        elif self.stack.indexOf(widget) < 0:
            self.stack.addWidget(widget)
        self._slots[mode] = widget
        if self._mode == mode:
            self.stack.setCurrentWidget(widget)

    def set_mode(self, mode: str, emit: bool = True) -> None:
        """Switch active docked mode between ``assistant`` and ``spec_search``."""
        normalized = mode if mode in VALID_MODES else MODE_ASSISTANT
        changed = normalized != self._mode
        self._mode = normalized

        self.mode_assistant_btn.blockSignals(True)
        self.mode_search_btn.blockSignals(True)
        self.mode_assistant_btn.setChecked(normalized == MODE_ASSISTANT)
        self.mode_search_btn.setChecked(normalized == MODE_SPEC_SEARCH)
        self.mode_assistant_btn.blockSignals(False)
        self.mode_search_btn.blockSignals(False)

        target = self._slots.get(normalized)
        if target is not None:
            self.stack.setCurrentWidget(target)

        if normalized == MODE_ASSISTANT:
            self.promote_btn.setToolTip("Open Assistant in full-screen workspace tab (Ctrl+9)")
        else:
            self.promote_btn.setToolTip("Open Spec Search in full-screen workspace tab (Ctrl+5)")

        if emit and changed:
            self.mode_changed.emit(normalized)

    def set_open(self, opened: bool, emit: bool = True) -> None:
        """Show or collapse the right side panel."""
        opened = bool(opened)
        changed = opened != self._is_open
        self._is_open = opened
        self.setVisible(opened)
        if opened:
            self.setMinimumWidth(MIN_PANEL_WIDTH)
            self.setMaximumWidth(MAX_PANEL_WIDTH)
        else:
            self.setMinimumWidth(0)
            self.setMaximumWidth(16777215)
        if emit and changed:
            self.toggled.emit(opened)

    def open_mode(self, mode: str) -> None:
        """Ensure the panel is open showing ``mode``."""
        self.set_mode(mode, emit=True)
        if not self._is_open:
            self.set_open(True, emit=True)

    def close_panel(self) -> None:
        """Collapse the right side panel."""
        self.set_open(False, emit=True)

    def toggle_mode(self, mode: str) -> None:
        """Toggle if already open in ``mode``; otherwise open/switch to ``mode``."""
        normalized = mode if mode in VALID_MODES else MODE_ASSISTANT
        if self._is_open and self._mode == normalized:
            self.close_panel()
        else:
            self.open_mode(normalized)

    def apply_splitter_sizes(self, splitter: QSplitter) -> None:
        """Synchronize the horizontal workspace ``QSplitter`` sizes cleanly."""
        if splitter is None:
            return
        total = splitter.width()
        if not self._is_open:
            self.setVisible(False)
            if total > 0:
                splitter.setSizes([total, 0])
            else:
                splitter.setSizes([1000, 0])
            return

        self.setVisible(True)
        if total <= 0:
            splitter.setSizes([max(500, 1000 - self._preferred_width), self._preferred_width])
            return

        handle_w = max(0, splitter.handleWidth())
        usable = max(0, total - handle_w)
        max_allowed = max(MIN_PANEL_WIDTH, int(usable * 0.48))
        target_w = min(self._preferred_width, max_allowed, MAX_PANEL_WIDTH)
        target_w = max(MIN_PANEL_WIDTH, target_w) if usable >= MIN_PANEL_WIDTH + 320 else max(220, int(usable * 0.42))
        main_w = max(240, usable - target_w)
        splitter.setSizes([main_w, target_w])

    def capture_width_from_splitter(self, splitter: QSplitter) -> None:
        """Persist user drag-resized width when panel is open."""
        if not self._is_open or splitter is None:
            return
        sizes = splitter.sizes()
        if len(sizes) >= 2 and sizes[1] >= MIN_PANEL_WIDTH:
            self._preferred_width = self.clamp_width(sizes[1])
