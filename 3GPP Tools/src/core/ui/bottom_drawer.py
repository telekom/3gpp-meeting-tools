# --- File: src/core/ui/bottom_drawer.py ---
"""
Collapsible bottom Console & Queue drawer for the 3GPP Delegate Tools shell (Phase 1.3).

Wraps `ConsolePanel` and `QueuePanel` in an expandable bottom pane toggled from
the status bar or keyboard shortcut, tracking unread error count and queued task count.
"""

from __future__ import annotations

import logging
from typing import List, Optional

from PyQt5.QtCore import QSize, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from core.ui.icons import lucide_icon
from core.ui.theme import C, R, S, T


class BottomDrawer(QFrame):
    """
    Slide-up / collapsible bottom drawer hosting ConsolePanel and QueuePanel.

    Emits:
      - `toggled(bool)`: True when opened, False when collapsed.
      - `badges_changed(int, int)`: `(active_queue_count, unread_error_count)`.
    """

    toggled = pyqtSignal(bool)
    badges_changed = pyqtSignal(int, int)

    def __init__(
        self,
        console_panel: QWidget,
        queue_panel: QWidget,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("bottomDrawerFrame")
        self.console_panel = console_panel
        self.queue_panel = queue_panel

        self._is_open: bool = False
        self._queue_count: int = 0
        self._unread_errors: int = 0
        self._saved_height: int = 220

        self._build_ui()
        self._apply_styles()
        self.set_open(False, emit=False)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(S.SM, S.XS, S.SM, S.SM)
        root.setSpacing(S.XS)

        # Drawer header strip
        header = QHBoxLayout()
        header.setContentsMargins(S.XS, 2, S.XS, 2)
        header.setSpacing(S.SM)

        icon_lbl = QLabel()
        icon_lbl.setPixmap(lucide_icon("terminal", color=C.TEXT_SECONDARY, size=14).pixmap(14, 14))
        header.addWidget(icon_lbl)

        title_lbl = QLabel("Console & Queue")
        title_lbl.setObjectName("drawerHeaderTitle")
        header.addWidget(title_lbl)

        self.queue_badge = QLabel("0 queued")
        self.queue_badge.setObjectName("drawerQueueBadge")
        header.addWidget(self.queue_badge)

        self.error_badge = QLabel("0 errors")
        self.error_badge.setObjectName("drawerErrorBadge")
        self.error_badge.setVisible(False)
        header.addWidget(self.error_badge)

        header.addStretch(1)

        self.close_btn = QPushButton("Hide")
        self.close_btn.setObjectName("drawerHideBtn")
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.setIcon(lucide_icon("chevron-down", color=C.TEXT_MUTED, active_color=C.TEXT_PRIMARY, size=14))
        self.close_btn.setIconSize(QSize(14, 14))
        self.close_btn.setToolTip("Collapse Console & Queue drawer (Ctrl+J)")
        self.close_btn.clicked.connect(lambda: self.set_open(False))
        header.addWidget(self.close_btn)

        root.addLayout(header)

        # Horizontal splitter inside drawer
        self.bottom_splitter = QSplitter(Qt.Horizontal, self)
        self.bottom_splitter.setObjectName("drawerInnerSplitter")
        self.bottom_splitter.addWidget(self.console_panel)
        self.bottom_splitter.addWidget(self.queue_panel)
        self.bottom_splitter.setSizes([680, 280])

        root.addWidget(self.bottom_splitter, 1)

    def _apply_styles(self) -> None:
        self.setStyleSheet(f"""
            QFrame#bottomDrawerFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.LG}px;
            }}
            QLabel#drawerHeaderTitle {{
                color: {C.TEXT_STRONG};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_TABLE}px;
                font-weight: 600;
            }}
            QLabel#drawerQueueBadge {{
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_SECONDARY};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.PILL}px;
                padding: 1px 8px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QLabel#drawerErrorBadge {{
                background-color: {C.DANGER_SOFT};
                color: {C.DANGER};
                border: 1px solid {C.DANGER_BORDER};
                border-radius: {R.PILL}px;
                padding: 1px 8px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QPushButton#drawerHideBtn {{
                background-color: transparent;
                color: {C.TEXT_MUTED};
                border: 1px solid transparent;
                border-radius: {R.SM}px;
                padding: 2px 8px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 500;
            }}
            QPushButton#drawerHideBtn:hover {{
                background-color: {C.BG_CANVAS};
                border-color: {C.BORDER_STRONG};
                color: {C.TEXT_PRIMARY};
            }}
        """)

    # --- State & Badge API ---

    @property
    def is_open(self) -> bool:
        return self._is_open

    @property
    def queue_count(self) -> int:
        return self._queue_count

    @property
    def unread_errors(self) -> int:
        return self._unread_errors

    @property
    def preferred_height(self) -> int:
        return self._saved_height

    @preferred_height.setter
    def preferred_height(self, height: int) -> None:
        if height >= 120:
            self._saved_height = height

    def set_open(self, opened: bool, emit: bool = True) -> None:
        """Expand or collapse the drawer cleanly."""
        changed = opened != self._is_open
        if not opened and self._is_open and self.height() >= 120:
            self._saved_height = self.height()

        self._is_open = bool(opened)
        self.setVisible(self._is_open)

        if self._is_open and self._unread_errors > 0:
            self._unread_errors = 0
            self._refresh_badges(emit=True)

        if emit and changed:
            self.toggled.emit(self._is_open)

    def toggle(self) -> bool:
        """Toggle drawer visibility and return new open state."""
        self.set_open(not self._is_open)
        return self._is_open

    def set_queue_count(self, count: int) -> None:
        """Update displayed queue item badge count."""
        clean = max(0, int(count))
        if clean != self._queue_count:
            self._queue_count = clean
            self._refresh_badges(emit=True)

    def record_log_event(self, message: str, level: int = logging.INFO) -> None:
        """Track unread errors when log events arrive while the drawer is closed."""
        is_error = level >= logging.ERROR or "❌" in (message or "")
        if is_error and not self._is_open:
            self._unread_errors += 1
            self._refresh_badges(emit=True)

    def _refresh_badges(self, emit: bool = True) -> None:
        noun_q = "item" if self._queue_count == 1 else "items"
        self.queue_badge.setText(f"{self._queue_count} {noun_q}" if self._queue_count else "Idle")

        noun_e = "error" if self._unread_errors == 1 else "errors"
        self.error_badge.setText(f"{self._unread_errors} {noun_e}")
        self.error_badge.setVisible(self._unread_errors > 0)

        if emit:
            self.badges_changed.emit(self._queue_count, self._unread_errors)

    def apply_splitter_sizes(self, outer_splitter: QSplitter) -> None:
        """Synchronize `outer_splitter` sizes cleanly when opening/collapsing."""
        if not self._is_open:
            total = max(1, outer_splitter.height() or sum(outer_splitter.sizes()) or 700)
            outer_splitter.setSizes([total, 0])
        else:
            total = max(320, outer_splitter.height() or sum(outer_splitter.sizes()) or 760)
            drawer_h = min(max(160, self._saved_height), max(180, total // 2))
            outer_splitter.setSizes([max(160, total - drawer_h), drawer_h])

    def inner_sizes(self) -> List[int]:
        return self.bottom_splitter.sizes()

    def set_inner_sizes(self, sizes: List[int]) -> None:
        if sizes and len(sizes) == 2 and sum(sizes) > 0:
            self.bottom_splitter.setSizes(sizes)
