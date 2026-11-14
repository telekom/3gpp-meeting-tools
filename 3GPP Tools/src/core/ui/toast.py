# --- File: src/core/ui/toast.py ---
"""
Non-blocking floating toast notification system for 3GPP Delegate Tools (Phase 1.5).

Displays transient, action-capable notifications (e.g., [Open], [Show in Folder])
stacked cleanly in the bottom-right corner of a host window, styled exclusively
from `core.ui.theme` tokens.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Sequence, Tuple, Union

from PyQt5.QtCore import QEvent, QObject, QPoint, QSize, Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.icons import lucide_icon
from core.ui.theme import C, R, S, T
from core.ui.ui_components import escape_qt_mnemonic

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ToastAction:
    """Interactive action button displayed inside a toast notification."""

    label: str
    callback: Callable[[], None]


_VARIANT_STYLES: Dict[str, Tuple[str, str, str, str]] = {
    "success": (C.BG_CARD, C.SUCCESS, C.SUCCESS_BORDER, "check-circle"),
    "info": (C.BG_CARD, C.INFO, C.INFO_BORDER, "info"),
    "warning": (C.BG_CARD, C.WARNING, C.WARNING_BORDER, "alert-triangle"),
    "danger": (C.BG_CARD, C.DANGER, C.DANGER_BORDER, "alert-circle"),
}


class ToastWidget(QFrame):
    """Floating notification card with semantic accent stripe and optional actions."""

    dismissed = pyqtSignal(object)

    def __init__(
        self,
        title: str,
        message: str = "",
        variant: str = "info",
        actions: Optional[Sequence[Union[ToastAction, Tuple[str, Callable[[], None]]]]] = None,
        duration_ms: int = 6000,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("appToastCard")
        self._variant = variant if variant in _VARIANT_STYLES else "info"
        self._duration_ms = max(0, int(duration_ms))
        self._closed = False
        self._action_buttons: List[QPushButton] = []

        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.dismiss)

        self._build_ui(title, message, actions or ())
        self._apply_styles()

        if self._duration_ms > 0:
            self._timer.start(self._duration_ms)

    def _build_ui(
        self,
        title: str,
        message: str,
        actions: Sequence[Union[ToastAction, Tuple[str, Callable[[], None]]]],
    ) -> None:
        self.setMinimumWidth(280)
        self.setMaximumWidth(400)

        bg, accent, border, icon_name = _VARIANT_STYLES[self._variant]
        _ = bg, border  # used in _apply_styles

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(18)
        shadow.setOffset(0, 4)
        shadow.setColor(QColor(0, 0, 0, 38))
        self.setGraphicsEffect(shadow)

        outer = QHBoxLayout(self)
        outer.setContentsMargins(S.MD, S.SM + 2, S.SM, S.SM + 2)
        outer.setSpacing(S.SM)

        icon_lbl = QLabel()
        icon_lbl.setPixmap(lucide_icon(icon_name, color=accent, size=18).pixmap(18, 18))
        icon_lbl.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        icon_lbl.setContentsMargins(0, 2, 0, 0)
        outer.addWidget(icon_lbl)

        body = QVBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(3)

        header_row = QHBoxLayout()
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.setSpacing(S.SM)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("toastTitleLabel")
        self.title_label.setWordWrap(True)
        header_row.addWidget(self.title_label, 1)

        self.close_btn = QPushButton()
        self.close_btn.setObjectName("toastCloseBtn")
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.setIcon(lucide_icon("x", color=C.TEXT_MUTED, active_color=C.TEXT_PRIMARY, size=14))
        self.close_btn.setIconSize(QSize(14, 14))
        self.close_btn.setToolTip("Dismiss notification")
        self.close_btn.clicked.connect(self.dismiss)
        header_row.addWidget(self.close_btn, 0, Qt.AlignTop)

        body.addLayout(header_row)

        self.message_label = QLabel(message)
        self.message_label.setObjectName("toastMessageLabel")
        self.message_label.setWordWrap(True)
        self.message_label.setVisible(bool(message))
        body.addWidget(self.message_label)

        if actions:
            action_row = QHBoxLayout()
            action_row.setContentsMargins(0, S.XS, 0, 0)
            action_row.setSpacing(S.XS)
            for item in actions:
                if isinstance(item, ToastAction):
                    act_label, act_cb = item.label, item.callback
                else:
                    act_label, act_cb = item[0], item[1]
                btn = QPushButton(escape_qt_mnemonic(act_label))
                btn.setObjectName("toastActionBtn")
                btn.setCursor(Qt.PointingHandCursor)
                btn.clicked.connect(lambda _checked=False, cb=act_cb: self._invoke_action(cb))
                action_row.addWidget(btn)
                self._action_buttons.append(btn)
            action_row.addStretch(1)
            body.addLayout(action_row)

        outer.addLayout(body, 1)

    def _apply_styles(self) -> None:
        bg, accent, border, _ = _VARIANT_STYLES[self._variant]
        self.setStyleSheet(f"""
            QFrame#appToastCard {{
                background-color: {bg};
                border: 1px solid {border};
                border-left: 4px solid {accent};
                border-radius: {R.LG}px;
            }}
            QLabel#toastTitleLabel {{
                color: {C.TEXT_PRIMARY};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_BODY}px;
                font-weight: 600;
            }}
            QLabel#toastMessageLabel {{
                color: {C.TEXT_SECONDARY};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_TABLE}px;
            }}
            QPushButton#toastCloseBtn {{
                background-color: transparent;
                border: none;
                border-radius: {R.SM}px;
                padding: 2px;
            }}
            QPushButton#toastCloseBtn:hover {{
                background-color: {C.BG_CANVAS};
            }}
            QPushButton#toastActionBtn {{
                background-color: {C.BG_SURFACE};
                color: {C.ACCENT_PRIMARY};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.SM}px;
                padding: 3px 10px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QPushButton#toastActionBtn:hover {{
                background-color: {C.ACCENT_SOFT};
                border-color: {C.ACCENT_PRIMARY};
            }}
        """)

    def _invoke_action(self, callback: Callable[[], None]) -> None:
        try:
            callback()
        finally:
            self.dismiss()

    def enterEvent(self, event) -> None:
        if self._timer.isActive():
            self._timer.stop()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:
        if not self._closed and self._duration_ms > 0:
            self._timer.start(self._duration_ms)
        super().leaveEvent(event)

    def dismiss(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._timer.stop()
        self.hide()
        self.dismissed.emit(self)
        self.deleteLater()


class ToastManager(QObject):
    """Manages stacked non-blocking toast overlays anchored inside a host window."""

    def __init__(
        self,
        host: QWidget,
        max_visible: int = 3,
        margin: int = 16,
        spacing: int = 8,
    ) -> None:
        super().__init__(host)
        self._host = host
        self._max_visible = max(1, int(max_visible))
        self._margin = margin
        self._spacing = spacing
        self._toasts: List[ToastWidget] = []
        self._host.installEventFilter(self)

    @property
    def active_toasts(self) -> List[ToastWidget]:
        return [t for t in self._toasts if not t._closed]

    def show_toast(
        self,
        title: str,
        message: str = "",
        variant: str = "info",
        actions: Optional[Sequence[Union[ToastAction, Tuple[str, Callable[[], None]]]]] = None,
        duration_ms: int = 6000,
    ) -> ToastWidget:
        toast = ToastWidget(
            title=title,
            message=message,
            variant=variant,
            actions=actions,
            duration_ms=duration_ms,
            parent=self._host,
        )
        toast.dismissed.connect(self._on_toast_dismissed)
        toast.adjustSize()

        self._toasts.append(toast)
        while len(self._toasts) > self._max_visible:
            oldest = self._toasts.pop(0)
            oldest.dismiss()

        toast.show()
        toast.raise_()
        self.reposition()
        return toast

    def dismiss_all(self) -> None:
        for toast in list(self._toasts):
            toast.dismiss()
        self._toasts.clear()

    def _on_toast_dismissed(self, toast: ToastWidget) -> None:
        if toast in self._toasts:
            self._toasts.remove(toast)
        self.reposition()

    def reposition(self) -> None:
        if not self._host:
            return

        host_rect = self._host.rect()
        bottom_offset = self._margin
        if isinstance(self._host, QMainWindow):
            sb = self._host.statusBar()
            if sb is not None and sb.isVisible():
                bottom_offset += sb.height()

        y_cursor = host_rect.height() - bottom_offset
        right_edge = host_rect.width() - self._margin

        # Avoid covering an open RightSidePanel dock (e.g., Assistant Send button)
        right_panel = getattr(self._host, "right_panel", None)
        if right_panel is not None and getattr(right_panel, "is_open", False) and right_panel.width() > 0:
            shifted_edge = right_edge - right_panel.width() - S.SM
            if self._host.isAncestorOf(right_panel):
                panel_left = right_panel.mapTo(self._host, QPoint(0, 0)).x()
                shifted_edge = min(shifted_edge, panel_left - S.SM)
            if shifted_edge >= self._margin + 280:
                right_edge = shifted_edge

        for toast in reversed(self._toasts):
            if toast._closed:
                continue
            toast.adjustSize()
            tw = min(max(toast.sizeHint().width(), toast.minimumWidth()), max(280, host_rect.width() - 2 * self._margin))
            toast.resize(tw, toast.sizeHint().height())
            th = toast.height()
            x = max(self._margin, right_edge - tw)
            y = max(self._margin, y_cursor - th)
            toast.move(QPoint(x, y))
            toast.raise_()
            y_cursor = y - self._spacing

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        if watched is self._host and event.type() in (QEvent.Resize, QEvent.Move, QEvent.Show):
            self.reposition()
        return super().eventFilter(watched, event)


def notify_toast(
    widget: Optional[QWidget],
    title: str,
    message: str = "",
    variant: str = "info",
    actions: Optional[Sequence[Union[ToastAction, Tuple[str, Callable[[], None]]]]] = None,
    duration_ms: int = 6000,
) -> Optional[ToastWidget]:
    """Show a non-blocking shell toast when hosted inside a ToastManager window.

    When running in standalone dialog or headless unit-test contexts without a
    top-level `show_toast` method, logs the message cleanly without blocking.
    """
    host = widget.window() if widget is not None else None
    if host is not None and callable(getattr(host, "show_toast", None)):
        return host.show_toast(
            title=title,
            message=message,
            variant=variant,
            actions=actions,
            duration_ms=duration_ms,
        )
    logger.info("[%s] %s%s", variant.upper(), title, f" — {message}" if message else "")
    return None


def notify_success(
    widget: Optional[QWidget],
    title: str,
    message: str = "",
    actions: Optional[Sequence[Union[ToastAction, Tuple[str, Callable[[], None]]]]] = None,
    duration_ms: int = 5500,
) -> Optional[ToastWidget]:
    """Convenience helper for non-blocking success notifications."""
    return notify_toast(
        widget=widget,
        title=title,
        message=message,
        variant="success",
        actions=actions,
        duration_ms=duration_ms,
    )


def notify_info(
    widget: Optional[QWidget],
    title: str,
    message: str = "",
    actions: Optional[Sequence[Union[ToastAction, Tuple[str, Callable[[], None]]]]] = None,
    duration_ms: int = 5000,
) -> Optional[ToastWidget]:
    """Convenience helper for non-blocking informational notifications."""
    return notify_toast(
        widget=widget,
        title=title,
        message=message,
        variant="info",
        actions=actions,
        duration_ms=duration_ms,
    )
