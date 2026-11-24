"""Live reader/editor Ctrl+Wheel & Ctrl+0 text zoom and compact help tips pill."""

from __future__ import annotations

import html as _html
from typing import Callable, List, Optional
from weakref import WeakKeyDictionary

from PyQt5.QtCore import QEvent, QObject, QSettings, Qt
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QLabel,
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QWidgetAction,
)

from core.ui.font_scaling.state import scale_px

_MIN_READER_ZOOM = -4
_MAX_READER_ZOOM = 10
_ZOOM_CHANGE_CALLBACKS: "WeakKeyDictionary[QWidget, Callable[[int], None]]" = WeakKeyDictionary()
_ZOOM_EVENT_FILTER: Optional[QObject] = None


def _unwrap_zoomable_widget(widget: Optional[QWidget]) -> Optional[QWidget]:
    """Walk parents from focus widget (e.g. QTextEdit viewport) to find a zoomable text widget."""
    cur = widget
    while cur is not None:
        if isinstance(cur, (QTextEdit, QPlainTextEdit, QTextBrowser)):
            return cur
        cur = cur.parentWidget()
    return None


def find_active_zoom_target(window: Optional[QWidget]) -> Optional[QWidget]:
    """
    Return the focused `QTextEdit`/`QPlainTextEdit`/`QTextBrowser`, or fall back
    to the primary visible text/reader widget in the active workspace tab.
    """
    focus_target = _unwrap_zoomable_widget(QApplication.focusWidget())
    if focus_target is not None:
        return focus_target

    if window is None:
        return None

    tabs = getattr(window, "tabs", None)
    active_page = tabs.currentWidget() if tabs is not None else window
    if active_page is None:
        return None

    candidates: List[QWidget] = []
    for cls in (QTextBrowser, QPlainTextEdit, QTextEdit):
        for w in active_page.findChildren(cls):
            if w.isVisible() and w.height() >= 80 and w.width() >= 120:
                candidates.append(w)
    if not candidates:
        return None
    return max(candidates, key=lambda w: w.width() * w.height())


def zoom_text_widget(widget: Optional[QWidget], delta_steps: int, reset: bool = False) -> Optional[int]:
    """
    Zoom a `QTextEdit` / `QPlainTextEdit` / `QTextBrowser` by `delta_steps`, or reset to 0.
    Handles pixel-sized global QSS fonts cleanly without collapsing `pointSizeF()` to <= 0.
    Returns the updated zoom step count, or None if `widget` is not zoomable.
    """
    target = _unwrap_zoomable_widget(widget)
    if target is None:
        return None

    cur_steps = int(target.property("_reader_zoom_steps") or 0)
    if reset:
        nxt_steps = 0
    else:
        nxt_steps = max(_MIN_READER_ZOOM, min(_MAX_READER_ZOOM, cur_steps + int(delta_steps)))

    base_pt_prop = target.property("_reader_zoom_base_pt")
    cur_font = target.font()
    if base_pt_prop is None:
        if cur_font.pointSizeF() > 0:
            base_pt = float(cur_font.pointSizeF())
        elif cur_font.pixelSize() > 0:
            dpi_y = float(target.logicalDpiY() or 96.0)
            base_pt = max(6.0, round(cur_font.pixelSize() * 72.0 / dpi_y, 2))
        else:
            base_pt = 9.5
        target.setProperty("_reader_zoom_base_pt", base_pt)
    else:
        base_pt = float(base_pt_prop)

    if nxt_steps != cur_steps or reset:
        new_pt = max(5.0, float(base_pt) + nxt_steps)
        cur_font.setPointSizeF(new_pt)
        target.setFont(cur_font)
        doc = target.document() if hasattr(target, "document") else None
        if doc is not None:
            doc.setDefaultFont(cur_font)

    target.setProperty("_reader_zoom_steps", nxt_steps)
    cb = _ZOOM_CHANGE_CALLBACKS.get(target)
    if cb is not None and (nxt_steps != cur_steps or reset):
        try:
            cb(nxt_steps)
        except Exception:
            pass
    return nxt_steps


def bind_reader_zoom_settings(widget: Optional[QWidget], settings: QSettings, key: str = "zoom_steps") -> None:
    """
    Install reader zoom, restore `zoom_steps` from `settings`, and flush updates immediately on every zoom change.
    """
    install_reader_zoom()
    target = _unwrap_zoomable_widget(widget)
    if target is None or settings is None:
        return

    try:
        saved_zoom = int(settings.value(key, 0, type=int) or 0)
    except Exception:
        saved_zoom = 0
    if saved_zoom:
        zoom_text_widget(target, saved_zoom)

    def _persist(steps: int, _s: QSettings = settings, _k: str = key) -> None:
        try:
            _s.setValue(_k, int(steps))
            _s.sync()
        except Exception:
            pass

    _ZOOM_CHANGE_CALLBACKS[target] = _persist


def create_help_tips_pill(
    tips: List[str],
    parent: Optional[QWidget] = None,
    *,
    button_text: str = "? Help",
    title: str = "Quick Tips & Shortcuts",
) -> QPushButton:
    """Creates a compact, discrete Help pill button that opens a clean tips popup on click."""
    from core.ui.theme import C

    btn = QPushButton(button_text, parent)
    btn.setCursor(Qt.PointingHandCursor)
    btn.setToolTip("\n".join(f"• {t}" for t in tips))
    btn.setStyleSheet(
        f"QPushButton {{ background: {C.BG_SURFACE}; color: {C.TEXT_MUTED}; "
        f"border: 1px solid {C.BORDER_STRONG}; border-radius: 10px; "
        "padding: 2px 9px; font-size: 11px; font-weight: 500; } "
        f"QPushButton:hover {{ background: {C.INFO_SOFT}; color: {C.INFO_DARK}; "
        f"border-color: {C.INFO_BORDER}; }}"
    )

    def _show_tips_popup():
        menu = QMenu(btn)
        menu.setStyleSheet(
            f"QMenu {{ background: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; "
            "border-radius: 6px; padding: 0px; }}"
        )
        card = QFrame(menu)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(4)
        hdr = QLabel(f"<b>{_html.escape(title)}</b>", card)
        hdr.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: 12px;")
        lay.addWidget(hdr)
        items_html = "<br>".join(f"• {_html.escape(t)}" for t in tips)
        body = QLabel(items_html, card)
        body.setWordWrap(True)
        body.setMaximumWidth(scale_px(380))
        body.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 11px;")
        lay.addWidget(body)

        wa = QWidgetAction(menu)
        wa.setDefaultWidget(card)
        menu.addAction(wa)
        menu.exec_(btn.mapToGlobal(btn.rect().bottomLeft()))

    btn.clicked.connect(_show_tips_popup)
    return btn


class _ReaderZoomEventFilter(QObject):
    """App-wide Ctrl+Wheel and Ctrl+0 reader zoom handler for QTextBrowser/QTextEdit/QPlainTextEdit."""

    def eventFilter(self, obj, event):  # type: ignore[override]
        try:
            etype = event.type()
            if etype == QEvent.Wheel and isinstance(obj, QWidget):
                if (event.modifiers() & Qt.ControlModifier) and not (
                    event.modifiers() & (Qt.ShiftModifier | Qt.AltModifier)
                ):
                    target = _unwrap_zoomable_widget(obj)
                    if target is not None:
                        dy = event.angleDelta().y()
                        if dy == 0:
                            dy = event.pixelDelta().y()
                        if dy != 0:
                            zoom_text_widget(target, 1 if dy > 0 else -1)
                            event.accept()
                            return True
            elif etype == QEvent.KeyPress and isinstance(obj, QWidget):
                if (event.modifiers() & Qt.ControlModifier) and event.key() == Qt.Key_0:
                    target = _unwrap_zoomable_widget(obj)
                    if target is not None:
                        zoom_text_widget(target, 0, reset=True)
                        event.accept()
                        return True
        except Exception:
            pass
        return super().eventFilter(obj, event)


def install_reader_zoom(app_or_widget: Optional[QObject] = None) -> None:
    """Installs the global Ctrl+Wheel / Ctrl+0 reader zoom filter once per QApplication instance."""
    global _ZOOM_EVENT_FILTER
    app = QApplication.instance()
    if app is None:
        return
    needs_install = _ZOOM_EVENT_FILTER is None
    if not needs_install:
        try:
            if _ZOOM_EVENT_FILTER.parent() is not app:  # type: ignore[union-attr]
                needs_install = True
        except RuntimeError:
            needs_install = True
    if needs_install:
        _ZOOM_EVENT_FILTER = _ReaderZoomEventFilter(app)
        app.installEventFilter(_ZOOM_EVENT_FILTER)
