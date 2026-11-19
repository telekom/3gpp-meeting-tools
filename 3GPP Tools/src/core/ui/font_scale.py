# --- File: src/core/ui/font_scale.py ---
"""
Global typography scaling, QSS/delegate scaling helpers, and live reader zoom.

Allows delegates with different visual needs or screen sizes to adjust the
application's font scale (90% Compact up to 140% Extra Large, or fine-tuned
85%–160%) with per-user QSettings persistence and instant live re-styling.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from PyQt5.QtCore import QEvent, QObject, QSettings, Qt, pyqtSignal
from PyQt5.QtGui import QFont, QTextDocument
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QHeaderView,
    QLabel,
    QPlainTextEdit,
    QTableView,
    QTextBrowser,
    QTextEdit,
    QTreeView,
    QWidget,
)

from core.ui.shell_state import create_shell_settings

logger = logging.getLogger(__name__)

SETTINGS_KEY_FONT_SCALE = "ui/font_scale"
DEFAULT_FONT_SCALE = 1.0
MIN_FONT_SCALE = 0.85
MAX_FONT_SCALE = 1.60


@dataclass(frozen=True)
class FontScalePreset:
    """Named accessibility/readability preset for the UI typography scale."""

    key: str
    label: str
    description: str
    scale: float


FONT_SCALE_PRESETS: Tuple[FontScalePreset, ...] = (
    FontScalePreset(
        key="compact",
        label="Compact (90%)",
        description="~12 px body · Fits more rows & columns on small displays",
        scale=0.90,
    ),
    FontScalePreset(
        key="standard",
        label="Standard (100% · Default)",
        description="13 px body · Balanced default desktop typography",
        scale=1.00,
    ),
    FontScalePreset(
        key="comfortable",
        label="Comfortable (115%)",
        description="~15 px body · Gently enlarged labels, menus & tables",
        scale=1.15,
    ),
    FontScalePreset(
        key="large",
        label="Large (125%)",
        description="~16 px body · High legibility for laptops & reading",
        scale=1.25,
    ),
    FontScalePreset(
        key="xlarge",
        label="Extra Large (140%)",
        description="~18 px body · Maximum readability across all tools",
        scale=1.40,
    ),
)

_FONT_SIZE_RE = re.compile(r"(font-size\s*:\s*)(\d+(?:\.\d+)?)\s*(px|pt)\b", re.IGNORECASE)
_HEIGHT_RE = re.compile(r"((?:min-height|max-height)\s*:\s*)(\d+(?:\.\d+)?)\s*px\b", re.IGNORECASE)

_CURRENT_SCALE: float = DEFAULT_FONT_SCALE
_SCALED_TO_BASE: Dict[str, str] = {}
_APP_BASE_QSS: Optional[str] = None
_ORIG_WIDGET_SET_STYLESHEET = None
_ORIG_APP_SET_STYLESHEET = None
_ORIG_LABEL_INIT = None
_ORIG_LABEL_SET_TEXT = None
_ORIG_TEXTEDIT_SET_HTML = None
_ORIG_TEXTEDIT_SET_PLAIN = None
_ORIG_TEXTEDIT_CLEAR = None
_ORIG_TEXTBROWSER_SET_HTML = None
_ORIG_TEXTDOC_SET_HTML = None
_NOTIFIER: Optional["_FontScaleNotifier"] = None


class _FontScaleNotifier(QObject):
    """Emits `scale_changed(float)` whenever the active UI font scale updates."""

    scale_changed = pyqtSignal(float)


def get_font_scale_notifier() -> _FontScaleNotifier:
    """Return the singleton QObject notifier for live font-scale updates."""
    global _NOTIFIER
    if _NOTIFIER is None:
        _NOTIFIER = _FontScaleNotifier()
    return _NOTIFIER


def clamp_font_scale(value: object) -> float:
    """Parse and clamp a font scale factor safely into [0.85, 1.60]."""
    try:
        parsed = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return DEFAULT_FONT_SCALE
    if parsed <= 0.0:
        return DEFAULT_FONT_SCALE
    return round(max(MIN_FONT_SCALE, min(MAX_FONT_SCALE, parsed)), 2)


def get_font_scale() -> float:
    """Return the currently active global UI font scale factor."""
    return _CURRENT_SCALE


def preset_for_scale(scale: Optional[float] = None) -> Optional[FontScalePreset]:
    """Return the matching `FontScalePreset` for `scale` if within 1%, else None."""
    target = clamp_font_scale(get_font_scale() if scale is None else scale)
    for preset in FONT_SCALE_PRESETS:
        if abs(preset.scale - target) < 0.015:
            return preset
    return None


def scale_px(base_px: int, minimum: Optional[int] = None, scale: Optional[float] = None) -> int:
    """Scale an integer pixel dimension using the active (or specified) font scale."""
    s = _CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    scaled = int(round(base_px * s))
    floor_val = 1 if minimum is None else int(minimum)
    return max(floor_val, scaled)


def scale_pt(base_pt: int, minimum: int = 7, scale: Optional[float] = None) -> int:
    """Scale an integer point font size using the active (or specified) font scale."""
    s = _CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    return max(int(minimum), int(round(base_pt * s)))


def scaled_font(
    base_font: Optional[QFont] = None,
    base_px: int = 12,
    *,
    bold: Optional[bool] = None,
    family: Optional[str] = None,
    minimum: int = 7,
    scale: Optional[float] = None,
) -> QFont:
    """Return a copy of `base_font` sized strictly in scaled logical pixels."""
    f = QFont(base_font) if base_font is not None else QFont()
    if family:
        f.setFamily(family)
    if bold is not None:
        f.setBold(bool(bold))
    f.setPixelSize(scale_px(base_px, minimum=minimum, scale=scale))
    return f


def scale_row_height(base_px: int, minimum: Optional[int] = None, scale: Optional[float] = None) -> int:
    """Scale a table/tree row height cleanly so enlarged text never clips vertically."""
    s = _CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    floor_val = max(22, int(round(base_px * 0.90))) if minimum is None else int(minimum)
    if s <= 1.0:
        return max(floor_val, int(round(base_px * max(0.92, s))))
    return max(base_px, int(round(base_px * s)))


def scale_stylesheet(qss: str, scale: Optional[float] = None) -> str:
    """
    Scale `font-size: Npx|Npt` rules inside a Qt stylesheet (or inline HTML CSS).

    When `scale > 1.0`, also scales `min-height` / `max-height` constraints in the
    16–48 px band so pills, inputs, and buttons expand instead of clipping text.
    """
    if not qss:
        return qss
    s = _CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    if abs(s - 1.0) < 0.005:
        return qss

    def _replace_font(match: re.Match) -> str:
        prefix, num_str, unit = match.group(1), match.group(2), match.group(3)
        val = float(num_str)
        if val < 6.0:
            return match.group(0)
        scaled = max(6, int(round(val * s)))
        return f"{prefix}{scaled}{unit}"

    updated = _FONT_SIZE_RE.sub(_replace_font, qss)

    if s > 1.0:
        def _replace_height(match: re.Match) -> str:
            prefix, num_str = match.group(1), match.group(2)
            val = float(num_str)
            if 16.0 <= val <= 48.0:
                return f"{prefix}{int(round(val * s))}px"
            return match.group(0)

        updated = _HEIGHT_RE.sub(_replace_height, updated)

    return updated


def scale_html_fonts(html_text: str, scale: Optional[float] = None) -> str:
    """Scale inline CSS `font-size: Npx|Npt` declarations inside rich-text HTML strings."""
    if not html_text or "font-size" not in html_text.lower():
        return html_text
    s = _CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    if abs(s - 1.0) < 0.005:
        return html_text

    def _replace_font(match: re.Match) -> str:
        prefix, num_str, unit = match.group(1), match.group(2), match.group(3)
        val = float(num_str)
        if val < 6.0:
            return match.group(0)
        scaled = max(6, int(round(val * s)))
        return f"{prefix}{scaled}{unit}"

    return _FONT_SIZE_RE.sub(_replace_font, html_text)


def load_font_scale(settings: Optional[QSettings] = None) -> float:
    """Load persisted UI font scale from per-user QSettings."""
    cfg = settings or create_shell_settings()
    raw = cfg.value(SETTINGS_KEY_FONT_SCALE, DEFAULT_FONT_SCALE)
    return clamp_font_scale(raw)


def save_font_scale(scale: float, settings: Optional[QSettings] = None) -> float:
    """Persist UI font scale into per-user QSettings and return the clamped value."""
    clamped = clamp_font_scale(scale)
    cfg = settings or create_shell_settings()
    cfg.setValue(SETTINGS_KEY_FONT_SCALE, f"{clamped:.2f}")
    cfg.sync()
    return clamped


def _resolve_unscaled_qss(incoming: str) -> str:
    """Recover the base unscaled stylesheet if `incoming` came from another widget's `.styleSheet()`."""
    if not incoming:
        return ""
    return _SCALED_TO_BASE.get(incoming, incoming)


def _record_scaled_pair(base_qss: str, scaled_qss: str) -> None:
    if scaled_qss and scaled_qss != base_qss:
        if len(_SCALED_TO_BASE) >= 4096:
            _SCALED_TO_BASE.clear()
        _SCALED_TO_BASE[scaled_qss] = base_qss


def install_font_scaling(app: Optional[QApplication]) -> None:
    """
    Hook `QWidget.setStyleSheet`, `QApplication.setStyleSheet`, and rich-text HTML
    setters (`QLabel`, `QTextEdit`, `QTextBrowser`, `QTextDocument`) once so both QSS
    and inline HTML `font-size: Npx` rules automatically respect `_CURRENT_SCALE`
    without double-scaling.
    """
    global _ORIG_WIDGET_SET_STYLESHEET, _ORIG_APP_SET_STYLESHEET
    global _ORIG_LABEL_INIT, _ORIG_LABEL_SET_TEXT
    global _ORIG_TEXTEDIT_SET_HTML, _ORIG_TEXTEDIT_SET_PLAIN, _ORIG_TEXTEDIT_CLEAR
    global _ORIG_TEXTBROWSER_SET_HTML, _ORIG_TEXTDOC_SET_HTML
    if app is None:
        return

    if _ORIG_WIDGET_SET_STYLESHEET is None:
        _ORIG_WIDGET_SET_STYLESHEET = QWidget.setStyleSheet

        def _scaled_widget_set_stylesheet(self_widget: QWidget, qss: str) -> None:
            raw_text = qss or ""
            if self_widget.property("_font_scale_ignore"):
                _ORIG_WIDGET_SET_STYLESHEET(self_widget, raw_text)
                return
            base_qss = _resolve_unscaled_qss(raw_text)
            self_widget.setProperty("_font_scale_base_qss", base_qss)
            scaled = scale_stylesheet(base_qss, _CURRENT_SCALE)
            _record_scaled_pair(base_qss, scaled)
            _ORIG_WIDGET_SET_STYLESHEET(self_widget, scaled)

        QWidget.setStyleSheet = _scaled_widget_set_stylesheet  # type: ignore[method-assign]

    if _ORIG_APP_SET_STYLESHEET is None:
        _ORIG_APP_SET_STYLESHEET = QApplication.setStyleSheet

        def _scaled_app_set_stylesheet(self_app: QApplication, qss: str) -> None:
            global _APP_BASE_QSS
            base_qss = _resolve_unscaled_qss(qss or "")
            _APP_BASE_QSS = base_qss
            scaled = scale_stylesheet(base_qss, _CURRENT_SCALE)
            _record_scaled_pair(base_qss, scaled)
            _ORIG_APP_SET_STYLESHEET(self_app, scaled)

        QApplication.setStyleSheet = _scaled_app_set_stylesheet  # type: ignore[method-assign]

    if _ORIG_LABEL_SET_TEXT is None:
        _ORIG_LABEL_SET_TEXT = QLabel.setText
        _ORIG_LABEL_INIT = QLabel.__init__

        def _scaled_label_set_text(self_label: QLabel, text: str) -> None:
            raw_text = text or ""
            if self_label.property("_font_scale_ignore") or "font-size" not in raw_text.lower():
                self_label.setProperty("_font_scale_base_html", None)
                _ORIG_LABEL_SET_TEXT(self_label, raw_text)
                return
            base_html = _resolve_unscaled_qss(raw_text)
            self_label.setProperty("_font_scale_base_html", base_html)
            scaled = scale_html_fonts(base_html, _CURRENT_SCALE)
            _record_scaled_pair(base_html, scaled)
            _ORIG_LABEL_SET_TEXT(self_label, scaled)

        def _scaled_label_init(self_label: QLabel, *args, **kwargs) -> None:
            _ORIG_LABEL_INIT(self_label, *args, **kwargs)
            if args and isinstance(args[0], str) and "font-size" in args[0].lower():
                _scaled_label_set_text(self_label, args[0])

        QLabel.setText = _scaled_label_set_text  # type: ignore[method-assign]
        QLabel.__init__ = _scaled_label_init  # type: ignore[method-assign]

    if _ORIG_TEXTEDIT_SET_HTML is None:
        _ORIG_TEXTEDIT_SET_HTML = QTextEdit.setHtml
        _ORIG_TEXTEDIT_SET_PLAIN = QTextEdit.setPlainText
        _ORIG_TEXTEDIT_CLEAR = QTextEdit.clear
        _ORIG_TEXTBROWSER_SET_HTML = QTextBrowser.setHtml

        def _make_html_setter(orig_fn):
            def _scaled_set_html(self_edit: QTextEdit, html_text: str) -> None:
                raw_text = html_text or ""
                if self_edit.property("_font_scale_ignore") or "font-size" not in raw_text.lower():
                    self_edit.setProperty("_font_scale_base_html", None)
                    orig_fn(self_edit, raw_text)
                    return
                base_html = _resolve_unscaled_qss(raw_text)
                self_edit.setProperty("_font_scale_base_html", base_html)
                scaled = scale_html_fonts(base_html, _CURRENT_SCALE)
                _record_scaled_pair(base_html, scaled)
                orig_fn(self_edit, scaled)

            return _scaled_set_html

        def _wrapped_set_plain(self_edit: QTextEdit, text: str) -> None:
            self_edit.setProperty("_font_scale_base_html", None)
            _ORIG_TEXTEDIT_SET_PLAIN(self_edit, text)

        def _wrapped_clear(self_edit: QTextEdit) -> None:
            self_edit.setProperty("_font_scale_base_html", None)
            _ORIG_TEXTEDIT_CLEAR(self_edit)

        QTextEdit.setHtml = _make_html_setter(_ORIG_TEXTEDIT_SET_HTML)  # type: ignore[method-assign]
        QTextBrowser.setHtml = _make_html_setter(_ORIG_TEXTBROWSER_SET_HTML)  # type: ignore[method-assign]
        QTextEdit.setPlainText = _wrapped_set_plain  # type: ignore[method-assign]
        QTextEdit.clear = _wrapped_clear  # type: ignore[method-assign]

    if _ORIG_TEXTDOC_SET_HTML is None:
        _ORIG_TEXTDOC_SET_HTML = QTextDocument.setHtml

        def _scaled_doc_set_html(self_doc: QTextDocument, html_text: str) -> None:
            raw_text = html_text or ""
            if "font-size" not in raw_text.lower():
                _ORIG_TEXTDOC_SET_HTML(self_doc, raw_text)
                return
            base_html = _resolve_unscaled_qss(raw_text)
            scaled = scale_html_fonts(base_html, _CURRENT_SCALE)
            _record_scaled_pair(base_html, scaled)
            _ORIG_TEXTDOC_SET_HTML(self_doc, scaled)

        QTextDocument.setHtml = _scaled_doc_set_html  # type: ignore[method-assign]

    install_reader_zoom(app)


def set_unscaled_stylesheet(widget: QWidget, qss: str) -> None:
    """
    Apply a literal stylesheet without automatic font-scale multiplication.

    Used by live preview widgets that explicitly format pixel sizes for a
    previewed target scale rather than the active application scale.
    """
    widget.setProperty("_font_scale_ignore", True)
    raw = qss or ""
    if _ORIG_WIDGET_SET_STYLESHEET is not None:
        _ORIG_WIDGET_SET_STYLESHEET(widget, raw)
    else:
        widget.setStyleSheet(raw)
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.updateGeometry()


def register_scaled_row_height(view_or_header: object, base_row_height: int, minimum: Optional[int] = None) -> None:
    """
    Set and register a table/tree vertical header so its default row height
    updates automatically both now and on live font-scale changes.
    """
    header: Optional[QHeaderView] = None
    if isinstance(view_or_header, QHeaderView):
        header = view_or_header
    elif hasattr(view_or_header, "verticalHeader"):
        header = view_or_header.verticalHeader()  # type: ignore[attr-defined]
    if header is None:
        return
    header.setProperty("_base_row_height", int(base_row_height))
    if minimum is not None:
        header.setProperty("_min_row_height", int(minimum))
    header.setDefaultSectionSize(scale_row_height(base_row_height, minimum=minimum))


def _refresh_registered_widgets(app: QApplication) -> None:
    """Re-apply scaled QSS and row heights across live widgets after a scale change."""
    if _APP_BASE_QSS is not None and _ORIG_APP_SET_STYLESHEET is not None:
        scaled_app = scale_stylesheet(_APP_BASE_QSS, _CURRENT_SCALE)
        _record_scaled_pair(_APP_BASE_QSS, scaled_app)
        _ORIG_APP_SET_STYLESHEET(app, scaled_app)

    for widget in app.allWidgets():
        try:
            if not widget.property("_font_scale_ignore") and _ORIG_WIDGET_SET_STYLESHEET is not None:
                base_qss = widget.property("_font_scale_base_qss")
                if isinstance(base_qss, str):
                    scaled_qss = scale_stylesheet(base_qss, _CURRENT_SCALE)
                    _record_scaled_pair(base_qss, scaled_qss)
                    _ORIG_WIDGET_SET_STYLESHEET(widget, scaled_qss)
                    widget.style().unpolish(widget)
                    widget.style().polish(widget)

                base_html = widget.property("_font_scale_base_html")
                if isinstance(widget, QLabel):
                    if not isinstance(base_html, str):
                        cur_txt = widget.text()
                        if cur_txt and "font-size" in cur_txt.lower():
                            base_html = _resolve_unscaled_qss(cur_txt)
                            widget.setProperty("_font_scale_base_html", base_html)
                    if isinstance(base_html, str) and _ORIG_LABEL_SET_TEXT is not None:
                        scaled_html = scale_html_fonts(base_html, _CURRENT_SCALE)
                        _record_scaled_pair(base_html, scaled_html)
                        _ORIG_LABEL_SET_TEXT(widget, scaled_html)
                elif isinstance(widget, QTextEdit) and isinstance(base_html, str):
                    sb = widget.verticalScrollBar()
                    scroll_pos = sb.value() if sb is not None else 0
                    scaled_html = scale_html_fonts(base_html, _CURRENT_SCALE)
                    _record_scaled_pair(base_html, scaled_html)
                    orig_html_fn = (
                        _ORIG_TEXTBROWSER_SET_HTML
                        if isinstance(widget, QTextBrowser) and _ORIG_TEXTBROWSER_SET_HTML is not None
                        else _ORIG_TEXTEDIT_SET_HTML
                    )
                    if orig_html_fn is not None:
                        orig_html_fn(widget, scaled_html)
                    if sb is not None:
                        sb.setValue(scroll_pos)

            if isinstance(widget, QComboBox):
                # Force Qt to invalidate cached AdjustToContentsOnFirstShow size hints.
                if widget.sizeAdjustPolicy() == QComboBox.AdjustToContentsOnFirstShow:
                    widget.setSizeAdjustPolicy(QComboBox.AdjustToContents)
                    widget.setSizeAdjustPolicy(QComboBox.AdjustToContentsOnFirstShow)
                if widget.view() is not None:
                    widget.view().doItemsLayout()
                    widget.view().updateGeometry()

            if isinstance(widget, QHeaderView):
                base_h = widget.property("_base_row_height")
                min_h = widget.property("_min_row_height")
                if base_h is not None:
                    widget.setDefaultSectionSize(
                        scale_row_height(int(base_h), minimum=(int(min_h) if min_h is not None else None))
                    )
            elif isinstance(widget, QAbstractItemView):
                widget.scheduleDelayedItemsLayout()
                widget.viewport().update()

            if widget.layout() is not None:
                widget.layout().invalidate()
            widget.updateGeometry()
        except RuntimeError:
            pass


def apply_font_scale(
    scale: float,
    app: Optional[QApplication] = None,
    save: bool = True,
    settings: Optional[QSettings] = None,
) -> float:
    """
    Set the active font scale, optionally persist it to QSettings, update
    `QApplication` font & live widget stylesheets, and notify listeners.
    """
    global _CURRENT_SCALE
    new_scale = clamp_font_scale(scale)
    changed = abs(new_scale - _CURRENT_SCALE) >= 0.005
    _CURRENT_SCALE = new_scale

    if save:
        save_font_scale(new_scale, settings=settings)

    target_app = app or QApplication.instance()
    if target_app is not None and isinstance(target_app, QApplication):
        install_font_scaling(target_app)
        from core.ui.theme import load_app_fonts

        load_app_fonts(target_app)
        _refresh_registered_widgets(target_app)

    if changed and _NOTIFIER is not None:
        _NOTIFIER.scale_changed.emit(new_scale)

    return new_scale


# ==========================================
# --- LIVE READER / EDITOR TEXT ZOOM ---
# ==========================================

_MIN_READER_ZOOM = -4
_MAX_READER_ZOOM = 10


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

    # Prefer large visible reader/editor panes inside the active page
    candidates: List[QWidget] = []
    for cls in (QTextBrowser, QPlainTextEdit, QTextEdit):
        for w in active_page.findChildren(cls):
            if w.isVisible() and w.height() >= 80 and w.width() >= 120:
                candidates.append(w)
    if not candidates:
        return None
    # Largest visible reader/editor wins
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
    return nxt_steps


_ZOOM_EVENT_FILTER: Optional[QObject] = None


class _ReaderZoomEventFilter(QObject):
    """App-wide Ctrl+Wheel and Ctrl+0 reader zoom handler for QTextBrowser/QTextEdit/QPlainTextEdit."""

    def eventFilter(self, obj, event):  # type: ignore[override]
        try:
            etype = event.type()
            if etype == QEvent.Wheel and isinstance(obj, QWidget):
                if (event.modifiers() & Qt.ControlModifier) and not (event.modifiers() & (Qt.ShiftModifier | Qt.AltModifier)):
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

