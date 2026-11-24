"""QSS/HTML font-size rewriting, Qt widget monkey-patch hooks, and live scale application."""

from __future__ import annotations

import re
from typing import Optional

from PyQt5.QtCore import QSettings
from PyQt5.QtGui import QTextDocument
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QHeaderView,
    QLabel,
    QTextBrowser,
    QTextEdit,
    QWidget,
)

from core.ui.font_scaling import state as _state
from core.ui.font_scaling.reader_zoom import install_reader_zoom
from core.ui.font_scaling.state import (
    clamp_font_scale,
    save_font_scale,
    scale_row_height,
)


def scale_stylesheet(qss: str, scale: Optional[float] = None) -> str:
    """
    Scale `font-size: Npx|Npt` rules inside a Qt stylesheet (or inline HTML CSS).

    When `scale > 1.0`, also scales `min-height` / `max-height` constraints in the
    16–48 px band so pills, inputs, and buttons expand instead of clipping text.
    """
    if not qss:
        return qss
    s = _state._CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    if abs(s - 1.0) < 0.005:
        return qss

    def _replace_font(match: re.Match) -> str:
        prefix, num_str, unit = match.group(1), match.group(2), match.group(3)
        val = float(num_str)
        if val < 6.0:
            return match.group(0)
        scaled = max(6, int(round(val * s)))
        return f"{prefix}{scaled}{unit}"

    updated = _state._FONT_SIZE_RE.sub(_replace_font, qss)

    if s > 1.0:
        def _replace_height(match: re.Match) -> str:
            prefix, num_str = match.group(1), match.group(2)
            val = float(num_str)
            if 16.0 <= val <= 48.0:
                return f"{prefix}{int(round(val * s))}px"
            return match.group(0)

        updated = _state._HEIGHT_RE.sub(_replace_height, updated)

    return updated


def scale_html_fonts(html_text: str, scale: Optional[float] = None) -> str:
    """Scale inline CSS `font-size: Npx|Npt` declarations inside rich-text HTML strings."""
    if not html_text or "font-size" not in html_text.lower():
        return html_text
    s = _state._CURRENT_SCALE if scale is None else clamp_font_scale(scale)
    if abs(s - 1.0) < 0.005:
        return html_text

    def _replace_font(match: re.Match) -> str:
        prefix, num_str, unit = match.group(1), match.group(2), match.group(3)
        val = float(num_str)
        if val < 6.0:
            return match.group(0)
        scaled = max(6, int(round(val * s)))
        return f"{prefix}{scaled}{unit}"

    return _state._FONT_SIZE_RE.sub(_replace_font, html_text)


def _resolve_unscaled_qss(incoming: str) -> str:
    """Recover the base unscaled stylesheet if `incoming` came from another widget's `.styleSheet()`."""
    if not incoming:
        return ""
    return _state._SCALED_TO_BASE.get(incoming, incoming)


def _record_scaled_pair(base_qss: str, scaled_qss: str) -> None:
    if scaled_qss and scaled_qss != base_qss:
        if len(_state._SCALED_TO_BASE) >= 4096:
            _state._SCALED_TO_BASE.clear()
        _state._SCALED_TO_BASE[scaled_qss] = base_qss


def install_font_scaling(app: Optional[QApplication]) -> None:
    """
    Hook `QWidget.setStyleSheet`, `QApplication.setStyleSheet`, and rich-text HTML
    setters (`QLabel`, `QTextEdit`, `QTextBrowser`, `QTextDocument`) once so both QSS
    and inline HTML `font-size: Npx` rules automatically respect `_CURRENT_SCALE`
    without double-scaling.
    """
    if app is None:
        return

    if _state._ORIG_WIDGET_SET_STYLESHEET is None:
        _state._ORIG_WIDGET_SET_STYLESHEET = QWidget.setStyleSheet

        def _scaled_widget_set_stylesheet(self_widget: QWidget, qss: str) -> None:
            raw_text = qss or ""
            if self_widget.property("_font_scale_ignore"):
                _state._ORIG_WIDGET_SET_STYLESHEET(self_widget, raw_text)
                return
            base_qss = _resolve_unscaled_qss(raw_text)
            self_widget.setProperty("_font_scale_base_qss", base_qss)
            scaled = scale_stylesheet(base_qss, _state._CURRENT_SCALE)
            _record_scaled_pair(base_qss, scaled)
            _state._ORIG_WIDGET_SET_STYLESHEET(self_widget, scaled)

        QWidget.setStyleSheet = _scaled_widget_set_stylesheet  # type: ignore[method-assign]

    if _state._ORIG_APP_SET_STYLESHEET is None:
        _state._ORIG_APP_SET_STYLESHEET = QApplication.setStyleSheet

        def _scaled_app_set_stylesheet(self_app: QApplication, qss: str) -> None:
            base_qss = _resolve_unscaled_qss(qss or "")
            _state._APP_BASE_QSS = base_qss
            scaled = scale_stylesheet(base_qss, _state._CURRENT_SCALE)
            _record_scaled_pair(base_qss, scaled)
            _state._ORIG_APP_SET_STYLESHEET(self_app, scaled)

        QApplication.setStyleSheet = _scaled_app_set_stylesheet  # type: ignore[method-assign]

    if _state._ORIG_LABEL_SET_TEXT is None:
        _state._ORIG_LABEL_SET_TEXT = QLabel.setText
        _state._ORIG_LABEL_INIT = QLabel.__init__

        def _scaled_label_set_text(self_label: QLabel, text: str) -> None:
            raw_text = text or ""
            if self_label.property("_font_scale_ignore") or "font-size" not in raw_text.lower():
                self_label.setProperty("_font_scale_base_html", None)
                _state._ORIG_LABEL_SET_TEXT(self_label, raw_text)
                return
            base_html = _resolve_unscaled_qss(raw_text)
            self_label.setProperty("_font_scale_base_html", base_html)
            scaled = scale_html_fonts(base_html, _state._CURRENT_SCALE)
            _record_scaled_pair(base_html, scaled)
            _state._ORIG_LABEL_SET_TEXT(self_label, scaled)

        def _scaled_label_init(self_label: QLabel, *args, **kwargs) -> None:
            _state._ORIG_LABEL_INIT(self_label, *args, **kwargs)
            if args and isinstance(args[0], str) and "font-size" in args[0].lower():
                _scaled_label_set_text(self_label, args[0])

        QLabel.setText = _scaled_label_set_text  # type: ignore[method-assign]
        QLabel.__init__ = _scaled_label_init  # type: ignore[method-assign]

    if _state._ORIG_TEXTEDIT_SET_HTML is None:
        _state._ORIG_TEXTEDIT_SET_HTML = QTextEdit.setHtml
        _state._ORIG_TEXTEDIT_SET_PLAIN = QTextEdit.setPlainText
        _state._ORIG_TEXTEDIT_CLEAR = QTextEdit.clear
        _state._ORIG_TEXTBROWSER_SET_HTML = QTextBrowser.setHtml

        def _make_html_setter(orig_fn):
            def _scaled_set_html(self_edit: QTextEdit, html_text: str) -> None:
                raw_text = html_text or ""
                if self_edit.property("_font_scale_ignore") or "font-size" not in raw_text.lower():
                    self_edit.setProperty("_font_scale_base_html", None)
                    orig_fn(self_edit, raw_text)
                    return
                base_html = _resolve_unscaled_qss(raw_text)
                self_edit.setProperty("_font_scale_base_html", base_html)
                scaled = scale_html_fonts(base_html, _state._CURRENT_SCALE)
                _record_scaled_pair(base_html, scaled)
                orig_fn(self_edit, scaled)

            return _scaled_set_html

        def _wrapped_set_plain(self_edit: QTextEdit, text: str) -> None:
            self_edit.setProperty("_font_scale_base_html", None)
            _state._ORIG_TEXTEDIT_SET_PLAIN(self_edit, text)

        def _wrapped_clear(self_edit: QTextEdit) -> None:
            self_edit.setProperty("_font_scale_base_html", None)
            _state._ORIG_TEXTEDIT_CLEAR(self_edit)

        QTextEdit.setHtml = _make_html_setter(_state._ORIG_TEXTEDIT_SET_HTML)  # type: ignore[method-assign]
        QTextBrowser.setHtml = _make_html_setter(_state._ORIG_TEXTBROWSER_SET_HTML)  # type: ignore[method-assign]
        QTextEdit.setPlainText = _wrapped_set_plain  # type: ignore[method-assign]
        QTextEdit.clear = _wrapped_clear  # type: ignore[method-assign]

    if _state._ORIG_TEXTDOC_SET_HTML is None:
        _state._ORIG_TEXTDOC_SET_HTML = QTextDocument.setHtml

        def _scaled_doc_set_html(self_doc: QTextDocument, html_text: str) -> None:
            raw_text = html_text or ""
            if "font-size" not in raw_text.lower():
                _state._ORIG_TEXTDOC_SET_HTML(self_doc, raw_text)
                return
            base_html = _resolve_unscaled_qss(raw_text)
            scaled = scale_html_fonts(base_html, _state._CURRENT_SCALE)
            _record_scaled_pair(base_html, scaled)
            _state._ORIG_TEXTDOC_SET_HTML(self_doc, scaled)

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
    if _state._ORIG_WIDGET_SET_STYLESHEET is not None:
        _state._ORIG_WIDGET_SET_STYLESHEET(widget, raw)
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
    if _state._APP_BASE_QSS is not None and _state._ORIG_APP_SET_STYLESHEET is not None:
        scaled_app = scale_stylesheet(_state._APP_BASE_QSS, _state._CURRENT_SCALE)
        _record_scaled_pair(_state._APP_BASE_QSS, scaled_app)
        _state._ORIG_APP_SET_STYLESHEET(app, scaled_app)

    for widget in app.allWidgets():
        try:
            if not widget.property("_font_scale_ignore") and _state._ORIG_WIDGET_SET_STYLESHEET is not None:
                base_qss = widget.property("_font_scale_base_qss")
                if isinstance(base_qss, str):
                    scaled_qss = scale_stylesheet(base_qss, _state._CURRENT_SCALE)
                    _record_scaled_pair(base_qss, scaled_qss)
                    _state._ORIG_WIDGET_SET_STYLESHEET(widget, scaled_qss)
                    widget.style().unpolish(widget)
                    widget.style().polish(widget)

                base_html = widget.property("_font_scale_base_html")
                if isinstance(widget, QLabel):
                    if not isinstance(base_html, str):
                        cur_txt = widget.text()
                        if cur_txt and "font-size" in cur_txt.lower():
                            base_html = _resolve_unscaled_qss(cur_txt)
                            widget.setProperty("_font_scale_base_html", base_html)
                    if isinstance(base_html, str) and _state._ORIG_LABEL_SET_TEXT is not None:
                        scaled_html = scale_html_fonts(base_html, _state._CURRENT_SCALE)
                        _record_scaled_pair(base_html, scaled_html)
                        _state._ORIG_LABEL_SET_TEXT(widget, scaled_html)
                elif isinstance(widget, QTextEdit) and isinstance(base_html, str):
                    sb = widget.verticalScrollBar()
                    scroll_pos = sb.value() if sb is not None else 0
                    scaled_html = scale_html_fonts(base_html, _state._CURRENT_SCALE)
                    _record_scaled_pair(base_html, scaled_html)
                    orig_html_fn = (
                        _state._ORIG_TEXTBROWSER_SET_HTML
                        if isinstance(widget, QTextBrowser) and _state._ORIG_TEXTBROWSER_SET_HTML is not None
                        else _state._ORIG_TEXTEDIT_SET_HTML
                    )
                    if orig_html_fn is not None:
                        orig_html_fn(widget, scaled_html)
                    if sb is not None:
                        sb.setValue(scroll_pos)

            if isinstance(widget, QComboBox):
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
    new_scale = clamp_font_scale(scale)
    changed = abs(new_scale - _state._CURRENT_SCALE) >= 0.005
    _state._CURRENT_SCALE = new_scale

    if save:
        save_font_scale(new_scale, settings=settings)

    target_app = app or QApplication.instance()
    if target_app is not None and isinstance(target_app, QApplication):
        install_font_scaling(target_app)
        from core.ui.theme import load_app_fonts

        load_app_fonts(target_app)
        _refresh_registered_widgets(target_app)

    if changed and _state._NOTIFIER is not None:
        _state._NOTIFIER.scale_changed.emit(new_scale)

    return new_scale
