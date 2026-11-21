"""Shared custom widgets, item delegates, and QSS builders for the Meeting Email Inbox."""

from __future__ import annotations

from PyQt5.QtCore import QItemSelectionModel, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QPainter
from PyQt5.QtWidgets import QPlainTextEdit, QStyledItemDelegate, QTableWidget

from core.ui.font_scale import scale_px
from core.ui.theme import C


_STATUS_COLOR_MAP = {
    "🔵 My TDoc": C.ACCENT_ALT,
    "🟢 Support": C.SUCCESS,
    "🔴 Object": C.DANGER_ALT,
    "🟡 Monitor": C.WARNING_ACCENT,
}

_STATUS_BADGE_BG_MAP = {
    "🔵 My TDoc": C.INFO_SOFT,
    "🟢 Support": C.SUCCESS_SOFT,
    "🔴 Object": C.DANGER_SOFT,
    "🟡 Monitor": C.WARNING_SOFT,
}


def _status_color_hex(status: str) -> str:
    return _STATUS_COLOR_MAP.get(str(status or "⚪ Neutral").strip(), "")


def get_compact_action_bar_qss(container_name: str = "InboxActionBar") -> str:
    """Shared compact button styling for email thread/TDoc reading panels."""
    return (
        f"#{container_name} QPushButton {{ padding: 3px 8px; border-radius: 4px; "
        f"border: 1px solid {C.BORDER_STRONG}; background-color: {C.BG_CARD}; color: {C.TEXT_BODY}; }} "
        f"#{container_name} QPushButton:hover {{ background-color: {C.BG_HOVER}; }} "
        f"#{container_name} QPushButton:checked {{ background-color: {C.INFO_SOFT}; border-color: {C.ACCENT_ALT}; color: {C.INFO}; font-weight: bold; }} "
        f"#{container_name} QPushButton:disabled {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_FAINT}; border-color: {C.BORDER_SUBTLE}; }}"
    )


def get_status_combo_qss(status_val: str) -> str:
    """Shared compact tint QSS for My Status selector combos."""
    st = str(status_val or "⚪ Neutral").strip() or "⚪ Neutral"
    accent = _STATUS_COLOR_MAP.get(st, C.BORDER_STRONG)
    bg = _STATUS_BADGE_BG_MAP.get(st, C.BG_CARD)
    return (
        f"QComboBox {{ background-color: {bg}; color: {C.TEXT_PRIMARY}; "
        f"border: 1px solid {accent}; border-radius: 4px; padding: 2px 6px; "
        f"font-weight: {'bold' if st != '⚪ Neutral' else 'normal'}; }}"
    )


class _InboxStatusStripDelegate(QStyledItemDelegate):
    """Paints a crisp left-edge status color stripe on the State column without widening the table."""

    def paint(self, painter: QPainter, option, index):
        super().paint(painter, option, index)
        color_hex = str(index.data(Qt.UserRole + 1) or "").strip()
        if not color_hex:
            return
        painter.save()
        bar_w = scale_px(5, minimum=4)
        rect = option.rect
        painter.fillRect(rect.left() + 1, rect.top() + 2, bar_w, max(4, rect.height() - 4), QColor(color_hex))
        painter.restore()


class _NotePlainTextEdit(QPlainTextEdit):
    """Compact inline note editor that emits save signals on blur or Ctrl+Enter."""

    focusLost = pyqtSignal()
    saveRequested = pyqtSignal()

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        self.focusLost.emit()

    def keyPressEvent(self, event):
        if (event.modifiers() & Qt.ControlModifier) and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.saveRequested.emit()
            event.accept()
            return
        super().keyPressEvent(event)


class _InboxTableWidget(QTableWidget):
    """QTableWidget with ExtendedSelection, deterministic selectRow(), and viewport resize signal."""

    viewportResized = pyqtSignal()

    def selectRow(self, row: int):
        sel_model = self.selectionModel()
        model = self.model()
        if sel_model is None or model is None or row < 0 or row >= self.rowCount():
            return
        idx = model.index(row, 0)
        sel_model.setCurrentIndex(idx, QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.viewportResized.emit()
