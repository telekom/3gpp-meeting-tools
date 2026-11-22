"""Shared custom widgets, item delegates, and QSS builders for the Meeting Email Inbox."""

from __future__ import annotations

from PyQt5.QtCore import QItemSelectionModel, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QPainter
from PyQt5.QtWidgets import (
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QStyledItemDelegate,
    QTableWidget,
    QToolButton,
)

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


def get_live_sync_pill_qss() -> str:
    """Returns QSS styling for the checkable continuous Live Auto-Sync pill button."""
    return (
        f"QPushButton {{ padding: 4px 10px; border-radius: 13px; border: 1px solid {C.BORDER_STRONG}; "
        f"background-color: {C.BG_CARD}; color: {C.TEXT_BODY}; font-size: 12px; }} "
        f"QPushButton:hover {{ background-color: {C.SUCCESS_SOFT}; border-color: {C.SUCCESS_BORDER_MEDIUM}; }} "
        f"QPushButton:checked {{ background-color: {C.SUCCESS_SOFT_ALT}; border-color: {C.SUCCESS}; "
        f"color: {C.SUCCESS_DARK}; font-weight: bold; }} "
        f"QPushButton:checked:hover {{ background-color: {C.SUCCESS_BORDER}; }}"
    )


def build_inbox_sync_controls(dialog):
    """Creates and attaches the Sync Outlook split button + Live (2m) toggle pill to `dialog`."""
    sync_btn = QToolButton(dialog)
    sync_btn.setText("⚡ Sync New")
    sync_btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
    sync_btn.setPopupMode(QToolButton.MenuButtonPopup)
    sync_btn.setToolTip(
        "Fast incremental Outlook sync of new Tracked Thread replies, TDoc emails, and read/flag updates since your last sync."
    )
    sync_menu = QMenu(sync_btn)
    act_selected = sync_menu.addAction("🔄 Sync Selected Thread")
    act_selected.setEnabled(False)
    act_selected.triggered.connect(dialog._on_sync_selected_clicked)
    act_filtered = sync_menu.addAction("⚡ Sync Filtered Only")
    act_filtered.triggered.connect(lambda: dialog._start_outlook_sync(scoped=True))
    sync_menu.addSeparator()
    act_full = sync_menu.addAction("📅 Full Meeting Sync (Date Range)...")
    act_full.triggered.connect(dialog._on_full_meeting_sync_clicked)
    sync_btn.setMenu(sync_menu)
    sync_btn.clicked.connect(dialog._on_sync_outlook_clicked)

    btn_live = QPushButton("🟢 Live (2m)", dialog)
    btn_live.setCheckable(True)
    btn_live.setChecked(bool(getattr(dialog, "_live_sync_enabled", False)))
    btn_live.setToolTip(
        "Continuous eMeeting Update Mode (every 2 minutes).\n"
        "Silently pulls newly arrived emails using the persisted watermark without rescanning old history.\n"
        "Persists across restarts and catches up the missing gap automatically when reopened."
    )
    btn_live.setStyleSheet(get_live_sync_pill_qss())
    btn_live.toggled.connect(dialog._on_live_sync_toggled)

    dialog.sync_outlook_btn = sync_btn
    dialog._sync_menu = sync_menu
    dialog._act_sync_selected = act_selected
    dialog._act_sync_filtered = act_filtered
    dialog._act_sync_full = act_full
    dialog._default_sync_scoped = False
    dialog.btn_live_sync = btn_live
    return sync_btn, btn_live


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
