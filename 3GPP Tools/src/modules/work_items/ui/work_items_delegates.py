"""Item delegates for the Work Items table view.

Contains custom renderers and interactive delegates for:
- TrackingDelegate: 1-click Star (Personal Interest) and Follow (Assigned Duty) toggles
- WgBadgeDelegate: TSG domain-colored badges for Working Groups
- WidDelegate: Work Item Description links with download/view indicators
- ReleaseBadgeDelegate: Release tags
- StatusBadgeDelegate: Active vs Finished indicators
- SpecsDelegate: Linked specifications pill and counter
- TimelineDelegate: Start and target completion date ranges
- RemarksDelegate: Latest remark with interactive history popup
"""

import logging
from PyQt5.QtCore import Qt, QRect, QEvent, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QPen, QPalette
from PyQt5.QtWidgets import (
    QApplication,
    QStyle,
    QStyledItemDelegate,
    QMenu,
)

from core.ui.ui_components import WG_DESCRIPTIONS, get_wg_colors

logger = logging.getLogger(__name__)


class TrackingDelegate(QStyledItemDelegate):
    """Renders 1-click Star (Personal Interest) and Follow (Assigned Duty) toggles."""
    star_toggled = pyqtSignal(str)
    follow_toggled = pyqtSignal(str)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        state = index.data(Qt.UserRole + 2) or {}
        is_starred = state.get("starred", False)
        is_following = state.get("following", False)

        rect = option.rect
        mid_x = rect.left() + rect.width() // 2
        star_rect = QRect(rect.left() + 2, rect.top() + 3, (rect.width() // 2) - 3, rect.height() - 6)
        follow_rect = QRect(mid_x + 1, rect.top() + 3, (rect.width() // 2) - 3, rect.height() - 6)

        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        # Draw Starred pill/icon
        if is_starred:
            painter.setBrush(QColor("#FEF3C7"))
            painter.setPen(QColor("#F59E0B"))
            painter.drawRoundedRect(star_rect, 4, 4)
            painter.setPen(QColor("#B45309"))
            font = option.font
            font.setPointSize(font.pointSize() if font.pointSize() > 8 else 9)
            painter.setFont(font)
            painter.drawText(star_rect, Qt.AlignCenter, "⭐")
        else:
            painter.setPen(QColor("#94A3B8"))
            font = option.font
            font.setPointSize(font.pointSize() if font.pointSize() > 8 else 9)
            painter.setFont(font)
            painter.drawText(star_rect, Qt.AlignCenter, "☆")

        # Draw Following pill/icon
        if is_following:
            painter.setBrush(QColor("#DBEAFE"))
            painter.setPen(QColor("#3B82F6"))
            painter.drawRoundedRect(follow_rect, 4, 4)
            painter.setPen(QColor("#1E40AF"))
            font = option.font
            font.setPointSize(font.pointSize() if font.pointSize() > 8 else 9)
            painter.setFont(font)
            painter.drawText(follow_rect, Qt.AlignCenter, "🎯")
        else:
            # Concentric circular target outline (avoids red Windows emoji glyphs when unfollowed)
            painter.setPen(QPen(QColor("#94A3B8"), 1.3))
            painter.setBrush(Qt.NoBrush)
            center = follow_rect.center()
            painter.drawEllipse(center, 6, 6)
            painter.drawEllipse(center, 2, 2)

        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            state = index.data(Qt.UserRole + 2) or {}
            wi_code = state.get("code", "")
            if not wi_code:
                return False
            mid_x = option.rect.left() + option.rect.width() // 2
            if event.pos().x() < mid_x:
                self.star_toggled.emit(wi_code)
            else:
                self.follow_toggled.emit(wi_code)
            return True
        return super().editorEvent(event, model, option, index)


class WgBadgeDelegate(QStyledItemDelegate):
    """Renders Working Groups as TSG domain-colored badges side-by-side."""

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        raw_text = index.data(Qt.DisplayRole)
        if not raw_text:
            return

        wgs = [w.strip() for w in str(raw_text).split(",") if w.strip()]
        if not wgs:
            return

        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        badge_font = QFont(option.font)
        badge_font.setPointSize(max(7, badge_font.pointSize() - 2))
        badge_font.setBold(True)
        painter.setFont(badge_font)

        badge_height = 18
        badge_padding_h = 5
        badge_gap = 3

        # Precompute widths
        badge_widths = []
        for wg in wgs:
            w = painter.fontMetrics().horizontalAdvance(wg) + (2 * badge_padding_h)
            badge_widths.append(w)

        total_width = sum(badge_widths) + badge_gap * (len(wgs) - 1)
        start_x = option.rect.left() + (option.rect.width() - total_width) // 2
        y = option.rect.top() + (option.rect.height() - badge_height) // 2

        cur_x = max(start_x, option.rect.left() + 2)
        for i, wg in enumerate(wgs):
            bg_col, fg_col = get_wg_colors(wg)
            w = badge_widths[i]

            badge_rect = QRect(cur_x, y, w, badge_height)
            painter.setBrush(QColor(bg_col))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(badge_rect, 3, 3)

            painter.setPen(QColor(fg_col))
            painter.drawText(badge_rect, Qt.AlignCenter, wg)

            cur_x += w + badge_gap
            if cur_x > option.rect.right() - 2:
                break

        painter.restore()


class WidDelegate(QStyledItemDelegate):
    """Draws WID as a clickable hyperlink pill with an external download icon."""
    wid_clicked = pyqtSignal(str)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text:
            return

        rect = option.rect.adjusted(4, 3, -4, -3)

        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        # Draw soft background pill
        painter.setBrush(QColor("#F1F5F9"))
        painter.setPen(QColor("#CBD5E1"))
        painter.drawRoundedRect(rect, 4, 4)

        # Draw text + icon
        painter.setPen(QColor("#2563EB"))
        font = option.font
        font.setUnderline(True)
        painter.setFont(font)

        display_str = f"{text} ↗"
        painter.drawText(rect, Qt.AlignCenter, display_str)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            text = index.data(Qt.DisplayRole)
            if text:
                self.wid_clicked.emit(text)
                return True
        return super().editorEvent(event, model, option, index)


class ReleaseBadgeDelegate(QStyledItemDelegate):
    """Renders 3GPP Release as a color-coded tag."""

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        rel = index.data(Qt.DisplayRole)
        if not rel:
            return

        rect = option.rect.adjusted(6, 4, -6, -4)
        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        # Distinct colors by Release generation
        rel_str = str(rel).upper()
        if any(r in rel_str for r in ["REL-19", "REL-20", "REL-21"]):
            bg, fg = "#ECFDF5", "#065F46"  # 6G prep: emerald
        elif any(r in rel_str for r in ["REL-15", "REL-16", "REL-17", "REL-18"]):
            bg, fg = "#EFF6FF", "#1E40AF"  # 5G/5G-Adv: blue
        elif any(r in rel_str for r in ["REL-8", "REL-9", "REL-10", "REL-11", "REL-12", "REL-13", "REL-14"]):
            bg, fg = "#F5F3FF", "#5B21B6"  # LTE: purple
        else:
            bg, fg = "#F1F5F9", "#475569"  # Legacy: slate

        painter.setBrush(QColor(bg))
        painter.setPen(QColor(fg).lighter(140))
        painter.drawRoundedRect(rect, 3, 3)

        painter.setPen(QColor(fg))
        font = option.font
        font.setPointSize(max(7, font.pointSize() - 1))
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(rect, Qt.AlignCenter, str(rel))
        painter.restore()


class StatusBadgeDelegate(QStyledItemDelegate):
    """Renders active vs finished work status as high-contrast indicator."""

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        status = index.data(Qt.DisplayRole)
        if not status:
            return

        rect = option.rect.adjusted(6, 4, -6, -4)
        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        if status == "Active":
            bg, fg = "#ECFDF5", "#047857"
            label = "● Active"
        else:
            bg, fg = "#F8FAFC", "#64748B"
            label = "✓ Done"

        painter.setBrush(QColor(bg))
        painter.setPen(QColor(fg).lighter(140))
        painter.drawRoundedRect(rect, 3, 3)

        painter.setPen(QColor(fg))
        font = option.font
        font.setPointSize(max(7, font.pointSize() - 1))
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(rect, Qt.AlignCenter, label)
        painter.restore()


class SpecsDelegate(QStyledItemDelegate):
    """Draws linked specifications count as a clickable badge."""
    specs_clicked = pyqtSignal(str)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text or text == "-":
            painter.save()
            painter.setPen(QColor("#94A3B8"))
            painter.drawText(option.rect, Qt.AlignCenter, "-")
            painter.restore()
            return

        rect = option.rect.adjusted(4, 4, -4, -4)
        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        painter.setBrush(QColor("#EFF6FF"))
        painter.setPen(QColor("#93C5FD"))
        painter.drawRoundedRect(rect, 4, 4)

        painter.setPen(QColor("#1D4ED8"))
        font = option.font
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(rect, Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            text = index.data(Qt.DisplayRole)
            if text and text != "-":
                row_data = model.get_row_data(index.row())
                wi_code = row_data.get("code")
                if wi_code:
                    self.specs_clicked.emit(str(wi_code))
                    return True
        return super().editorEvent(event, model, option, index)


class TimelineDelegate(QStyledItemDelegate):
    """Renders timeline dates with clean tabular typography."""

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text or text == "-":
            painter.save()
            painter.setPen(QColor("#94A3B8"))
            painter.drawText(option.rect, Qt.AlignCenter, "-")
            painter.restore()
            return

        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        painter.setPen(QColor("#334155"))
        font = option.font
        font.setPointSize(max(7, font.pointSize() - 1))
        painter.setFont(font)
        painter.drawText(option.rect, Qt.AlignCenter, text)
        painter.restore()


class RemarksDelegate(QStyledItemDelegate):
    """Draws the latest remark text with a clickable history badge."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.button_width = 45
        self.button_margin = 4

    def get_button_rect(self, option):
        rect = option.rect
        return QRect(
            rect.right() - self.button_width - self.button_margin,
            rect.top() + self.button_margin,
            self.button_width,
            rect.height() - (2 * self.button_margin)
        )

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)

        remarks_list = index.data(Qt.UserRole + 1)
        if not remarks_list:
            return

        latest_remark = remarks_list[0]
        count = len(remarks_list)

        text_rect = option.rect.adjusted(6, 0, -(self.button_width + 10), 0)
        if option.state & QStyle.State_Selected:
            painter.setPen(option.palette.color(QPalette.HighlightedText))
        else:
            painter.setPen(option.palette.color(QPalette.Text))

        elided_text = option.fontMetrics.elidedText(latest_remark, Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, elided_text)

        btn_rect = self.get_button_rect(option)
        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        painter.setBrush(QColor("#EBF3FC"))
        painter.setPen(QColor("#BFDBFE"))
        painter.drawRoundedRect(btn_rect, 4, 4)
        painter.setPen(QColor("#1E5C99"))
        painter.drawText(btn_rect, Qt.AlignCenter, f"💬 {count}")
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            btn_rect = self.get_button_rect(option)
            if btn_rect.contains(event.pos()):
                remarks_list = index.data(Qt.UserRole + 1)
                if remarks_list:
                    menu = QMenu()
                    for remark in remarks_list:
                        menu.addAction(remark)
                    menu.exec_(event.globalPos())
                return True
        return super().editorEvent(event, model, option, index)
