"""Item delegates for the Work Items table view."""

import logging
from PyQt5.QtCore import Qt, QRect, QEvent, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QPen, QPalette, QBrush
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
        if not raw_text or raw_text == "-":
            return

        wgs = [w.strip() for w in str(raw_text).split(",") if w.strip()]
        if not wgs:
            return

        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        font = option.font
        font.setBold(True)
        font.setPointSize(font.pointSize() - 1 if font.pointSize() > 8 else 8)
        painter.setFont(font)
        fm = painter.fontMetrics()

        badges = []
        total_width = 0
        badge_h = 18
        spacing = 3

        for wg in wgs:
            tw = fm.horizontalAdvance(wg) if hasattr(fm, 'horizontalAdvance') else fm.width(wg)
            w = max(tw + 8, 28)
            badges.append((wg, w))
            total_width += w
        total_width += spacing * (len(badges) - 1)

        start_x = option.rect.left() + max(3, (option.rect.width() - total_width) // 2)
        y = option.rect.top() + (option.rect.height() - badge_h) // 2

        cur_x = start_x
        for wg, w in badges:
            colors = get_wg_colors(wg)
            bg_col = QColor(colors.get("bg", "#F1F5F9"))
            border_col = QColor(colors.get("border", "#CBD5E1"))
            text_col = QColor(colors.get("text", "#475569"))

            badge_rect = QRect(cur_x, y, w, badge_h)
            painter.setBrush(bg_col)
            painter.setPen(border_col)
            painter.drawRoundedRect(badge_rect, 3, 3)

            painter.setPen(text_col)
            painter.drawText(badge_rect, Qt.AlignCenter, wg)
            cur_x += w + spacing

        painter.restore()


class WidDelegate(QStyledItemDelegate):
    """Renders Latest WID as an interactive document chip."""
    link_clicked = pyqtSignal(str, str)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text:
            return

        badge_w = min(max(80, option.rect.width() - 12), 95)
        badge_h = 20
        badge_rect = QRect(
            option.rect.left() + (option.rect.width() - badge_w) // 2,
            option.rect.top() + (option.rect.height() - badge_h) // 2,
            badge_w,
            badge_h
        )

        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        is_selected = bool(option.state & QStyle.State_Selected)

        if is_selected:
            painter.setBrush(QColor("#DBEAFE"))
            painter.setPen(QColor("#2563EB"))
        else:
            painter.setBrush(QColor("#EFF6FF"))
            painter.setPen(QColor("#BFDBFE"))
        painter.drawRoundedRect(badge_rect, 4, 4)

        painter.setPen(QColor("#1E40AF" if is_selected else "#1E5C99"))
        font = option.font
        font.setBold(True)
        font.setPointSize(font.pointSize() - 1 if font.pointSize() > 8 else 8)
        painter.setFont(font)
        painter.drawText(badge_rect, Qt.AlignCenter, f"📄 {text}")
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            text = index.data(Qt.DisplayRole)
            if text:
                self.link_clicked.emit(text, "open_doc")
                return True
        return super().editorEvent(event, model, option, index)


class ReleaseBadgeDelegate(QStyledItemDelegate):
    """Renders Release as a subtle standardized pill badge."""

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text or text == "-":
            return

        badge_w = min(max(54, option.rect.width() - 16), 64)
        badge_h = 20
        badge_rect = QRect(
            option.rect.left() + (option.rect.width() - badge_w) // 2,
            option.rect.top() + (option.rect.height() - badge_h) // 2,
            badge_w,
            badge_h
        )

        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        painter.setBrush(QColor("#F1F5F9"))
        painter.setPen(QColor("#CBD5E1"))
        painter.drawRoundedRect(badge_rect, 4, 4)

        painter.setPen(QColor("#334155"))
        font = option.font
        font.setBold(True)
        font.setPointSize(font.pointSize() - 1 if font.pointSize() > 8 else 8)
        painter.setFont(font)
        painter.drawText(badge_rect, Qt.AlignCenter, text)
        painter.restore()


class StatusBadgeDelegate(QStyledItemDelegate):
    """Renders Status as a modern pill badge (Active in emerald, Finished in slate)."""

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text:
            return

        is_active = (str(text).strip().lower() == "active")
        badge_w = 78
        badge_h = 20
        badge_rect = QRect(
            option.rect.left() + (option.rect.width() - badge_w) // 2,
            option.rect.top() + (option.rect.height() - badge_h) // 2,
            badge_w,
            badge_h
        )

        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        if is_active:
            painter.setBrush(QColor("#ECFDF5"))
            painter.setPen(QColor("#A7F3D0"))
            text_color = QColor("#065F46")
            label = "🟢 Active"
        else:
            painter.setBrush(QColor("#F1F5F9"))
            painter.setPen(QColor("#CBD5E1"))
            text_color = QColor("#475569")
            label = "🏁 Finished"

        painter.drawRoundedRect(badge_rect, 10, 10)

        painter.setPen(text_color)
        font = option.font
        font.setBold(True)
        font.setPointSize(font.pointSize() - 1 if font.pointSize() > 8 else 8)
        painter.setFont(font)
        painter.drawText(badge_rect, Qt.AlignCenter, label)
        painter.restore()


class SpecsDelegate(QStyledItemDelegate):
    """Renders the linked specifications count as a standardized pill button."""
    specs_clicked = pyqtSignal(int)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text or text == "-":
            painter.save()
            painter.setPen(QColor("#94A3B8"))
            painter.drawText(option.rect, Qt.AlignCenter, "-")
            painter.restore()
            return

        badge_w = min(max(44, option.rect.width() - 12), 52)
        badge_h = 20
        badge_rect = QRect(
            option.rect.left() + (option.rect.width() - badge_w) // 2,
            option.rect.top() + (option.rect.height() - badge_h) // 2,
            badge_w,
            badge_h
        )

        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        painter.setBrush(QColor("#EBF3FC"))
        painter.setPen(QColor("#BFDBFE"))
        painter.drawRoundedRect(badge_rect, 10, 10)

        painter.setPen(QColor("#1E5C99"))
        font = option.font
        font.setBold(True)
        font.setPointSize(font.pointSize() - 1 if font.pointSize() > 8 else 8)
        painter.setFont(font)
        painter.drawText(badge_rect, Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            text = index.data(Qt.DisplayRole)
            if text and text != "-":
                self.specs_clicked.emit(index.row())
                return True
        return super().editorEvent(event, model, option, index)


class TimelineDelegate(QStyledItemDelegate):
    """Renders compact Year-Month timeline."""

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
        if option.state & QStyle.State_Selected:
            painter.setPen(option.palette.color(QPalette.HighlightedText))
        else:
            painter.setPen(QColor("#334155"))
        font = option.font
        font.setPointSize(font.pointSize() - 1 if font.pointSize() > 8 else 8)
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

