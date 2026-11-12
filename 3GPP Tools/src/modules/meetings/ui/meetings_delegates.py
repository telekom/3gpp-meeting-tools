# --- File: src/modules/meetings/ui/meetings_delegates.py ---
import logging

from PyQt5.QtCore import Qt, QEvent
from PyQt5.QtGui import QBrush, QColor, QPen
from PyQt5.QtWidgets import QMenu, QStyle, QStyledItemDelegate

from core.ui.ui_components import get_wg_colors
from core.ui.theme import C

logger = logging.getLogger(__name__)


class WGTagDelegate(QStyledItemDelegate):
    """Paints styled TSG domain badges for the Working Group column."""
    def __init__(self, parent_tab):
        super().__init__(parent_tab.table)
        self.parent_tab = parent_tab

    def paint(self, painter, option, index):
        wg = index.data(Qt.DisplayRole)
        if not wg:
            super().paint(painter, option, index)
            return

        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        if option.state & QStyle.State_Selected:
            painter.fillRect(option.rect, option.palette.highlight())

        colors = get_wg_colors(wg)
        bg = colors.get("bg", C.BG_CANVAS)
        border = colors.get("border", C.BORDER_STRONG)
        text_col = colors.get("text", C.TEXT_SECONDARY)

        if option.state & QStyle.State_MouseOver:
            border_pen = QPen(QColor(text_col), 1.5)
        else:
            border_pen = QPen(QColor(border), 1)

        rect = option.rect
        badge_rect = rect.adjusted(6, 6, -6, -6)

        painter.setBrush(QBrush(QColor(bg)))
        painter.setPen(border_pen)
        painter.drawRoundedRect(badge_rect, 4, 4)

        font = painter.font()
        font.setPointSize(9)
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(QPen(QColor(text_col)))
        painter.drawText(badge_rect, Qt.AlignCenter, wg)

        painter.restore()


class TDocsButtonDelegate(QStyledItemDelegate):
    def __init__(self, parent_tab):
        super().__init__(parent_tab.table)
        self.parent_tab = parent_tab

    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        row_data = index.model().data(index, Qt.UserRole)
        if not row_data:
            return

        status = row_data.get('tdoc_btn_status', 'na')
        painter.save()
        painter.setRenderHint(painter.Antialiasing)

        if status == 'na':
            bg, border, text_col, text = C.BG_SURFACE, C.BORDER_SUBTLE, C.TEXT_FAINT, "N/A"
        elif status == 'open':
            bg, border, text_col, text = C.WG_RAN_BG, C.WG_RAN_BORDER, C.WG_RAN_TEXT, "✔ Open"
        elif status == 'fetching':
            bg, border, text_col, text = C.WARNING_SOFT, C.WARNING_BORDER, C.WARNING, "⏳ Fetching"
        else:
            bg, border, text_col, text = C.ACCENT_SOFT, C.INFO_BORDER, C.ACCENT_PRIMARY, "⬇ Get"

        if option.state & QStyle.State_MouseOver and status != 'na':
            border = text_col

        rect = option.rect
        btn_rect = rect.adjusted(4, 6, -4, -6)

        painter.setBrush(QBrush(QColor(bg)))
        painter.setPen(QPen(QColor(border), 1))
        painter.drawRoundedRect(btn_rect, 6, 6)

        font = painter.font()
        font.setPointSize(8)
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(QPen(QColor(text_col)))
        painter.drawText(btn_rect, Qt.AlignCenter, text)

        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            row_data = model.data(index, Qt.UserRole)
            if not row_data:
                return False

            status = row_data.get('tdoc_btn_status', 'na')
            if status == 'open':
                self.parent_tab._open_tdocs_window(row_data, row_data.get('tdoc_filepath'))
            elif status == 'get':
                self.parent_tab._download_and_open_tdocs(row_data, index.row())
            return True

        return super().editorEvent(event, model, option, index)


class DynamicHoverMenuDelegate(QStyledItemDelegate):
    def __init__(self, parent_tab):
        super().__init__(parent_tab.table)
        self.parent_tab = parent_tab

    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        painter.save()
        font = painter.font()
        font.setPointSize(16)
        font.setBold(True)
        painter.setFont(font)

        if option.state & QStyle.State_MouseOver:
            painter.setPen(QColor(C.ACCENT_PRIMARY))
        else:
            painter.setPen(QColor(C.TEXT_FAINT))

        painter.drawText(option.rect, Qt.AlignCenter, "⋮")
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            row_data = model.data(index, Qt.UserRole)
            if not row_data:
                return False

            menu = QMenu(self.parent_tab)
            self.parent_tab._populate_dynamic_menu(menu, row_data, index.row())
            menu.exec_(event.globalPos())
            return True

        return super().editorEvent(event, model, option, index)
