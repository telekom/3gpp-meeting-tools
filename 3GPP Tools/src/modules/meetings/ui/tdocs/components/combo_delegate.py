# --- File: src/modules/meetings/ui/tdocs/components/combo_delegate.py ---
"""Theme-aware item delegate for CheckableComboBox popups."""

from PyQt5.QtCore import QPoint, QRectF, QSize, Qt
from PyQt5.QtGui import QBrush, QColor, QPainter, QPen
from PyQt5.QtWidgets import QStyle, QStyledItemDelegate, QStyleOptionViewItem

from core.ui.font_scale import scale_px
from core.ui.theme import C


class CheckableComboItemDelegate(QStyledItemDelegate):
    """
    Draws crisp theme-aware checkboxes and text inside CheckableComboBox popups.

    Using an explicit item delegate prevents Qt's QStyleSheetStyle + Fusion
    PE_IndicatorItemViewItemCheck interaction from painting solid black boxes
    when ancestor widgets have transparent/borderless QWidget stylesheets.
    """

    def paint(self, painter, option, index):
        if not (index.flags() & Qt.ItemIsUserCheckable):
            super().paint(painter, option, index)
            return

        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        is_checked = index.data(Qt.CheckStateRole) in (Qt.Checked, 2)
        is_enabled = bool(opt.state & QStyle.State_Enabled)
        is_hovered = bool(opt.state & QStyle.State_MouseOver)
        is_selected = bool(opt.state & QStyle.State_Selected)

        raw_text = opt.text
        fg_brush = index.data(Qt.ForegroundRole)
        opt.text = ""
        opt.features &= ~QStyleOptionViewItem.HasCheckIndicator
        style = opt.widget.style() if opt.widget else None
        if style is not None:
            style.drawControl(QStyle.CE_ItemViewItem, opt, painter, opt.widget)

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)

        rect = opt.rect
        box_sz = scale_px(15, minimum=14)
        left_pad = scale_px(8)
        gap = scale_px(8)
        box_y = rect.top() + (rect.height() - box_sz) / 2.0
        box_rect = QRectF(rect.left() + left_pad, box_y, box_sz, box_sz)

        if is_checked:
            fill_col = QColor(C.ACCENT_PRIMARY if is_enabled else C.BORDER_STRONG)
            border_col = fill_col
            painter.setBrush(fill_col)
            painter.setPen(QPen(border_col, 1.2))
            painter.drawRoundedRect(box_rect, 3.0, 3.0)

            check_pen = QPen(
                QColor(C.TEXT_ON_ACCENT),
                max(1.6, box_sz * 0.12),
                Qt.SolidLine,
                Qt.RoundCap,
                Qt.RoundJoin,
            )
            painter.setPen(check_pen)
            x1, y1 = box_rect.left() + box_sz * 0.22, box_rect.top() + box_sz * 0.52
            x2, y2 = box_rect.left() + box_sz * 0.43, box_rect.top() + box_sz * 0.73
            x3, y3 = box_rect.left() + box_sz * 0.78, box_rect.top() + box_sz * 0.28
            painter.drawLine(QPoint(int(round(x1)), int(round(y1))), QPoint(int(round(x2)), int(round(y2))))
            painter.drawLine(QPoint(int(round(x2)), int(round(y2))), QPoint(int(round(x3)), int(round(y3))))
        else:
            fill_col = QColor(C.BG_CARD)
            border_col = QColor(C.BORDER_HOVER if is_hovered else C.BORDER_STRONG)
            painter.setBrush(fill_col)
            painter.setPen(QPen(border_col, 1.4))
            painter.drawRoundedRect(box_rect, 3.0, 3.0)

        text_color = QColor(C.TEXT_STRONG)
        if isinstance(fg_brush, QBrush) and fg_brush.style() != Qt.NoBrush:
            text_color = fg_brush.color()
        elif isinstance(fg_brush, QColor) and fg_brush.isValid():
            text_color = fg_brush
        elif is_selected:
            text_color = QColor(C.ACCENT_PRIMARY)

        painter.setPen(text_color)
        painter.setFont(opt.font)
        text_rect = rect.adjusted(int(round(left_pad + box_sz + gap)), 0, -scale_px(6), 0)
        painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft | Qt.TextSingleLine, raw_text)
        painter.restore()

    def sizeHint(self, option, index):
        base = super().sizeHint(option, index)
        fm = option.fontMetrics
        box_sz = scale_px(15, minimum=14)
        extra_w = scale_px(8) + box_sz + scale_px(8) + scale_px(12)
        text = str(index.data(Qt.DisplayRole) or "")
        text_w = fm.horizontalAdvance(text)
        h = max(base.height(), fm.height() + scale_px(8), scale_px(24))
        return QSize(max(base.width(), text_w + extra_w), h)
