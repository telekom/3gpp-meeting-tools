"""Custom Qt item delegates for the Spec Search Evolution Matrix."""

from PyQt5.QtCore import QModelIndex, QRect, Qt
from PyQt5.QtGui import QBrush, QColor, QFontMetrics, QPainter, QPen
from PyQt5.QtWidgets import QStyle, QStyledItemDelegate, QStyleOptionViewItem

from core.ui.font_scale import scaled_font, scale_px
from core.ui.theme import C, T


class SpecEvolutionStatusDelegate(QStyledItemDelegate):
    """
    Renders status cells in the Evolution Matrix as sleek, centered, rounded pill badges
    (🟢 Added, ⚡ Post-Cutoff, ✓ Present, or —) without obscuring row selection.
    """

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex):
        col = index.column()

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)

        is_selected = bool(option.state & QStyle.State_Selected)
        is_hover = bool(option.state & QStyle.State_MouseOver)

        # 1. Background fill
        if is_selected:
            bg_brush = QBrush(QColor(C.ACCENT_SOFT))
        elif is_hover:
            bg_brush = QBrush(QColor(C.BG_SURFACE))
        elif index.row() % 2 == 1:
            bg_brush = QBrush(QColor(C.BG_SURFACE))
        else:
            bg_brush = QBrush(QColor(C.BG_CARD))
        painter.fillRect(option.rect, bg_brush)

        # Subtle bottom separator line
        painter.setPen(QPen(QColor(C.BG_CANVAS), 1))
        painter.drawLine(option.rect.bottomLeft(), option.rect.bottomRight())

        # 2. Cell Content
        if col == 0:
            val = str(index.data(Qt.DisplayRole) or "")
            text_rect = option.rect.adjusted(8, 0, -8, 0)
            painter.setFont(scaled_font(option.font, T.SIZE_TABLE, bold=True))
            painter.setPen(QColor(C.ACCENT_SKY if not is_selected else C.ACCENT_SKY_DARK))
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, val)
        elif col == 1:
            val = str(index.data(Qt.DisplayRole) or "")
            text_rect = option.rect.adjusted(8, 0, -8, 0)
            painter.setFont(scaled_font(option.font, T.SIZE_TABLE, bold=False))
            painter.setPen(QColor(C.TEXT_PRIMARY if not is_selected else C.TEXT_STRONG))
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, val)
        else:
            display_text = str(index.data(Qt.DisplayRole) or "").strip()
            if not display_text or display_text in ("-", "—"):
                painter.setFont(scaled_font(option.font, T.SIZE_BODY, bold=False))
                painter.setPen(QColor(C.BORDER_STRONG))
                painter.drawText(option.rect, Qt.AlignCenter, "—")
            else:
                if "Post-Cutoff" in display_text or "⚡" in display_text:
                    pill_bg = QColor(C.WARNING_SOFT_ALT)
                    pill_border = QColor(C.WARNING_BORDER)
                    pill_fg = QColor(C.WARNING)
                    badge_label = "⚡ Post-Cutoff"
                elif "Added" in display_text or "🟢" in display_text:
                    pill_bg = QColor(C.SUCCESS_SOFT_ALT)
                    pill_border = QColor(C.SUCCESS_BORDER_MEDIUM)
                    pill_fg = QColor(C.SUCCESS)
                    badge_label = "🟢 Added"
                elif "Present" in display_text or "✓" in display_text:
                    pill_bg = QColor(C.BG_CANVAS)
                    pill_border = QColor(C.BORDER_SUBTLE)
                    pill_fg = QColor(C.TEXT_SECONDARY)
                    badge_label = "✓ Present"
                else:
                    pill_bg = QColor(C.BG_SURFACE)
                    pill_border = QColor(C.BORDER_SUBTLE)
                    pill_fg = QColor(C.TEXT_BODY)
                    badge_label = display_text

                pill_font = scaled_font(option.font, T.SIZE_CAPTION, bold=True)
                fm = QFontMetrics(pill_font)
                text_w = fm.horizontalAdvance(badge_label)
                pill_h = scale_px(20)
                pill_w = max(scale_px(68), min(option.rect.width() - 8, text_w + scale_px(14)))

                cx = option.rect.center().x()
                cy = option.rect.center().y()
                pill_rect = QRect(cx - pill_w // 2, cy - pill_h // 2, pill_w, pill_h)

                painter.setBrush(QBrush(pill_bg))
                painter.setPen(QPen(pill_border, 1))
                painter.drawRoundedRect(pill_rect, 10, 10)

                painter.setFont(pill_font)
                painter.setPen(pill_fg)
                painter.drawText(pill_rect, Qt.AlignCenter, badge_label)

        painter.restore()
