"""Item delegates for Protocol Library messages and Evolution Matrix tree cells."""

from __future__ import annotations

from typing import Optional

from PyQt5.QtCore import QRect, QRectF, QSize, Qt
from PyQt5.QtGui import QBrush, QColor, QFont, QFontMetrics, QPainter, QPainterPath, QPen
from PyQt5.QtWidgets import (
    QApplication,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QWidget,
)

from core.ui.font_scale import scale_px, scaled_font
from core.ui.theme import C, T
from core.ui.ui_components import get_wg_colors
from modules.nas.ui.nas_dialogs import get_spec_metadata


class NASMessageItemDelegate(QStyledItemDelegate):
    """
    Renders protocol messages and information elements in a clean, scannable 2-line layout:
    - Line 1 (Top): Message / IE Name (regular weight unselected; semibold + accent when selected)
    - Line 2 (Bottom): Compact WG chip + muted 'Spec · Acronym' subtitle, and right-aligned Clause
    """

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)

    def sizeHint(self, option, index) -> QSize:
        return QSize(option.rect.width(), scale_px(42, minimum=42))

    def paint(self, painter: QPainter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)

        rect = option.rect
        is_selected = bool(option.state & QStyle.State_Selected)
        is_hovered = bool(option.state & QStyle.State_MouseOver)

        # 1. Background rendering
        if is_selected:
            bg_brush = QBrush(QColor(C.ACCENT_SOFT))
            border_pen = QPen(QColor(C.ACCENT_RING), 1.0)
        elif is_hovered:
            bg_brush = QBrush(QColor(C.BG_SURFACE))
            border_pen = QPen(QColor(C.BORDER_SUBTLE), 1.0)
        else:
            bg_brush = QBrush(QColor(C.BG_CARD))
            border_pen = None

        bg_rect = QRectF(rect.x() + 2, rect.y() + 1, rect.width() - 4, rect.height() - 2)
        bg_path = QPainterPath()
        bg_path.addRoundedRect(bg_rect, 4.0, 4.0)
        painter.fillPath(bg_path, bg_brush)
        if border_pen is not None:
            painter.setPen(border_pen)
            painter.drawPath(bg_path)
        else:
            painter.setPen(QPen(QColor(C.BORDER_ROW), 1.0))
            painter.drawLine(rect.x() + 8, rect.bottom(), rect.right() - 8, rect.bottom())

        # Active selection accent marker on left border
        if is_selected:
            accent_path = QPainterPath()
            accent_rect = QRectF(rect.x() + 2, rect.y() + 3, 3.0, rect.height() - 6)
            accent_path.addRoundedRect(accent_rect, 1.5, 1.5)
            painter.fillPath(accent_path, QBrush(QColor(C.ACCENT_PRIMARY)))

        # 2. Extract item metadata
        msg_name = str(index.data(Qt.DisplayRole) or "")
        spec_num = str(index.data(Qt.UserRole + 2) or "").replace("TS", "").strip()
        clause = str(index.data(Qt.UserRole + 3) or "").strip()

        meta = get_spec_metadata(spec_num) if spec_num else {}
        wg = meta.get("wg", "")
        acronym = str(index.data(Qt.UserRole + 4) or meta.get("acronym", "")).strip()

        left_x = rect.x() + (10 if not is_selected else 12)
        right_limit = rect.right() - 8

        # 3. Line 1: Message / IE Name (regular when unselected; semibold when selected)
        name_font = scaled_font(option.font, T.SIZE_TABLE)
        name_font.setWeight(QFont.DemiBold if is_selected else QFont.Normal)
        painter.setFont(name_font)
        if is_selected:
            painter.setPen(QColor(C.ACCENT_PRIMARY))
        elif is_hovered:
            painter.setPen(QColor(C.TEXT_PRIMARY))
        else:
            painter.setPen(QColor(C.TEXT_BODY))

        line1_h = max(scale_px(17), (rect.height() - 6) // 2)
        line1_rect = QRect(left_x, rect.y() + 3, max(40, right_limit - left_x), line1_h)
        elided_text = QFontMetrics(name_font).elidedText(msg_name, Qt.ElideRight, line1_rect.width())
        painter.drawText(line1_rect, Qt.AlignLeft | Qt.AlignVCenter, elided_text)

        # 4. Line 2: Compact WG chip + clean 'Spec · Acronym' subtitle, Right-aligned Clause
        chip_y = rect.y() + 3 + line1_h + 1
        chip_h = max(scale_px(15), min(scale_px(16), rect.bottom() - chip_y - 3))

        sub_font = scaled_font(option.font, T.SIZE_CAPTION)
        sub_font.setWeight(QFont.Normal)
        sub_fm = QFontMetrics(sub_font)

        clause_left_edge = right_limit
        if clause and clause != msg_name:
            clause_text = f"Cl. {clause}"
            clause_w = sub_fm.horizontalAdvance(clause_text) + 4
            clause_x = right_limit - clause_w
            if clause_x > left_x + 40:
                painter.setFont(sub_font)
                painter.setPen(QColor(C.TEXT_MUTED))
                painter.drawText(QRect(clause_x, chip_y, clause_w, chip_h), Qt.AlignRight | Qt.AlignVCenter, clause_text)
                clause_left_edge = clause_x - 6

        cur_x = left_x
        if wg:
            wg_colors = get_wg_colors(wg)
            wg_font = scaled_font(option.font, 10)
            wg_font.setWeight(QFont.DemiBold)
            wg_fm = QFontMetrics(wg_font)
            wg_w = wg_fm.horizontalAdvance(wg) + 8
            if cur_x + wg_w <= clause_left_edge:
                wg_path = QPainterPath()
                wg_path.addRoundedRect(QRectF(cur_x, chip_y + 0.5, wg_w, chip_h - 1), 3.0, 3.0)
                painter.fillPath(wg_path, QBrush(QColor(wg_colors.get("bg", C.BG_SURFACE))))
                painter.setPen(QPen(QColor(wg_colors.get("border", C.BORDER_STRONG)), 0.8))
                painter.drawPath(wg_path)

                painter.setFont(wg_font)
                painter.setPen(QColor(wg_colors.get("text", C.TEXT_BODY)))
                painter.drawText(QRect(cur_x, chip_y, wg_w, chip_h), Qt.AlignCenter, wg)
                cur_x += wg_w + 6

        subtitle_parts = []
        if spec_num:
            subtitle_parts.append(spec_num)
        if acronym and acronym != f"TS {spec_num}":
            subtitle_parts.append(acronym)
        subtitle_str = " · ".join(subtitle_parts)

        if subtitle_str:
            avail_sub = clause_left_edge - cur_x
            if avail_sub > 24:
                sub_elided = sub_fm.elidedText(subtitle_str, Qt.ElideRight, avail_sub)
                painter.setFont(sub_font)
                painter.setPen(QColor(C.TEXT_SECONDARY))
                painter.drawText(
                    QRect(cur_x, chip_y, avail_sub, chip_h),
                    Qt.AlignLeft | Qt.AlignVCenter,
                    sub_elided,
                )

        painter.restore()


class NASMatrixTreeDelegate(QStyledItemDelegate):
    """Renders crisp geometric tree guide lines for nested fields in the Evolution Matrix."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)

    def paint(self, painter: QPainter, option, index):
        guide_info = index.data(Qt.UserRole + 1)
        depth = index.data(Qt.UserRole)
        if guide_info is None or depth is None or depth <= 0:
            super().paint(painter, option, index)
            return

        ancestor_continues, is_last_sibling, raw_label = guide_info

        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        opt.text = ""

        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.CE_ItemViewItem, opt, painter, opt.widget)

        painter.save()
        rect = opt.rect
        pad_left = scale_px(6, minimum=4)
        indent_w = scale_px(16, minimum=14)
        mid_y = rect.center().y()

        painter.setPen(QPen(QColor(C.BORDER_STRONG), 1.0))

        for lvl_idx, continues in enumerate(ancestor_continues):
            if continues:
                cx = rect.left() + pad_left + lvl_idx * indent_w + indent_w // 2
                painter.drawLine(cx, rect.top(), cx, rect.bottom())

        branch_cx = rect.left() + pad_left + (depth - 1) * indent_w + indent_w // 2
        branch_end_y = mid_y if is_last_sibling else rect.bottom()
        painter.drawLine(branch_cx, rect.top(), branch_cx, branch_end_y)
        painter.drawLine(branch_cx, mid_y, branch_cx + indent_w // 2 + 2, mid_y)

        text_x = rect.left() + pad_left + depth * indent_w + 4
        text_rect = QRect(text_x, rect.top(), max(20, rect.right() - text_x - 4), rect.height())
        painter.setFont(opt.font)
        is_selected = bool(opt.state & QStyle.State_Selected)
        painter.setPen(QColor(C.TEXT_PRIMARY) if is_selected else QColor(C.TEXT_BODY))
        elided = QFontMetrics(opt.font).elidedText(str(raw_label), Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, elided)
        painter.restore()
