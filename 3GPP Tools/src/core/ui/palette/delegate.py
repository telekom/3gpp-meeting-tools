"""Custom Qt item delegate for Command Palette headers and action rows."""

from __future__ import annotations

from typing import Dict, Optional, Tuple

from PyQt5.QtCore import QRect, QSize, Qt
from PyQt5.QtGui import QColor, QFont, QIcon, QPen
from PyQt5.QtWidgets import QStyle, QStyledItemDelegate, QWidget

from core.ui.font_scale import scale_px, scale_row_height
from core.ui.icons import lucide_icon
from core.ui.palette.models import CommandItem, _KIND_HEADER, _ROLE_CMD, _ROLE_KIND
from core.ui.theme import C, R, T


class _PaletteRowDelegate(QStyledItemDelegate):
    """Crisp custom delegate for group headers (26px) and action rows (36px)."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._icon_cache: Dict[Tuple[str, bool], QIcon] = {}

    def _get_icon(self, name: str, active: bool) -> QIcon:
        key = (name or "compass", active)
        if key not in self._icon_cache:
            color = C.ACCENT_PRIMARY if active else C.TEXT_MUTED
            self._icon_cache[key] = lucide_icon(key[0], color=color, size=15)
        return self._icon_cache[key]

    def paint(self, painter, option, index) -> None:
        painter.save()
        kind = index.data(_ROLE_KIND)
        rect = option.rect

        if kind == _KIND_HEADER:
            painter.setPen(QColor(C.TEXT_FAINT))
            f = QFont(T.FONT_FAMILY_UI)
            f.setPixelSize(scale_px(10, minimum=8))
            f.setWeight(QFont.Bold)
            painter.setFont(f)
            text = str(index.data(Qt.DisplayRole) or "").upper()
            painter.drawText(rect.adjusted(14, 4, -12, -2), Qt.AlignVCenter | Qt.AlignLeft, text)
            painter.restore()
            return

        cmd: Optional[CommandItem] = index.data(_ROLE_CMD)
        if cmd is None:
            painter.restore()
            return

        selected = bool(option.state & QStyle.State_Selected)
        hovered = bool(option.state & QStyle.State_MouseOver)

        row_rect = rect.adjusted(6, 2, -6, -2)
        if selected:
            painter.setBrush(QColor(C.ACCENT_SOFT))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(row_rect, R.MD, R.MD)
        elif hovered:
            painter.setBrush(QColor(C.BG_SURFACE))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(row_rect, R.MD, R.MD)

        # Icon on left
        icon = self._get_icon(cmd.icon_name, selected)
        icon_rect = QRect(row_rect.left() + 8, row_rect.top() + (row_rect.height() - 16) // 2, 16, 16)
        icon.paint(painter, icon_rect, Qt.AlignCenter)

        # Right-side shortcut / badge pill width calculation
        right_x = row_rect.right() - 8
        pill_text = cmd.shortcut or cmd.badge or ""
        if pill_text:
            pf = QFont(T.FONT_FAMILY_MONO if cmd.shortcut else T.FONT_FAMILY_UI)
            pf.setPixelSize(scale_px(11, minimum=9))
            pf.setWeight(QFont.DemiBold)
            painter.setFont(pf)
            fm_p = painter.fontMetrics()
            pw = fm_p.horizontalAdvance(pill_text) + 12
            ph = scale_px(20, minimum=18)
            pill_rect = QRect(right_x - pw, row_rect.top() + (row_rect.height() - ph) // 2, pw, ph)
            if cmd.shortcut:
                painter.setBrush(QColor(C.BG_SURFACE))
                painter.setPen(QPen(QColor(C.BORDER_SUBTLE), 1))
            else:
                painter.setBrush(QColor(C.ACCENT_SOFT if selected else C.BG_SURFACE))
                painter.setPen(QPen(QColor(C.BORDER_SUBTLE), 1))
            painter.drawRoundedRect(pill_rect, R.SM, R.SM)
            painter.setPen(QColor(C.ACCENT_PRIMARY if selected and not cmd.shortcut else C.TEXT_MUTED))
            painter.drawText(pill_rect, Qt.AlignCenter, pill_text)
            right_x = pill_rect.left() - 8

        text_left = icon_rect.right() + 10
        avail_w = max(80, right_x - text_left)

        # Title + inline muted subtitle
        tf = QFont(T.FONT_FAMILY_UI)
        tf.setPixelSize(scale_px(T.SIZE_BODY))
        tf.setWeight(QFont.DemiBold if selected else QFont.Medium)
        painter.setFont(tf)
        fm_t = painter.fontMetrics()
        title_elided = fm_t.elidedText(cmd.title, Qt.ElideRight, avail_w)
        painter.setPen(QColor(C.ACCENT_PRIMARY if selected else C.TEXT_PRIMARY))
        painter.drawText(
            QRect(text_left, row_rect.top(), avail_w, row_rect.height()),
            Qt.AlignVCenter | Qt.AlignLeft,
            title_elided,
        )

        title_w = fm_t.horizontalAdvance(title_elided)
        rem_w = avail_w - title_w - 10
        if cmd.subtitle and rem_w > 48:
            sf = QFont(T.FONT_FAMILY_UI)
            sf.setPixelSize(scale_px(T.SIZE_CAPTION))
            painter.setFont(sf)
            fm_s = painter.fontMetrics()
            sub_elided = fm_s.elidedText(cmd.subtitle, Qt.ElideRight, rem_w)
            painter.setPen(QColor(C.TEXT_MUTED))
            painter.drawText(
                QRect(text_left + title_w + 10, row_rect.top(), rem_w, row_rect.height()),
                Qt.AlignVCenter | Qt.AlignLeft,
                sub_elided,
            )

        painter.restore()

    def sizeHint(self, option, index) -> QSize:
        kind = index.data(_ROLE_KIND)
        if kind == _KIND_HEADER:
            return QSize(option.rect.width(), scale_row_height(24, minimum=22))
        return QSize(option.rect.width(), scale_row_height(36, minimum=32))
