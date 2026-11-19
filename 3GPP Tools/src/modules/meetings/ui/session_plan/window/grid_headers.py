# --- File: src/modules/meetings/ui/session_plan/window/grid_headers.py ---
"""
Custom-painted Overview grid header views and table (time-first vertical header,
day horizontal header, NOW highlight).
"""
from typing import TYPE_CHECKING, Optional, Tuple
from PyQt5.QtCore import (
    QRect,
    QSize,
    Qt,
)
from PyQt5.QtGui import (
    QColor,
    QFont,
    QFontMetrics,
    QPainter,
    QPen,
)
from PyQt5.QtWidgets import QHeaderView, QTableWidget
from core.ui.font_scale import scale_px
from core.ui.theme import C
from modules.meetings.ui.session_plan.window.constants import DAY_NAMES, GridRowDescriptor

if TYPE_CHECKING:  # pragma: no cover
    from modules.meetings.ui.session_plan.window.main_window import SessionPlanWindow


class SessionPlanVerticalHeader(QHeaderView):
    """
    Custom time-first vertical header for the Session Plan Overview Grid.

    Renders the time range and period code (`Q1 · 09:00 – 10:30`) as the bold
    primary anchor on line 1, with the color-coded stream/room and optional
    local-time subtitle on line 2, plus a live `NOW` highlight when active.
    """

    def __init__(self, window: "SessionPlanWindow", parent=None):
        super().__init__(Qt.Vertical, parent)
        self._window = window
        self.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.setSectionsClickable(False)

    def _row_desc(self, logical_index: int) -> Optional[GridRowDescriptor]:
        rows = getattr(self._window, "grid_rows", None)
        if rows and 0 <= logical_index < len(rows):
            return rows[logical_index]
        return None

    def _fonts(self) -> Tuple[QFont, QFont, QFont]:
        f_time = QFont(self.font())
        f_time.setBold(True)
        f_time.setPixelSize(max(11, scale_px(12)))

        f_sub = QFont(self.font())
        f_sub.setBold(False)
        f_sub.setPixelSize(max(9, scale_px(10)))

        f_badge = QFont(self.font())
        f_badge.setBold(True)
        f_badge.setPixelSize(max(8, scale_px(9)))
        return f_time, f_sub, f_badge

    def sectionSizeFromContents(self, logicalIndex: int) -> QSize:
        r_desc = self._row_desc(logicalIndex)
        if r_desc is None:
            return super().sectionSizeFromContents(logicalIndex)

        f_time, f_sub, f_badge = self._fonts()
        fm_time = QFontMetrics(f_time)
        fm_sub = QFontMetrics(f_sub)
        fm_badge = QFontMetrics(f_badge)

        t_start, t_end = r_desc.time_range
        if r_desc.row_type == "break":
            title_txt = r_desc.display_label.splitlines()[0]
            line1_w = fm_time.horizontalAdvance(title_txt) + scale_px(24)
            line2_txt = f"⏰ {t_start} – {t_end}"
            line2_w = fm_sub.horizontalAdvance(line2_txt) + scale_px(24)
        else:
            line1_txt = f"{r_desc.period} · {t_start} – {t_end}"
            now_extra = fm_badge.horizontalAdvance("NOW") + scale_px(18) if r_desc.is_live_period else 0
            line1_w = fm_time.horizontalAdvance(line1_txt) + now_extra + scale_px(24)
            line2_txt = f"Str {r_desc.stream_number} · {r_desc.short_room}"
            line2_w = fm_sub.horizontalAdvance(line2_txt) + scale_px(32)

        w = max(scale_px(156), line1_w, line2_w)
        h = max(scale_px(44), fm_time.height() + fm_sub.height() + scale_px(14))
        return QSize(w, h)

    def sizeHint(self) -> QSize:
        rows = getattr(self._window, "grid_rows", None) or []
        if not rows:
            return QSize(scale_px(156), scale_px(44))
        max_w = scale_px(156)
        for idx in range(len(rows)):
            max_w = max(max_w, self.sectionSizeFromContents(idx).width())
        return QSize(max_w, super().sizeHint().height())

    def paintSection(self, painter: QPainter, rect: QRect, logicalIndex: int):
        r_desc = self._row_desc(logicalIndex)
        if r_desc is None or not rect.isValid():
            super().paintSection(painter, rect, logicalIndex)
            return

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)

        # Colors by stream / break / live status
        stream_accents = {
            1: C.INDIGO_DARK,
            2: C.EMERALD_DARK,
            3: C.WARNING,
        }
        accent_hex = C.WARNING_AMBER if r_desc.row_type == "break" else stream_accents.get(r_desc.stream_number, C.BORDER_STRONG)
        if r_desc.is_live_period:
            bg_hex = C.SUCCESS_SOFT
            line1_color = C.SUCCESS_DARK
        elif r_desc.row_type == "break":
            bg_hex = C.WARNING_SOFT
            line1_color = C.WARNING_DARK
        else:
            bg_hex = C.BG_SURFACE
            line1_color = C.TEXT_PRIMARY

        painter.fillRect(rect, QColor(bg_hex))

        # Left color accent bar
        bar_w = max(3, scale_px(4))
        painter.fillRect(QRect(rect.left(), rect.top(), bar_w, rect.height()), QColor(accent_hex))

        # Bottom & right grid borders (strong 2px divider at Quarter/Period boundaries)
        rows = getattr(self._window, "grid_rows", None) or []
        tv = getattr(self._window, "grid_table", None)
        is_period_end = True
        for nxt in range(logicalIndex + 1, len(rows)):
            if tv is None or not tv.isRowHidden(nxt):
                is_period_end = (rows[nxt].period != r_desc.period)
                break
        if is_period_end:
            painter.setPen(QPen(QColor(C.TEXT_FAINT), 2))
            painter.drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom())
        else:
            painter.setPen(QPen(QColor(C.BORDER_SUBTLE), 1))
            painter.drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom())
        painter.setPen(QPen(QColor(C.TEXT_FAINT), 2))
        painter.drawLine(rect.right(), rect.top(), rect.right(), rect.bottom())

        f_time, f_sub, f_badge = self._fonts()
        fm_time = QFontMetrics(f_time)
        fm_sub = QFontMetrics(f_sub)
        fm_badge = QFontMetrics(f_badge)

        pad_l = bar_w + scale_px(7)
        pad_r = scale_px(6)
        total_txt_h = fm_time.height() + scale_px(2) + fm_sub.height()
        top_y = rect.top() + max(scale_px(4), (rect.height() - total_txt_h) // 2)

        t_start, t_end = r_desc.time_range

        if r_desc.row_type == "break":
            title_txt = r_desc.display_label.splitlines()[0]
            r1 = QRect(rect.left() + pad_l, top_y, rect.width() - pad_l - pad_r, fm_time.height())
            painter.setFont(f_time)
            painter.setPen(QColor(line1_color))
            painter.drawText(r1, Qt.AlignLeft | Qt.AlignVCenter, title_txt)

            time_str = f"⏰ {t_start} – {t_end}"
            r2 = QRect(rect.left() + pad_l, top_y + fm_time.height() + scale_px(2), rect.width() - pad_l - pad_r, fm_sub.height())
            painter.setFont(f_sub)
            f_sub_bold = QFont(f_sub)
            f_sub_bold.setBold(True)
            painter.setFont(f_sub_bold)
            painter.setPen(QColor(C.WARNING))
            painter.drawText(r2, Qt.AlignLeft | Qt.AlignVCenter, fm_sub.elidedText(time_str, Qt.ElideRight, r2.width()))
        else:
            now_w = 0
            if r_desc.is_live_period:
                badge_txt = "NOW"
                bw = fm_badge.horizontalAdvance(badge_txt) + scale_px(10)
                bh = fm_badge.height() + scale_px(2)
                bx = rect.right() - pad_r - bw
                by = top_y + max(0, (fm_time.height() - bh) // 2)
                b_rect = QRect(bx, by, bw, bh)
                painter.setPen(QPen(QColor(C.SUCCESS_BORDER_MEDIUM), 1))
                painter.setBrush(QColor(C.SUCCESS_SOFT_ALT))
                painter.drawRoundedRect(b_rect, 4, 4)
                painter.setFont(f_badge)
                painter.setPen(QColor(C.SUCCESS_DARK))
                painter.drawText(b_rect, Qt.AlignCenter, badge_txt)
                now_w = bw + scale_px(4)

            line1_txt = f"{r_desc.period} · {t_start} – {t_end}"
            r1 = QRect(rect.left() + pad_l, top_y, rect.width() - pad_l - pad_r - now_w, fm_time.height())
            painter.setFont(f_time)
            painter.setPen(QColor(line1_color))
            painter.drawText(r1, Qt.AlignLeft | Qt.AlignVCenter, fm_time.elidedText(line1_txt, Qt.ElideRight, r1.width()))

            # Stream color dot + Stream / Room
            line2_y = top_y + fm_time.height() + scale_px(2)
            dot_sz = max(6, scale_px(7))
            dot_y = line2_y + max(0, (fm_sub.height() - dot_sz) // 2)
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(accent_hex))
            painter.drawEllipse(rect.left() + pad_l, dot_y, dot_sz, dot_sz)

            sub_x = rect.left() + pad_l + dot_sz + scale_px(5)
            sub_w = rect.right() - pad_r - sub_x
            sub_txt = f"Str {r_desc.stream_number} · {r_desc.short_room}"
            r2 = QRect(sub_x, line2_y, max(20, sub_w), fm_sub.height())
            painter.setFont(f_sub)
            painter.setPen(QColor(C.TEXT_SECONDARY))
            painter.drawText(r2, Qt.AlignLeft | Qt.AlignVCenter, fm_sub.elidedText(sub_txt, Qt.ElideRight, r2.width()))

        painter.restore()


class SessionPlanHorizontalHeader(QHeaderView):
    """
    2-tier merged horizontal header for the Session Plan Overview Grid.

    Top tier merges each visible Mon–Fri column pair into one Day Banner
    (with alternating slate header tints and an emerald Today highlight).
    Bottom tier renders `1st Half` and `2nd Half` separated by a faint dashed
    split, while day boundaries use a strong 2px divider.
    """

    def __init__(self, window: "SessionPlanWindow", parent=None):
        super().__init__(Qt.Horizontal, parent)
        self._window = window
        self.setSectionsClickable(False)
        self.setDefaultAlignment(Qt.AlignCenter)

    def sizeHint(self) -> QSize:
        base = super().sizeHint()
        return QSize(base.width(), max(scale_px(44), 44))

    def paintEvent(self, event):
        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.Antialiasing, True)
        vp_rect = self.viewport().rect()
        total_h = vp_rect.height()
        top_h = max(scale_px(24), int(total_h * 0.56))
        bot_h = max(1, total_h - top_h)

        f_day = QFont(self.font())
        f_day.setBold(True)
        f_day.setPixelSize(max(11, scale_px(12)))
        fm_day = QFontMetrics(f_day)

        f_sub = QFont(self.font())
        f_sub.setBold(True)
        f_sub.setPixelSize(max(9, scale_px(10)))

        f_badge = QFont(self.font())
        f_badge.setBold(True)
        f_badge.setPixelSize(max(8, scale_px(9)))
        fm_badge = QFontMetrics(f_badge)

        painter.fillRect(vp_rect, QColor(C.BG_SURFACE))

        for day_idx in range(len(DAY_NAMES)):
            c1 = day_idx * 2
            c2 = c1 + 1
            if c2 >= self.count():
                continue
            v1 = not self.isSectionHidden(c1)
            v2 = not self.isSectionHidden(c2)
            if not v1 and not v2:
                continue

            x1 = self.sectionViewportPosition(c1)
            w1 = self.sectionSize(c1) if v1 else 0
            x2 = self.sectionViewportPosition(c2)
            w2 = self.sectionSize(c2) if v2 else 0

            left = x1 if v1 else x2
            right = (x2 + w2) if v2 else (x1 + w1)
            if right <= left:
                continue

            day_rect = QRect(left, 0, right - left, top_h)
            day_lbl, is_today = self._window._get_day_header_info(day_idx)

            if is_today:
                banner_bg = C.WG_RAN_BG
                accent_col = C.EMERALD_ACCENT
                accent_h = max(3, scale_px(3))
                txt_col = C.WG_RAN_TEXT
                sub_bg = C.SUCCESS_SOFT
            elif day_idx % 2 == 1:
                banner_bg = C.BG_HOVER
                accent_col = C.TEXT_FAINT
                accent_h = max(2, scale_px(2))
                txt_col = C.TEXT_STRONG
                sub_bg = C.BG_SURFACE
            else:
                banner_bg = C.BG_SURFACE
                accent_col = C.TEXT_MUTED
                accent_h = max(2, scale_px(2))
                txt_col = C.TEXT_PRIMARY
                sub_bg = C.BG_CARD

            painter.fillRect(day_rect, QColor(banner_bg))
            painter.fillRect(QRect(left, 0, right - left, accent_h), QColor(accent_col))

            if is_today:
                badge_txt = "Today"
                bw = fm_badge.horizontalAdvance(badge_txt) + scale_px(10)
                bh = fm_badge.height() + scale_px(2)
                txt_w = fm_day.horizontalAdvance(day_lbl)
                gap = scale_px(6)
                combined_w = txt_w + gap + bw
                start_x = left + max(scale_px(4), ((right - left) - combined_w) // 2)
                painter.setFont(f_day)
                painter.setPen(QColor(txt_col))
                painter.drawText(
                    QRect(start_x, accent_h, txt_w + 2, top_h - accent_h),
                    Qt.AlignVCenter | Qt.AlignLeft,
                    day_lbl,
                )
                by = accent_h + max(1, ((top_h - accent_h) - bh) // 2)
                b_rect = QRect(start_x + txt_w + gap, by, bw, bh)
                painter.setPen(QPen(QColor(C.SUCCESS_BORDER_MEDIUM), 1))
                painter.setBrush(QColor(C.SUCCESS_SOFT_ALT))
                painter.drawRoundedRect(b_rect, 4, 4)
                painter.setFont(f_badge)
                painter.setPen(QColor(C.SUCCESS))
                painter.drawText(b_rect, Qt.AlignCenter, badge_txt)
            else:
                painter.setFont(f_day)
                painter.setPen(QColor(txt_col))
                painter.drawText(
                    QRect(left + 2, accent_h, max(10, right - left - 4), top_h - accent_h),
                    Qt.AlignCenter,
                    fm_day.elidedText(day_lbl, Qt.ElideRight, max(10, right - left - 6)),
                )

            painter.setPen(QPen(QColor(C.BORDER_STRONG), 1))
            painter.drawLine(left, top_h - 1, right - 1, top_h - 1)

            for sub_col, sub_lbl in ((c1, "1st Half"), (c2, "2nd Half")):
                if self.isSectionHidden(sub_col):
                    continue
                sx = self.sectionViewportPosition(sub_col)
                sw = self.sectionSize(sub_col)
                sub_rect = QRect(sx, top_h, sw, bot_h)
                painter.fillRect(sub_rect, QColor(sub_bg))
                painter.setFont(f_sub)
                painter.setPen(QColor(C.SUCCESS_DARK if is_today else C.TEXT_MUTED))
                painter.drawText(sub_rect, Qt.AlignCenter, sub_lbl)

                if sub_col == c1 and v2:
                    dash_pen = QPen(QColor(C.BORDER_STRONG), 1, Qt.DashLine)
                    painter.setPen(dash_pen)
                    painter.drawLine(sx + sw - 1, top_h + 2, sx + sw - 1, total_h - 2)

            painter.setPen(QPen(QColor(C.BORDER_STRONG), 2))
            painter.drawLine(left, total_h - 1, right - 1, total_h - 1)

            painter.setPen(QPen(QColor(C.TEXT_FAINT), 2))
            painter.drawLine(right - 1, 0, right - 1, total_h - 1)


class SessionPlanGridTable(QTableWidget):
    """
    Custom QTableWidget for the Session Plan Overview Grid that paints semantic
    separators instead of uniform gridlines:
    - Strong vertical separators (2px #94A3B8) between days
    - Faint dashed internal separators (1px #CBD5E1) between 1st Half and 2nd Half
    - Strong horizontal separators (2px #94A3B8) between Quarter/Break periods
    - Light horizontal separators (1px #E2E8F0) between streams in the same Quarter
    """

    def __init__(self, window: "SessionPlanWindow", parent=None):
        super().__init__(parent)
        self._window = window
        self.setShowGrid(False)

    def paintEvent(self, event):
        super().paintEvent(event)
        rows = getattr(self._window, "grid_rows", None) or []
        if not rows or self.rowCount() == 0:
            return

        vp = self.viewport()
        painter = QPainter(vp)
        vp_w = vp.width()
        vp_h = vp.height()

        visible_rows = [r for r in range(self.rowCount()) if not self.isRowHidden(r)]
        for day_idx in range(len(DAY_NAMES)):
            c1 = day_idx * 2
            c2 = c1 + 1
            if c2 >= self.columnCount():
                continue
            v1 = not self.isColumnHidden(c1)
            v2 = not self.isColumnHidden(c2)
            if not v1 and not v2:
                continue

            if v1 and v2:
                split_x = self.columnViewportPosition(c1) + self.columnWidth(c1) - 1
                if 0 <= split_x < vp_w:
                    painter.setPen(QPen(QColor(C.BORDER_STRONG), 1, Qt.DashLine))
                    for r in visible_rows:
                        if self.columnSpan(r, c1) == 1:
                            ry = self.rowViewportPosition(r)
                            rh = self.rowHeight(r)
                            painter.drawLine(split_x, ry, split_x, ry + rh)

            end_col = c2 if v2 else c1
            day_right_x = self.columnViewportPosition(end_col) + self.columnWidth(end_col) - 1
            if 0 <= day_right_x < vp_w:
                painter.setPen(QPen(QColor(C.TEXT_FAINT), 2))
                painter.drawLine(day_right_x, 0, day_right_x, vp_h)

        for idx, r in enumerate(visible_rows):
            ry_bot = self.rowViewportPosition(r) + self.rowHeight(r) - 1
            if ry_bot < 0 or ry_bot > vp_h:
                continue
            cur_period = rows[r].period if r < len(rows) else ""
            next_r = visible_rows[idx + 1] if idx + 1 < len(visible_rows) else None
            next_period = rows[next_r].period if (next_r is not None and next_r < len(rows)) else None
            is_period_boundary = (next_period is None) or (cur_period != next_period)
            if is_period_boundary:
                painter.setPen(QPen(QColor(C.TEXT_FAINT), 2))
            else:
                painter.setPen(QPen(QColor(C.BORDER_SUBTLE), 1))
            painter.drawLine(0, ry_bot, vp_w, ry_bot)
