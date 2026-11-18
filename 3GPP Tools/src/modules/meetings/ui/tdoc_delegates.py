# --- File: modules/meetings/ui/tdoc_delegates.py ---
from PyQt5.QtCore import pyqtSignal, QRectF, QSize, QEvent, Qt, QPoint
from PyQt5.QtGui import QTextDocument, QPainter, QColor, QPen
from PyQt5.QtWidgets import QStyledItemDelegate, QStyleOptionViewItem, QStyle
from core.ui.font_scale import scaled_font, scale_px
from core.ui.theme import C, T


class HtmlDelegate(QStyledItemDelegate):
    linkClicked = pyqtSignal(str)
    linkRightClicked = pyqtSignal(str, QPoint)  # <--- NEW: Signal for Right-Clicks

    @staticmethod
    def _build_document(options: QStyleOptionViewItem, text_width: float) -> QTextDocument:
        doc = QTextDocument()
        doc.setDocumentMargin(4)
        doc.setDefaultFont(options.font)
        doc.setTextWidth(max(20.0, float(text_width)))
        doc.setHtml(options.text)
        return doc

    @staticmethod
    def _vertical_offset(rect_height: float, doc_height: float) -> float:
        if doc_height < rect_height:
            return max(0.0, (rect_height - doc_height) / 2.0)
        return 0.0

    def paint(self, painter, option, index):
        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        painter.save()
        doc = self._build_document(options, options.rect.width())

        options.text = ""
        options.widget.style().drawControl(QStyle.CE_ItemViewItem, options, painter)

        y_offset = self._vertical_offset(options.rect.height(), doc.size().height())
        painter.translate(options.rect.left(), options.rect.top() + y_offset)
        clip = QRectF(0, 0, options.rect.width(), max(0.0, options.rect.height() - y_offset))
        doc.drawContents(painter, clip)
        painter.restore()

    def sizeHint(self, option, index):
        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        col_w = options.rect.width()
        if col_w <= 0 and options.widget is not None and hasattr(options.widget, "columnWidth"):
            col_w = options.widget.columnWidth(index.column())
        if col_w <= 0:
            col_w = scale_px(220)
        doc = self._build_document(options, col_w)
        return QSize(int(col_w), max(scale_px(48), int(doc.size().height()) + scale_px(6)))

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease:
            options = QStyleOptionViewItem(option)
            self.initStyleOption(options, index)
            doc = self._build_document(options, options.rect.width())
            y_offset = self._vertical_offset(options.rect.height(), doc.size().height())

            pos = event.pos() - QPoint(option.rect.left(), int(round(option.rect.top() + y_offset)))
            anchor = doc.documentLayout().anchorAt(pos)
            if anchor:
                # ---> THE FIX: Distinguish between Left and Right Clicks
                if event.button() == Qt.RightButton:
                    self.linkRightClicked.emit(anchor, event.globalPos())
                else:
                    self.linkClicked.emit(anchor)
                return True
        return super().editorEvent(event, model, option, index)


class TDocActionDelegate(QStyledItemDelegate):
    actionClicked = pyqtSignal(str)

    @staticmethod
    def _button_rect(cell_rect) -> QRectF:
        pill = cell_rect.adjusted(4, 4, -4, -4)
        max_h = scale_px(40, minimum=36)
        if pill.height() > max_h:
            top = cell_rect.top() + (cell_rect.height() - max_h) / 2.0
            return QRectF(pill.left(), top, pill.width(), max_h)
        return QRectF(pill)

    def paint(self, painter, option, index):
        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        options.text = ""
        options.widget.style().drawControl(QStyle.CE_ItemViewItem, options, painter)

        tdoc = index.data(Qt.UserRole)
        state = index.data(Qt.UserRole + 1)
        has_revs = index.data(Qt.UserRole + 2)
        if not tdoc: return

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        rect = self._button_rect(option.rect)

        if state == 'THREAD':
            bg_color, border_color, text_color = QColor("#EEF2FF"), QColor("#A5B4FC"), QColor("#4338CA")
            text = "📧 View"
        elif state == 'EXISTS':
            bg_color, border_color, text_color = QColor("#E6F4E6"), QColor("#A3DDA3"), QColor("#0C6B0C")
            text = "✓ Open"
        elif state == 'LOADING':
            bg_color, border_color, text_color = QColor("#FFF4CE"), QColor("#F3C74C"), QColor("#B85C00")
            text = "⏳ Fetching"
        else:
            bg_color, border_color, text_color = QColor("#E1F0FF"), QColor("#99C9FF"), QColor(C.OFFICE_BLUE)
            text = "⬇ Download"

        if has_revs and state != 'LOADING':
            text += " (+Rev)"

        painter.setBrush(bg_color)
        painter.setPen(QPen(border_color, 1))
        painter.drawRoundedRect(rect, 10, 10)

        painter.setPen(text_color)
        painter.setFont(scaled_font(painter.font(), T.SIZE_CAPTION, bold=True))
        painter.drawText(rect, Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease:
            rect = self._button_rect(option.rect)
            if rect.contains(event.pos()):
                tdoc = index.data(Qt.UserRole)
                if tdoc:
                    self.actionClicked.emit(tdoc)
                    return True
        return super().editorEvent(event, model, option, index)



class AgendaItemDelegate(QStyledItemDelegate):
    """
    Renders Agenda Item numbers centered, and when a Sub-Item is present,
    renders a compact, styled rounded pill badge beneath the AI number.
    """
    def paint(self, painter, option, index):
        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        options.text = ""
        options.widget.style().drawControl(QStyle.CE_ItemViewItem, options, painter)

        ai_text = str(index.data(Qt.DisplayRole) or "").strip()
        sub_item = str(index.data(Qt.UserRole + 4) or "").strip()

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)

        rect = option.rect

        if not sub_item:
            # Single-line: AI number centered
            painter.setPen(
                options.palette.color(options.palette.Normal, options.palette.Text)
                if not (options.state & QStyle.State_Selected)
                else options.palette.highlightedText().color()
            )
            font = painter.font()
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignCenter, ai_text)
        else:
            # Two-line: AI number top, Sub-Item pill bottom (vertically centered as a block)
            ai_h = scale_px(20, minimum=18)
            gap = scale_px(2, minimum=2)
            pill_h = scale_px(18, minimum=16)
            total_h = ai_h + gap + pill_h
            top_y = rect.top() + max(2.0, (rect.height() - total_h) / 2.0)

            # 1. AI Number (top)
            ai_rect = QRectF(rect.left(), top_y, rect.width(), ai_h)
            painter.setPen(
                options.palette.color(options.palette.Normal, options.palette.Text)
                if not (options.state & QStyle.State_Selected)
                else options.palette.highlightedText().color()
            )
            font = painter.font()
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(ai_rect, Qt.AlignCenter, ai_text)

            # 2. Sub-Item pill (bottom)
            pill_margin = 3
            pill_rect = QRectF(
                rect.left() + pill_margin,
                top_y + ai_h + gap,
                rect.width() - (2 * pill_margin),
                pill_h,
            )

            if pill_rect.width() > 16:
                bg_color = QColor(C.BG_CANVAS) if not (options.state & QStyle.State_Selected) else QColor(C.INFO_SOFT_STRONG)
                border_color = QColor(C.BORDER_STRONG) if not (options.state & QStyle.State_Selected) else QColor(C.INFO_BORDER_MEDIUM)
                text_color = QColor(C.TEXT_BODY) if not (options.state & QStyle.State_Selected) else QColor(C.INFO_DARK)

                painter.setBrush(bg_color)
                painter.setPen(QPen(border_color, 1))
                painter.drawRoundedRect(pill_rect, 4, 4)

                pill_font = painter.font()
                pill_font.setBold(False)
                base_px = pill_font.pixelSize()
                if base_px <= 0:
                    base_px = scale_px(T.SIZE_TABLE)
                pill_font.setPixelSize(max(8, base_px - scale_px(2, minimum=2)))
                painter.setFont(pill_font)
                painter.setPen(text_color)

                fm = painter.fontMetrics()
                available_w = max(10, int(pill_rect.width() - 6))
                elided_sub = fm.elidedText(sub_item, Qt.ElideRight, available_w)
                text_rect = pill_rect.adjusted(3, 0, -3, 0)
                painter.drawText(text_rect, Qt.AlignCenter, elided_sub)

        painter.restore()


class SourceCompanyDelegate(QStyledItemDelegate):
    """
    Renders the TDoc 'Source' column with visual differentiation between confirmed contributors
    (normal foreground text) and unconfirmed / tentative carry-over contributors marked with
    '?' or '[...]' (rendered greyed-out and italicized).
    Preserves standard plain-text painting and fast eliding when no tentative sources exist.
    """
    def paint(self, painter, option, index):
        from core.utils.company_sanitizer import CompanySanitizer

        raw_source = str(index.data(Qt.UserRole) or index.data(Qt.DisplayRole) or "").strip()
        if not raw_source or not CompanySanitizer.has_tentative_sources(raw_source):
            super().paint(painter, option, index)
            return

        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        options.text = ""
        style = options.widget.style() if options.widget else None
        if style:
            style.drawControl(QStyle.CE_ItemViewItem, options, painter)

        is_selected = bool(options.state & QStyle.State_Selected)
        normal_color = (
            options.palette.highlightedText().color().name()
            if is_selected
            else options.palette.color(options.palette.Normal, options.palette.Text).name()
        )
        tentative_color = C.TEXT_MUTED if is_selected else C.TEXT_FAINT

        import html as _html
        parts = []
        for seg_text, is_tent in CompanySanitizer.tokenize_source_segments(raw_source):
            esc = _html.escape(seg_text)
            if is_tent:
                parts.append(f'<span style="color: {tentative_color}; font-style: italic;">{esc}</span>')
            else:
                parts.append(f'<span style="color: {normal_color};">{esc}</span>')

        painter.save()
        doc = QTextDocument()
        doc.setDocumentMargin(4)
        doc.setDefaultFont(options.font)
        doc.setTextWidth(max(20, options.rect.width()))
        doc.setHtml("".join(parts))

        painter.translate(options.rect.left(), options.rect.top())
        clip = QRectF(0, 0, options.rect.width(), options.rect.height())
        doc.drawContents(painter, clip)
        painter.restore()

