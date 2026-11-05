# --- File: modules/meetings/ui/tdoc_delegates.py ---
from PyQt5.QtCore import pyqtSignal, QRectF, QSize, QEvent, Qt, QPoint
from PyQt5.QtGui import QTextDocument, QPainter, QColor, QPen
from PyQt5.QtWidgets import QStyledItemDelegate, QStyleOptionViewItem, QStyle


class HtmlDelegate(QStyledItemDelegate):
    linkClicked = pyqtSignal(str)
    linkRightClicked = pyqtSignal(str, QPoint)  # <--- NEW: Signal for Right-Clicks

    def paint(self, painter, option, index):
        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        painter.save()
        doc = QTextDocument()
        doc.setDocumentMargin(4)
        doc.setDefaultFont(options.font)
        doc.setHtml(options.text)

        options.text = ""
        options.widget.style().drawControl(QStyle.CE_ItemViewItem, options, painter)

        painter.translate(options.rect.left(), options.rect.top())
        clip = QRectF(0, 0, options.rect.width(), options.rect.height())
        doc.drawContents(painter, clip)
        painter.restore()

    def sizeHint(self, option, index):
        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        doc = QTextDocument()
        doc.setDocumentMargin(4)
        doc.setDefaultFont(options.font)
        doc.setHtml(options.text)
        return QSize(int(doc.idealWidth()), int(doc.size().height()))

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease:
            options = QStyleOptionViewItem(option)
            self.initStyleOption(options, index)
            doc = QTextDocument()
            doc.setDocumentMargin(4)
            doc.setDefaultFont(options.font)
            doc.setHtml(options.text)

            pos = event.pos() - option.rect.topLeft()
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
        rect = option.rect.adjusted(4, 4, -4, -4)

        if state == 'EXISTS':
            bg_color, border_color, text_color = QColor("#E6F4E6"), QColor("#A3DDA3"), QColor("#0C6B0C")
            text = "✓ Open"
        elif state == 'LOADING':
            bg_color, border_color, text_color = QColor("#FFF4CE"), QColor("#F3C74C"), QColor("#B85C00")
            text = "⏳ Fetching"
        else:
            bg_color, border_color, text_color = QColor("#E1F0FF"), QColor("#99C9FF"), QColor("#005A9E")
            text = "⬇ Download"

        if has_revs and state != 'LOADING':
            text += " (+Rev)"

        painter.setBrush(bg_color)
        painter.setPen(QPen(border_color, 1))
        painter.drawRoundedRect(rect, 10, 10)

        painter.setPen(text_color)
        font = painter.font()
        font.setBold(True)
        font.setPointSize(8)
        painter.setFont(font)
        painter.drawText(rect, Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease:
            rect = option.rect.adjusted(4, 4, -4, -4)
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
            # Two-line: AI number top, Sub-Item pill bottom
            # 1. AI Number (top)
            ai_rect = QRectF(rect.left(), rect.top() + 3, rect.width(), 20)
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
            pill_h = 18
            pill_margin = 3
            pill_rect = QRectF(rect.left() + pill_margin, rect.top() + 24, rect.width() - (2 * pill_margin), pill_h)

            if pill_rect.width() > 16:
                bg_color = QColor("#F1F5F9") if not (options.state & QStyle.State_Selected) else QColor("#DBEAFE")
                border_color = QColor("#CBD5E1") if not (options.state & QStyle.State_Selected) else QColor("#93C5FD")
                text_color = QColor("#334155") if not (options.state & QStyle.State_Selected) else QColor("#1E40AF")

                painter.setBrush(bg_color)
                painter.setPen(QPen(border_color, 1))
                painter.drawRoundedRect(pill_rect, 4, 4)

                pill_font = painter.font()
                pill_font.setBold(False)
                if pill_font.pointSize() > 8:
                    pill_font.setPointSize(pill_font.pointSize() - 2)
                elif pill_font.pixelSize() > 10:
                    pill_font.setPixelSize(pill_font.pixelSize() - 2)
                else:
                    pill_font.setPointSize(8)
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
        tentative_color = "#64748B" if is_selected else "#94A3B8"

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

