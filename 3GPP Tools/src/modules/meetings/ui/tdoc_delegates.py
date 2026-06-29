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