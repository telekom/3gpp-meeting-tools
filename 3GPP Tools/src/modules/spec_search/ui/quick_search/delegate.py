"""Result list delegate and HTML snippet formatting for QuickSpecSearchPanel."""

from __future__ import annotations

import html
import re
from typing import List

from PyQt5.QtCore import QRectF, QSize, Qt
from PyQt5.QtGui import QAbstractTextDocumentLayout, QColor, QPen, QTextDocument
from PyQt5.QtWidgets import QStyle, QStyledItemDelegate

from core.ui.font_scale import scale_px
from core.ui.theme import C, R

_ROLE_HTML = Qt.UserRole + 10
_ROLE_META = Qt.UserRole + 11


class _QuickSearchResultDelegate(QStyledItemDelegate):
    """Lightweight HTML delegate rendering citation header and <mark> snippet."""

    def paint(self, painter, option, index) -> None:
        painter.save()
        rect = option.rect
        selected = bool(option.state & QStyle.State_Selected)
        hovered = bool(option.state & QStyle.State_MouseOver)

        if selected:
            bg = C.ACCENT_SOFT
            border = C.ACCENT_PRIMARY
        elif hovered:
            bg = C.BG_SURFACE
            border = C.BORDER_SUBTLE
        else:
            bg = C.BG_CARD
            border = C.BORDER_SUBTLE

        painter.setPen(Qt.NoPen)
        painter.setBrush(option.palette.base())
        card_rect = rect.adjusted(4, 3, -4, -3)

        painter.setBrush(QColor(bg))
        painter.setPen(QPen(QColor(border), 1))
        painter.drawRoundedRect(card_rect, R.MD, R.MD)

        html_text = index.data(_ROLE_HTML) or ""
        doc = QTextDocument()
        doc.setDocumentMargin(0)
        doc.setDefaultFont(option.font)
        doc.setTextWidth(max(120, card_rect.width() - 16))
        doc.setHtml(html_text)

        painter.translate(card_rect.left() + 8, card_rect.top() + 6)
        ctx = QAbstractTextDocumentLayout.PaintContext()
        clip = QRectF(0, 0, card_rect.width() - 16, card_rect.height() - 12)
        painter.setClipRect(clip)
        doc.documentLayout().draw(painter, ctx)
        painter.restore()

    def sizeHint(self, option, index) -> QSize:
        html_text = index.data(_ROLE_HTML) or ""
        doc = QTextDocument()
        doc.setDocumentMargin(0)
        doc.setDefaultFont(option.font)
        width = max(240, option.rect.width() - 24) if option.rect.width() > 60 else 320
        doc.setTextWidth(width)
        doc.setHtml(html_text)
        height = int(doc.size().height()) + scale_px(18)
        return QSize(width, max(scale_px(54), min(scale_px(130), height)))


def _format_snippet_html(snippet: str) -> str:
    """Escape snippet HTML safely while styling FTS5 <mark>...</mark> hits."""
    if not snippet:
        return ""
    tokens = re.split(r"(<mark>.*?</mark>)", str(snippet), flags=re.IGNORECASE | re.DOTALL)
    parts: List[str] = []
    for tok in tokens:
        low = tok.lower()
        if low.startswith("<mark>") and low.endswith("</mark>"):
            inner = html.escape(tok[6:-7])
            parts.append(
                f'<span style="background-color:{C.WARNING_SOFT};color:{C.TEXT_STRONG};'
                f'font-weight:600;border-radius:2px;">{inner}</span>'
            )
        else:
            parts.append(html.escape(tok))
    return "".join(parts)
