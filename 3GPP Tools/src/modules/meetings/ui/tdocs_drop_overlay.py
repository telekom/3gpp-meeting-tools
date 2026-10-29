# --- File: src/modules/meetings/ui/tdocs_drop_overlay.py ---
"""
Semi-transparent visual overlay for drag & drop file ingestion across Meetings & Agendas.
Aligned with ModernDropZone Fluent design tokens.
"""

from typing import List, Optional

from PyQt5.QtCore import QRectF, Qt
from PyQt5.QtGui import QBrush, QColor, QPainter, QPainterPath, QPen
from PyQt5.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget


class _OverlayCard(QFrame):
    """Card widget that paints an anti-aliased rounded dashed border."""

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = QRectF(self.rect()).adjusted(2.0, 2.0, -2.0, -2.0)

        # Background fill
        painter.setBrush(QBrush(QColor(240, 247, 255, 245)))  # Blue 50 with 96% opacity

        # Anti-aliased dashed border
        pen = QPen(QColor("#2563EB"), 2.0, Qt.DashLine)
        pen.setDashPattern([8, 5])
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(pen)

        painter.drawRoundedRect(rect, 16.0, 16.0)


class DropOverlayWidget(QWidget):
    """Semi-transparent visual overlay aligned with ModernDropZone design tokens."""

    def __init__(
        self,
        parent=None,
        title: str = "Drop Chairman's Notes / Agenda file(s) to import",
        supported_exts: Optional[List[str]] = None,
    ):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.hide()

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(40, 40, 40, 40)

        card = _OverlayCard()
        card.setMinimumWidth(480)
        card.setMaximumWidth(680)

        card_layout = QVBoxLayout(card)
        card_layout.setAlignment(Qt.AlignCenter)
        card_layout.setContentsMargins(36, 28, 36, 28)
        card_layout.setSpacing(10)

        icon_lbl = QLabel("📥")
        icon_lbl.setStyleSheet("""
            QLabel {
                font-size: 36px;
                background-color: rgba(37, 99, 235, 0.12);
                border-radius: 30px;
                padding: 10px;
            }
        """)
        icon_lbl.setFixedSize(60, 60)
        icon_lbl.setAlignment(Qt.AlignCenter)

        text_lbl = QLabel(f"<b>{title}</b>")
        text_lbl.setStyleSheet("font-size: 15px; color: #0F172A; border: none; background: transparent;")
        text_lbl.setAlignment(Qt.AlignCenter)

        sub_lbl = QLabel("Supports single or multiple files (smart automatic extraction & parsing)")
        sub_lbl.setStyleSheet("font-size: 12px; color: #475569; border: none; background: transparent;")
        sub_lbl.setAlignment(Qt.AlignCenter)

        badges_row = QHBoxLayout()
        badges_row.setSpacing(6)
        badges_row.setAlignment(Qt.AlignCenter)
        for ext in (supported_exts or [".docx", ".doc", ".htm", ".zip"]):
            b = QLabel(ext)
            b.setStyleSheet("""
                QLabel {
                    background-color: #EFF6FF;
                    color: #1D4ED8;
                    border: 1px solid #BFDBFE;
                    border-radius: 4px;
                    font-size: 10px;
                    font-weight: 600;
                    padding: 2px 7px;
                }
            """)
            badges_row.addWidget(b)

        card_layout.addWidget(icon_lbl, 0, Qt.AlignCenter)
        card_layout.addWidget(text_lbl)
        card_layout.addWidget(sub_lbl)
        card_layout.addLayout(badges_row)

        layout.addWidget(card)

    def paintEvent(self, event):
        """Paint soft semi-transparent scrim across parent window."""
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(15, 23, 42, 85))  # Slate-900 at ~33% opacity

