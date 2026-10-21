# --- File: src/modules/meetings/ui/tdocs_drop_overlay.py ---
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QFrame, QLabel, QVBoxLayout, QWidget


class DropOverlayWidget(QWidget):
    """Semi-transparent visual overlay aligned with InteractiveDropLabel design tokens."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.hide()

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: rgba(235, 243, 252, 0.95);
                border: 3px dashed #395396;
                border-radius: 12px;
                padding: 30px;
            }
        """)
        card_layout = QVBoxLayout(card)
        card_layout.setAlignment(Qt.AlignCenter)
        card_layout.setSpacing(10)

        icon_lbl = QLabel("📥")
        icon_lbl.setStyleSheet("font-size: 48px; border: none; background: transparent;")
        icon_lbl.setAlignment(Qt.AlignCenter)

        text_lbl = QLabel("<b>Drop Chairman's Notes / Agenda file(s) to import</b><br>"
                          "<span style='font-size: 12px; color: #475569;'>"
                          "Supports single or multiple <code>.docx</code>, <code>.doc</code>, and <code>.htm</code> files</span>")
        text_lbl.setStyleSheet("font-size: 16px; color: #1E5C99; border: none; background: transparent;")
        text_lbl.setAlignment(Qt.AlignCenter)

        card_layout.addWidget(icon_lbl)
        card_layout.addWidget(text_lbl)
        layout.addWidget(card)
