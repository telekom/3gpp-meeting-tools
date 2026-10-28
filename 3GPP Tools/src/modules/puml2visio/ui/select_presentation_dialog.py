# --- File: src/modules/puml2visio/ui/select_presentation_dialog.py ---
"""Dialog allowing users to select which open PowerPoint presentation to append a PlantUML diagram slide to."""

import logging
from typing import Optional, List, Dict, Any

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QListWidget,
    QListWidgetItem,
    QFrame,
    QWidget,
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY

logger = logging.getLogger(__name__)

LIST_STYLE_CLEAN = """
QListWidget {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    font-size: 12px;
    outline: none;
    padding: 4px;
}
QListWidget::item {
    min-height: 52px;
    padding: 4px;
    border-radius: 4px;
    border-bottom: 1px solid #F1F5F9;
}
QListWidget::item:selected {
    background-color: #EBF3FC;
    color: #0F172A;
    border: 1px solid #93C5FD;
}
QListWidget::item:hover:!selected {
    background-color: #F8FAFC;
}
"""


class PresentationItemWidget(QWidget):
    """Custom row widget rendering presentation title, slide count, and active status."""

    def __init__(self, pres: Dict[str, Any], parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(3)

        # Top row: icon + title + [Active Deck] badge
        top_row = QHBoxLayout()
        top_row.setSpacing(6)

        title_lbl = QLabel(f"📊 <b>{pres.get('name', 'Presentation')}</b>")
        title_lbl.setStyleSheet("color: #0F172A; font-size: 13px;")
        top_row.addWidget(title_lbl)

        if pres.get("is_active"):
            active_badge = QLabel("Active Deck")
            active_badge.setStyleSheet(
                "background-color: #DCFCE7; color: #166534; font-weight: 600; "
                "font-size: 10px; padding: 1px 7px; border-radius: 8px; border: 1px solid #86EFAC;"
            )
            top_row.addWidget(active_badge)

        top_row.addStretch()
        layout.addLayout(top_row)

        # Bottom row: slide count and path
        slides = pres.get("slides_count", 0)
        slide_text = f"{slides} slide{'s' if slides != 1 else ''}"
        full_path = pres.get("full_name", "")
        if len(full_path) > 60:
            full_path = "..." + full_path[-57:]

        detail_text = f"{slide_text}  •  {full_path}"
        detail_lbl = QLabel(detail_text)
        detail_lbl.setStyleSheet("color: #64748B; font-size: 11px;")
        layout.addWidget(detail_lbl)


class SelectPresentationDialog(QDialog):
    """Modal dialog prompting delegate to choose one of multiple open PowerPoint presentations."""

    def __init__(self, presentations: List[Dict[str, Any]], parent=None):
        super().__init__(parent)
        self.presentations = presentations or []
        self._selected_presentation: Optional[Dict[str, Any]] = None

        self.setWindowTitle("Select PowerPoint Presentation")
        self.resize(520, 380)
        self.setModal(True)
        self._setup_ui()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 18, 18, 18)
        main_layout.setSpacing(12)

        # Header
        header_layout = QVBoxLayout()
        header_layout.setSpacing(4)
        title_lbl = QLabel("Select Target PowerPoint Slide Deck")
        title_lbl.setStyleSheet("font-size: 15px; font-weight: bold; color: #0F172A;")
        desc_lbl = QLabel(
            "Multiple presentations are currently open in Microsoft PowerPoint. "
            "Choose which slide deck to append this diagram to:"
        )
        desc_lbl.setStyleSheet("font-size: 12px; color: #475569;")
        desc_lbl.setWordWrap(True)
        header_layout.addWidget(title_lbl)
        header_layout.addWidget(desc_lbl)
        main_layout.addLayout(header_layout)

        # List Widget
        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet(LIST_STYLE_CLEAN)
        self.list_widget.setSelectionMode(QListWidget.SingleSelection)

        active_idx = 0
        for idx, pres in enumerate(self.presentations):
            item = QListWidgetItem(self.list_widget)
            widget = PresentationItemWidget(pres)
            item.setSizeHint(widget.sizeHint())
            item.setData(Qt.UserRole, pres)
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, widget)

            if pres.get("is_active"):
                active_idx = idx

        if self.presentations:
            self.list_widget.setCurrentRow(active_idx)

        self.list_widget.itemDoubleClicked.connect(self._on_item_double_clicked)
        main_layout.addWidget(self.list_widget)

        # Informational Note
        info_frame = QFrame()
        info_frame.setStyleSheet("""
            QFrame {
                background-color: #F0F9FF;
                border: 1px solid #BAE6FD;
                border-radius: 6px;
            }
        """)
        info_layout = QHBoxLayout(info_frame)
        info_layout.setContentsMargins(10, 8, 10, 8)
        info_layout.setSpacing(8)

        info_icon = QLabel("ℹ️")
        info_icon.setStyleSheet("font-size: 14px;")
        info_text = QLabel(
            "The diagram will be added as a new blank slide at the end of the presentation with "
            "native editable shapes, automatic aspect ratio scaling, and embedded source code in Speaker Notes."
        )
        info_text.setWordWrap(True)
        info_text.setStyleSheet("color: #0369A1; font-size: 11px;")
        info_layout.addWidget(info_icon, 0, Qt.AlignTop)
        info_layout.addWidget(info_text, 1)
        main_layout.addWidget(info_frame)

        # Footer Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.clicked.connect(self.reject)

        self.add_btn = QPushButton("➕ Add Slide")
        self.add_btn.setObjectName("primaryBtn")
        self.add_btn.setStyleSheet(
            "background-color: #1E5C99; color: white; font-weight: bold; "
            "padding: 6px 16px; border-radius: 4px; font-size: 12px;"
        )
        self.add_btn.setCursor(Qt.PointingHandCursor)
        self.add_btn.setDefault(True)
        self.add_btn.clicked.connect(self._on_accept_clicked)

        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(self.add_btn)
        main_layout.addLayout(btn_layout)

    def _on_item_double_clicked(self, item: QListWidgetItem):
        self._on_accept_clicked()

    def _on_accept_clicked(self):
        curr_item = self.list_widget.currentItem()
        if curr_item:
            self._selected_presentation = curr_item.data(Qt.UserRole)
        self.accept()

    def get_selected_presentation(self) -> Optional[Dict[str, Any]]:
        """Return the user-selected presentation dictionary, or None."""
        return self._selected_presentation
