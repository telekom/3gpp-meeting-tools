# --- File: src/modules/puml2visio/ui/select_presentation_dialog.py ---
"""Dialog allowing users to select which open PowerPoint presentation to append a PlantUML diagram slide to."""

import logging
from typing import Optional, List, Dict, Any

from PyQt5.QtCore import Qt, QSize
from PyQt5.QtWidgets import (
    QButtonGroup,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY

logger = logging.getLogger(__name__)

LIST_STYLE_CARDS = """
QListWidget#presentationList {
    background-color: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    outline: none;
    padding: 6px;
}
QListWidget#presentationList::item {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 6px;
    margin-bottom: 6px;
    min-height: 68px;
}
QListWidget#presentationList::item:hover {
    background-color: #FFFFFF;
    border: 1.5px solid #CBD5E1;
}
QListWidget#presentationList::item:selected {
    background-color: #FFF7ED;
    border: 1.5px solid #EA580C;
}
"""

PRIMARY_BTN_STYLE = """
QPushButton#primaryBtn {
    background-color: #D24726;
    color: #FFFFFF;
    font-weight: bold;
    border: 1px solid #B73D20;
    border-radius: 5px;
    padding: 7px 20px;
    font-size: 12px;
}
QPushButton#primaryBtn:hover {
    background-color: #B73D20;
}
QPushButton#primaryBtn:pressed {
    background-color: #9C3319;
}
QPushButton#primaryBtn:disabled {
    background-color: #E2E8F0;
    color: #94A3B8;
    border: 1px solid #CBD5E1;
}
"""


def _format_display_path(full_path: str, pres_name: str, max_chars: int = 46) -> str:
    """Format and shorten file paths cleanly to prevent card overflow."""
    if not full_path or full_path == pres_name:
        if "/" not in full_path and "\\" not in full_path:
            return "New presentation (not saved to disk yet)"
    if len(full_path) <= max_chars:
        return full_path
    prefix_len = 16
    suffix_len = max_chars - prefix_len - 3
    return full_path[:prefix_len] + "..." + full_path[-suffix_len:]


class PresentationItemWidget(QWidget):
    """Custom row card widget rendering presentation icon, title, slide count, path, and selection state."""

    def __init__(self, pres: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.pres = pres
        self._is_selected = False
        self._setup_ui()

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(12)

        # Left: Office PowerPoint slide badge (deep orange brand gradient)
        self.icon_badge = QLabel("P")
        self.icon_badge.setFixedSize(36, 36)
        self.icon_badge.setAlignment(Qt.AlignCenter)
        self.icon_badge.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #E05238, stop:1 #C43E1C);"
            "color: #FFFFFF; font-weight: 800; font-size: 17px; border-radius: 6px; "
            "font-family: 'Segoe UI', Arial, sans-serif;"
        )
        self.icon_badge.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        layout.addWidget(self.icon_badge, 0, Qt.AlignVCenter)

        # Center: Metadata Column (Top: Title + Badges, Bottom: Slides + Path)
        info_layout = QVBoxLayout()
        info_layout.setContentsMargins(0, 0, 0, 0)
        info_layout.setSpacing(4)

        # Top row: presentation name & status badges
        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.setSpacing(8)

        pres_name = self.pres.get("name", "Presentation.pptx")
        self.title_lbl = QLabel(pres_name)
        self.title_lbl.setStyleSheet("color: #0F172A; font-size: 13px; font-weight: 700;")
        self.title_lbl.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        top_row.addWidget(self.title_lbl)

        if self.pres.get("is_active"):
            active_badge = QLabel("⭐ Active Deck")
            active_badge.setStyleSheet(
                "background-color: #DCFCE7; color: #166534; font-weight: 600; "
                "font-size: 10px; padding: 2px 8px; border-radius: 10px; border: 1px solid #86EFAC;"
            )
            active_badge.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            top_row.addWidget(active_badge)

        if not self.pres.get("saved", True):
            unsaved_badge = QLabel("🟠 Unsaved")
            unsaved_badge.setStyleSheet(
                "background-color: #FEF3C7; color: #92400E; font-weight: 600; "
                "font-size: 10px; padding: 2px 7px; border-radius: 10px; border: 1px solid #FCD34D;"
            )
            unsaved_badge.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            top_row.addWidget(unsaved_badge)

        top_row.addStretch()
        info_layout.addLayout(top_row)

        # Bottom row: slide count and path
        bottom_row = QHBoxLayout()
        bottom_row.setContentsMargins(0, 0, 0, 0)
        bottom_row.setSpacing(6)

        slides = self.pres.get("slides_count", 0)
        slide_text = f"📑 {slides} slide{'s' if slides != 1 else ''}"
        self.slides_lbl = QLabel(slide_text)
        self.slides_lbl.setStyleSheet("color: #475569; font-size: 11px; font-weight: 600;")
        self.slides_lbl.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        bottom_row.addWidget(self.slides_lbl)

        dot_lbl = QLabel("•")
        dot_lbl.setStyleSheet("color: #CBD5E1; font-size: 11px;")
        dot_lbl.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        bottom_row.addWidget(dot_lbl)

        full_path = self.pres.get("full_name", "")
        display_path = _format_display_path(full_path, pres_name)

        self.path_lbl = QLabel(f"📁 {display_path}")
        self.path_lbl.setStyleSheet("color: #64748B; font-size: 11px;")
        self.path_lbl.setToolTip(full_path if full_path else "Unsaved presentation")
        self.path_lbl.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        bottom_row.addWidget(self.path_lbl)

        bottom_row.addStretch()
        info_layout.addLayout(bottom_row)

        layout.addLayout(info_layout, 1)

        # Right: Selection pill indicator
        self.status_pill = QLabel("○")
        self.status_pill.setStyleSheet("color: #CBD5E1; font-size: 13px; font-weight: normal; padding: 2px 6px;")
        self.status_pill.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        layout.addWidget(self.status_pill, 0, Qt.AlignVCenter)

    def set_selected(self, selected: bool):
        """Update visual selection indicator."""
        self._is_selected = selected
        if selected:
            self.status_pill.setText("✓ Selected")
            self.status_pill.setStyleSheet(
                "background-color: #FFEDD5; color: #C2410C; font-weight: bold; "
                "font-size: 10px; padding: 3px 9px; border-radius: 10px; border: 1px solid #FDBA74;"
            )
        else:
            self.status_pill.setText("○")
            self.status_pill.setStyleSheet(
                "color: #CBD5E1; font-size: 13px; font-weight: normal; padding: 2px 6px;"
            )

    def sizeHint(self) -> QSize:
        return QSize(480, 68)

    def minimumSizeHint(self) -> QSize:
        return QSize(220, 64)


class SelectPresentationDialog(QDialog):
    """Modal dialog prompting delegate to choose one of multiple open PowerPoint presentations."""

    def __init__(self, presentations: List[Dict[str, Any]], parent=None):
        super().__init__(parent)
        self.presentations = presentations or []
        self._selected_presentation: Optional[Dict[str, Any]] = None
        self._create_new: bool = False

        self.setWindowTitle("Select PowerPoint Presentation")
        # Remove old Windows help '?' button from dialog title bar
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setMinimumSize(560, 390)
        self.resize(575, 410)
        self.setModal(True)
        self._setup_ui()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 16, 18, 16)
        main_layout.setSpacing(11)

        # Header Frame with PowerPoint Tile & Refresh Button
        header_layout = QHBoxLayout()
        header_layout.setSpacing(12)

        header_icon = QLabel("P")
        header_icon.setFixedSize(38, 38)
        header_icon.setAlignment(Qt.AlignCenter)
        header_icon.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #E05238, stop:1 #C43E1C);"
            "color: #FFFFFF; font-weight: 800; font-size: 18px; border-radius: 8px; "
            "font-family: 'Segoe UI', Arial, sans-serif;"
        )
        header_layout.addWidget(header_icon, 0, Qt.AlignVCenter)

        header_text_layout = QVBoxLayout()
        header_text_layout.setSpacing(2)

        title_lbl = QLabel("Select Target PowerPoint Slide Deck")
        title_lbl.setStyleSheet("font-size: 15px; font-weight: bold; color: #0F172A;")

        desc_lbl = QLabel(
            "Multiple presentations are currently open in Microsoft PowerPoint. "
            "Choose which slide deck to append this diagram to:"
        )
        desc_lbl.setStyleSheet("font-size: 12px; color: #64748B;")
        desc_lbl.setWordWrap(True)

        header_text_layout.addWidget(title_lbl)
        header_text_layout.addWidget(desc_lbl)
        header_layout.addLayout(header_text_layout, 1)

        # Refresh button to re-scan running PowerPoint decks
        self.btn_refresh = QPushButton("🔄 Refresh")
        self.btn_refresh.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_refresh.setCursor(Qt.PointingHandCursor)
        self.btn_refresh.setToolTip("Re-scan Microsoft PowerPoint for newly opened or closed presentations")
        self.btn_refresh.clicked.connect(self._on_refresh_clicked)
        header_layout.addWidget(self.btn_refresh, 0, Qt.AlignTop)

        main_layout.addLayout(header_layout)

        # Filter bar (shown when >= 3 presentations exist)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("🔍 Filter presentations by name or path...")
        self.filter_edit.setClearButtonEnabled(True)
        self.filter_edit.setStyleSheet("""
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 5px;
                padding: 4px 8px;
                font-size: 12px;
            }
            QLineEdit:focus {
                border: 1px solid #EA580C;
            }
        """)
        self.filter_edit.textChanged.connect(self._on_filter_changed)
        main_layout.addWidget(self.filter_edit)
        self.filter_edit.setVisible(len(self.presentations) >= 3)

        # Presentation List Widget
        self.list_widget = QListWidget()
        self.list_widget.setObjectName("presentationList")
        self.list_widget.setStyleSheet(LIST_STYLE_CARDS)
        self.list_widget.setSelectionMode(QListWidget.SingleSelection)
        self.list_widget.currentItemChanged.connect(self._on_current_item_changed)
        self.list_widget.itemDoubleClicked.connect(self._on_item_double_clicked)
        main_layout.addWidget(self.list_widget, 1)

        # Slide Placement Frame (End vs After Active Slide)
        placement_frame = QFrame()
        placement_frame.setObjectName("placementFrame")
        placement_frame.setStyleSheet("""
            QFrame#placementFrame {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
            }
            QLabel, QRadioButton {
                border: none;
                background: transparent;
            }
        """)
        placement_layout = QHBoxLayout(placement_frame)
        placement_layout.setContentsMargins(12, 7, 12, 7)
        placement_layout.setSpacing(14)

        place_lbl = QLabel("<b>Slide Placement:</b>")
        place_lbl.setStyleSheet("color: #334155; font-size: 11px;")
        placement_layout.addWidget(place_lbl)

        self.radio_group = QButtonGroup(self)
        self.radio_end = QRadioButton("Append to end of presentation")
        self.radio_end.setChecked(True)
        self.radio_end.setStyleSheet("font-size: 11px; color: #1E293B;")
        self.radio_group.addButton(self.radio_end)
        placement_layout.addWidget(self.radio_end)

        self.radio_after_active = QRadioButton("Insert after active slide")
        self.radio_after_active.setStyleSheet("font-size: 11px; color: #1E293B;")
        self.radio_group.addButton(self.radio_after_active)
        placement_layout.addWidget(self.radio_after_active)
        placement_layout.addStretch()

        main_layout.addWidget(placement_frame)

        # Informational Callout Frame (Exported Slide Specifications)
        info_frame = QFrame()
        info_frame.setObjectName("infoFrame")
        info_frame.setStyleSheet("""
            QFrame#infoFrame {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)
        info_layout = QHBoxLayout(info_frame)
        info_layout.setContentsMargins(12, 7, 12, 7)
        info_layout.setSpacing(8)

        info_lbl = QLabel(
            "💡 <b>Slide Output:</b> Native editable Office shapes, auto 16:9/4:3 scaling, and PlantUML preserved in notes."
        )
        info_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        info_layout.addWidget(info_lbl)
        info_layout.addStretch()

        main_layout.addWidget(info_frame)

        # Footer Buttons
        footer_layout = QHBoxLayout()
        footer_layout.setSpacing(8)

        self.btn_new_pres = QPushButton("+ New Presentation Instead")
        self.btn_new_pres.setStyleSheet(
            "QPushButton { background-color: #F8FAFC; border: 1px solid #CBD5E1; color: #1E293B; "
            "font-size: 11px; padding: 6px 12px; border-radius: 4px; font-weight: 500; } "
            "QPushButton:hover { background-color: #F1F5F9; border-color: #94A3B8; }"
        )
        self.btn_new_pres.setCursor(Qt.PointingHandCursor)
        self.btn_new_pres.setToolTip("Export diagram to a fresh standalone presentation instead of appending")
        self.btn_new_pres.clicked.connect(self._on_new_presentation_clicked)
        footer_layout.addWidget(self.btn_new_pres)

        footer_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.clicked.connect(self.reject)
        footer_layout.addWidget(cancel_btn)

        self.add_btn = QPushButton("+ Add Slide to Deck")
        self.add_btn.setObjectName("primaryBtn")
        self.add_btn.setStyleSheet(PRIMARY_BTN_STYLE)
        self.add_btn.setCursor(Qt.PointingHandCursor)
        self.add_btn.setDefault(True)
        self.add_btn.clicked.connect(self._on_accept_clicked)
        footer_layout.addWidget(self.add_btn)

        main_layout.addLayout(footer_layout)

        # Populate cards
        self._populate_list()

    def _populate_list(self):
        self.list_widget.clear()
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

        has_presentations = len(self.presentations) > 0
        self.add_btn.setEnabled(has_presentations)

        if has_presentations:
            self.list_widget.setCurrentRow(active_idx)
            curr = self.list_widget.currentItem()
            if curr:
                self._selected_presentation = curr.data(Qt.UserRole)
                self._update_selection_indicator(curr)

    def _on_current_item_changed(self, current: Optional[QListWidgetItem], previous: Optional[QListWidgetItem]):
        if current:
            self._selected_presentation = current.data(Qt.UserRole)
        self._update_selection_indicator(current)

    def _update_selection_indicator(self, current: Optional[QListWidgetItem]):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            widget = self.list_widget.itemWidget(item)
            if isinstance(widget, PresentationItemWidget):
                widget.set_selected(item == current)

        if current:
            pres = current.data(Qt.UserRole) or {}
            slides = pres.get("slides_count", 0)
            self.radio_end.setText(f"Append at end of presentation (as Slide {slides + 1})")
            self.add_btn.setEnabled(True)
        else:
            self.radio_end.setText("Append to end of presentation")
            self.add_btn.setEnabled(False)

    def _on_filter_changed(self, text: str):
        query = text.strip().lower()
        first_visible = None
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            pres = item.data(Qt.UserRole) or {}
            name = pres.get("name", "").lower()
            path = pres.get("full_name", "").lower()
            match = not query or (query in name or query in path)
            item.setHidden(not match)
            if match and first_visible is None:
                first_visible = item

        curr = self.list_widget.currentItem()
        if first_visible and (curr is None or curr.isHidden()):
            self.list_widget.setCurrentItem(first_visible)

    def _on_refresh_clicked(self):
        try:
            from modules.puml2visio.core.powerpoint_converter import get_open_powerpoint_presentations
            fresh = get_open_powerpoint_presentations()
            self.presentations = fresh
            self.filter_edit.setVisible(len(self.presentations) >= 3)
            self._populate_list()
        except Exception as e:
            logger.warning(f"Failed to refresh PowerPoint presentations: {e}")

    def _on_new_presentation_clicked(self):
        """User chose to create a new standalone PowerPoint presentation instead of appending."""
        self._create_new = True
        self.accept()

    def _on_item_double_clicked(self, item: QListWidgetItem):
        self._on_accept_clicked()

    def _on_accept_clicked(self):
        curr_item = self.list_widget.currentItem()
        if curr_item:
            self._selected_presentation = curr_item.data(Qt.UserRole)
        self.accept()

    def is_new_presentation_requested(self) -> bool:
        """Return True if user explicitly requested exporting to a new presentation instead."""
        return self._create_new

    def get_selected_presentation(self) -> Optional[Dict[str, Any]]:
        """Return the user-selected presentation dictionary with 'insert_at' preference, or None."""
        pres = self._selected_presentation
        if not pres:
            curr_item = self.list_widget.currentItem()
            if curr_item:
                pres = curr_item.data(Qt.UserRole)
        if not pres:
            return None
        res = dict(pres)
        res["insert_at"] = "after_active" if self.radio_after_active.isChecked() else "end"
        return res
