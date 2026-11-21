"""Stacked empty and loaded card page builders for ModernDropZone."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from .styles import (
    CLEAR_BTN_STYLE,
    DROP_ZONE_ACTION_BTN_STYLE,
    EMPTY_ICON_CANVAS_STYLE,
    EMPTY_ICON_COMPACT_STYLE,
    EXT_BADGE_STYLE,
    FILE_SIZE_BADGE_STYLE,
    SENSITIVITY_BADGE_STYLE,
    SUBTITLE_CANVAS_STYLE,
    SUBTITLE_COMPACT_STYLE,
    TITLE_LABEL_STYLE,
)


class _DropZonePagesMixin:
    """Builds the stacked empty prompt page and loaded document card page."""

    def _init_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self.stack = QStackedWidget(self)
        self.stack.setStyleSheet("background: transparent; border: none;")

        # Page 0: Empty / Drop Prompt Page
        self.empty_page = self._create_empty_page()
        self.stack.addWidget(self.empty_page)

        # Page 1: File Loaded / Ready Card Page
        self.loaded_page = self._create_loaded_page()
        self.stack.addWidget(self.loaded_page)

        root_layout.addWidget(self.stack)

        if self.compact:
            self.setFixedHeight(60)
        else:
            # Give the canvas drop zone enough height so icon, multi-line title,
            # subtitle, and extension pills never collide on high-DPI laptop screens.
            self.setMinimumHeight(140)
            self.setMaximumHeight(260)

    def _create_empty_page(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet("background: transparent;")

        if self.compact:
            # Compact Ribbon Layout (Horizontal)
            layout = QHBoxLayout(page)
            layout.setContentsMargins(14, 8, 14, 8)
            layout.setSpacing(12)

            self.empty_icon_lbl = QLabel("📥")
            self.empty_icon_lbl.setStyleSheet(EMPTY_ICON_COMPACT_STYLE)
            self.empty_icon_lbl.setFixedSize(36, 36)
            self.empty_icon_lbl.setAlignment(Qt.AlignCenter)
            layout.addWidget(self.empty_icon_lbl)

            text_col = QVBoxLayout()
            text_col.setSpacing(2)
            text_col.setAlignment(Qt.AlignVCenter)

            self.lbl_empty_title = QLabel(self.primary_title)
            self.lbl_empty_title.setStyleSheet(TITLE_LABEL_STYLE)
            text_col.addWidget(self.lbl_empty_title)

            exts_hint = f"Supports {', '.join(self.accepted_extensions)}" if self.accepted_extensions else ""
            sub_text = (
                f"{exts_hint} • {self.subtitle}"
                if (exts_hint and self.subtitle)
                else (exts_hint or self.subtitle or "")
            )
            self.lbl_empty_sub = QLabel(sub_text)
            self.lbl_empty_sub.setStyleSheet(SUBTITLE_COMPACT_STYLE)
            text_col.addWidget(self.lbl_empty_sub)

            layout.addLayout(text_col, 1)

            self.btn_empty_browse = QPushButton("📂 Browse...")
            self.btn_empty_browse.setStyleSheet(DROP_ZONE_ACTION_BTN_STYLE)
            self.btn_empty_browse.setCursor(Qt.PointingHandCursor)
            self.btn_empty_browse.clicked.connect(self._on_browse_clicked)
            layout.addWidget(self.btn_empty_browse)

        else:
            # Canvas Stacked Layout (Spacious)
            layout = QVBoxLayout(page)
            layout.setContentsMargins(18, 16, 18, 16)
            layout.setSpacing(6)
            layout.setAlignment(Qt.AlignVCenter)

            self.empty_icon_lbl = QLabel("📥")
            self.empty_icon_lbl.setStyleSheet(EMPTY_ICON_CANVAS_STYLE)
            self.empty_icon_lbl.setFixedSize(36, 36)
            self.empty_icon_lbl.setAlignment(Qt.AlignCenter)
            layout.addWidget(self.empty_icon_lbl, 0, Qt.AlignHCenter)

            self.lbl_empty_title = QLabel(self.primary_title)
            self.lbl_empty_title.setStyleSheet(TITLE_LABEL_STYLE)
            self.lbl_empty_title.setWordWrap(True)
            self.lbl_empty_title.setAlignment(Qt.AlignCenter)
            layout.addWidget(self.lbl_empty_title)

            self.lbl_empty_sub = QLabel(self.subtitle or "")
            self.lbl_empty_sub.setStyleSheet(SUBTITLE_CANVAS_STYLE)
            self.lbl_empty_sub.setWordWrap(True)
            self.lbl_empty_sub.setAlignment(Qt.AlignCenter)
            layout.addWidget(self.lbl_empty_sub)

            # Optional format badges pill row
            if self.accepted_extensions:
                badges_row = QHBoxLayout()
                badges_row.setSpacing(5)
                badges_row.setContentsMargins(0, 2, 0, 0)
                badges_row.setAlignment(Qt.AlignCenter)
                for ext in self.accepted_extensions[:5]:
                    badge = QLabel(ext)
                    badge.setStyleSheet(EXT_BADGE_STYLE)
                    badges_row.addWidget(badge)
                layout.addLayout(badges_row)

        return page

    def _create_loaded_page(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet("background: transparent;")

        layout = QHBoxLayout(page)
        layout.setContentsMargins(14, 8, 14, 8)
        layout.setSpacing(12)

        # File Archetype Icon Badge
        self.loaded_icon_lbl = QLabel("📄")
        self.loaded_icon_lbl.setFixedSize(40, 40)
        self.loaded_icon_lbl.setAlignment(Qt.AlignCenter)
        self.loaded_icon_lbl.setStyleSheet("font-size: 20px; border-radius: 8px;")
        layout.addWidget(self.loaded_icon_lbl)

        # File Details (Title + Metadata row)
        details_col = QVBoxLayout()
        details_col.setSpacing(3)
        details_col.setAlignment(Qt.AlignVCenter)

        self.lbl_file_name = QLabel()
        self.lbl_file_name.setStyleSheet(TITLE_LABEL_STYLE)
        details_col.addWidget(self.lbl_file_name)

        meta_row = QHBoxLayout()
        meta_row.setSpacing(6)
        meta_row.setAlignment(Qt.AlignLeft)

        self.lbl_file_size = QLabel()
        self.lbl_file_size.setStyleSheet(FILE_SIZE_BADGE_STYLE)
        meta_row.addWidget(self.lbl_file_size)

        self.lbl_file_desc = QLabel()
        self.lbl_file_desc.setStyleSheet(SUBTITLE_COMPACT_STYLE)
        meta_row.addWidget(self.lbl_file_desc)

        self.lbl_sensitivity = QLabel()
        self.lbl_sensitivity.setStyleSheet(SENSITIVITY_BADGE_STYLE)
        self.lbl_sensitivity.hide()
        meta_row.addWidget(self.lbl_sensitivity)

        details_col.addLayout(meta_row)
        layout.addLayout(details_col, 1)

        # Action Buttons: Change & Clear
        actions_row = QHBoxLayout()
        actions_row.setSpacing(6)
        actions_row.setAlignment(Qt.AlignVCenter)

        self.btn_change = QPushButton("🔄 Replace")
        self.btn_change.setStyleSheet(DROP_ZONE_ACTION_BTN_STYLE)
        self.btn_change.setCursor(Qt.PointingHandCursor)
        self.btn_change.clicked.connect(self._on_browse_clicked)
        actions_row.addWidget(self.btn_change)

        self.btn_clear = QPushButton("✕")
        self.btn_clear.setToolTip("Clear selected document")
        self.btn_clear.setFixedSize(28, 28)
        self.btn_clear.setStyleSheet(CLEAR_BTN_STYLE)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.clicked.connect(self.clear)
        actions_row.addWidget(self.btn_clear)

        layout.addLayout(actions_row)
        return page
