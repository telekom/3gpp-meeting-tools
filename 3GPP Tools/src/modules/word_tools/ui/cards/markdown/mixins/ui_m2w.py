"""Markdown -> Word page UI builder mixin for WordMarkdownCard."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, CARD_FRAME_STYLE
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane


class MarkdownM2wUiMixin:
    """Constructs Page 1 (Markdown -> Word) controls and compilation result banner."""

    def _setup_page_m2w(self, container: QWidget):
        m2w_layout = QVBoxLayout(container)
        m2w_layout.setContentsMargins(0, 0, 0, 0)
        m2w_layout.setSpacing(10)

        self.pane_m2w = DocumentSelectorPane(
            "📝 MARKDOWN DOCUMENT TO CONVERT (.md)",
            accepted_extensions=[".md", ".markdown", ".txt", ".docx", ".doc"],
            drop_prompt="Drag & drop Markdown file (.md) here",
        )
        # Hide Word-COM-only controls on the Markdown input pane
        self.pane_m2w.quick_active_btn.setVisible(False)
        if hasattr(self.pane_m2w.tabs, "setTabVisible"):
            self.pane_m2w.tabs.setTabVisible(1, False)
        else:
            self.pane_m2w.tabs.setTabEnabled(1, False)
        self.pane_m2w.url_input.setPlaceholderText("https://... (Markdown .md file URL)")
        self.pane_m2w.file_selected.connect(self._on_m2w_file_selected)
        m2w_layout.addWidget(self.pane_m2w, 1)

        m2w_opts_frame = QFrame()
        m2w_opts_frame.setObjectName("cardFrame")
        m2w_opts_frame.setStyleSheet(CARD_FRAME_STYLE)
        opts_layout = QVBoxLayout(m2w_opts_frame)
        opts_layout.setContentsMargins(14, 12, 14, 12)
        opts_layout.setSpacing(8)

        hdr_lbl = QLabel("<b>Conversion Options:</b>")
        hdr_lbl.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 12px;")
        opts_layout.addWidget(hdr_lbl)

        grid = QGridLayout()
        grid.setHorizontalSpacing(24)
        grid.setVerticalSpacing(6)

        self.m2w_media_cb = QCheckBox("Auto-detect companion media folder ({stem}_media/)")
        self.m2w_media_cb.setChecked(True)
        self.m2w_media_cb.setToolTip(
            "Automatically embeds figures and drawings located in the companion media folder or relative paths."
        )
        grid.addWidget(self.m2w_media_cb, 0, 0)

        self.m2w_diagrams_cb = QCheckBox("Render PlantUML diagrams (```plantuml``` code blocks)")
        self.m2w_diagrams_cb.setChecked(True)
        self.m2w_diagrams_cb.setToolTip(
            "Compiles embedded PlantUML sequence and architecture code blocks into high-resolution figures on the fly."
        )
        grid.addWidget(self.m2w_diagrams_cb, 0, 1)

        self.m2w_styles_cb = QCheckBox("Apply 3GPP contribution styles (3gpp_70.dotx && TableGrid)")
        self.m2w_styles_cb.setChecked(True)
        self.m2w_styles_cb.setToolTip(
            "Injects canonical 3GPP headings, typography, margins, and TableGrid formatting into the OpenXML package."
        )
        grid.addWidget(self.m2w_styles_cb, 1, 0)

        self.m2w_auto_open_cb = QCheckBox("Auto-open in Word after conversion")
        self.m2w_auto_open_cb.setChecked(True)
        self.m2w_auto_open_cb.setToolTip(
            "Launches the generated document in Microsoft Word immediately upon completion."
        )
        grid.addWidget(self.m2w_auto_open_cb, 1, 1)

        opts_layout.addLayout(grid)

        dest_row = QHBoxLayout()
        self.m2w_custom_dest_cb = QCheckBox("Save to custom output directory:")
        self.m2w_custom_dest_cb.toggled.connect(self._on_m2w_custom_dest_toggled)
        self.m2w_dest_edit = QLineEdit()
        self.m2w_dest_edit.setEnabled(False)
        self.m2w_dest_edit.setPlaceholderText(
            "Defaults to same directory as source Markdown document"
        )
        self.m2w_dest_browse_btn = QPushButton("Browse...")
        self.m2w_dest_browse_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.m2w_dest_browse_btn.setEnabled(False)
        self.m2w_dest_browse_btn.clicked.connect(self._browse_m2w_dest)

        dest_row.addWidget(self.m2w_custom_dest_cb)
        dest_row.addWidget(self.m2w_dest_edit, 1)
        dest_row.addWidget(self.m2w_dest_browse_btn)
        opts_layout.addLayout(dest_row)

        m2w_layout.addWidget(m2w_opts_frame)

        act_row = QHBoxLayout()
        self.run_m2w_btn = QPushButton("🚀 Convert to Word (.docx)")
        self.run_m2w_btn.setObjectName("primaryBtn")
        self.run_m2w_btn.setCursor(Qt.PointingHandCursor)
        self.run_m2w_btn.clicked.connect(self._trigger_m2w_conversion)
        act_row.addWidget(self.run_m2w_btn)
        act_row.addStretch()
        m2w_layout.addLayout(act_row)

        self.m2w_progress = QProgressBar()
        self.m2w_progress.setRange(0, 0)
        self.m2w_progress.setVisible(False)
        self.m2w_progress.setStyleSheet("QProgressBar { max-height: 4px; }")
        m2w_layout.addWidget(self.m2w_progress)

        self.m2w_result_frame = QFrame()
        self.m2w_result_frame.setObjectName("cardFrame")
        self.m2w_result_frame.setStyleSheet(
            CARD_FRAME_STYLE
            + f"""
            QFrame#cardFrame {{
                background-color: {C.SUCCESS_SOFT};
                border: 1px solid {C.SUCCESS_BORDER_MEDIUM};
                border-radius: 6px;
                padding: 6px 12px;
            }}
        """
        )
        self.m2w_result_frame.setVisible(False)
        m2w_res_layout = QHBoxLayout(self.m2w_result_frame)
        m2w_res_layout.setContentsMargins(6, 6, 6, 6)
        m2w_res_layout.setSpacing(8)

        self.m2w_result_icon = QLabel("✅")
        self.m2w_result_lbl = QLabel()
        self.m2w_result_lbl.setStyleSheet(
            f"color: {C.SUCCESS_DARK}; font-size: 11px; font-weight: bold;"
        )
        m2w_res_layout.addWidget(self.m2w_result_icon)
        m2w_res_layout.addWidget(self.m2w_result_lbl, 1)

        self.m2w_open_folder_btn = QPushButton("📂 Open Folder")
        self.m2w_open_folder_btn.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.m2w_open_folder_btn.clicked.connect(self._open_m2w_folder)

        self.m2w_open_file_btn = QPushButton("👁️ Open in Word")
        self.m2w_open_file_btn.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.m2w_open_file_btn.clicked.connect(self._open_m2w_file)

        self.m2w_copy_path_btn = QPushButton("🔗 Copy Path")
        self.m2w_copy_path_btn.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.m2w_copy_path_btn.clicked.connect(self._copy_m2w_path_to_clipboard)

        m2w_res_layout.addWidget(self.m2w_open_folder_btn)
        m2w_res_layout.addWidget(self.m2w_open_file_btn)
        m2w_res_layout.addWidget(self.m2w_copy_path_btn)

        m2w_layout.addWidget(self.m2w_result_frame)
        m2w_layout.addStretch()
