"""Word -> Markdown page UI builder mixin for WordMarkdownCard."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, CARD_FRAME_STYLE
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane


class MarkdownW2mUiMixin:
    """Constructs Page 0 (Word -> Markdown) controls and result preview card."""

    def _setup_page_w2m(self, container: QWidget):
        md_layout = QVBoxLayout(container)
        md_layout.setContentsMargins(0, 0, 0, 0)
        md_layout.setSpacing(10)

        self.pane_markdown = DocumentSelectorPane(
            "📄 DOCUMENT / PRESENTATION TO CONVERT (.docx / .doc / .pptx / .ppt)",
            accepted_extensions=[
                ".docx",
                ".doc",
                ".docm",
                ".pptx",
                ".ppt",
                ".pdf",
                ".md",
                ".markdown",
            ],
            drop_prompt="Drag & drop Word document, PowerPoint deck, or PDF here",
        )
        self.pane_markdown.file_selected.connect(self._on_w2m_file_selected)
        md_layout.addWidget(self.pane_markdown, 1)

        md_options_frame = QFrame()
        md_options_frame.setObjectName("cardFrame")
        md_options_frame.setStyleSheet(CARD_FRAME_STYLE)
        opts_layout = QVBoxLayout(md_options_frame)
        opts_layout.setContentsMargins(14, 12, 14, 12)
        opts_layout.setSpacing(8)

        hdr_lbl = QLabel("<b>Conversion Options:</b>")
        hdr_lbl.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 12px;")
        opts_layout.addWidget(hdr_lbl)

        grid = QGridLayout()
        grid.setHorizontalSpacing(24)
        grid.setVerticalSpacing(6)

        tc_box = QHBoxLayout()
        tc_box.setSpacing(6)
        tc_lbl = QLabel("Track Changes:")
        tc_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_BODY};")
        self.md_revisions_combo = QComboBox()
        self.md_revisions_combo.addItem("Inline Markup [INSERTED / DELETED]", "markup")
        self.md_revisions_combo.addItem("Accept All Changes (Clean Final View)", "accept")
        self.md_revisions_combo.addItem("Reject All Changes (Original Baseline)", "reject")
        tc_box.addWidget(tc_lbl)
        tc_box.addWidget(self.md_revisions_combo, 1)
        grid.addLayout(tc_box, 0, 0)

        self.md_media_cb = QCheckBox("Extract figures and drawings to media/ folder")
        self.md_media_cb.setChecked(True)
        self.md_media_cb.setToolTip(
            "Extracts embedded images, converts vector EMF/WMF drawings to PNG, and links Visio (.vsdx) sources."
        )
        grid.addWidget(self.md_media_cb, 0, 1)

        self.md_header_cb = QCheckBox("Include summary header")
        self.md_header_cb.setChecked(True)
        self.md_header_cb.setToolTip(
            "Prepends document metadata (author, date, word count) to the generated markdown."
        )
        grid.addWidget(self.md_header_cb, 1, 0)

        opts_layout.addLayout(grid)

        dest_row = QHBoxLayout()
        self.md_custom_dest_cb = QCheckBox("Save to custom output directory:")
        self.md_custom_dest_cb.toggled.connect(self._on_md_custom_dest_toggled)
        self.md_dest_edit = QLineEdit()
        self.md_dest_edit.setEnabled(False)
        self.md_dest_edit.setPlaceholderText("Defaults to same directory as source Word document")
        self.md_dest_browse_btn = QPushButton("Browse...")
        self.md_dest_browse_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.md_dest_browse_btn.setEnabled(False)
        self.md_dest_browse_btn.clicked.connect(self._browse_md_dest)

        dest_row.addWidget(self.md_custom_dest_cb)
        dest_row.addWidget(self.md_dest_edit, 1)
        dest_row.addWidget(self.md_dest_browse_btn)
        opts_layout.addLayout(dest_row)

        md_layout.addWidget(md_options_frame)

        act_row = QHBoxLayout()
        self.run_markdown_btn = QPushButton("🚀 Convert to Markdown (.md)")
        self.run_markdown_btn.setObjectName("primaryBtn")
        self.run_markdown_btn.setCursor(Qt.PointingHandCursor)
        self.run_markdown_btn.clicked.connect(self._trigger_markdown_conversion)
        act_row.addWidget(self.run_markdown_btn)
        act_row.addStretch()
        md_layout.addLayout(act_row)

        self.md_progress = QProgressBar()
        self.md_progress.setRange(0, 0)
        self.md_progress.setVisible(False)
        self.md_progress.setStyleSheet("QProgressBar { max-height: 4px; }")
        md_layout.addWidget(self.md_progress)

        self.md_result_frame = QFrame()
        self.md_result_frame.setObjectName("cardFrame")
        self.md_result_frame.setStyleSheet(
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
        self.md_result_frame.setVisible(False)
        res_layout = QVBoxLayout(self.md_result_frame)
        res_layout.setContentsMargins(6, 6, 6, 6)
        res_layout.setSpacing(6)

        res_top_row = QHBoxLayout()
        self.md_result_icon = QLabel("✅")
        self.md_result_lbl = QLabel()
        self.md_result_lbl.setStyleSheet(
            f"color: {C.SUCCESS_DARK}; font-size: 11px; font-weight: bold;"
        )
        res_top_row.addWidget(self.md_result_icon)
        res_top_row.addWidget(self.md_result_lbl, 1)

        self.md_open_folder_btn = QPushButton("📂 Open Folder")
        self.md_open_folder_btn.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.md_open_folder_btn.clicked.connect(self._open_md_folder)

        self.md_open_file_btn = QPushButton("👁️ View File")
        self.md_open_file_btn.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.md_open_file_btn.clicked.connect(self._open_md_file)

        self.md_copy_btn = QPushButton("📋 Copy Markdown")
        self.md_copy_btn.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.md_copy_btn.clicked.connect(self._copy_md_to_clipboard)

        self.md_preview_toggle_btn = QPushButton("📝 Show Preview ▾")
        self.md_preview_toggle_btn.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.md_preview_toggle_btn.clicked.connect(self._toggle_md_preview)

        res_top_row.addWidget(self.md_open_folder_btn)
        res_top_row.addWidget(self.md_open_file_btn)
        res_top_row.addWidget(self.md_copy_btn)
        res_top_row.addWidget(self.md_preview_toggle_btn)
        res_layout.addLayout(res_top_row)

        self.md_preview_edit = QPlainTextEdit()
        self.md_preview_edit.setReadOnly(True)
        self.md_preview_edit.setMaximumHeight(200)
        self.md_preview_edit.setStyleSheet(
            "font-family: Consolas, monospace; font-size: 11px; background: white;"
        )
        self.md_preview_edit.setVisible(False)
        res_layout.addWidget(self.md_preview_edit)

        md_layout.addWidget(self.md_result_frame)
        md_layout.addStretch()
