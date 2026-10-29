"""Card: Word to Markdown Conversion."""

import logging
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
)
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane

logger = logging.getLogger(__name__)


class WordMarkdownCard(QWidget):
    """Card: Fast headless conversion of Word documents to Markdown with revision tracking & media extraction."""

    convert_markdown_requested = pyqtSignal(str, dict)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._last_md_output_path: Optional[str] = None
        self._last_converted_md: Optional[str] = None
        self._setup_ui()
        self.pane_md = self.pane_markdown

    def _setup_ui(self):
        md_layout = QVBoxLayout(self)
        md_layout.setContentsMargins(10, 10, 10, 10)
        md_layout.setSpacing(10)

        self.pane_markdown = DocumentSelectorPane("📄 WORD DOCUMENT TO CONVERT")
        md_layout.addWidget(self.pane_markdown)

        md_options_frame = QFrame()
        md_options_frame.setObjectName("cardFrame")
        md_options_frame.setStyleSheet(CARD_FRAME_STYLE)
        opts_layout = QVBoxLayout(md_options_frame)
        opts_layout.setContentsMargins(12, 10, 12, 10)
        opts_layout.setSpacing(8)

        opts_row1 = QHBoxLayout()
        opts_row1.addWidget(QLabel("<b>Track Changes:</b>"))
        self.md_revisions_combo = QComboBox()
        self.md_revisions_combo.addItem("Inline Markup [INSERTED / DELETED]", "markup")
        self.md_revisions_combo.addItem("Accept All Changes (Clean Final View)", "accept")
        self.md_revisions_combo.addItem("Reject All Changes (Original Baseline)", "reject")
        opts_row1.addWidget(self.md_revisions_combo)
        opts_row1.addSpacing(16)

        self.md_media_cb = QCheckBox("Extract figures and drawings to media/ folder")
        self.md_media_cb.setChecked(True)
        self.md_media_cb.setToolTip(
            "Extracts embedded images, converts vector EMF/WMF drawings to PNG, and links Visio (.vsdx) sources."
        )
        opts_row1.addWidget(self.md_media_cb)

        self.md_header_cb = QCheckBox("Include summary header")
        self.md_header_cb.setChecked(True)
        opts_row1.addWidget(self.md_header_cb)
        opts_row1.addStretch()
        opts_layout.addLayout(opts_row1)

        dest_row = QHBoxLayout()
        self.md_custom_dest_cb = QCheckBox("Save to custom output directory:")
        self.md_custom_dest_cb.toggled.connect(self._on_md_custom_dest_toggled)
        self.md_dest_edit = QLineEdit()
        self.md_dest_edit.setEnabled(False)
        self.md_dest_edit.setPlaceholderText("Defaults to same directory as source document")
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
        self.md_result_frame.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #F0FDF4;
                border: 1px solid #86EFAC;
                border-radius: 6px;
                padding: 6px 12px;
            }
        """)
        self.md_result_frame.setVisible(False)
        res_layout = QVBoxLayout(self.md_result_frame)
        res_layout.setContentsMargins(6, 6, 6, 6)
        res_layout.setSpacing(6)

        res_top_row = QHBoxLayout()
        self.md_result_icon = QLabel("✅")
        self.md_result_lbl = QLabel()
        self.md_result_lbl.setStyleSheet("color: #166534; font-size: 11px; font-weight: bold;")
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
        self.md_preview_edit.setStyleSheet("font-family: Consolas, monospace; font-size: 11px; background: white;")
        self.md_preview_edit.setVisible(False)
        res_layout.addWidget(self.md_preview_edit)

        md_layout.addWidget(self.md_result_frame)
        md_layout.addStretch()

    def _trigger_markdown_conversion(self):
        inputs = self.pane_markdown.get_inputs()
        if not inputs:
            QMessageBox.warning(self, "No Document", "Please select or drop a Word document to convert to Markdown.")
            return

        source_file = inputs[0]
        options = {
            "revisions_mode": self.md_revisions_combo.currentData() or "markup",
            "extract_media": self.md_media_cb.isChecked(),
            "include_header": self.md_header_cb.isChecked(),
            "output_dir": self.md_dest_edit.text().strip() if self.md_custom_dest_cb.isChecked() else None,
        }

        self.md_result_frame.setVisible(False)
        self.md_progress.setVisible(True)
        self.run_markdown_btn.setEnabled(False)
        self.convert_markdown_requested.emit(source_file, options)

    def on_conversion_success(self, out_path: str):
        p = Path(out_path).resolve()
        if p.suffix.lower() not in (".md", ".markdown"):
            return
        self._last_md_output_path = str(p)
        self._last_converted_md = str(p)
        self.md_progress.setVisible(False)
        self.run_markdown_btn.setEnabled(True)
        self.md_result_icon.setText("✅")
        self.md_result_lbl.setStyleSheet("font-weight: bold; color: #166534;")
        self.md_result_lbl.setText(f"Successfully converted to Markdown: {p.name}")
        self.md_open_folder_btn.setEnabled(True)
        self.md_open_file_btn.setEnabled(True)
        self.md_copy_btn.setEnabled(True)
        self.md_preview_toggle_btn.setEnabled(True)
        self.md_result_frame.setVisible(True)

        if p.is_file():
            try:
                with open(p, "r", encoding="utf-8", errors="ignore") as f:
                    preview_lines = [f.readline() for _ in range(60)]
                self.md_preview_edit.setPlainText("".join(preview_lines))
            except Exception:
                self.md_preview_edit.setPlainText("")

    def on_conversion_failed(self, error_msg: str):
        self.md_progress.setVisible(False)
        self.run_markdown_btn.setEnabled(True)
        self.md_result_icon.setText("❌")
        self.md_result_lbl.setStyleSheet("font-weight: bold; color: #DC2626;")
        self.md_result_lbl.setText(f"Conversion failed: {error_msg}")
        self.md_open_folder_btn.setEnabled(False)
        self.md_open_file_btn.setEnabled(False)
        self.md_copy_btn.setEnabled(False)
        self.md_preview_toggle_btn.setEnabled(False)
        self.md_preview_edit.setVisible(False)
        self.md_result_frame.setVisible(True)

    def on_processing_state_changed(self, is_processing: bool):
        self.run_markdown_btn.setEnabled(not is_processing)
        self.md_progress.setVisible(is_processing)

    def _open_md_folder(self):
        if self._last_md_output_path:
            import os
            folder = Path(self._last_md_output_path).parent
            if folder.is_dir() and hasattr(os, "startfile"):
                os.startfile(str(folder))

    def _open_md_file(self):
        if self._last_md_output_path and Path(self._last_md_output_path).is_file():
            import os
            if hasattr(os, "startfile"):
                os.startfile(str(self._last_md_output_path))

    def _copy_md_to_clipboard(self):
        if self._last_md_output_path and Path(self._last_md_output_path).is_file():
            try:
                with open(self._last_md_output_path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                QApplication.clipboard().setText(content)
                self.md_copy_btn.setText("✓ Copied!")
            except Exception as e:
                QMessageBox.warning(self, "Copy Failed", f"Could not read file: {e}")

    def _toggle_md_preview(self):
        is_vis = self.md_preview_edit.isVisible()
        self.md_preview_edit.setVisible(not is_vis)
        self.md_preview_toggle_btn.setText("📝 Hide Preview ▴" if not is_vis else "📝 Show Preview ▾")

    def _on_md_custom_dest_toggled(self, checked: bool):
        self.md_dest_edit.setEnabled(checked)
        self.md_dest_browse_btn.setEnabled(checked)

    def _browse_md_dest(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory for Markdown")
        if folder:
            self.md_dest_edit.setText(folder)
