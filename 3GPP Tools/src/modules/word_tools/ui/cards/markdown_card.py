"""Card: Bidirectional Markdown ⇄ Word Studio."""

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
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    TOGGLE_PILL_STYLE,
)
from modules.word_tools.core.document_opener import open_document
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane

logger = logging.getLogger(__name__)


class WordMarkdownCard(QWidget):
    """
    Card: Bidirectional Markdown ⇄ Word Studio.
    Supports fast headless conversion of Word documents to Markdown (with revision tracking & media extraction)
    and compilation of Markdown documents to Word (.docx) with embedded media, PlantUML diagrams, and 3GPP styles.
    """

    convert_markdown_requested = pyqtSignal(str, dict)
    convert_markdown_to_word_requested = pyqtSignal(str, dict)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._last_md_output_path: Optional[str] = None
        self._last_converted_md: Optional[str] = None
        self._last_docx_output_path: Optional[str] = None
        self._setup_ui()

        # Backward compatibility aliases
        self.pane_md = self.pane_markdown

    def _setup_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(10)

        # ---------------------------------------------------------------------
        # Direction Navigation Strip
        # ---------------------------------------------------------------------
        nav_frame = QFrame()
        nav_frame.setObjectName("cardFrame")
        nav_frame.setStyleSheet(CARD_FRAME_STYLE)
        nav_layout = QHBoxLayout(nav_frame)
        nav_layout.setContentsMargins(12, 6, 12, 6)
        nav_layout.setSpacing(8)

        lbl_dir = QLabel("<b>Direction:</b>")
        lbl_dir.setStyleSheet("color: #475569; font-size: 11px;")
        nav_layout.addWidget(lbl_dir)

        self.btn_mode_w2m = QPushButton("📄 Word → Markdown")
        self.btn_mode_m2w = QPushButton("📝 Markdown → Word")
        for btn in (self.btn_mode_w2m, self.btn_mode_m2w):
            btn.setCheckable(True)
            btn.setStyleSheet(TOGGLE_PILL_STYLE)
            btn.setCursor(Qt.PointingHandCursor)

        self.btn_mode_w2m.setChecked(True)
        self.btn_mode_w2m.clicked.connect(lambda: self.set_studio_mode(0))
        self.btn_mode_m2w.clicked.connect(lambda: self.set_studio_mode(1))
        nav_layout.addWidget(self.btn_mode_w2m)
        nav_layout.addWidget(self.btn_mode_m2w)
        nav_layout.addStretch()

        self.btn_switch_direction = QPushButton("⇄ Switch Direction")
        self.btn_switch_direction.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_switch_direction.setCursor(Qt.PointingHandCursor)
        self.btn_switch_direction.setToolTip("Quickly toggle between Word-to-Markdown and Markdown-to-Word modes.")
        self.btn_switch_direction.clicked.connect(self._toggle_direction)
        nav_layout.addWidget(self.btn_switch_direction)

        root_layout.addWidget(nav_frame)

        # ---------------------------------------------------------------------
        # Stacked Pages: [0] Word → Markdown, [1] Markdown → Word
        # ---------------------------------------------------------------------
        self.studio_stack = QStackedWidget()

        # Page 0: Word to Markdown
        page_w2m = QWidget()
        self._setup_page_w2m(page_w2m)
        self.studio_stack.addWidget(page_w2m)

        # Page 1: Markdown to Word
        page_m2w = QWidget()
        self._setup_page_m2w(page_m2w)
        self.studio_stack.addWidget(page_m2w)

        root_layout.addWidget(self.studio_stack)

    # =========================================================================
    # Page 0: Word → Markdown UI
    # =========================================================================
    def _setup_page_w2m(self, container: QWidget):
        md_layout = QVBoxLayout(container)
        md_layout.setContentsMargins(0, 0, 0, 0)
        md_layout.setSpacing(10)

        self.pane_markdown = DocumentSelectorPane(
            "📄 WORD DOCUMENT TO CONVERT",
            accepted_extensions=[".docx", ".doc", ".docm", ".md", ".markdown"],
        )
        self.pane_markdown.file_selected.connect(self._on_w2m_file_selected)
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
        self.md_result_frame.setStyleSheet(
            CARD_FRAME_STYLE
            + """
            QFrame#cardFrame {
                background-color: #F0FDF4;
                border: 1px solid #86EFAC;
                border-radius: 6px;
                padding: 6px 12px;
            }
        """
        )
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

    # =========================================================================
    # Page 1: Markdown → Word UI
    # =========================================================================
    def _setup_page_m2w(self, container: QWidget):
        m2w_layout = QVBoxLayout(container)
        m2w_layout.setContentsMargins(0, 0, 0, 0)
        m2w_layout.setSpacing(10)

        self.pane_m2w = DocumentSelectorPane(
            "📝 MARKDOWN DOCUMENT TO CONVERT (.md)",
            accepted_extensions=[".md", ".markdown", ".txt", ".docx", ".doc"],
        )
        self.pane_m2w.file_selected.connect(self._on_m2w_file_selected)
        m2w_layout.addWidget(self.pane_m2w)

        m2w_opts_frame = QFrame()
        m2w_opts_frame.setObjectName("cardFrame")
        m2w_opts_frame.setStyleSheet(CARD_FRAME_STYLE)
        opts_layout = QVBoxLayout(m2w_opts_frame)
        opts_layout.setContentsMargins(12, 10, 12, 10)
        opts_layout.setSpacing(8)

        row1 = QHBoxLayout()
        self.m2w_media_cb = QCheckBox("Auto-detect companion media folder ({stem}_media/)")
        self.m2w_media_cb.setChecked(True)
        self.m2w_media_cb.setToolTip(
            "Automatically embeds figures and drawings located in the companion media folder or relative paths."
        )
        row1.addWidget(self.m2w_media_cb)

        self.m2w_diagrams_cb = QCheckBox("Render PlantUML diagrams (```plantuml)")
        self.m2w_diagrams_cb.setChecked(True)
        self.m2w_diagrams_cb.setToolTip(
            "Compiles embedded PlantUML sequence and architecture code blocks into high-resolution figures on the fly."
        )
        row1.addWidget(self.m2w_diagrams_cb)
        row1.addStretch()
        opts_layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.m2w_styles_cb = QCheckBox("Apply 3GPP contribution styles (3gpp_70.dotx & TableGrid)")
        self.m2w_styles_cb.setChecked(True)
        self.m2w_styles_cb.setToolTip(
            "Injects canonical 3GPP headings, typography, margins, and TableGrid formatting into the OpenXML package."
        )
        row2.addWidget(self.m2w_styles_cb)

        self.m2w_auto_open_cb = QCheckBox("Auto-open in Word after conversion")
        self.m2w_auto_open_cb.setChecked(True)
        self.m2w_auto_open_cb.setToolTip("Launches the generated document in Microsoft Word immediately upon completion.")
        row2.addWidget(self.m2w_auto_open_cb)
        row2.addStretch()
        opts_layout.addLayout(row2)

        dest_row = QHBoxLayout()
        self.m2w_custom_dest_cb = QCheckBox("Save to custom output directory:")
        self.m2w_custom_dest_cb.toggled.connect(self._on_m2w_custom_dest_toggled)
        self.m2w_dest_edit = QLineEdit()
        self.m2w_dest_edit.setEnabled(False)
        self.m2w_dest_edit.setPlaceholderText("Defaults to same directory as source Markdown document")
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
            + """
            QFrame#cardFrame {
                background-color: #F0FDF4;
                border: 1px solid #86EFAC;
                border-radius: 6px;
                padding: 6px 12px;
            }
        """
        )
        self.m2w_result_frame.setVisible(False)
        m2w_res_layout = QHBoxLayout(self.m2w_result_frame)
        m2w_res_layout.setContentsMargins(6, 6, 6, 6)
        m2w_res_layout.setSpacing(8)

        self.m2w_result_icon = QLabel("✅")
        self.m2w_result_lbl = QLabel()
        self.m2w_result_lbl.setStyleSheet("color: #166534; font-size: 11px; font-weight: bold;")
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

    # =========================================================================
    # Mode Switching & Smart Routing
    # =========================================================================
    def set_studio_mode(self, index: int):
        """Sets the active studio mode (0: Word → Markdown, 1: Markdown → Word)."""
        idx = max(0, min(index, 1))
        self.studio_stack.setCurrentIndex(idx)
        self.btn_mode_w2m.setChecked(idx == 0)
        self.btn_mode_m2w.setChecked(idx == 1)

    def _toggle_direction(self):
        new_idx = 1 if self.studio_stack.currentIndex() == 0 else 0
        self.set_studio_mode(new_idx)

    def _on_w2m_file_selected(self, files: list):
        if not files:
            return
        ext = Path(files[0]).suffix.lower()
        if ext in (".md", ".markdown"):
            self.set_studio_mode(1)
            self.pane_m2w.selected_files = files
            self.pane_m2w.drop_zone.setText(f"📄 {Path(files[0]).name}")

    def _on_m2w_file_selected(self, files: list):
        if not files:
            return
        ext = Path(files[0]).suffix.lower()
        if ext in (".docx", ".doc", ".docm"):
            self.set_studio_mode(0)
            self.pane_markdown.selected_files = files
            self.pane_markdown.drop_zone.setText(f"📄 {Path(files[0]).name}")

    # =========================================================================
    # Word → Markdown Execution & Results
    # =========================================================================
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

    # =========================================================================
    # Markdown → Word Execution & Results
    # =========================================================================
    def _trigger_m2w_conversion(self):
        inputs = self.pane_m2w.get_inputs()
        if not inputs:
            QMessageBox.warning(self, "No Document", "Please select or drop a Markdown (.md) document to convert to Word.")
            return

        source_file = inputs[0]
        options = {
            "companion_media_dir": None,
            "render_diagrams": self.m2w_diagrams_cb.isChecked(),
            "apply_3gpp_styles": self.m2w_styles_cb.isChecked(),
            "auto_open": self.m2w_auto_open_cb.isChecked(),
            "output_dir": self.m2w_dest_edit.text().strip() if self.m2w_custom_dest_cb.isChecked() else None,
        }

        self.m2w_result_frame.setVisible(False)
        self.m2w_progress.setVisible(True)
        self.run_m2w_btn.setEnabled(False)
        self.convert_markdown_to_word_requested.emit(source_file, options)

    def _open_m2w_folder(self):
        if self._last_docx_output_path:
            import os
            folder = Path(self._last_docx_output_path).parent
            if folder.is_dir() and hasattr(os, "startfile"):
                os.startfile(str(folder))

    def _open_m2w_file(self):
        if self._last_docx_output_path and Path(self._last_docx_output_path).is_file():
            try:
                open_document(self._last_docx_output_path)
            except Exception as e:
                QMessageBox.warning(self, "Open Failed", f"Could not open document in Word: {e}")

    def _copy_m2w_path_to_clipboard(self):
        if self._last_docx_output_path:
            QApplication.clipboard().setText(self._last_docx_output_path)
            self.m2w_copy_path_btn.setText("✓ Copied!")

    def _on_m2w_custom_dest_toggled(self, checked: bool):
        self.m2w_dest_edit.setEnabled(checked)
        self.m2w_dest_browse_btn.setEnabled(checked)

    def _browse_m2w_dest(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory for Word Document")
        if folder:
            self.m2w_dest_edit.setText(folder)

    # =========================================================================
    # Universal Task Callbacks
    # =========================================================================
    def on_conversion_success(self, out_path: str):
        p = Path(out_path).resolve()
        suffix = p.suffix.lower()

        # Word → Markdown completion
        if suffix in (".md", ".markdown"):
            self.set_studio_mode(0)
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

        # Markdown → Word completion
        elif suffix in (".docx", ".doc"):
            self.set_studio_mode(1)
            self._last_docx_output_path = str(p)
            self.m2w_progress.setVisible(False)
            self.run_m2w_btn.setEnabled(True)
            self.m2w_result_icon.setText("✅")
            self.m2w_result_lbl.setStyleSheet("font-weight: bold; color: #166534;")
            self.m2w_result_lbl.setText(f"Successfully compiled Word document: {p.name}")
            self.m2w_open_folder_btn.setEnabled(True)
            self.m2w_open_file_btn.setEnabled(True)
            self.m2w_copy_path_btn.setEnabled(True)
            self.m2w_result_frame.setVisible(True)

    def on_conversion_failed(self, error_msg: str):
        # Reset progress in active mode
        if self.studio_stack.currentIndex() == 0:
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
        else:
            self.m2w_progress.setVisible(False)
            self.run_m2w_btn.setEnabled(True)
            self.m2w_result_icon.setText("❌")
            self.m2w_result_lbl.setStyleSheet("font-weight: bold; color: #DC2626;")
            self.m2w_result_lbl.setText(f"Compilation failed: {error_msg}")
            self.m2w_open_folder_btn.setEnabled(False)
            self.m2w_open_file_btn.setEnabled(False)
            self.m2w_copy_path_btn.setEnabled(False)
            self.m2w_result_frame.setVisible(True)

    def on_processing_state_changed(self, is_processing: bool):
        self.run_markdown_btn.setEnabled(not is_processing)
        self.run_m2w_btn.setEnabled(not is_processing)
        if self.studio_stack.currentIndex() == 0:
            self.md_progress.setVisible(is_processing)
        else:
            self.m2w_progress.setVisible(is_processing)
