"""Modern 2-Tier Header Ribbon & Segmented Navigation for Word Tools.

Houses:
- Row 1: System Ribbon (Active Word monitor, 3GPP DOTX template badge, LibreOffice pill, Settings dropdown)
- Row 2: Segmented Workbench Navigation Pills (Compare, Markdown, CR/pCR, Sanitizer, Excerpt, Converter, Visio)
- Operation cards orchestrated via QStackedWidget (implemented in modules.word_tools.ui.cards).
"""

import logging
import os
from pathlib import Path
import shutil
from typing import Optional
import webbrowser

from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QAction,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    COMBOBOX_STYLE_TOOLBAR,
    InteractiveDropLabel,
    TOGGLE_PILL_STYLE,
)
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.document_opener import get_active_word_document
from modules.word_tools.core.styles_3gpp import (
    install_3gpp_dotx_template,
    is_3gpp_dotx_installed,
    uninstall_3gpp_dotx_template,
)
from modules.word_tools.ui.cards import (
    WordCompareCard,
    WordConverterCard,
    WordCrPcrCard,
    WordExcerptCard,
    WordMarkdownCard,
    WordSanitizerCard,
    WordVisioCard,
)

logger = logging.getLogger(__name__)


class WordExtractorTab(QWidget):
    """Modernized Word Tools tab with 2-tier ribbon, live Word COM monitor, and modular operation cards."""

    # Public signals connected by MainWindow
    extract_visio_requested = pyqtSignal(str)
    split_doc_requested = pyqtSignal(str)
    compare_doc_requested = pyqtSignal(str, str, dict)
    convert_doc_requested = pyqtSignal(str, str)
    convert_markdown_requested = pyqtSignal(str, dict)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._active_word_file: Optional[Path] = None
        self._setup_ui()

        # Active Word background polling (every 3 seconds)
        self.word_poll_timer = QTimer(self)
        self.word_poll_timer.timeout.connect(self._poll_active_word)
        self.word_poll_timer.start(3000)
        self._poll_active_word()

        # Environment status checks
        self._check_libreoffice_status()
        self._check_dotx_status()

    # -------------------------------------------------------------------------
    # UI Setup
    # -------------------------------------------------------------------------
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # ---------------------------------------------------------------------
        # 2-Tier Modern Ribbon
        # ---------------------------------------------------------------------
        self.ribbon_frame = QFrame()
        self.ribbon_frame.setObjectName("cardFrame")
        self.ribbon_frame.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 4px 8px;
            }
        """)
        ribbon_layout = QVBoxLayout(self.ribbon_frame)
        ribbon_layout.setContentsMargins(8, 4, 8, 4)
        ribbon_layout.setSpacing(6)

        # Row 1: System Monitors & Settings
        row1 = QHBoxLayout()
        row1.setSpacing(8)

        self.btn_active_word = QPushButton("🪟 Active Word: None Detected")
        self.btn_active_word.setStyleSheet("""
            QPushButton {
                background-color: #F1F5F9;
                color: #475569;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 3px 10px;
                font-weight: 500;
                font-size: 11px;
            }
        """)
        self.btn_active_word.setToolTip("Refreshes automatically. Click to open the active document folder.")
        self.btn_active_word.clicked.connect(self._poll_active_word)
        row1.addWidget(self.btn_active_word)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet("color: #CBD5E1;")
        row1.addWidget(sep1)

        self.dotx_badge = QLabel("3GPP Macro Template (3gpp_70.dotx): Checking...")
        self.dotx_badge.setStyleSheet("font-size: 11px; color: #475569;")
        row1.addWidget(self.dotx_badge)

        self.dotx_btn = QPushButton("Install Template")
        self.dotx_btn.setStyleSheet("font-size: 10px; padding: 2px 6px;")
        self.dotx_btn.clicked.connect(self._on_toggle_dotx_clicked)
        row1.addWidget(self.dotx_btn)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet("color: #CBD5E1;")
        row1.addWidget(sep2)

        self.lo_badge = QLabel("LibreOffice: Checking...")
        self.lo_badge.setStyleSheet("font-size: 11px; color: #475569;")
        row1.addWidget(self.lo_badge)

        row1.addStretch()

        self.settings_btn = QPushButton("⚙️ Word Settings ▾")
        self.settings_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self._setup_settings_menu()
        row1.addWidget(self.settings_btn)

        ribbon_layout.addLayout(row1)

        hsep = QFrame()
        hsep.setFrameShape(QFrame.HLine)
        hsep.setStyleSheet("color: #E2E8F0;")
        ribbon_layout.addWidget(hsep)

        # Row 2: Segmented Workbench Navigation
        row2 = QHBoxLayout()
        row2.setSpacing(6)

        lbl_nav = QLabel("<b>Workbench:</b>")
        lbl_nav.setStyleSheet("color: #475569; font-size: 11px;")
        row2.addWidget(lbl_nav)

        self.nav_btn_compare = QPushButton("⚖️ Compare Documents")
        self.nav_btn_markdown = QPushButton("📝 Word → Markdown")
        self.nav_btn_cr_pcr = QPushButton("🏛️ 3GPP CR && Templates")
        self.nav_btn_author = QPushButton("🛠️ Sanitizer && Revisions")
        self.nav_btn_author.setToolTip("Unify proofing language, purge language Track Changes clutter, consolidate revision authors, and sanitize metadata.")
        self.nav_btn_slice = QPushButton("✂️ Slice && Excerpt")
        self.nav_btn_convert = QPushButton("🔄 Format Converter")
        self.nav_btn_visio = QPushButton("🖼️ Visio && Media Extractor")

        self._nav_buttons = [
            (self.nav_btn_compare, 4),
            (self.nav_btn_markdown, 6),
            (self.nav_btn_cr_pcr, 7),
            (self.nav_btn_author, 8),
            (self.nav_btn_slice, 3),
            (self.nav_btn_convert, 5),
            (self.nav_btn_visio, 1),
        ]

        for btn, idx in self._nav_buttons:
            btn.setCheckable(True)
            btn.setStyleSheet(TOGGLE_PILL_STYLE)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda checked, i=idx: self.set_active_operation(i))
            row2.addWidget(btn)

        row2.addStretch()

        row2.addWidget(QLabel("<b>Jump:</b>"))
        self.op_combo = QComboBox()
        self.op_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.op_combo.addItems([
            "Convert Legacy .doc to .docx (Explicit LibreOffice)",
            "Extract Embedded Visio Diagrams",
            "Subtractive Slicing (Split by Clause)",
            "Extract Sections by Heading",
            "Compare Documents (Native Word Diff)",
            "Convert Document Format (Auto / Word)",
            "Convert Word to Markdown (.md)",
            "Draft 3GPP Contribution (CR / pCR)",
            "Consolidate / Anonymize Track Changes Authors",
        ])
        self.op_combo.currentIndexChanged.connect(self._on_op_combo_changed)
        row2.addWidget(self.op_combo)

        ribbon_layout.addLayout(row2)
        layout.addWidget(self.ribbon_frame)

        # ---------------------------------------------------------------------
        # Operation Cards (Stacked Widget)
        # ---------------------------------------------------------------------
        self.stack = QStackedWidget()

        # Index 0: Legacy LibreOffice Drop Target
        self.card_lo_convert = QWidget()
        lo_layout = QVBoxLayout(self.card_lo_convert)
        self.lo_drop = InteractiveDropLabel("📥 Drag & Drop legacy .doc files here for batch conversion via LibreOffice", [".doc"])
        self.lo_drop.file_dropped.connect(self._on_lo_dropped)
        lo_layout.addWidget(self.lo_drop)
        self.stack.addWidget(self.card_lo_convert)

        # Index 1: Visio Extractor Card
        self.card_visio = WordVisioCard()
        self.card_visio.extract_visio_requested.connect(self.extract_visio_requested)
        self.stack.addWidget(self.card_visio)

        # Index 2: Split Document Card
        self.card_split = QWidget()
        split_layout = QVBoxLayout(self.card_split)
        self.split_drop = InteractiveDropLabel("📥 Drag & Drop Word document (.docx) here to split by major heading clauses", [".docx"])
        self.split_drop.file_dropped.connect(self._on_split_dropped)
        split_layout.addWidget(self.split_drop)
        self.stack.addWidget(self.card_split)

        # Index 3: Heading Excerpt Card
        self.card_excerpt = WordExcerptCard()
        self.stack.addWidget(self.card_excerpt)

        # Index 4: Compare Documents Card
        self.card_compare = WordCompareCard()
        self.card_compare.compare_doc_requested.connect(self.compare_doc_requested)
        self.stack.addWidget(self.card_compare)

        # Index 5: Multi-Format Converter Card
        self.card_convert = WordConverterCard()
        self.card_convert.convert_doc_requested.connect(self.convert_doc_requested)
        self.stack.addWidget(self.card_convert)

        # Index 6: Word to Markdown Card
        self.card_markdown = WordMarkdownCard()
        self.card_markdown.convert_markdown_requested.connect(self.convert_markdown_requested)
        self.stack.addWidget(self.card_markdown)

        # Index 7: 3GPP CR / pCR Card
        self.card_cr_pcr = WordCrPcrCard()
        self.stack.addWidget(self.card_cr_pcr)

        # Index 8: Document Sanitizer & Revision Hub
        self.card_author_renamer = WordSanitizerCard()
        self.stack.addWidget(self.card_author_renamer)

        layout.addWidget(self.stack)

        # Default to Compare Documents (Index 4)
        self.set_active_operation(4)

    # -------------------------------------------------------------------------
    # Ribbon & Operation Switching
    # -------------------------------------------------------------------------
    def set_active_operation(self, index: int):
        self.stack.setCurrentIndex(index)
        self._update_nav_pills(index)
        if self.op_combo.currentIndex() != index:
            self.op_combo.blockSignals(True)
            self.op_combo.setCurrentIndex(index)
            self.op_combo.blockSignals(False)

    def _on_op_combo_changed(self, index: int):
        self.set_active_operation(index)

    def _update_nav_pills(self, active_index: int):
        for btn, idx in self._nav_buttons:
            is_active = (idx == active_index)
            btn.setChecked(is_active)
            if is_active:
                btn.setStyleSheet(
                    "background-color: #EBF3FC; border: 2px solid #395396; color: #395396; font-weight: 700; border-radius: 4px; padding: 4px 10px; font-size: 11px;"
                )
            else:
                btn.setStyleSheet(TOGGLE_PILL_STYLE)

    def _poll_active_word(self):
        try:
            active_path = get_active_word_document()
            if active_path and active_path.is_file():
                self._active_word_file = active_path
                self.btn_active_word.setText(f"🪟 Active Word: {active_path.name}")
                self.btn_active_word.setStyleSheet("""
                    QPushButton {
                        background-color: #ECFDF5;
                        color: #065F46;
                        border: 1px solid #6EE7B7;
                        border-radius: 4px;
                        padding: 3px 10px;
                        font-weight: 600;
                        font-size: 11px;
                    }
                """)
                self.btn_active_word.setToolTip(f"Active in Microsoft Word:\n{active_path.resolve()}\n\nClick to reveal in folder.")
            else:
                self._active_word_file = None
                self.btn_active_word.setText("🪟 Active Word: None Detected")
                self.btn_active_word.setStyleSheet("""
                    QPushButton {
                        background-color: #F1F5F9;
                        color: #64748B;
                        border: 1px solid #CBD5E1;
                        border-radius: 4px;
                        padding: 3px 10px;
                        font-weight: 500;
                        font-size: 11px;
                    }
                """)
                self.btn_active_word.setToolTip("No active document currently open in Microsoft Word.")
        except Exception as exc:
            logger.debug("Active Word poll failed: %s", exc)

    def _setup_settings_menu(self):
        menu = QMenu(self)

        act_settings = QAction("⚙️ Preferences & Sensitivity Labels...", self)
        act_settings.triggered.connect(self._open_settings_dialog)
        menu.addAction(act_settings)

        act_batch_label = QAction("🏷️ Batch Apply Sensitivity Labels...", self)
        act_batch_label.triggered.connect(self._open_batch_label_dialog)
        menu.addAction(act_batch_label)

        menu.addSeparator()

        act_templates = QAction("📂 Open Word Templates Folder", self)
        act_templates.triggered.connect(self._open_word_templates_folder)
        menu.addAction(act_templates)

        self.settings_btn.setMenu(menu)

    def _open_settings_dialog(self):
        from modules.word_tools.ui.word_settings_dialog import WordSettingsDialog
        dlg = WordSettingsDialog(parent=self)
        dlg.exec_()
        self._check_dotx_status()

    def _open_batch_label_dialog(self):
        from modules.word_tools.ui.batch_label_dialog import BatchSensitivityLabelDialog
        dlg = BatchSensitivityLabelDialog(parent=self)
        dlg.exec_()

    def _open_word_templates_folder(self):
        roaming = os.environ.get("APPDATA", "")
        templates_dir = Path(roaming) / "Microsoft" / "Templates"
        if templates_dir.is_dir() and hasattr(os, "startfile"):
            os.startfile(str(templates_dir))
        else:
            QMessageBox.information(self, "Templates Folder", f"Microsoft Word Templates directory:\n{templates_dir}")

    def _check_libreoffice_status(self):
        lo_path = WordConfig.load().get("libreoffice_path", "")
        if lo_path and Path(lo_path).is_file():
            self.lo_badge.setText("🟢 LibreOffice: Configured")
            self.lo_badge.setStyleSheet("color: #166534; font-weight: 600; font-size: 11px;")
        else:
            found = shutil.which("soffice") or shutil.which("libreoffice")
            if found:
                self.lo_badge.setText("🟢 LibreOffice: In PATH")
                self.lo_badge.setStyleSheet("color: #166534; font-weight: 600; font-size: 11px;")
            else:
                self.lo_badge.setText("🟠 LibreOffice: Not Found")
                self.lo_badge.setStyleSheet("color: #92400E; font-weight: 600; font-size: 11px;")

    def _check_dotx_status(self):
        installed = is_3gpp_dotx_installed()
        if installed:
            self.dotx_badge.setText("🟢 3GPP Macro Template: Installed")
            self.dotx_badge.setStyleSheet("color: #166534; font-weight: 600; font-size: 11px;")
            self.dotx_btn.setText("Uninstall Template")
            self.dotx_btn.setStyleSheet("font-size: 10px; padding: 2px 6px; color: #991B1B;")
        else:
            self.dotx_badge.setText("⚪ 3GPP Macro Template: Not Installed")
            self.dotx_badge.setStyleSheet("color: #475569; font-size: 11px;")
            self.dotx_btn.setText("Install Template")
            self.dotx_btn.setStyleSheet("font-size: 10px; padding: 2px 6px; color: #1E40AF;")

    def _on_toggle_dotx_clicked(self):
        if is_3gpp_dotx_installed():
            reply = QMessageBox.question(
                self,
                "Uninstall 3GPP Template",
                "Remove 3gpp_70.dotx from your Microsoft Word Templates folder?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                uninstall_3gpp_dotx_template()
                self._check_dotx_status()
        else:
            ok, msg = install_3gpp_dotx_template()
            if ok:
                QMessageBox.information(self, "Template Installed", msg)
            else:
                QMessageBox.warning(self, "Installation Failed", msg)
            self._check_dotx_status()

    # -------------------------------------------------------------------------
    # Drop Handlers for simple operations
    # -------------------------------------------------------------------------
    def _on_lo_dropped(self, files: list):
        if files:
            self.convert_doc_requested.emit(files[0], "docx")

    def _on_split_dropped(self, files: list):
        if files:
            self.split_doc_requested.emit(files[0])

    # -------------------------------------------------------------------------
    # External Worker Lifecycle Hooks (Invoked by MainWindow)
    # -------------------------------------------------------------------------
    def on_conversion_success(self, out_path: str):
        if hasattr(self.card_markdown, "on_conversion_success"):
            self.card_markdown.on_conversion_success(out_path)
        if hasattr(self.card_convert, "on_conversion_success"):
            self.card_convert.on_conversion_success(out_path)
        if hasattr(self.card_visio, "on_extraction_success"):
            self.card_visio.on_extraction_success(out_path)

    def on_conversion_failed(self, error_msg: str):
        if hasattr(self.card_markdown, "on_conversion_failed"):
            self.card_markdown.on_conversion_failed(error_msg)
        QMessageBox.critical(self, "Conversion Error", f"Operation failed:\n{error_msg}")

    def on_processing_state_changed(self, is_processing: bool):
        if hasattr(self.card_markdown, "on_processing_state_changed"):
            self.card_markdown.on_processing_state_changed(is_processing)
