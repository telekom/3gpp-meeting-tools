"""Card: 3GPP Contribution & Template Workbench (CR / pCR Generator)."""

import logging
from pathlib import Path
from typing import Optional
import webbrowser

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    COMBOBOX_STYLE_TOOLBAR,
)
from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane

logger = logging.getLogger(__name__)


class WordCrPcrCard(QWidget):
    """Card: First-class 3GPP Change Request (CR) and Pseudo-CR authoring workbench."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        card_layout = QVBoxLayout(self)
        card_layout.setContentsMargins(10, 10, 10, 10)
        card_layout.setSpacing(10)

        # 1. 3GPP Contribution Builder Frame
        builder_frame = QFrame()
        builder_frame.setObjectName("cardFrame")
        builder_frame.setStyleSheet(CARD_FRAME_STYLE)
        bf_layout = QVBoxLayout(builder_frame)
        bf_layout.setContentsMargins(14, 12, 14, 12)
        bf_layout.setSpacing(10)

        title_row = QHBoxLayout()
        bf_title = QLabel("<b>📝 3GPP Contribution Builder (CR & Pseudo-CR Generator)</b>")
        bf_title.setStyleSheet("font-size: 13px; color: #1E293B;")
        title_row.addWidget(bf_title)
        title_row.addStretch()

        self.cr_mode_combo = QComboBox()
        self.cr_mode_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.cr_mode_combo.addItem("Formal Change Request (CR)", "cr")
        self.cr_mode_combo.addItem("Pseudo-CR (pCR for Work Items/TRs)", "pcr")
        title_row.addWidget(QLabel("Contribution Type:"))
        title_row.addWidget(self.cr_mode_combo)
        bf_layout.addLayout(title_row)

        desc_lbl = QLabel(
            "Rapidly assemble formal 3GPP Change Requests directly from specification documents. "
            "Features automated meeting metadata pre-fill, change-block contiguous grouping, "
            "authoritative 3GPP Word style enforcement, and corporate sensitivity labeling."
        )
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        bf_layout.addWidget(desc_lbl)

        # Document Selector Pane specifically for CR drafting
        self.pane_cr = DocumentSelectorPane("📄 BASELINE SPECIFICATION DOCUMENT")
        bf_layout.addWidget(self.pane_cr)

        act_row = QHBoxLayout()
        self.btn_launch_cr = QPushButton("🚀 Open Contribution Builder (CR / pCR)...")
        self.btn_launch_cr.setObjectName("primaryBtn")
        self.btn_launch_cr.setCursor(Qt.PointingHandCursor)
        self.btn_launch_cr.clicked.connect(self._launch_cr_pcr_from_pane)
        act_row.addWidget(self.btn_launch_cr)
        act_row.addStretch()
        bf_layout.addLayout(act_row)

        card_layout.addWidget(builder_frame)

        # 2. 3GPP Templates & Styles Reference Frame
        tmpl_frame = QFrame()
        tmpl_frame.setObjectName("cardFrame")
        tmpl_frame.setStyleSheet(CARD_FRAME_STYLE)
        tf_layout = QVBoxLayout(tmpl_frame)
        tf_layout.setContentsMargins(14, 12, 14, 12)
        tf_layout.setSpacing(8)

        tf_title = QLabel("<b>🏛️ Official 3GPP Contribution Templates & Specification Forms</b>")
        tf_title.setStyleSheet("font-size: 13px; color: #1E293B;")
        tf_layout.addWidget(tf_title)

        tf_desc = QLabel(
            "Download official Word templates and specification drafting covers directly from 3GPP servers. "
            "Click any template below to download the authoritative release form."
        )
        tf_desc.setWordWrap(True)
        tf_desc.setStyleSheet("color: #475569; font-size: 11px;")
        tf_layout.addWidget(tf_desc)

        forms_row = QHBoxLayout()
        forms_row.setSpacing(8)

        btn_cr_form = QPushButton("📄 Download CR Template (Template_3GPP_CR.docx)")
        btn_cr_form.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_cr_form.clicked.connect(lambda: webbrowser.open("https://www.3gpp.org/ftp/Information/CR_forms/Template_3GPP_CR.docx"))

        btn_pcr_form = QPushButton("📄 Download Pseudo-CR Form (Pseudo-CR-Form-v0-1.zip)")
        btn_pcr_form.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_pcr_form.clicked.connect(lambda: webbrowser.open("https://www.3gpp.org/ftp/Information/CR_forms/Pseudo-CR-Form-v0-1.zip"))

        btn_open_tmpls = QPushButton("📂 Open Word Templates Folder")
        btn_open_tmpls.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_open_tmpls.clicked.connect(self._open_word_templates_folder)

        forms_row.addWidget(btn_cr_form)
        forms_row.addWidget(btn_pcr_form)
        forms_row.addWidget(btn_open_tmpls)
        forms_row.addStretch()
        tf_layout.addLayout(forms_row)

        card_layout.addWidget(tmpl_frame)
        card_layout.addStretch()

    def _launch_cr_pcr_from_pane(self):
        inputs = self.pane_cr.get_inputs()
        initial_file = inputs[0] if inputs else ""

        mode = self.cr_mode_combo.currentData() or "cr"
        initial_spec = ""
        initial_version = ""

        if initial_file:
            stem = Path(initial_file).stem
            import re
            m = re.search(r"(\d{2})[\.\-_](\d{3})", stem)
            if m:
                initial_spec = f"TS {m.group(1)}.{m.group(2)}"
            v_match = re.search(r"[vV](\d+)[\.\-_](\d+)[\.\-_](\d+)", stem)
            if v_match:
                initial_version = f"{v_match.group(1)}.{v_match.group(2)}.{v_match.group(3)}"

        dlg = CrPcrDialog(
            parent=self,
            initial_file=initial_file,
            mode=mode,
            initial_spec=initial_spec,
            initial_version=initial_version,
        )
        dlg.exec_()

    def _open_word_templates_folder(self):
        import os
        roaming = os.environ.get("APPDATA", "")
        templates_dir = Path(roaming) / "Microsoft" / "Templates"
        if templates_dir.is_dir() and hasattr(os, "startfile"):
            os.startfile(str(templates_dir))
        else:
            QMessageBox.information(self, "Templates Folder", f"Microsoft Word Templates directory:\n{templates_dir}")
