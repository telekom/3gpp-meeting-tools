# --- File: src/modules/meetings/ui/dossier_export_dialog.py ---
"""
dossier_export_dialog.py - Configuration dialog for the Navigable Dossier Exporter (F-04).
Allows delegates to configure HTML & PDF export options, companion AI prompt generation,
and target directories.
"""

from __future__ import annotations

import datetime
import os
import re
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QGroupBox,
    QRadioButton,
    QCheckBox,
    QPushButton,
    QLineEdit,
    QFileDialog,
    QButtonGroup,
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY

STYLE_COMPACT_BTN = """
    QPushButton {
        background-color: #F8FAFC;
        color: #1E293B;
        border: 1px solid #CBD5E1;
        border-radius: 4px;
        padding: 4px 10px;
        font-size: 11px;
        font-weight: 500;
    }
    QPushButton:hover {
        background-color: #F1F5F9;
        border-color: #94A3B8;
    }
    QPushButton:pressed {
        background-color: #E2E8F0;
    }
"""

STYLE_PRIMARY_BTN = """
    QPushButton {
        background-color: #005A9E;
        color: #FFFFFF;
        font-weight: bold;
        border-radius: 4px;
        padding: 7px 18px;
        font-size: 12px;
        border: 1px solid #004578;
    }
    QPushButton:hover {
        background-color: #004578;
    }
    QPushButton:pressed {
        background-color: #00335A;
    }
"""


class DossierExportDialog(QDialog):
    """
    Configuration dialog for exporting selected or visible TDocs to a consolidated
    Navigable HTML Dossier and Bookmarked PDF.
    """

    def __init__(
        self,
        parent,
        meeting_dir: Path,
        tdocs_count: int,
        mtg_info: dict,
        scope_label: str = "selected",
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.tdocs_count = tdocs_count
        self.mtg_info = mtg_info or {}
        self.scope_label = scope_label

        self.setWindowTitle("📑 Export Navigable Dossier (HTML & PDF)")
        self.resize(540, 480)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(18, 18, 18, 18)

        # Header Info Banner
        wg = str(self.mtg_info.get("wg_name", "3GPP")).strip()
        mtg = str(self.mtg_info.get("meeting_number", "")).strip()
        mtg_str = f"{wg} #{mtg}" if wg or mtg else "Current Meeting"

        header_lbl = QLabel(
            f"<h3 style='margin:0 0 4px 0; color:#005A9E;'>📑 Compile Technical Dossier</h3>"
            f"<span style='color:#475569; font-size:12px;'>"
            f"Ready to compile <b>{self.tdocs_count} {self.scope_label}</b> TDoc(s) from <b>{mtg_str}</b> "
            f"into a structured, navigable technical dossier."
            f"</span>"
        )
        header_lbl.setWordWrap(True)
        layout.addWidget(header_lbl)

        # 1. Format Selection
        fmt_group = QGroupBox("1. Output Formats")
        fmt_layout = QVBoxLayout(fmt_group)
        fmt_layout.setSpacing(8)

        self.btn_group_fmt = QButtonGroup(self)

        self.radio_both = QRadioButton("Both Interactive HTML && Bookmarked PDF (Recommended)")
        self.radio_both.setToolTip(
            "Produces a self-contained responsive HTML dossier for fast laptop reading, "
            "and a vector PDF with hierarchical bookmarks for offline review and NotebookLM/Claude."
        )
        self.radio_both.setChecked(True)
        self.btn_group_fmt.addButton(self.radio_both, 1)
        fmt_layout.addWidget(self.radio_both)

        self.radio_html = QRadioButton("Interactive HTML Dossier only (.html)")
        self.radio_html.setToolTip("Fastest export. Generates a standalone, responsive web page with sticky navigation.")
        self.btn_group_fmt.addButton(self.radio_html, 2)
        fmt_layout.addWidget(self.radio_html)

        self.radio_pdf = QRadioButton("Bookmarked Vector PDF only (.pdf)")
        self.radio_pdf.setToolTip("Generates a single consolidated PDF with internal links and hierarchical outline bookmarks.")
        self.btn_group_fmt.addButton(self.radio_pdf, 3)
        fmt_layout.addWidget(self.radio_pdf)

        layout.addWidget(fmt_group)

        # 2. Options
        opts_group = QGroupBox("2. Dossier Options")
        opts_layout = QVBoxLayout(opts_group)
        opts_layout.setSpacing(8)

        self.chk_prompt = QCheckBox("Generate Companion AI Prompt && Page Citation Map (for NotebookLM / Claude)")
        self.chk_prompt.setToolTip(
            "Creates dossier_prompt.txt detailing exact 1-indexed page boundaries for every TDoc, "
            "with tailored technical evaluation directives."
        )
        self.chk_prompt.setChecked(True)
        opts_layout.addWidget(self.chk_prompt)

        self.chk_b64 = QCheckBox("Embed Images Directly into HTML (Self-Contained Single File)")
        self.chk_b64.setToolTip(
            "Embeds all extracted call-flow diagrams and figures as Base64 data URIs so the HTML file can be moved or emailed independently."
        )
        self.chk_b64.setChecked(True)
        opts_layout.addWidget(self.chk_b64)

        self.chk_download = QCheckBox("Download missing TDocs automatically (via 3GPP FTP / portal)")
        self.chk_download.setChecked(True)
        opts_layout.addWidget(self.chk_download)

        self.chk_force = QCheckBox("Force re-extraction of Word documents (ignore existing cache)")
        self.chk_force.setChecked(False)
        opts_layout.addWidget(self.chk_force)

        self.chk_auto_open = QCheckBox("Open destination directory upon completion")
        self.chk_auto_open.setChecked(True)
        opts_layout.addWidget(self.chk_auto_open)

        layout.addWidget(opts_group)

        # 3. Output Directory
        dest_group = QGroupBox("3. Destination Directory")
        dest_layout = QHBoxLayout(dest_group)
        dest_layout.setSpacing(6)

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        clean_wg = re.sub(r'[^\w\-]', '_', wg)
        clean_mtg = re.sub(r'[^\w\-]', '_', mtg)
        folder_name = f"Dossier_{clean_wg}_{clean_mtg}_{timestamp}" if (clean_wg or clean_mtg) else f"Dossier_{timestamp}"
        default_dir = self.meeting_dir / "Export" / folder_name

        self.dest_edit = QLineEdit(str(default_dir))
        dest_layout.addWidget(self.dest_edit)

        btn_browse = QPushButton("📁 Browse...")
        btn_browse.setStyleSheet(STYLE_COMPACT_BTN)
        btn_browse.clicked.connect(self._on_browse_clicked)
        dest_layout.addWidget(btn_browse)

        layout.addWidget(dest_group)

        layout.addStretch()

        # Action Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_cancel = QPushButton("Cancel")
        btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)

        btn_start = QPushButton("🚀 Start Export")
        btn_start.setStyleSheet(STYLE_PRIMARY_BTN)
        btn_start.setDefault(True)
        btn_start.clicked.connect(self.accept)
        btn_layout.addWidget(btn_start)

        layout.addLayout(btn_layout)

    def _on_browse_clicked(self):
        curr = self.dest_edit.text().strip() or str(self.meeting_dir / "Export")
        folder = QFileDialog.getExistingDirectory(self, "Select Dossier Destination Directory", curr)
        if folder:
            self.dest_edit.setText(folder)

    def get_config(self) -> dict:
        """Returns the user configuration dictionary."""
        fmt = "both"
        if self.radio_html.isChecked():
            fmt = "html"
        elif self.radio_pdf.isChecked():
            fmt = "pdf"

        dest = self.dest_edit.text().strip()
        out_dir = Path(dest) if dest else self.meeting_dir / "Export"

        return {
            "export_format": fmt,
            "include_prompt": self.chk_prompt.isChecked(),
            "embed_base64_images": self.chk_b64.isChecked(),
            "download_missing": self.chk_download.isChecked(),
            "force_reconvert": self.chk_force.isChecked(),
            "auto_open": self.chk_auto_open.isChecked(),
            "output_dir": out_dir,
        }
