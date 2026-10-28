"""Cards: Document Format Converter & Embedded Visio Extractor."""

import logging
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt, pyqtSignal
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
    CARD_FRAME_STYLE,
    COMBOBOX_STYLE_TOOLBAR,
)
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane

logger = logging.getLogger(__name__)


class WordVisioCard(QWidget):
    """Card: Extract embedded Visio diagrams (.vsdx) from Microsoft Word documents."""

    extract_visio_requested = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._last_visio_output_dir: Optional[str] = None
        self._setup_ui()

    def _setup_ui(self):
        visio_layout = QVBoxLayout(self)
        visio_layout.setContentsMargins(10, 10, 10, 10)
        visio_layout.setSpacing(10)

        visio_hdr = QFrame()
        visio_hdr.setObjectName("cardFrame")
        visio_hdr.setStyleSheet(CARD_FRAME_STYLE)
        vh_layout = QVBoxLayout(visio_hdr)
        vh_layout.setContentsMargins(14, 12, 14, 12)
        vh_layout.setSpacing(4)
        vh_title = QLabel("<b>🖼️ Embedded Visio & Media Extractor</b>")
        vh_title.setStyleSheet("font-size: 13px; color: #1E293B;")
        vh_desc = QLabel(
            "Extracts embedded Microsoft Visio drawings (.vsdx), OLE diagram objects, and vector media "
            "directly from Word documents into native editable Visio files."
        )
        vh_desc.setWordWrap(True)
        vh_desc.setStyleSheet("color: #475569; font-size: 11px;")
        vh_layout.addWidget(vh_title)
        vh_layout.addWidget(vh_desc)
        visio_layout.addWidget(visio_hdr)

        self.pane_visio = DocumentSelectorPane("📄 WORD DOCUMENT CONTAINING VISIO OBJECTS")
        visio_layout.addWidget(self.pane_visio)

        visio_act_row = QHBoxLayout()
        self.run_visio_btn = QPushButton("🚀 Extract Visio Diagrams")
        self.run_visio_btn.setObjectName("primaryBtn")
        self.run_visio_btn.setCursor(Qt.PointingHandCursor)
        self.run_visio_btn.clicked.connect(self._trigger_visio_extraction)
        visio_act_row.addWidget(self.run_visio_btn)
        visio_act_row.addStretch()
        visio_layout.addLayout(visio_act_row)

        self.visio_result_frame = QFrame()
        self.visio_result_frame.setObjectName("cardFrame")
        self.visio_result_frame.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #F0FDF4;
                border: 1px solid #86EFAC;
                border-radius: 6px;
                padding: 6px 12px;
            }
        """)
        self.visio_result_frame.setVisible(False)
        vr_layout = QHBoxLayout(self.visio_result_frame)
        vr_layout.setContentsMargins(4, 4, 4, 4)

        self.visio_result_icon = QLabel("✅")
        self.visio_result_lbl = QLabel()
        self.visio_result_lbl.setStyleSheet("color: #166534; font-size: 11px; font-weight: bold;")
        self.visio_open_folder_btn = QPushButton("📂 Open Output Folder")
        self.visio_open_folder_btn.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.visio_open_folder_btn.clicked.connect(self._open_visio_folder)

        vr_layout.addWidget(self.visio_result_icon)
        vr_layout.addWidget(self.visio_result_lbl, 1)
        vr_layout.addWidget(self.visio_open_folder_btn)
        visio_layout.addWidget(self.visio_result_frame)
        visio_layout.addStretch()

    def _trigger_visio_extraction(self):
        inputs = self.pane_visio.get_inputs()
        if not inputs:
            QMessageBox.warning(self, "No Document Selected", "Please select or drop a Word document containing Visio objects.")
            return
        self.visio_result_frame.setVisible(False)
        self.extract_visio_requested.emit(inputs[0])

    def on_extraction_success(self, output_path: str):
        self._last_visio_output_dir = output_path
        self.visio_result_lbl.setText(f"Visio extraction complete: {output_path}")
        self.visio_result_frame.setVisible(True)

    def _open_visio_folder(self):
        if self._last_visio_output_dir:
            import os
            folder = Path(self._last_visio_output_dir)
            if folder.is_file():
                folder = folder.parent
            if folder.is_dir() and hasattr(os, "startfile"):
                os.startfile(str(folder))


class WordConverterCard(QWidget):
    """Card: Multi-format Word converter (DOCX, PDF, HTML, TXT, MD)."""

    convert_doc_requested = pyqtSignal(str, str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._last_conv_output_path: Optional[str] = None
        self._setup_ui()

    def _setup_ui(self):
        convert_layout = QVBoxLayout(self)
        convert_layout.setContentsMargins(10, 10, 10, 10)
        convert_layout.setSpacing(10)

        conv_hdr = QFrame()
        conv_hdr.setObjectName("cardFrame")
        conv_hdr.setStyleSheet(CARD_FRAME_STYLE)
        ch_layout = QVBoxLayout(conv_hdr)
        ch_layout.setContentsMargins(14, 12, 14, 12)
        ch_layout.setSpacing(4)
        ch_title = QLabel("<b>🔄 Multi-Format Document Converter</b>")
        ch_title.setStyleSheet("font-size: 13px; color: #1E293B;")
        ch_desc = QLabel(
            "Converts Microsoft Word documents (.docx, legacy .doc) into modern formats: "
            "PDF, HTML, Plain Text, or Markdown."
        )
        ch_desc.setWordWrap(True)
        ch_desc.setStyleSheet("color: #475569; font-size: 11px;")
        ch_layout.addWidget(ch_title)
        ch_layout.addWidget(ch_desc)
        convert_layout.addWidget(conv_hdr)

        self.pane_convert = DocumentSelectorPane("📄 WORD DOCUMENT TO CONVERT")
        convert_layout.addWidget(self.pane_convert)

        conv_opts_frame = QFrame()
        conv_opts_frame.setObjectName("cardFrame")
        conv_opts_frame.setStyleSheet(CARD_FRAME_STYLE)
        co_layout = QHBoxLayout(conv_opts_frame)
        co_layout.setContentsMargins(12, 8, 12, 8)
        co_layout.setSpacing(12)

        co_layout.addWidget(QLabel("<b>Target Format:</b>"))
        self.format_combo = QComboBox()
        self.format_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        for fmt in ["PDF", "DOCX", "MD", "HTML", "XPS", "RTF", "TXT"]:
            self.format_combo.addItem(fmt, fmt.lower())
        self.format_combo.currentIndexChanged.connect(self._on_format_changed)
        co_layout.addWidget(self.format_combo)

        self.conv_engine_lbl = QLabel("Engine: 🟢 Microsoft Word COM")
        self.conv_engine_lbl.setStyleSheet("font-size: 11px; color: #065F46; font-weight: 500;")
        co_layout.addWidget(self.conv_engine_lbl)
        co_layout.addStretch()
        convert_layout.addWidget(conv_opts_frame)

        conv_act_row = QHBoxLayout()
        self.run_convert_btn = QPushButton("🚀 Convert Document")
        self.run_convert_btn.setObjectName("primaryBtn")
        self.run_convert_btn.setCursor(Qt.PointingHandCursor)
        self.run_convert_btn.clicked.connect(self._trigger_conversion)
        conv_act_row.addWidget(self.run_convert_btn)
        conv_act_row.addStretch()
        convert_layout.addLayout(conv_act_row)

        self.conv_result_frame = QFrame()
        self.conv_result_frame.setObjectName("cardFrame")
        self.conv_result_frame.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #F0FDF4;
                border: 1px solid #86EFAC;
                border-radius: 6px;
                padding: 6px 12px;
            }
        """)
        self.conv_result_frame.setVisible(False)
        cr_layout = QHBoxLayout(self.conv_result_frame)
        cr_layout.setContentsMargins(4, 4, 4, 4)

        self.conv_result_icon = QLabel("✅")
        self.conv_result_lbl = QLabel()
        self.conv_result_lbl.setStyleSheet("color: #166534; font-size: 11px; font-weight: bold;")
        self.conv_open_folder_btn = QPushButton("📂 Open Output Folder")
        self.conv_open_folder_btn.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.conv_open_folder_btn.clicked.connect(self._open_conv_folder)

        cr_layout.addWidget(self.conv_result_icon)
        cr_layout.addWidget(self.conv_result_lbl, 1)
        cr_layout.addWidget(self.conv_open_folder_btn)
        convert_layout.addWidget(self.conv_result_frame)
        convert_layout.addStretch()

    def _on_format_changed(self, index: int = 0):
        fmt = self.format_combo.currentText().upper()
        if fmt == "DOCX":
            self.conv_engine_lbl.setText("Engine: 🟠 LibreOffice (for legacy .doc) / Word COM")
            self.conv_engine_lbl.setStyleSheet("font-size: 11px; color: #B45309; font-weight: 500;")
        elif fmt == "MD":
            self.conv_engine_lbl.setText("Engine: 🟢 Native High-Fidelity Markdown (with Media)")
            self.conv_engine_lbl.setStyleSheet("font-size: 11px; color: #0F766E; font-weight: 500;")
        else:
            self.conv_engine_lbl.setText("Engine: 🟢 Microsoft Word COM")
            self.conv_engine_lbl.setStyleSheet("font-size: 11px; color: #065F46; font-weight: 500;")

    def _trigger_conversion(self):
        inputs = self.pane_convert.get_inputs()
        if not inputs:
            QMessageBox.warning(self, "No Document Selected", "Please select or drop a Word document to convert.")
            return
        target_format = self.format_combo.currentData() or "docx"
        self.conv_result_frame.setVisible(False)
        self.convert_doc_requested.emit(inputs[0], target_format)

    def on_conversion_success(self, output_path: str):
        self._last_conv_output_path = output_path
        self.conv_result_lbl.setText(f"Conversion complete: {Path(output_path).name}")
        self.conv_result_frame.setVisible(True)

    def _open_conv_folder(self):
        if self._last_conv_output_path:
            import os
            folder = Path(self._last_conv_output_path).parent
            if folder.is_dir() and hasattr(os, "startfile"):
                os.startfile(str(folder))
