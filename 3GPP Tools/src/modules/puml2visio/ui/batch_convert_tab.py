# --- File: src/modules/puml2visio/ui/batch_convert_tab.py ---
"""Modernized Visio & Office Conversion Studio.

Features:
- Real-time System Diagnostics Ribbon (Visio COM, PowerPoint COM, PlantUML Engine, Output path)
- Multi-Card Operations Architecture:
  * Card 1: Batch PlantUML Converter (multi-target: VSDX, PPTX, SVG, ASCII; staging queue with diagram type detection)
  * Card 2: PowerPoint ⇄ Visio Bridge (two-way vector conversion via EMF bridging)
  * Card 3: Visio Source Extractor & Reverse Engineering (instant headless OpenXML extraction + syntax preview)
  * Card 4: 3GPP Stencils & Architecture Reference (5GC, NG-RAN, EPS NF entity catalog with 1-click insertion)
- Recent Exports Drawer (1-click Open, Show in Folder, Copy Path)
"""

import datetime
import logging
import os
import re
import subprocess
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import Qt, QSize, QTimer, pyqtSignal
from PyQt5.QtGui import QColor, QCursor, QFont, QIcon
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    TABLE_STYLE_CLEAN,
    InteractiveDropLabel,
)
from core.utils.paths import get_project_root
from core.utils.utils import get_detailed_java_info
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.core.visio_converter import VisioReaderThread
from modules.puml2visio.templates.plantuml_snippets import (
    PLANTUML_SNIPPETS,
    PlantUMLSnippet,
    register_snippets_reloaded_callback,
    reload_snippets,
)
from modules.puml2visio.ui.preset_dialog import (
    UserPresetDialog,
    open_presets_folder,
)
from modules.puml2visio.ui.syntax_highlighter import PlantUMLHighlighter
from modules.puml2visio.utils.paths import get_puml2visio_asset_path
from modules.puml2visio.utils.utils import (
    get_plantuml_version,
    strip_watermark,
)

logger = logging.getLogger(__name__)


# =============================================================================
# Helper Utilities
# =============================================================================

def is_visio_registered() -> bool:
    """Non-fatal check whether Microsoft Visio COM server is registered in Windows."""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Visio.Application")
        winreg.CloseKey(key)
        return True
    except (FileNotFoundError, OSError, ImportError):
        return False


def is_powerpoint_registered() -> bool:
    """Non-fatal check whether Microsoft PowerPoint COM server is registered in Windows."""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application")
        winreg.CloseKey(key)
        return True
    except (FileNotFoundError, OSError, ImportError):
        return False


def detect_puml_diagram_type(file_path: Path) -> str:
    """Quickly scans the opening lines of a PlantUML file to infer the diagram type."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = [f.readline() for _ in range(35)]
        content = " ".join(lines).lower()
        if "sequence" in content or "->" in content or "autonumber" in content:
            return "Sequence"
        elif "class " in content or "interface " in content:
            return "Class"
        elif "start" in content and "stop" in content:
            return "Activity"
        elif "[*]" in content or "state " in content:
            return "State"
        elif "usecase" in content or "actor " in content:
            return "Use Case"
        elif "package " in content or "node " in content or "component " in content:
            return "Component / Arch"
    except Exception:
        pass
    return "PlantUML"


def extract_plantuml_from_vsdx_fast(vsdx_path: Path) -> Optional[str]:
    """Headless OpenXML extraction of embedded PlantUML source from a .vsdx archive.

    Runs in under 5ms without spawning any COM Visio process.
    """
    try:
        with zipfile.ZipFile(vsdx_path, "r") as z:
            names = z.namelist()
            for name in names:
                if name.startswith("visio/pages/page") and name.endswith(".xml"):
                    content = z.read(name)
                    if b"@startuml" in content:
                        root = ET.fromstring(content)
                        text = "".join(root.itertext())
                        if "@startuml" in text:
                            return strip_watermark(text)
    except Exception as e:
        logger.debug(f"Fast vsdx extract failed for {vsdx_path.name}: {e}")
    return None


def format_file_size(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    elif num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    else:
        return f"{num_bytes / (1024 * 1024):.1f} MB"


# =============================================================================
# Main Class: BatchConvertTab
# =============================================================================

class BatchConvertTab(QWidget):
    """Modernized Visio & Office Conversion Studio with multi-card workflow,

    system diagnostics ribbon, staging table, and reverse-engineering inspector.
    """

    # Backward-compatible & new custom signals
    files_dropped = pyqtSignal(list, str)
    send_to_editor_requested = pyqtSignal(str)
    insert_snippet_requested = pyqtSignal(object)
    java_maintenance_requested = pyqtSignal()
    open_folder_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._export_dir = get_project_root().parent / "output"
        self._staged_puml_files: List[Path] = []
        self._staged_bridge_files: List[Path] = []
        self._reader_thread: Optional[VisioReaderThread] = None

        self._setup_ui()
        self._refresh_system_diagnostics()
        self._refresh_recent_exports()

    # -------------------------------------------------------------------------
    # UI Setup
    # -------------------------------------------------------------------------
    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        # 1. Top System Diagnostics & Configuration Ribbon
        self._setup_diagnostics_ribbon(main_layout)

        # 2. Operation Mode Switcher Bar
        self._setup_mode_switcher(main_layout)

        # 3. Stacked Operation Cards
        self.stack = QStackedWidget()
        self._setup_card_puml_batch()
        self._setup_card_office_bridge()
        self._setup_card_vsdx_extractor()
        self._setup_card_stencils()

        main_layout.addWidget(self.stack, stretch=1)

        # 4. Bottom Recent Exports Strip
        self._setup_recent_exports_strip(main_layout)

    # -------------------------------------------------------------------------
    # 1. Diagnostics Ribbon
    # -------------------------------------------------------------------------
    def _setup_diagnostics_ribbon(self, parent_layout: QVBoxLayout):
        ribbon_frame = QFrame()
        ribbon_frame.setObjectName("cardFrame")
        ribbon_frame.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #F8FAFC;
                padding: 4px 10px;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
            }
        """)
        ribbon_layout = QHBoxLayout(ribbon_frame)
        ribbon_layout.setContentsMargins(8, 5, 8, 5)
        ribbon_layout.setSpacing(10)

        # Visio COM Status Badge
        self.visio_badge = QLabel("🔍 Checking Visio...")
        self.visio_badge.setStyleSheet(
            "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #E2E8F0; color: #475569;"
        )
        ribbon_layout.addWidget(self.visio_badge)

        # PowerPoint COM Status Badge
        self.ppt_badge = QLabel("🔍 Checking PowerPoint...")
        self.ppt_badge.setStyleSheet(
            "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #E2E8F0; color: #475569;"
        )
        ribbon_layout.addWidget(self.ppt_badge)

        # PlantUML Engine & Java Badge (Clickable)
        self.java_badge_btn = QPushButton("☕ PlantUML Engine")
        self.java_badge_btn.setStyleSheet("""
            QPushButton {
                background-color: #EFF6FF;
                color: #1D4ED8;
                border: 1px solid #BFDBFE;
                border-radius: 12px;
                padding: 3px 10px;
                font-weight: 600;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #DBEAFE;
                border-color: #93C5FD;
            }
        """)
        self.java_badge_btn.setToolTip("Click to inspect Java runtime and PlantUML JAR updates.")
        self.java_badge_btn.clicked.connect(self.java_maintenance_requested.emit)
        ribbon_layout.addWidget(self.java_badge_btn)

        ribbon_layout.addStretch()

        # Output Folder Display & Action
        out_icon = QLabel("📁")
        out_lbl_title = QLabel("Output:")
        out_lbl_title.setStyleSheet("font-weight: bold; color: #64748B; font-size: 11px;")
        self.out_path_lbl = QLabel(str(self._export_dir.name) + "/")
        self.out_path_lbl.setStyleSheet("color: #1E293B; font-size: 12px; font-family: monospace;")
        self.out_path_lbl.setToolTip(str(self._export_dir.resolve()))

        self.open_folder_btn = QPushButton("📂 Open Folder")
        self.open_folder_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.open_folder_btn.clicked.connect(self._on_open_output_folder)

        self.change_folder_btn = QPushButton("⚙️ Change...")
        self.change_folder_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.change_folder_btn.setToolTip("Change destination folder for exported diagrams")
        self.change_folder_btn.clicked.connect(self._on_change_output_folder)

        ribbon_layout.addWidget(out_icon)
        ribbon_layout.addWidget(out_lbl_title)
        ribbon_layout.addWidget(self.out_path_lbl)
        ribbon_layout.addWidget(self.open_folder_btn)
        ribbon_layout.addWidget(self.change_folder_btn)

        parent_layout.addWidget(ribbon_frame)

    # -------------------------------------------------------------------------
    # 2. Mode Switcher
    # -------------------------------------------------------------------------
    def _setup_mode_switcher(self, parent_layout: QVBoxLayout):
        switcher_layout = QHBoxLayout()
        switcher_layout.setSpacing(6)

        mode_title = QLabel("<b>⚙️ Operation Mode:</b>")
        mode_title.setStyleSheet("color: #334155; font-size: 12px;")
        switcher_layout.addWidget(mode_title)

        self.mode_btn_group = QButtonGroup(self)
        self.mode_btn_group.setExclusive(True)

        modes = [
            ("📊 Batch PlantUML Converter", 0),
            ("⇄ Office Bridge (PPTX ⇄ Visio)", 1),
            ("🔍 Visio Source Extractor", 2),
            ("📐 3GPP Stencils Reference", 3),
        ]

        btn_style = """
            QPushButton {
                background-color: #F8FAFC;
                color: #475569;
                border: 1px solid #CBD5E1;
                border-radius: 5px;
                padding: 5px 12px;
                font-weight: 600;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                color: #1E293B;
            }
            QPushButton:checked {
                background-color: #EBF3FC;
                border: 2px solid #395396;
                color: #395396;
            }
        """

        for label, idx in modes:
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setStyleSheet(btn_style)
            if idx == 0:
                btn.setChecked(True)
            btn.clicked.connect(lambda checked, i=idx: self.stack.setCurrentIndex(i))
            self.mode_btn_group.addButton(btn, idx)
            switcher_layout.addWidget(btn)

        switcher_layout.addStretch()
        parent_layout.addLayout(switcher_layout)

    # -------------------------------------------------------------------------
    # Card 1: Batch PlantUML Converter
    # -------------------------------------------------------------------------
    def _setup_card_puml_batch(self):
        card = QWidget()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Options & Target Format Ribbon
        top_bar = QHBoxLayout()
        top_bar.setSpacing(10)

        top_bar.addWidget(QLabel("<b>Target Format:</b>"))
        self.puml_target_combo = QComboBox()
        self.puml_target_combo.addItems([
            "Microsoft Visio (.vsdx)",
            "Microsoft PowerPoint (.pptx)",
            "Scalable Vector Graphic (.svg)",
            "Unicode Text Art (.txt)",
        ])
        self.puml_target_combo.setStyleSheet("font-size: 12px; padding: 3px 8px;")
        top_bar.addWidget(self.puml_target_combo)

        self.chk_embed_source = QCheckBox("Embed PlantUML in Visio page")
        self.chk_embed_source.setChecked(True)
        self.chk_embed_source.setToolTip("Embeds original PlantUML source code into a hidden Visio page for 100% lossless reverse recovery.")
        top_bar.addWidget(self.chk_embed_source)

        self.chk_auto_open = QCheckBox("Auto-open output upon completion")
        self.chk_auto_open.setChecked(True)
        top_bar.addWidget(self.chk_auto_open)

        top_bar.addStretch()
        layout.addLayout(top_bar)

        # Drop Zone (Compact)
        self.drop_label_puml = InteractiveDropLabel(
            "📥 Drag & Drop PlantUML file(s) (.puml, .plantuml, .txt) here to stage for batch conversion",
            [".puml", ".plantuml", ".txt"],
        )
        self.drop_label_puml.setMaximumHeight(80)
        self.drop_label_puml.file_dropped.connect(self._on_puml_files_dropped)
        layout.addWidget(self.drop_label_puml)

        # Buttons Toolbar
        btn_bar = QHBoxLayout()
        btn_bar.setSpacing(6)

        self.add_files_btn = QPushButton("📂 Add Files...")
        self.add_files_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.add_files_btn.clicked.connect(self._browse_puml_files)

        self.add_folder_btn = QPushButton("📁 Add Folder...")
        self.add_folder_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.add_folder_btn.clicked.connect(self._browse_puml_folder)

        self.clear_staged_btn = QPushButton("🗑️ Clear Staging")
        self.clear_staged_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_staged_btn.clicked.connect(self._clear_staged_puml)

        self.staged_count_lbl = QLabel("0 diagrams staged")
        self.staged_count_lbl.setStyleSheet("color: #64748B; font-size: 11px; margin-left: 6px;")

        btn_bar.addWidget(self.add_files_btn)
        btn_bar.addWidget(self.add_folder_btn)
        btn_bar.addWidget(self.clear_staged_btn)
        btn_bar.addWidget(self.staged_count_lbl)
        btn_bar.addStretch()

        self.convert_puml_btn = QPushButton("🚀 Convert All Staged Diagrams")
        self.convert_puml_btn.setObjectName("primaryBtn")
        self.convert_puml_btn.setEnabled(False)
        self.convert_puml_btn.clicked.connect(self._start_puml_batch_conversion)
        btn_bar.addWidget(self.convert_puml_btn)

        layout.addLayout(btn_bar)

        # Staging Table
        self.puml_table = QTableWidget(0, 5)
        self.puml_table.setHorizontalHeaderLabels(["#", "Filename", "Diagram Type", "Size", "Action"])
        self.puml_table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.puml_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.puml_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        header = self.puml_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        layout.addWidget(self.puml_table, stretch=1)
        self.stack.addWidget(card)

    # -------------------------------------------------------------------------
    # Card 2: PowerPoint ⇄ Visio Bridge
    # -------------------------------------------------------------------------
    def _setup_card_office_bridge(self):
        card = QWidget()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        # Direction Switcher & Banner
        dir_frame = QFrame()
        dir_frame.setObjectName("cardFrame")
        dir_frame.setStyleSheet(CARD_FRAME_STYLE + "QFrame#cardFrame { background-color: #F8FAFC; padding: 8px; }")
        dir_layout = QVBoxLayout(dir_frame)
        dir_layout.setSpacing(6)

        opts_layout = QHBoxLayout()
        opts_layout.addWidget(QLabel("<b>Conversion Direction:</b>"))

        self.rb_pptx2visio = QRadioButton("📊 PowerPoint (.pptx) ➔ Microsoft Visio (.vsdx)")
        self.rb_pptx2visio.setChecked(True)
        self.rb_pptx2visio.toggled.connect(self._on_bridge_direction_changed)
        opts_layout.addWidget(self.rb_pptx2visio)

        self.rb_visio2pptx = QRadioButton("📐 Microsoft Visio (.vsdx) ➔ PowerPoint Presentation (.pptx)")
        self.rb_visio2pptx.toggled.connect(self._on_bridge_direction_changed)
        opts_layout.addWidget(self.rb_visio2pptx)
        opts_layout.addStretch()

        dir_layout.addLayout(opts_layout)

        info_lbl = QLabel(
            "💡 <b>EMF Lossless Vector Bridging:</b> Converts multi-page presentations or Visio drawings into "
            "native editable Office vector shapes with automated ungrouping and text shrink-wrapping."
        )
        info_lbl.setWordWrap(True)
        info_lbl.setStyleSheet("color: #475569; font-size: 11px;")
        dir_layout.addWidget(info_lbl)

        layout.addWidget(dir_frame)

        # Drop Zone for Bridge
        self.drop_label_bridge = InteractiveDropLabel(
            "📥 Drag & Drop PowerPoint presentation (.pptx) here to convert to multi-page Visio (.vsdx)",
            [".pptx"],
        )
        self.drop_label_bridge.file_dropped.connect(self._on_bridge_files_dropped)
        layout.addWidget(self.drop_label_bridge)

        # Actions Row
        bridge_btn_bar = QHBoxLayout()
        self.add_bridge_file_btn = QPushButton("📂 Browse File(s)...")
        self.add_bridge_file_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.add_bridge_file_btn.clicked.connect(self._browse_bridge_files)

        self.clear_bridge_btn = QPushButton("🗑️ Clear")
        self.clear_bridge_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_bridge_btn.clicked.connect(self._clear_staged_bridge)

        self.bridge_count_lbl = QLabel("0 presentation(s) staged")
        self.bridge_count_lbl.setStyleSheet("color: #64748B; font-size: 11px;")

        bridge_btn_bar.addWidget(self.add_bridge_file_btn)
        bridge_btn_bar.addWidget(self.clear_bridge_btn)
        bridge_btn_bar.addWidget(self.bridge_count_lbl)
        bridge_btn_bar.addStretch()

        self.convert_bridge_btn = QPushButton("🚀 Convert Office Presentation")
        self.convert_bridge_btn.setObjectName("primaryBtn")
        self.convert_bridge_btn.setEnabled(False)
        self.convert_bridge_btn.clicked.connect(self._start_bridge_conversion)
        bridge_btn_bar.addWidget(self.convert_bridge_btn)

        layout.addLayout(bridge_btn_bar)

        # Staged Bridge Table
        self.bridge_table = QTableWidget(0, 4)
        self.bridge_table.setHorizontalHeaderLabels(["#", "File Path", "Size", "Action"])
        self.bridge_table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.bridge_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.bridge_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        b_header = self.bridge_table.horizontalHeader()
        b_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        b_header.setSectionResizeMode(1, QHeaderView.Stretch)
        b_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        b_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        layout.addWidget(self.bridge_table, stretch=1)
        self.stack.addWidget(card)

    # -------------------------------------------------------------------------
    # Card 3: Visio PlantUML Extractor & Reverse Engineering
    # -------------------------------------------------------------------------
    def _setup_card_vsdx_extractor(self):
        card = QWidget()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Header Info Banner
        header_box = QHBoxLayout()
        header_desc = QLabel(
            "<b>Reverse Engineering:</b> Extract original PlantUML source code embedded in .vsdx diagrams generated by 3GPP Tools."
        )
        header_desc.setStyleSheet("color: #334155; font-size: 12px;")
        header_box.addWidget(header_desc)
        header_box.addStretch()

        layout.addLayout(header_box)

        # Drop Zone & File Selector
        drop_row = QHBoxLayout()
        self.drop_label_extractor = InteractiveDropLabel(
            "📥 Drag & Drop a generated Visio (.vsdx) file here to extract its PlantUML source code",
            [".vsdx"],
        )
        self.drop_label_extractor.setMaximumHeight(65)
        self.drop_label_extractor.file_dropped.connect(self._on_extractor_file_dropped)
        drop_row.addWidget(self.drop_label_extractor, 1)

        self.browse_vsdx_btn = QPushButton("📂 Browse .vsdx...")
        self.browse_vsdx_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.browse_vsdx_btn.clicked.connect(self._browse_vsdx_for_extraction)
        drop_row.addWidget(self.browse_vsdx_btn)

        layout.addLayout(drop_row)

        # Extractor Status & Actions Bar
        status_bar = QHBoxLayout()
        self.extractor_status_lbl = QLabel("No file loaded.")
        self.extractor_status_lbl.setStyleSheet("font-weight: 600; color: #64748B; font-size: 11px;")
        status_bar.addWidget(self.extractor_status_lbl)
        status_bar.addStretch()

        self.send_to_editor_btn = QPushButton("📝 Open in PlantUML IDE Tab")
        self.send_to_editor_btn.setObjectName("primaryBtn")
        self.send_to_editor_btn.setEnabled(False)
        self.send_to_editor_btn.setToolTip("Switch to the PlantUML editor tab with this code loaded for live preview and editing.")
        self.send_to_editor_btn.clicked.connect(self._on_send_to_editor_clicked)

        self.copy_extracted_btn = QPushButton("📋 Copy Code")
        self.copy_extracted_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_extracted_btn.setEnabled(False)
        self.copy_extracted_btn.clicked.connect(self._on_copy_extracted_clicked)

        self.save_extracted_btn = QPushButton("💾 Save as .puml...")
        self.save_extracted_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.save_extracted_btn.setEnabled(False)
        self.save_extracted_btn.clicked.connect(self._on_save_extracted_clicked)

        status_bar.addWidget(self.send_to_editor_btn)
        status_bar.addWidget(self.copy_extracted_btn)
        status_bar.addWidget(self.save_extracted_btn)

        layout.addLayout(status_bar)

        # Syntax-Highlighted Code Viewer
        self.extracted_text_edit = QPlainTextEdit()
        self.extracted_text_edit.setPlaceholderText("Extracted PlantUML source code will appear here...")
        self.extracted_text_edit.setStyleSheet("""
            QPlainTextEdit {
                font-family: Consolas, 'Cascadia Code', 'Courier New', monospace;
                font-size: 12px;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                background-color: #FFFFFF;
                color: #1E293B;
                line-height: 1.4;
            }
        """)
        self.extracted_highlighter = PlantUMLHighlighter(self.extracted_text_edit.document())
        layout.addWidget(self.extracted_text_edit, stretch=1)

        self.stack.addWidget(card)

    # -------------------------------------------------------------------------
    # Card 4: 3GPP Stencils & Architecture Reference
    # -------------------------------------------------------------------------
    def _setup_card_stencils(self):
        card = QWidget()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Header description
        desc_lbl = QLabel(
            "<b>3GPP Architecture Presets & Stencils:</b> Quick-insert standardized 5G Core, NG-RAN, EPS, and IMS "
            "Network Function definitions into your PlantUML diagrams. Supports custom user-defined presets "
            "(saved to Git-ignored config)."
        )
        desc_lbl.setStyleSheet("color: #334155; font-size: 12px;")
        layout.addWidget(desc_lbl)

        # Toolbar: Filter search, category combo, add preset, reload, open folder
        tb_layout = QHBoxLayout()
        tb_layout.setSpacing(8)

        self.stencils_search_edit = QLineEdit()
        self.stencils_search_edit.setPlaceholderText("Filter stencils (e.g. UE, AMF, gNB, 5GS)...")
        self.stencils_search_edit.setStyleSheet("padding: 4px 8px; font-size: 11px; border: 1px solid #CBD5E1; border-radius: 4px;")
        self.stencils_search_edit.textChanged.connect(self._populate_stencils_table)
        tb_layout.addWidget(self.stencils_search_edit, stretch=2)

        self.stencils_cat_combo = QComboBox()
        self.stencils_cat_combo.addItems([
            "All Categories",
            "3GPP Network Entities & NFs",
            "3GPP Architecture Presets",
            "Participants & Grouping",
            "⭐ User Defined Presets",
        ])
        self.stencils_cat_combo.setStyleSheet("padding: 4px 8px; font-size: 11px;")
        self.stencils_cat_combo.currentIndexChanged.connect(self._populate_stencils_table)
        tb_layout.addWidget(self.stencils_cat_combo, stretch=1)

        btn_add = QPushButton("➕ Add User Preset...")
        btn_add.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_add.setToolTip("Create a custom PlantUML preset saved to Git-ignored configuration")
        btn_add.clicked.connect(self._open_add_stencil_dialog)
        tb_layout.addWidget(btn_add)

        btn_reload = QPushButton("🔄 Reload")
        btn_reload.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_reload.setToolTip("Reload presets from disk (plantuml_presets.json and user presets)")
        btn_reload.clicked.connect(self._reload_stencils)
        tb_layout.addWidget(btn_reload)

        btn_folder = QPushButton("📂 Open Config")
        btn_folder.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_folder.setToolTip("Open configuration folder in Windows Explorer")
        btn_folder.clicked.connect(open_presets_folder)
        tb_layout.addWidget(btn_folder)

        layout.addLayout(tb_layout)

        # Stencils Table / List
        self.stencils_table = QTableWidget(0, 4)
        self.stencils_table.setHorizontalHeaderLabels(["Category", "Preset Name", "Description", "Actions"])
        self.stencils_table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.stencils_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.stencils_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        s_header = self.stencils_table.horizontalHeader()
        s_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        s_header.setSectionResizeMode(2, QHeaderView.Stretch)
        s_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        # Populate with 3GPP snippets
        self._populate_stencils_table()

        # Listen for global reload
        register_snippets_reloaded_callback(self._populate_stencils_table)

        layout.addWidget(self.stencils_table, stretch=1)
        self.stack.addWidget(card)

    def _open_add_stencil_dialog(self):
        dialog = UserPresetDialog(parent=self)
        if dialog.exec_():
            self._populate_stencils_table()

    def _reload_stencils(self):
        reload_snippets()
        self._populate_stencils_table()

    def _edit_stencil_dialog(self, snippet: PlantUMLSnippet):
        dialog = UserPresetDialog(existing_snippet=snippet, parent=self)
        if dialog.exec_():
            self._populate_stencils_table()

    def _populate_stencils_table(self):
        self.stencils_table.setRowCount(0)
        query = self.stencils_search_edit.text().strip().lower() if hasattr(self, 'stencils_search_edit') else ""
        selected_cat = self.stencils_cat_combo.currentText() if hasattr(self, 'stencils_cat_combo') else "All Categories"

        relevant_snippets = []
        for s in PLANTUML_SNIPPETS:
            # Filter by category combo
            if selected_cat == "⭐ User Defined Presets":
                if not s.is_user_defined:
                    continue
            elif selected_cat != "All Categories":
                if s.category != selected_cat:
                    continue
            else:
                # In All Categories mode, show 3GPP and participant stencils plus any user-defined
                if not ("3gpp" in s.category.lower() or "participant" in s.category.lower() or s.is_user_defined):
                    continue

            # Filter by search text
            if query:
                combined = f"{s.label} {s.category} {s.description} {s.code}".lower()
                if query not in combined:
                    continue

            relevant_snippets.append(s)

        for s in relevant_snippets:
            row = self.stencils_table.rowCount()
            self.stencils_table.insertRow(row)

            cat_item = QTableWidgetItem(s.category)
            display_name = f"⭐ {s.label} [User]" if s.is_user_defined else s.label
            name_item = QTableWidgetItem(display_name)
            name_item.setFont(QFont("Segoe UI", 9, QFont.Bold))
            if s.is_user_defined:
                name_item.setForeground(QColor("#B45309"))

            desc_item = QTableWidgetItem(s.description)
            desc_item.setToolTip(s.code)

            self.stencils_table.setItem(row, 0, cat_item)
            self.stencils_table.setItem(row, 1, name_item)
            self.stencils_table.setItem(row, 2, desc_item)

            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(2, 2, 2, 2)
            action_layout.setSpacing(4)

            btn_copy = QPushButton("📋 Copy")
            btn_copy.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            btn_copy.setToolTip("Copy this PlantUML snippet to clipboard")
            btn_copy.clicked.connect(lambda checked, snip=s: self._copy_snippet_to_clipboard(snip))
            action_layout.addWidget(btn_copy)

            btn_insert = QPushButton("📝 Insert in Editor")
            btn_insert.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            btn_insert.setToolTip("Switch to PlantUML editor and insert this element at the cursor")
            btn_insert.clicked.connect(lambda checked, snip=s: self._insert_snippet_into_editor(snip))
            action_layout.addWidget(btn_insert)

            if s.is_user_defined:
                btn_edit = QPushButton("✏️ Edit")
                btn_edit.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
                btn_edit.setToolTip("Edit this user-defined preset")
                btn_edit.clicked.connect(lambda checked, snip=s: self._edit_stencil_dialog(snip))
                action_layout.addWidget(btn_edit)

            self.stencils_table.setCellWidget(row, 3, action_widget)

    # -------------------------------------------------------------------------
    # 4. Recent Exports Strip
    # -------------------------------------------------------------------------
    def _setup_recent_exports_strip(self, parent_layout: QVBoxLayout):
        strip_frame = QFrame()
        strip_frame.setObjectName("cardFrame")
        strip_frame.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 4px 8px;
            }
        """)
        strip_layout = QVBoxLayout(strip_frame)
        strip_layout.setContentsMargins(6, 4, 6, 4)
        strip_layout.setSpacing(4)

        header_row = QHBoxLayout()
        header_lbl = QLabel("<b>🕒 Recent Exports:</b>")
        header_lbl.setStyleSheet("color: #475569; font-size: 11px;")

        self.recent_count_lbl = QLabel("(0 files)")
        self.recent_count_lbl.setStyleSheet("color: #94A3B8; font-size: 11px;")

        self.refresh_recent_btn = QPushButton("🔄 Refresh")
        self.refresh_recent_btn.setStyleSheet("font-size: 10px; padding: 1px 6px;")
        self.refresh_recent_btn.clicked.connect(self._refresh_recent_exports)

        header_row.addWidget(header_lbl)
        header_row.addWidget(self.recent_count_lbl)
        header_row.addStretch()
        header_row.addWidget(self.refresh_recent_btn)
        strip_layout.addLayout(header_row)

        self.recent_container = QVBoxLayout()
        self.recent_container.setSpacing(2)
        strip_layout.addLayout(self.recent_container)

        parent_layout.addWidget(strip_frame)

    # -------------------------------------------------------------------------
    # Logic: System Diagnostics
    # -------------------------------------------------------------------------
    def _refresh_system_diagnostics(self):
        # 1. Visio COM
        if is_visio_registered():
            self.visio_badge.setText("🟢 Visio Ready")
            self.visio_badge.setStyleSheet(
                "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #DCFCE7; color: #166534;"
            )
            self.visio_badge.setToolTip("Microsoft Visio COM server is registered and ready.")
        else:
            self.visio_badge.setText("🟡 Visio Not Detected")
            self.visio_badge.setStyleSheet(
                "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #FEF3C7; color: #92400E;"
            )
            self.visio_badge.setToolTip("Microsoft Visio is not registered on this system. PlantUML, SVG, and PPTX conversions remain available.")

        # 2. PowerPoint COM
        if is_powerpoint_registered():
            self.ppt_badge.setText("🟢 PowerPoint Ready")
            self.ppt_badge.setStyleSheet(
                "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #DCFCE7; color: #166534;"
            )
            self.ppt_badge.setToolTip("Microsoft PowerPoint COM server is registered and ready.")
        else:
            self.ppt_badge.setText("⚪ PowerPoint Not Detected")
            self.ppt_badge.setStyleSheet(
                "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #F1F5F9; color: #64748B;"
            )
            self.ppt_badge.setToolTip("Microsoft PowerPoint is not registered on this system.")

        # 3. Java & PlantUML JAR
        java_info = get_detailed_java_info()
        jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
        if java_info.get("installed") and jar_path.exists():
            v = get_plantuml_version(java_info.get("path"), jar_path)
            self.java_badge_btn.setText(f"☕ PlantUML Engine: v{v or 'Detected'}")
            self.java_badge_btn.setStyleSheet("""
                QPushButton {
                    background-color: #DCFCE7;
                    color: #166534;
                    border: 1px solid #BBF7D0;
                    border-radius: 12px;
                    padding: 3px 10px;
                    font-weight: 600;
                    font-size: 11px;
                }
                QPushButton:hover {
                    background-color: #BBF7D0;
                }
            """)
        else:
            self.java_badge_btn.setText("⚠️ PlantUML Engine: Check Java/JAR")
            self.java_badge_btn.setStyleSheet("""
                QPushButton {
                    background-color: #FEF3C7;
                    color: #92400E;
                    border: 1px solid #FDE68A;
                    border-radius: 12px;
                    padding: 3px 10px;
                    font-weight: 600;
                    font-size: 11px;
                }
                QPushButton:hover {
                    background-color: #FDE68A;
                }
            """)

    # -------------------------------------------------------------------------
    # Logic: Card 1 (Batch PlantUML)
    # -------------------------------------------------------------------------
    def _on_puml_files_dropped(self, paths: list):
        for p in paths:
            path = Path(p)
            if path.is_file() and path not in self._staged_puml_files:
                self._staged_puml_files.append(path)
        self._update_puml_table()

    def _browse_puml_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PlantUML Diagram Files",
            str(self._export_dir.resolve()),
            "PlantUML Diagrams (*.puml *.plantuml *.txt);;All Files (*.*)",
        )
        if files:
            self._on_puml_files_dropped(files)

    def _browse_puml_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Folder with PlantUML Diagrams", str(self._export_dir.resolve())
        )
        if folder:
            p_folder = Path(folder)
            found = []
            for ext in (".puml", ".plantuml", ".txt"):
                found.extend(list(p_folder.glob(f"*{ext}")))
            if found:
                self._on_puml_files_dropped(found)
            else:
                QMessageBox.information(
                    self, "No Diagrams Found", f"No .puml or .txt diagram files found in:\n{folder}"
                )

    def _clear_staged_puml(self):
        self._staged_puml_files.clear()
        self._update_puml_table()

    def _remove_staged_puml_index(self, idx: int):
        if 0 <= idx < len(self._staged_puml_files):
            self._staged_puml_files.pop(idx)
            self._update_puml_table()

    def _update_puml_table(self):
        self.puml_table.setRowCount(0)
        for i, file_path in enumerate(self._staged_puml_files, start=1):
            row = self.puml_table.rowCount()
            self.puml_table.insertRow(row)

            # 0. Index
            idx_item = QTableWidgetItem(str(i))
            idx_item.setTextAlignment(Qt.AlignCenter)

            # 1. Filename
            fn_item = QTableWidgetItem(file_path.name)
            fn_item.setToolTip(str(file_path.resolve()))

            # 2. Diagram Type
            dtype = detect_puml_diagram_type(file_path)
            dt_item = QTableWidgetItem(dtype)
            dt_item.setTextAlignment(Qt.AlignCenter)

            # 3. Size
            sz_str = format_file_size(file_path.stat().st_size) if file_path.exists() else "—"
            sz_item = QTableWidgetItem(sz_str)
            sz_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            # 4. Remove Action
            btn_remove = QPushButton("✕")
            btn_remove.setFixedSize(24, 22)
            btn_remove.setStyleSheet("font-size: 11px; padding: 0px; color: #EF4444; font-weight: bold;")
            btn_remove.setToolTip("Remove from staging")
            btn_remove.clicked.connect(lambda checked, idx=row: self._remove_staged_puml_index(idx))

            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(4, 2, 4, 2)
            action_layout.setAlignment(Qt.AlignCenter)
            action_layout.addWidget(btn_remove)

            self.puml_table.setItem(row, 0, idx_item)
            self.puml_table.setItem(row, 1, fn_item)
            self.puml_table.setItem(row, 2, dt_item)
            self.puml_table.setItem(row, 3, sz_item)
            self.puml_table.setCellWidget(row, 4, action_widget)

        count = len(self._staged_puml_files)
        self.staged_count_lbl.setText(f"{count} diagram(s) staged")
        self.convert_puml_btn.setEnabled(count > 0)

    def _start_puml_batch_conversion(self):
        if not self._staged_puml_files:
            return

        fmt_choice = self.puml_target_combo.currentText()
        if ".vsdx" in fmt_choice:
            target_fmt = "vsdx"
        elif ".pptx" in fmt_choice:
            target_fmt = "pptx"
        elif ".svg" in fmt_choice:
            target_fmt = "svg"
        else:
            target_fmt = "ascii"

        paths_to_send = list(self._staged_puml_files)
        self.files_dropped.emit(paths_to_send, target_fmt)
        self._staged_puml_files.clear()
        self._update_puml_table()

    # -------------------------------------------------------------------------
    # Logic: Card 2 (Office Bridge)
    # -------------------------------------------------------------------------
    def _on_bridge_direction_changed(self):
        if self.rb_pptx2visio.isChecked():
            self.drop_label_bridge.accepted_extensions = [".pptx"]
            self.drop_label_bridge.setText(
                "📥 Drag & Drop PowerPoint presentation (.pptx) here to convert to multi-page Visio (.vsdx)"
            )
            self.convert_bridge_btn.setText("🚀 Convert PowerPoint to Visio (.vsdx)")
        else:
            self.drop_label_bridge.accepted_extensions = [".vsdx"]
            self.drop_label_bridge.setText(
                "📥 Drag & Drop Visio diagram (.vsdx) here to convert to PowerPoint presentation (.pptx)"
            )
            self.convert_bridge_btn.setText("🚀 Convert Visio to PowerPoint (.pptx)")
        self._staged_bridge_files.clear()
        self._update_bridge_table()

    def _on_bridge_files_dropped(self, paths: list):
        for p in paths:
            path = Path(p)
            if path.is_file() and path not in self._staged_bridge_files:
                self._staged_bridge_files.append(path)
        self._update_bridge_table()

    def _browse_bridge_files(self):
        if self.rb_pptx2visio.isChecked():
            filt = "PowerPoint Presentations (*.pptx);;All Files (*.*)"
            title = "Select PowerPoint Presentations"
        else:
            filt = "Visio Diagrams (*.vsdx);;All Files (*.*)"
            title = "Select Visio Diagrams"

        files, _ = QFileDialog.getOpenFileNames(
            self, title, str(self._export_dir.resolve()), filt
        )
        if files:
            self._on_bridge_files_dropped(files)

    def _clear_staged_bridge(self):
        self._staged_bridge_files.clear()
        self._update_bridge_table()

    def _remove_staged_bridge_index(self, idx: int):
        if 0 <= idx < len(self._staged_bridge_files):
            self._staged_bridge_files.pop(idx)
            self._update_bridge_table()

    def _update_bridge_table(self):
        self.bridge_table.setRowCount(0)
        for i, file_path in enumerate(self._staged_bridge_files, start=1):
            row = self.bridge_table.rowCount()
            self.bridge_table.insertRow(row)

            idx_item = QTableWidgetItem(str(i))
            idx_item.setTextAlignment(Qt.AlignCenter)

            fn_item = QTableWidgetItem(file_path.name)
            fn_item.setToolTip(str(file_path.resolve()))

            sz_str = format_file_size(file_path.stat().st_size) if file_path.exists() else "—"
            sz_item = QTableWidgetItem(sz_str)
            sz_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            btn_remove = QPushButton("✕")
            btn_remove.setFixedSize(24, 22)
            btn_remove.setStyleSheet("font-size: 11px; padding: 0px; color: #EF4444; font-weight: bold;")
            btn_remove.clicked.connect(lambda checked, idx=row: self._remove_staged_bridge_index(idx))

            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(4, 2, 4, 2)
            action_layout.setAlignment(Qt.AlignCenter)
            action_layout.addWidget(btn_remove)

            self.bridge_table.setItem(row, 0, idx_item)
            self.bridge_table.setItem(row, 1, fn_item)
            self.bridge_table.setItem(row, 2, sz_item)
            self.bridge_table.setCellWidget(row, 3, action_widget)

        count = len(self._staged_bridge_files)
        self.bridge_count_lbl.setText(f"{count} presentation(s) staged")
        self.convert_bridge_btn.setEnabled(count > 0)

    def _start_bridge_conversion(self):
        if not self._staged_bridge_files:
            return
        target_fmt = "pptx_to_visio" if self.rb_pptx2visio.isChecked() else "vsdx_to_pptx"
        paths_to_send = list(self._staged_bridge_files)
        self.files_dropped.emit(paths_to_send, target_fmt)
        self._staged_bridge_files.clear()
        self._update_bridge_table()

    # -------------------------------------------------------------------------
    # Logic: Card 3 (Visio Extractor)
    # -------------------------------------------------------------------------
    def _on_extractor_file_dropped(self, paths: list):
        if paths:
            self._load_vsdx_for_extraction(Path(paths[0]))

    def _browse_vsdx_for_extraction(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Generated Visio File",
            str(self._export_dir.resolve()),
            "Visio Diagrams (*.vsdx);;All Files (*.*)",
        )
        if file_path:
            self._load_vsdx_for_extraction(Path(file_path))

    def _load_vsdx_for_extraction(self, vsdx_path: Path):
        self.extractor_status_lbl.setText(f"⏳ Reading {vsdx_path.name}...")
        self.extractor_status_lbl.setStyleSheet("color: #395396; font-weight: 600; font-size: 11px;")

        # Try fast headless extraction first (sub-millisecond OpenXML)
        fast_code = extract_plantuml_from_vsdx_fast(vsdx_path)
        if fast_code:
            self._display_extracted_code(fast_code, vsdx_path.name, is_fast=True)
            return

        # Fall back to Visio COM reader thread if fast extraction could not find it
        self._reader_thread = VisioReaderThread(str(vsdx_path.resolve()))
        self._reader_thread.text_extracted.connect(
            lambda code: self._display_extracted_code(code, vsdx_path.name, is_fast=False)
        )
        self._reader_thread.error_occurred.connect(self._on_extractor_error)
        self._reader_thread.start()

    def _display_extracted_code(self, code: str, filename: str, is_fast: bool = True):
        self.extracted_text_edit.setPlainText(code)
        method = "Headless OpenXML" if is_fast else "Visio COM"
        self.extractor_status_lbl.setText(f"✅ Extracted source from '{filename}' ({method})")
        self.extractor_status_lbl.setStyleSheet("color: #166534; font-weight: 600; font-size: 11px;")
        self.send_to_editor_btn.setEnabled(True)
        self.copy_extracted_btn.setEnabled(True)
        self.save_extracted_btn.setEnabled(True)

    def _on_extractor_error(self, err_msg: str):
        self.extractor_status_lbl.setText(f"❌ {err_msg}")
        self.extractor_status_lbl.setStyleSheet("color: #DC2626; font-weight: 600; font-size: 11px;")
        self.send_to_editor_btn.setEnabled(False)
        self.copy_extracted_btn.setEnabled(False)
        self.save_extracted_btn.setEnabled(False)

    def _on_send_to_editor_clicked(self):
        code = self.extracted_text_edit.toPlainText().strip()
        if code:
            self.send_to_editor_requested.emit(code)

    def _on_copy_extracted_clicked(self):
        code = self.extracted_text_edit.toPlainText()
        if code:
            from PyQt5.QtWidgets import QApplication
            QApplication.clipboard().setText(code)
            QToolTip.showText(QCursor.pos(), "📋 PlantUML code copied to clipboard!", self, msecShowTime=1800)

    def _on_save_extracted_clicked(self):
        code = self.extracted_text_edit.toPlainText()
        if not code:
            return
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Extracted PlantUML Diagram",
            str(self._export_dir / "extracted_diagram.puml"),
            "PlantUML Diagram (*.puml);;Text File (*.txt)",
        )
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(code)
            QToolTip.showText(QCursor.pos(), f"💾 Saved to {Path(save_path).name}", self, msecShowTime=2000)

    # -------------------------------------------------------------------------
    # Logic: Card 4 (Stencils)
    # -------------------------------------------------------------------------
    def _copy_snippet_to_clipboard(self, snippet: PlantUMLSnippet):
        from PyQt5.QtWidgets import QApplication
        QApplication.clipboard().setText(snippet.code)
        QToolTip.showText(QCursor.pos(), f"📋 Copied '{snippet.label}' snippet!", self, msecShowTime=1800)

    def _insert_snippet_into_editor(self, snippet: PlantUMLSnippet):
        self.insert_snippet_requested.emit(snippet)

    # -------------------------------------------------------------------------
    # Logic: Recent Exports Strip
    # -------------------------------------------------------------------------
    def _refresh_recent_exports(self):
        # Clear existing items in recent_container
        while self.recent_container.count():
            item = self.recent_container.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        if not self._export_dir.exists():
            return

        valid_exts = {".vsdx", ".pptx", ".svg", ".txt"}
        recent_files = []
        try:
            for f in self._export_dir.iterdir():
                if f.is_file() and f.suffix.lower() in valid_exts:
                    recent_files.append((f, f.stat().st_mtime))
        except Exception:
            pass

        recent_files.sort(key=lambda x: x[1], reverse=True)
        top_files = recent_files[:5]

        self.recent_count_lbl.setText(f"({len(top_files)} files)")

        if not top_files:
            empty_lbl = QLabel("No diagrams exported yet in output/ directory.")
            empty_lbl.setStyleSheet("color: #94A3B8; font-style: italic; font-size: 11px;")
            self.recent_container.addWidget(empty_lbl)
            return

        for f, mtime in top_files:
            dt_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(2, 1, 2, 1)
            row_layout.setSpacing(6)

            # Icon based on ext
            ext = f.suffix.lower()
            icon_str = "📄" if ext == ".vsdx" else ("📊" if ext == ".pptx" else ("🌐" if ext == ".svg" else "🔤"))

            icon_lbl = QLabel(icon_str)
            name_lbl = QLabel(f.name)
            name_lbl.setStyleSheet("font-weight: 500; color: #1E293B; font-size: 11px;")
            name_lbl.setToolTip(str(f.resolve()))

            time_lbl = QLabel(dt_str)
            time_lbl.setStyleSheet("color: #94A3B8; font-size: 10px; margin-left: 4px;")

            size_str = format_file_size(f.stat().st_size)
            size_lbl = QLabel(size_str)
            size_lbl.setStyleSheet("color: #64748B; font-size: 10px;")

            btn_open = QPushButton("👁️ Open")
            btn_open.setStyleSheet("font-size: 10px; padding: 1px 6px;")
            btn_open.clicked.connect(lambda checked, path=f: self._open_file(path))

            btn_folder = QPushButton("📂 Folder")
            btn_folder.setStyleSheet("font-size: 10px; padding: 1px 6px;")
            btn_folder.clicked.connect(lambda checked, path=f: self._show_in_folder(path))

            btn_copy = QPushButton("🔗 Copy Path")
            btn_copy.setStyleSheet("font-size: 10px; padding: 1px 6px;")
            btn_copy.clicked.connect(lambda checked, path=f: self._copy_file_path(path))

            row_layout.addWidget(icon_lbl)
            row_layout.addWidget(name_lbl)
            row_layout.addWidget(time_lbl)
            row_layout.addWidget(size_lbl)
            row_layout.addStretch()
            row_layout.addWidget(btn_open)
            row_layout.addWidget(btn_folder)
            row_layout.addWidget(btn_copy)

            self.recent_container.addWidget(row_widget)

    def _open_file(self, file_path: Path):
        try:
            os.startfile(str(file_path.resolve()))
        except Exception as e:
            QMessageBox.warning(self, "Error Opening File", f"Could not open {file_path.name}:\n{e}")

    def _show_in_folder(self, file_path: Path):
        try:
            subprocess.run(["explorer", f"/select,{str(file_path.resolve())}"])
        except Exception:
            try:
                os.startfile(str(file_path.parent.resolve()))
            except Exception:
                pass

    def _copy_file_path(self, file_path: Path):
        from PyQt5.QtWidgets import QApplication
        QApplication.clipboard().setText(str(file_path.resolve()))
        QToolTip.showText(QCursor.pos(), f"🔗 Copied path:\n{file_path.name}", self, msecShowTime=1800)

    # -------------------------------------------------------------------------
    # Destination & Folder Actions
    # -------------------------------------------------------------------------
    def _on_open_output_folder(self):
        self._export_dir.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(self._export_dir.resolve()))
        except Exception as e:
            logger.error(f"Failed to open export directory: {e}")

    def _on_change_output_folder(self):
        new_dir = QFileDialog.getExistingDirectory(
            self, "Select Default Output Folder", str(self._export_dir.resolve())
        )
        if new_dir:
            self._export_dir = Path(new_dir)
            self.out_path_lbl.setText(str(self._export_dir.name) + "/")
            self.out_path_lbl.setToolTip(str(self._export_dir.resolve()))
            self._refresh_recent_exports()

    # -------------------------------------------------------------------------
    # Public & Backward-Compatible API
    # -------------------------------------------------------------------------
    def set_state(self, state: str, text: Optional[str] = None):
        """Maintains full backward compatibility with main_window.py set_state calls."""
        if hasattr(self, "drop_label_puml"):
            self.drop_label_puml.set_state(state, text)
        self._refresh_system_diagnostics()

    def on_conversion_success(self, out_path: str):
        """Called by main_window.py when a diagram conversion finishes successfully."""
        self._refresh_recent_exports()
