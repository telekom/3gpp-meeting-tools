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

import logging
import os
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QButtonGroup,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import CARD_FRAME_STYLE
from core.utils.paths import get_project_root
from core.utils.utils import get_detailed_java_info
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.ui.cards import (
    OfficeBridgeCard,
    PumlBatchCard,
    RecentExportsStrip,
    StencilsCard,
    VsdxExtractorCard,
    detect_puml_diagram_type,
    extract_plantuml_from_vsdx_fast,
    format_file_size,
    is_powerpoint_registered,
    is_visio_registered,
)
from modules.puml2visio.utils.paths import get_puml2visio_asset_path
from modules.puml2visio.utils.utils import get_plantuml_version

logger = logging.getLogger(__name__)

# Re-export utilities for backward compatibility
__all__ = [
    "is_visio_registered",
    "is_powerpoint_registered",
    "detect_puml_diagram_type",
    "extract_plantuml_from_vsdx_fast",
    "format_file_size",
    "BatchConvertTab",
]


class BatchConvertTab(QWidget):
    """Modernized Visio & Office Conversion Studio with multi-card workflow,

    system diagnostics ribbon, staging table, and reverse-engineering inspector.
    """

    # Backward-compatible & custom signals
    files_dropped = pyqtSignal(list, str)
    send_to_editor_requested = pyqtSignal(str)
    insert_snippet_requested = pyqtSignal(object)
    java_maintenance_requested = pyqtSignal()
    open_folder_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._export_dir = get_project_root().parent / "output"
        self._setup_ui()
        self._refresh_system_diagnostics()

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

        self.card_puml = PumlBatchCard(self._get_export_dir)
        self.card_puml.files_dropped.connect(self.files_dropped)

        self.card_bridge = OfficeBridgeCard(self._get_export_dir)
        self.card_bridge.files_dropped.connect(self.files_dropped)

        self.card_extractor = VsdxExtractorCard(self._get_export_dir)
        self.card_extractor.send_to_editor_requested.connect(self.send_to_editor_requested)

        self.card_stencils = StencilsCard()
        self.card_stencils.insert_snippet_requested.connect(self.insert_snippet_requested)

        self.stack.addWidget(self.card_puml)
        self.stack.addWidget(self.card_bridge)
        self.stack.addWidget(self.card_extractor)
        self.stack.addWidget(self.card_stencils)

        # Backward compatibility aliases
        self.drop_label_puml = self.card_puml.drop_label_puml
        self.puml_table = self.card_puml.puml_table
        self.bridge_table = self.card_bridge.bridge_table
        self.stencils_search_edit = self.card_stencils.stencils_search_edit
        self.stencils_table = self.card_stencils.stencils_table

        main_layout.addWidget(self.stack, stretch=1)

        # 4. Bottom Recent Exports Strip
        self.recent_strip = RecentExportsStrip(self._get_export_dir)
        main_layout.addWidget(self.recent_strip)

    def _get_export_dir(self) -> Path:
        return self._export_dir

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
                color: #1E40AF;
                border: 1px solid #BFDBFE;
                border-radius: 12px;
                padding: 3px 10px;
                font-weight: 600;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #DBEAFE;
            }
        """)
        self.java_badge_btn.setToolTip("Click to view Java Runtime Diagnostics & PlantUML Maintenance")
        self.java_badge_btn.clicked.connect(self.java_maintenance_requested)
        ribbon_layout.addWidget(self.java_badge_btn)

        ribbon_layout.addStretch()

        # Output Directory Indicator & Controls
        out_lbl = QLabel("📂 Output:")
        out_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold;")
        ribbon_layout.addWidget(out_lbl)

        self.output_dir_btn = QPushButton(f"output/ ({self._export_dir.name})")
        self.output_dir_btn.setStyleSheet("""
            QPushButton {
                background-color: #F1F5F9;
                color: #0F172A;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #E2E8F0;
            }
        """)
        self.output_dir_btn.setToolTip(f"Open export folder: {self._export_dir.resolve()}")
        self.output_dir_btn.clicked.connect(self._on_open_output_folder)
        ribbon_layout.addWidget(self.output_dir_btn)

        self.change_dir_btn = QPushButton("Change...")
        self.change_dir_btn.setStyleSheet("font-size: 11px; padding: 2px 6px;")
        self.change_dir_btn.clicked.connect(self._on_change_output_folder)
        ribbon_layout.addWidget(self.change_dir_btn)

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
    # System Diagnostics & Output Folder Handlers
    # -------------------------------------------------------------------------
    def _refresh_system_diagnostics(self):
        # 1. Visio COM
        if is_visio_registered():
            self.visio_badge.setText("🟢 Visio COM: Active")
            self.visio_badge.setStyleSheet(
                "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #DCFCE7; color: #166534;"
            )
            self.visio_badge.setToolTip("Microsoft Visio COM is registered and available for native shape creation.")
        else:
            self.visio_badge.setText("🟠 Visio COM: Not Found")
            self.visio_badge.setStyleSheet(
                "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #FEF3C7; color: #92400E;"
            )
            self.visio_badge.setToolTip("Visio COM automation is not detected. Diagrams can still be exported to SVG, PPTX, or text.")

        # 2. PowerPoint COM
        if is_powerpoint_registered():
            self.ppt_badge.setText("🟢 PowerPoint COM: Active")
            self.ppt_badge.setStyleSheet(
                "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #DCFCE7; color: #166534;"
            )
            self.ppt_badge.setToolTip("Microsoft PowerPoint COM is registered and available for presentation generation.")
        else:
            self.ppt_badge.setText("🟠 PowerPoint COM: Not Found")
            self.ppt_badge.setStyleSheet(
                "padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; background-color: #FEF3C7; color: #92400E;"
            )
            self.ppt_badge.setToolTip("PowerPoint COM automation is not detected.")

        # 3. PlantUML JAR & Java Engine
        try:
            jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
            puml_ver = get_plantuml_version(jar_path) if jar_path.exists() else "Missing"
        except Exception:
            puml_ver = "Unknown"

        java_info = get_detailed_java_info()
        java_ver = java_info.get("version", "No Java")
        self.java_badge_btn.setText(f"☕ PlantUML {puml_ver} (Java {java_ver})")

    def _on_open_output_folder(self):
        try:
            self._export_dir.mkdir(parents=True, exist_ok=True)
            os.startfile(str(self._export_dir.resolve()))
        except Exception as e:
            logger.warning(f"Could not open export folder: {e}")
        self.open_folder_requested.emit()

    def _on_change_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Directory", str(self._export_dir.resolve())
        )
        if folder:
            self._export_dir = Path(folder)
            self.output_dir_btn.setText(f"output/ ({self._export_dir.name})")
            self.output_dir_btn.setToolTip(f"Export folder: {self._export_dir.resolve()}")
            self.recent_strip.refresh()

    def set_state(self, state, text: Optional[str] = None):
        """Maintains full backward compatibility with main_window.py set_state calls.

        Accepts either:
        - state: str ('ready', 'idle', 'busy', 'error') and optional text: str
        - is_converting: bool (True for busy, False for ready)
        """
        if isinstance(state, bool):
            state = "busy" if state else "ready"

        # If legacy generic ready prompt is passed, substitute with card-specific text
        if state in ("ready", "idle") and text and "Batch converts between Visio & PowerPoint" in text:
            text = "📥 Drag & Drop PlantUML file(s) (.puml, .plantuml, .txt) here to stage for batch conversion"

        if hasattr(self, "card_puml") and hasattr(self.card_puml, "set_state"):
            self.card_puml.set_state(state, text)
        elif hasattr(self, "drop_label_puml"):
            self.drop_label_puml.set_state(state, text)

        if hasattr(self, "card_bridge") and hasattr(self.card_bridge, "drop_label_bridge"):
            if state == "busy":
                self.card_bridge.drop_label_bridge.set_state("busy")
            elif state in ("ready", "idle"):
                self.card_bridge.drop_label_bridge.set_state("ready")

        self._refresh_system_diagnostics()

    def on_conversion_success(self, output_path: str = ""):
        self.recent_strip.refresh()
