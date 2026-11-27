"""Protocol Clause 9 / IE definition inspector card widget."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtCore import QPoint, Qt, QUrl, pyqtSignal
from PyQt5.QtWidgets import (
    QAction,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import scale_px
from core.ui.theme import C
from modules.nas.core.nas_db import NASDatabase
from modules.nas.core.nas_exporter import copy_definition_to_clipboard
from .inspector_render import NASInspectorRenderMixin
from .packet_stack_widget import PacketStackWidget


class _InspectorHeaderFrame(QFrame):
    """Header frame emitting double-click for maximize/restore focus toggle."""

    double_clicked = pyqtSignal()

    def mouseDoubleClickEvent(self, event) -> None:  # type: ignore[override]
        self.double_clicked.emit()
        super().mouseDoubleClickEvent(event)


class NASInspectorWidget(NASInspectorRenderMixin, QFrame):
    """Modern Card Frame rendering Clause 9 diagrams, ASN.1 syntax, field descriptions, and UP Packet Stack."""

    jump_to_message_requested = pyqtSignal(str)
    filter_by_ie_requested = pyqtSignal(str)
    import_spec_requested = pyqtSignal(str, int)  # (spec_number, major_version)
    open_spec_requested = pyqtSignal(str, str, str, str)  # (spec_number, version, clause_number, doc_format)
    show_in_specs_tab_requested = pyqtSignal(str, str)     # (spec_number, version)
    search_in_spec_search_requested = pyqtSignal(str, str) # (spec_number, version)
    send_plantuml_to_editor_requested = pyqtSignal(str)
    header_double_clicked = pyqtSignal()

    def __init__(self, db: Optional[NASDatabase] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.db = db
        self._current_clause_defs: Dict[str, Dict[str, Any]] = {}
        self._current_ie_clause: Optional[str] = None
        self._current_ie_name: Optional[str] = None
        self._current_ie_spec: Optional[str] = None
        self._current_containing_msgs: List[Dict[str, Any]] = []
        self._current_html_content: str = ""
        self._updating_combo: bool = False
        self._inspector_mode: str = "definition"
        self._current_message_name_hint: str = ""
        self._packet_available: bool = False

        self.setObjectName("cardFrame")
        self.setStyleSheet(f"""
            QFrame#cardFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
            }}
        """)

        self._setup_ui()

    def set_database(self, db: NASDatabase):
        self.db = db
        if hasattr(self, "packet_stack_widget"):
            self.packet_stack_widget.set_database(db)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 1. Header Toolbar
        header_frame = _InspectorHeaderFrame()
        header_frame.setObjectName("inspHeaderFrame")
        header_frame.double_clicked.connect(self.header_double_clicked.emit)
        header_frame.setStyleSheet(f"""
            QFrame#inspHeaderFrame {{
                background-color: {C.BG_SURFACE};
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 4px 8px;
            }}
        """)
        insp_header = QHBoxLayout(header_frame)
        insp_header.setContentsMargins(6, 4, 6, 4)
        insp_header.setSpacing(6)
        self.insp_header_layout = insp_header

        card_title = QLabel("🔍 Field & Structure Inspector:")
        card_title.setStyleSheet(f"font-weight: bold; color: {C.TEXT_STRONG}; font-size: 11px;")
        insp_header.addWidget(card_title)

        self.title_lbl = QLabel("No Selection")
        self.title_lbl.setMaximumWidth(scale_px(300))
        self.title_lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.title_lbl.setStyleSheet(f"""
            font-weight: bold;
            color: {C.ACCENT_SKY};
            font-size: 11px;
            background-color: {C.SKY_SOFT};
            padding: 2px 8px;
            border-radius: 4px;
            border: 1px solid {C.SKY_BORDER};
        """)
        insp_header.addWidget(self.title_lbl)

        self.mode_segment = QFrame(header_frame)
        self.mode_segment.setObjectName("inspModeSegment")
        self.mode_segment.setStyleSheet(
            f"#inspModeSegment {{ background-color: {C.BG_CANVAS}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 5px; }}"
        )
        mode_lay = QHBoxLayout(self.mode_segment)
        mode_lay.setContentsMargins(2, 2, 2, 2)
        mode_lay.setSpacing(2)
        seg_qss = (
            f"QPushButton {{ font-size: 10.5px; font-weight: 500; color: {C.TEXT_SECONDARY}; "
            f"background-color: transparent; border: 1px solid transparent; border-radius: 4px; padding: 1px 7px; }} "
            f"QPushButton:hover {{ color: {C.TEXT_PRIMARY}; background-color: {C.BG_SURFACE}; }} "
            f"QPushButton:checked {{ font-weight: 600; color: {C.ACCENT_PRIMARY}; "
            f"background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; }}"
        )
        self.btn_mode_definition = QPushButton("📖 Definition", self.mode_segment)
        self.btn_mode_definition.setCheckable(True)
        self.btn_mode_definition.setChecked(True)
        self.btn_mode_definition.setCursor(Qt.PointingHandCursor)
        self.btn_mode_definition.setStyleSheet(seg_qss)
        self.btn_mode_definition.clicked.connect(lambda: self.set_inspector_mode("definition"))
        mode_lay.addWidget(self.btn_mode_definition)

        self.btn_mode_packet = QPushButton("📦 Full Packet", self.mode_segment)
        self.btn_mode_packet.setCheckable(True)
        self.btn_mode_packet.setChecked(False)
        self.btn_mode_packet.setCursor(Qt.PointingHandCursor)
        self.btn_mode_packet.setToolTip(
            "Inspect full multi-layer packet stack (5G UP: TS 29.281 + TS 38.415; N2 CP: TS 38.413 NGAP; 5GS NAS: TS 24.501 + URSP/5GSM)"
        )
        self.btn_mode_packet.setStyleSheet(seg_qss)
        self.btn_mode_packet.clicked.connect(lambda: self.set_inspector_mode("packet"))
        mode_lay.addWidget(self.btn_mode_packet)

        self.mode_segment.setVisible(False)
        insp_header.addWidget(self.mode_segment)

        insp_header.addStretch()

        self.usage_btn = QPushButton("Used in: 0 msgs ▾")
        self.usage_btn.setVisible(False)
        self.usage_btn.setCursor(Qt.PointingHandCursor)
        self.usage_btn.setToolTip("View other messages that contain this Information Element")
        self.usage_btn.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: bold;
                color: {C.ACCENT_SKY_DARK};
                background-color: {C.SKY_SOFT};
                border: 1px solid {C.SKY_BORDER};
                border-radius: 4px;
                padding: 2px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.SKY_BORDER};
                border-color: {C.ACCENT_SKY};
            }}
        """)
        self.usage_btn.clicked.connect(self._show_usage_menu)
        insp_header.addWidget(self.usage_btn)

        self.version_combo = QComboBox()
        self.version_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.version_combo.setMinimumWidth(scale_px(120))
        self.version_combo.view().setMinimumWidth(scale_px(220))
        self.version_combo.setToolTip("Switch specification release for this definition")
        self.version_combo.setStyleSheet(f"""
            QComboBox {{
                combobox-popup: 0;
                font-weight: 600;
                font-size: 11px;
                padding: 2px 8px;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background-color: {C.BG_CARD};
            }}
            QComboBox:hover {{
                border-color: {C.ACCENT_SKY};
            }}
        """)
        self.version_combo.currentIndexChanged.connect(self._on_version_changed)
        insp_header.addWidget(self.version_combo)

        self.copy_btn = QPushButton("📋 Copy Definition ▾")
        self.copy_btn.setCursor(Qt.PointingHandCursor)
        self.copy_btn.setToolTip("Copy definition to clipboard (Word-compatible rich text, Markdown, plain text, or citation)")
        self.copy_btn.setEnabled(False)
        self.copy_btn.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: 500;
                color: {C.TEXT_BODY};
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 2px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_PRIMARY};
                border-color: {C.TEXT_FAINT};
            }}
            QPushButton:disabled {{
                color: {C.TEXT_FAINT};
                background-color: {C.BG_SURFACE};
                border-color: {C.BORDER_SUBTLE};
            }}
        """)

        copy_menu = QMenu(self)
        copy_menu.setStyleSheet(f"""
            QMenu {{
                font-size: 11px;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                padding: 4px;
            }}
            QMenu::item {{
                padding: 5px 18px 5px 10px;
                border-radius: 3px;
            }}
            QMenu::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
            }}
        """)
        act_cp_word = copy_menu.addAction("📋 Copy for Word / Outlook (Rich Text)")
        act_cp_word.triggered.connect(lambda: self._copy_formatted(mode="word"))
        act_cp_md = copy_menu.addAction("📝 Copy as Markdown")
        act_cp_md.triggered.connect(lambda: self._copy_formatted(mode="markdown"))
        act_cp_plain = copy_menu.addAction("📄 Copy Plain Text")
        act_cp_plain.triggered.connect(lambda: self._copy_formatted(mode="text"))
        copy_menu.addSeparator()
        act_cp_cite = copy_menu.addAction("🔖 Copy Citation Only")
        act_cp_cite.triggered.connect(lambda: self._copy_formatted(mode="citation"))

        self.copy_btn.setMenu(copy_menu)
        insp_header.addWidget(self.copy_btn)

        self.open_spec_btn = QPushButton("📖 Open Spec ▾")
        self.open_spec_btn.setCursor(Qt.PointingHandCursor)
        self.open_spec_btn.setToolTip("Open this specification release in Microsoft Word at this clause")
        self.open_spec_btn.setEnabled(False)
        self.open_spec_btn.setStyleSheet(f"""
            QPushButton {{
                font-size: 11px;
                font-weight: 500;
                color: {C.ACCENT_PRIMARY};
                background-color: {C.SKY_SURFACE};
                border: 1px solid {C.SKY_BORDER};
                border-radius: 4px;
                padding: 2px 8px;
            }}
            QPushButton:hover {{
                background-color: {C.SKY_SOFT};
                border-color: {C.ACCENT_SKY};
                color: {C.ACCENT_SKY_DARK};
            }}
            QPushButton:disabled {{
                color: {C.TEXT_FAINT};
                background-color: {C.BG_SURFACE};
                border-color: {C.BORDER_SUBTLE};
            }}
        """)
        self.open_spec_btn.clicked.connect(self._on_open_spec_btn_clicked)
        insp_header.addWidget(self.open_spec_btn)

        layout.addWidget(header_frame)

        # 2. Content Stack: Empty State (0) vs Text Browser (1)
        self.stack = QStackedWidget()

        # Empty State Placeholder
        self.empty_widget = QWidget()
        empty_layout = QVBoxLayout(self.empty_widget)
        empty_layout.setAlignment(Qt.AlignCenter)
        empty_layout.setSpacing(6)

        empty_icon = QLabel("🔍")
        empty_icon.setStyleSheet(f"font-size: 26px; color: {C.TEXT_FAINT};")
        empty_icon.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_icon)

        empty_title = QLabel("Information Element & Field Inspector")
        empty_title.setStyleSheet(f"font-weight: bold; font-size: 13px; color: {C.TEXT_BODY};")
        empty_title.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_title)

        empty_sub = QLabel(
            "Click on any Information Element or field in the Evolution Matrix above\n"
            "to inspect its bit-level octet layout, ASN.1 syntax, and field descriptions."
        )
        empty_sub.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; line-height: 1.4;")
        empty_sub.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_sub)

        self.stack.addWidget(self.empty_widget)

        # Content Container (Index 1): Definition TextBrowser (0) vs Full Packet Stack (1)
        self.content_stack = QStackedWidget()

        # Text Browser View
        self.text_area = QTextBrowser()
        self.text_area.setReadOnly(True)
        self.text_area.setOpenLinks(False)
        self.text_area.setOpenExternalLinks(False)
        self.text_area.anchorClicked.connect(self._on_anchor_clicked)
        self.text_area.setContextMenuPolicy(Qt.CustomContextMenu)
        self.text_area.customContextMenuRequested.connect(self._show_text_area_context_menu)
        self.text_area.setStyleSheet(f"""
            QTextBrowser {{
                border: none;
                padding: 8px;
                background-color: {C.BG_CARD};
            }}
        """)
        self.content_stack.addWidget(self.text_area)

        self.packet_stack_widget = PacketStackWidget(db=self.db, parent=self)
        self.packet_stack_widget.open_spec_requested.connect(self.open_spec_requested.emit)
        self.packet_stack_widget.inspect_field_requested.connect(self._on_inspect_packet_field)
        self.packet_stack_widget.send_plantuml_to_editor_requested.connect(self.send_plantuml_to_editor_requested.emit)
        self.content_stack.addWidget(self.packet_stack_widget)

        self.stack.addWidget(self.content_stack)

        layout.addWidget(self.stack)

    def _copy_formatted(self, mode: str = "word"):
        """Copies the active IE definition to the clipboard in Word, Markdown, or plain text format."""
        spec, ver, clause = self._get_active_spec_and_version()
        name = self._current_ie_name or self.title_lbl.text().strip()
        html_content = getattr(self, "_current_html_content", "") or self.text_area.toHtml()

        if not html_content and not name:
            return

        copy_definition_to_clipboard(
            html_content=html_content,
            spec_number=spec,
            version=ver,
            ie_name=name,
            clause=clause,
            mode=mode,
        )

        self.copy_btn.setText("✓ Copied!")
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(1500, lambda: self.copy_btn.setText("📋 Copy Definition ▾"))

    def _copy_current_text(self):
        """Legacy fallback calling Word-compatible formatted copy."""
        self._copy_formatted(mode="word")

    def _show_text_area_context_menu(self, pos: QPoint):
        """Context menu inside definition text browser offering formatted Word/Markdown export."""
        menu = self.text_area.createStandardContextMenu()
        menu.setStyleSheet(f"""
            QMenu {{
                font-size: 11px;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                padding: 4px;
            }}
            QMenu::item {{
                padding: 5px 18px 5px 10px;
                border-radius: 3px;
            }}
            QMenu::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
            }}
        """)
        menu.addSeparator()
        act_cp_word = menu.addAction("📋 Copy Definition for Word / Outlook")
        act_cp_word.triggered.connect(lambda: self._copy_formatted("word"))
        act_cp_md = menu.addAction("📝 Copy Definition as Markdown")
        act_cp_md.triggered.connect(lambda: self._copy_formatted("markdown"))
        act_cp_cite = menu.addAction("🔖 Copy Citation")
        act_cp_cite.triggered.connect(lambda: self._copy_formatted("citation"))
        menu.exec_(self.text_area.mapToGlobal(pos))

    def _on_anchor_clicked(self, url: QUrl):
        url_str = url.toString()
        if url_str.startswith("action:import:"):
            parts = url_str.replace("action:import:", "").split(":")
            spec_num = parts[0]
            major_ver = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
            self.import_spec_requested.emit(spec_num, major_ver)
        elif url_str.startswith("action:open_spec:"):
            parts = url_str.replace("action:open_spec:", "").split(":")
            spec_num = parts[0]
            ver = parts[1] if len(parts) > 1 else ""
            clause = parts[2] if len(parts) > 2 else ""
            self.open_spec_requested.emit(spec_num, ver, clause, "word")

    def _get_active_spec_and_version(self) -> Tuple[str, str, str]:
        spec = self._current_ie_spec or ""
        version = ""
        clause = self._current_ie_clause or ""

        combo_val = self.version_combo.currentData()
        if combo_val and " v" in str(combo_val):
            parts = str(combo_val).split(" v")
            spec = parts[0].strip()
            version = parts[1].strip()

        return spec, version, clause

    def _update_open_spec_menu(self):
        spec, ver, clause = self._get_active_spec_and_version()
        if not spec:
            self.open_spec_btn.setEnabled(False)
            self.open_spec_btn.setMenu(None)
            self.open_spec_btn.setText("📖 Open Spec")
            return

        self.open_spec_btn.setEnabled(True)
        ver_tag = f" v{ver}" if ver else ""
        self.open_spec_btn.setText(f"📖 Open TS {spec}{ver_tag} ▾")
        self.open_spec_btn.setToolTip(f"Open TS {spec}{ver_tag} at Clause {clause} in Microsoft Word")

        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                font-size: 11px;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                padding: 4px;
            }}
            QMenu::item {{
                padding: 5px 18px 5px 10px;
                border-radius: 3px;
            }}
            QMenu::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
            }}
        """)

        act_word = QAction(f"📝 Open TS {spec}{ver_tag} in Word (Clause {clause or 'definition'})", self)
        act_word.triggered.connect(lambda: self.open_spec_requested.emit(spec, ver, clause, "word"))
        menu.addAction(act_word)

        act_pdf = QAction(f"📕 Open TS {spec}{ver_tag} in PDF", self)
        act_pdf.triggered.connect(lambda: self.open_spec_requested.emit(spec, ver, clause, "pdf"))
        menu.addAction(act_pdf)

        act_folder = QAction("📂 Open Local Folder in Explorer", self)
        act_folder.triggered.connect(lambda: self.open_spec_requested.emit(spec, ver, "", "folder"))
        menu.addAction(act_folder)

        menu.addSeparator()

        act_specs = QAction(f"📚 Show TS {spec}{ver_tag} in Specifications Tab", self)
        act_specs.triggered.connect(lambda: self.show_in_specs_tab_requested.emit(spec, ver))
        menu.addAction(act_specs)

        act_search = QAction(f"🔎 Search TS {spec} in Spec Search Tab", self)
        act_search.triggered.connect(lambda: self.search_in_spec_search_requested.emit(spec, ver))
        menu.addAction(act_search)

        self.open_spec_btn.setMenu(menu)

    def _on_open_spec_btn_clicked(self):
        spec, ver, clause = self._get_active_spec_and_version()
        if spec:
            self.open_spec_requested.emit(spec, ver, clause, "word")
