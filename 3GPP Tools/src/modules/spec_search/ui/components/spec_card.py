"""ReleasePillWidget and SpecCardWidget controls for Spec Search version filtering."""

import re
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QPoint, Qt, pyqtSignal
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from core.ui.theme import C, T
from modules.spec_search.core.spec_search_db import extract_major_release, parse_version_tuple


class ReleasePillWidget(QPushButton):
    """
    Compact clickable pill representing a single specification version release.
    Features distinct active/inactive color-coding and rich tooltip.
    """
    def __init__(self, version_data: Dict[str, Any], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.version_data = version_data
        self.spec_number = str(version_data.get("spec_number", ""))
        self.version = str(version_data.get("version", ""))
        self.version_id = version_data.get("id")
        self.release_date = str(version_data.get("release_date") or "")
        self.total_clauses = version_data.get("total_clauses", 0)

        rel = extract_major_release(self.version)
        short_rel = rel.replace("Rel-", "R") if rel.startswith("Rel-") else rel
        self.setText(f"{short_rel}: {self.version}")
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(scale_px(22, minimum=20))

        # Build informative tooltip
        date_str = f"Date: {self.release_date}\n" if self.release_date else ""
        clauses_str = f"Clauses: {self.total_clauses:,}\n" if self.total_clauses else ""
        self.setToolTip(
            f"TS {self.spec_number} v{self.version}\n"
            f"Release: {rel}\n"
            f"{date_str}{clauses_str}"
            f"(Click to toggle search inclusion | Right-click for options)"
        )
        self._update_style()
        self.toggled.connect(lambda _: self._update_style())

    def _update_style(self):
        if self.isChecked():
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.ACCENT_SKY};
                    color: {C.BG_CARD};
                    border: 1px solid {C.ACCENT_SKY};
                    border-radius: 11px;
                    padding: 1px 7px;
                    font-size: 11px;
                    font-weight: 600;
                }}
                QPushButton:hover {{
                    background-color: {C.ACCENT_SKY_DARK};
                    border-color: {C.ACCENT_SKY_DARK};
                }}
            """)
        else:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.BG_SURFACE};
                    color: {C.TEXT_MUTED};
                    border: 1px solid {C.BORDER_STRONG};
                    border-radius: 11px;
                    padding: 1px 7px;
                    font-size: 11px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {C.BG_CANVAS};
                    color: {C.TEXT_PRIMARY};
                    border-color: {C.TEXT_FAINT};
                }}
            """)


class SpecCardWidget(QFrame):
    """
    Compact specification card with header controls and an auto-wrapping flow of release pills.
    """
    selection_changed = pyqtSignal()
    delete_version_requested = pyqtSignal(str, str)
    delete_spec_requested = pyqtSignal(str, int)
    prune_spec_requested = pyqtSignal(str)
    smart_import_requested = pyqtSignal(str)

    def __init__(
        self,
        spec_number: str,
        title: str,
        versions: List[Dict[str, Any]],
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.spec_number = spec_number
        self.raw_title = title
        self.versions = versions
        self.pills: List[ReleasePillWidget] = []
        self._updating_checks = False

        self.setObjectName("SpecCard")
        self.setStyleSheet(f"""
            QFrame#SpecCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 6px;
            }}
            QFrame#SpecCard:hover {{
                border-color: {C.BORDER_STRONG};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        # Header Row
        hdr = QHBoxLayout()
        hdr.setContentsMargins(0, 0, 0, 0)
        hdr.setSpacing(6)

        self.chk_all = QCheckBox()
        self.chk_all.setCursor(Qt.PointingHandCursor)
        self.chk_all.clicked.connect(self._on_header_checkbox_clicked)
        hdr.addWidget(self.chk_all)

        lbl_spec = QLabel(f"TS {spec_number}")
        lbl_spec.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: {T.SIZE_TABLE}px; font-weight: bold;")
        hdr.addWidget(lbl_spec)

        clean_title = re.sub(r'^(NR;\s*|LTE;\s*|Universal Mobile Telecommunications System \(UMTS\);\s*)', '', title, flags=re.IGNORECASE).strip()
        parts = [p.strip() for p in clean_title.split(';') if p.strip()]
        display_title = parts[0] if parts else clean_title
        if len(display_title) > 36:
            display_title = display_title[:35].rstrip() + "…"

        self.lbl_title = QLabel(display_title)
        self.lbl_title.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        if title:
            self.lbl_title.setToolTip(f"TS {spec_number}: {title}\n{len(versions)} indexed release(s)")
        hdr.addWidget(self.lbl_title, stretch=1)

        self.badge = QLabel(str(len(versions)))
        self.badge.setStyleSheet(f"""
            background-color: {C.BG_CANVAS};
            color: {C.TEXT_SECONDARY};
            border: 1px solid {C.BORDER_SUBTLE};
            border-radius: 8px;
            padding: 1px 6px;
            font-size: 10px;
            font-weight: bold;
        """)
        hdr.addWidget(self.badge)

        self.btn_opt = QToolButton()
        self.btn_opt.setText("⋮")
        self.btn_opt.setCursor(Qt.PointingHandCursor)
        self.btn_opt.setStyleSheet(f"""
            QToolButton {{
                border: none;
                background: transparent;
                color: {C.TEXT_MUTED};
                font-weight: bold;
                font-size: 12px;
                padding: 0px 2px;
            }}
            QToolButton:hover {{
                color: {C.TEXT_PRIMARY};
                background-color: {C.BG_CANVAS};
                border-radius: 3px;
            }}
        """)
        self.btn_opt.clicked.connect(self._show_options_menu)
        hdr.addWidget(self.btn_opt)

        layout.addLayout(hdr)

        # Flow Layout of Pills
        self.pills_container = QWidget()
        self.flow = FlowLayout(self.pills_container, margin=0, h_spacing=4, v_spacing=4)

        for v in sorted(versions, key=lambda x: parse_version_tuple(x["version"]), reverse=True):
            pill = ReleasePillWidget(v, self)
            pill.toggled.connect(self._on_pill_toggled)
            pill.setContextMenuPolicy(Qt.CustomContextMenu)
            pill.customContextMenuRequested.connect(lambda pos, p=pill: self._show_pill_menu(p, pos))
            self.pills.append(pill)
            self.flow.addWidget(pill)

        layout.addWidget(self.pills_container)

    def _show_options_menu(self):
        menu = QMenu(self)
        act_smart = menu.addAction(f"🔄 Check && Ingest Latest Releases for TS {self.spec_number}...")
        act_smart.triggered.connect(lambda: self.smart_import_requested.emit(self.spec_number))
        menu.addSeparator()
        act_sel_all = menu.addAction("Select All Releases")
        act_sel_all.triggered.connect(self.select_all)
        act_desel = menu.addAction("Deselect All Releases")
        act_desel.triggered.connect(self.deselect_all)
        menu.addSeparator()

        # Check if multiple versions in any release exist
        rels = [extract_major_release(v["version"]) for v in self.versions]
        has_duplicates = len(rels) > len(set(rels))
        if has_duplicates:
            act_prune = menu.addAction(f"🧹 Prune Older Versions in TS {self.spec_number}")
            act_prune.triggered.connect(lambda: self.prune_spec_requested.emit(self.spec_number))

        act_del = menu.addAction(f"🗑️ Delete TS {self.spec_number} ({len(self.versions)} versions)")
        act_del.triggered.connect(lambda: self.delete_spec_requested.emit(self.spec_number, len(self.versions)))
        menu.exec_(QCursor.pos())

    def _show_pill_menu(self, pill: ReleasePillWidget, pos: QPoint):
        menu = QMenu(self)
        act_del = menu.addAction(f"🗑️ Delete TS {pill.spec_number} v{pill.version}")
        act_del.triggered.connect(lambda: self.delete_version_requested.emit(pill.spec_number, pill.version))
        menu.exec_(pill.mapToGlobal(pos))

    def _on_header_checkbox_clicked(self):
        if self._updating_checks:
            return
        all_checked = all(p.isChecked() for p in self.pills)
        new_state = not all_checked
        self._updating_checks = True
        for p in self.pills:
            p.setChecked(new_state)
        self._updating_checks = False
        self._update_header_state()
        self.selection_changed.emit()

    def _on_pill_toggled(self):
        if self._updating_checks:
            return
        self._update_header_state()
        self.selection_changed.emit()

    def _update_header_state(self):
        self._updating_checks = True
        checked_count = sum(1 for p in self.pills if p.isChecked())
        if checked_count == len(self.pills):
            self.chk_all.setCheckState(Qt.Checked)
        elif checked_count == 0:
            self.chk_all.setCheckState(Qt.Unchecked)
        else:
            self.chk_all.setCheckState(Qt.PartiallyChecked)
        self._updating_checks = False

    def select_all(self):
        self._updating_checks = True
        for p in self.pills:
            p.setChecked(True)
        self._updating_checks = False
        self._update_header_state()
        self.selection_changed.emit()

    def deselect_all(self):
        self._updating_checks = True
        for p in self.pills:
            p.setChecked(False)
        self._updating_checks = False
        self._update_header_state()
        self.selection_changed.emit()

    def set_pill_state(self, version: str, checked: bool):
        for p in self.pills:
            if p.version == version:
                p.setChecked(checked)
                break
        self._update_header_state()
