"""Release pill and specification card widgets for the Indexed Protocols library."""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QPoint, Qt, pyqtSignal
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QAction,
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
from core.ui.ui_components import get_wg_colors
from modules.nas.core.nas_db import parse_version_tuple
from modules.nas.ui.nas_dialogs import get_spec_metadata, get_spec_tooltip


def _extract_major_release(version: str) -> str:
    """Extract major release label (e.g. 'Rel-20') from a 3GPP version string."""
    clean_v = str(version or "").lstrip("vV").strip()
    first = clean_v.split(".")[0] if clean_v else ""
    if first.isdigit():
        return f"Rel-{first}"
    return "Rel-?"


class NASReleasePillWidget(QPushButton):
    """
    Compact clickable pill representing a single protocol specification release.
    Displays e.g. 'R20: 20.0.0' with active/inactive highlight, tooltip, and double-click open.
    """

    open_word_requested = pyqtSignal(str, str)

    def __init__(self, version_data: Dict[str, Any], meta: Dict[str, str], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.version_data = version_data
        self.spec_number = str(version_data.get("spec_number", ""))
        self.version = str(version_data.get("version", ""))
        self.version_id = version_data.get("id")

        rel = _extract_major_release(self.version)
        short_rel = rel.replace("Rel-", "R") if rel.startswith("Rel-") else rel
        self.setText(f"{short_rel}: {self.version}")
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(scale_px(20, minimum=18))

        self.setToolTip(
            f"TS {self.spec_number} v{self.version} ({meta.get('title', '')})\n"
            f"Working Group: {meta.get('wg', '3GPP')} · Release: {rel}\n"
            f"(Click to toggle comparison | Double-click to open Word | Right-click for options)"
        )
        self._update_style()
        self.toggled.connect(lambda _: self._update_style())

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton and self.spec_number and self.version:
            self.open_word_requested.emit(self.spec_number, self.version)
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def _update_style(self):
        if self.isChecked():
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.ACCENT_SKY};
                    color: {C.BG_CARD};
                    border: 1px solid {C.ACCENT_SKY};
                    border-radius: 10px;
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
                    border-radius: 10px;
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


class NASSpecCardWidget(QFrame):
    """
    Modern protocol specification card featuring WG color pill, acronym label,
    tri-state selection checkbox, collapsible release pills, and options menu.
    """

    selection_changed = pyqtSignal()
    collapsed_changed = pyqtSignal()
    reload_version_requested = pyqtSignal(str, str)
    reload_spec_requested = pyqtSignal(str, int)
    delete_version_requested = pyqtSignal(str, str)
    delete_spec_requested = pyqtSignal(str, int)
    open_spec_requested = pyqtSignal(str, str, str, str)
    show_in_specs_requested = pyqtSignal(str, str)
    search_in_spec_search_requested = pyqtSignal(str, str)

    def __init__(
        self,
        spec_number: str,
        versions: List[Dict[str, Any]],
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.spec_number = spec_number
        self.versions = sorted(versions, key=lambda x: parse_version_tuple(x["version"]), reverse=True)
        self.meta = get_spec_metadata(spec_number)
        self.pills: List[NASReleasePillWidget] = []
        self._updating_checks = False
        self._is_collapsed = False

        self.setObjectName("NASSpecCard")
        self.setStyleSheet(f"""
            QFrame#NASSpecCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 6px;
            }}
            QFrame#NASSpecCard:hover {{
                border-color: {C.BORDER_STRONG};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 5, 6, 5)
        layout.setSpacing(4)

        # Card Header Row: [☑] [WG] TS 24.501  5GS NAS   [Count] [▾] [⋮]
        hdr = QHBoxLayout()
        hdr.setContentsMargins(0, 0, 0, 0)
        hdr.setSpacing(5)

        self.chk_all = QCheckBox()
        self.chk_all.setCursor(Qt.PointingHandCursor)
        self.chk_all.clicked.connect(self._on_header_checkbox_clicked)
        hdr.addWidget(self.chk_all)

        wg = str(self.meta.get("wg", "") or "3GPP")
        wg_colors = get_wg_colors(wg)
        wg_badge = QLabel(wg)
        wg_badge.setStyleSheet(f"""
            background-color: {wg_colors.get('bg', C.BG_SURFACE)};
            color: {wg_colors.get('text', C.TEXT_BODY)};
            border: 1px solid {wg_colors.get('border', C.BORDER_STRONG)};
            border-radius: 3px;
            padding: 1px 5px;
            font-size: 10px;
            font-weight: bold;
        """)
        hdr.addWidget(wg_badge)

        lbl_spec = QLabel(f"TS {spec_number}")
        lbl_spec.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: {T.SIZE_TABLE}px; font-weight: bold;")
        lbl_spec.setToolTip(get_spec_tooltip(spec_number))
        hdr.addWidget(lbl_spec)

        acronym = str(self.meta.get("acronym", "") or "")
        lbl_acr = QLabel(acronym)
        lbl_acr.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        lbl_acr.setToolTip(get_spec_tooltip(spec_number))
        hdr.addWidget(lbl_acr, stretch=1)

        self.count_badge = QLabel(str(len(self.versions)))
        self.count_badge.setStyleSheet(f"""
            background-color: {C.BG_CANVAS};
            color: {C.TEXT_SECONDARY};
            border: 1px solid {C.BORDER_SUBTLE};
            border-radius: 8px;
            padding: 1px 6px;
            font-size: 10px;
            font-weight: 600;
        """)
        hdr.addWidget(self.count_badge)

        self.btn_collapse = QToolButton()
        self.btn_collapse.setText("▾")
        self.btn_collapse.setCursor(Qt.PointingHandCursor)
        self.btn_collapse.setToolTip(f"Collapse / expand releases for TS {spec_number}")
        self.btn_collapse.setStyleSheet(f"""
            QToolButton {{
                border: none;
                background: transparent;
                color: {C.TEXT_MUTED};
                font-weight: bold;
                font-size: 11px;
                padding: 0px 3px;
            }}
            QToolButton:hover {{
                color: {C.TEXT_PRIMARY};
                background-color: {C.BG_CANVAS};
                border-radius: 3px;
            }}
        """)
        self.btn_collapse.clicked.connect(lambda: self.set_collapsed(not self._is_collapsed))
        hdr.addWidget(self.btn_collapse)

        self.btn_opt = QToolButton()
        self.btn_opt.setText("⋮")
        self.btn_opt.setCursor(Qt.PointingHandCursor)
        self.btn_opt.setToolTip(f"Actions for TS {spec_number}")
        self.btn_opt.setStyleSheet(f"""
            QToolButton {{
                border: none;
                background: transparent;
                color: {C.TEXT_MUTED};
                font-weight: bold;
                font-size: 12px;
                padding: 0px 3px;
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

        # Flow Layout of Release Pills (explicit .QWidget scope prevents cascading into children)
        self.pills_container = QWidget()
        self.pills_container.setStyleSheet(".QWidget { background: transparent; border: none; }")
        flow = FlowLayout(self.pills_container, margin=0, h_spacing=4, v_spacing=3)

        for v in self.versions:
            pill = NASReleasePillWidget(v, self.meta, self)
            pill.toggled.connect(self._on_pill_toggled)
            pill.open_word_requested.connect(
                lambda sp, vr: self.open_spec_requested.emit(sp, vr, "", "word")
            )
            pill.setContextMenuPolicy(Qt.CustomContextMenu)
            pill.customContextMenuRequested.connect(lambda pos, p=pill: self._show_pill_menu(p, pos))
            self.pills.append(pill)
            flow.addWidget(pill)

        layout.addWidget(self.pills_container)

    def _show_options_menu(self):
        latest_v = self.versions[0]["version"] if self.versions else ""
        menu = QMenu(self)
        act_word = QAction(f"📝 Open Latest TS {self.spec_number} (v{latest_v}) in Word", self)
        act_word.triggered.connect(lambda: self.open_spec_requested.emit(self.spec_number, latest_v, "", "word"))
        menu.addAction(act_word)

        act_folder = QAction("📂 Open Local Folder in Explorer", self)
        act_folder.triggered.connect(lambda: self.open_spec_requested.emit(self.spec_number, latest_v, "", "folder"))
        menu.addAction(act_folder)
        menu.addSeparator()

        act_reload = QAction(f"🔄 Reload all versions of TS {self.spec_number} ({len(self.versions)} versions)", self)
        act_reload.triggered.connect(lambda: self.reload_spec_requested.emit(self.spec_number, len(self.versions)))
        menu.addAction(act_reload)
        menu.addSeparator()

        act_sel_all = QAction("Select All Releases", self)
        act_sel_all.triggered.connect(self.select_all)
        menu.addAction(act_sel_all)

        act_desel = QAction("Deselect All Releases", self)
        act_desel.triggered.connect(self.deselect_all)
        menu.addAction(act_desel)
        menu.addSeparator()

        act_specs = QAction(f"📚 Show TS {self.spec_number} in Specifications Tab", self)
        act_specs.triggered.connect(lambda: self.show_in_specs_requested.emit(self.spec_number, latest_v))
        menu.addAction(act_specs)

        act_search = QAction(f"🔎 Search TS {self.spec_number} in Spec Search Tab", self)
        act_search.triggered.connect(lambda: self.search_in_spec_search_requested.emit(self.spec_number, latest_v))
        menu.addAction(act_search)
        menu.addSeparator()

        act_del = QAction(f"🗑️ Delete all versions of TS {self.spec_number} ({len(self.versions)} versions)", self)
        act_del.triggered.connect(lambda: self.delete_spec_requested.emit(self.spec_number, len(self.versions)))
        menu.addAction(act_del)
        menu.exec_(QCursor.pos())

    def _show_pill_menu(self, pill: NASReleasePillWidget, pos: QPoint):
        menu = QMenu(self)
        act_word = QAction(f"📝 Open TS {pill.spec_number} v{pill.version} (Word)", self)
        act_word.triggered.connect(lambda: self.open_spec_requested.emit(pill.spec_number, pill.version, "", "word"))
        menu.addAction(act_word)

        act_pdf = QAction(f"📕 Open TS {pill.spec_number} v{pill.version} (PDF)", self)
        act_pdf.triggered.connect(lambda: self.open_spec_requested.emit(pill.spec_number, pill.version, "", "pdf"))
        menu.addAction(act_pdf)

        act_folder = QAction("📂 Open Containing Folder in Explorer", self)
        act_folder.triggered.connect(lambda: self.open_spec_requested.emit(pill.spec_number, pill.version, "", "folder"))
        menu.addAction(act_folder)
        menu.addSeparator()

        act_reload = QAction(f"🔄 Reload TS {pill.spec_number} v{pill.version}", self)
        act_reload.triggered.connect(lambda: self.reload_version_requested.emit(pill.spec_number, pill.version))
        menu.addAction(act_reload)
        menu.addSeparator()

        act_specs = QAction(f"📚 Show TS {pill.spec_number} v{pill.version} in Specifications Tab", self)
        act_specs.triggered.connect(lambda: self.show_in_specs_requested.emit(pill.spec_number, pill.version))
        menu.addAction(act_specs)

        act_search = QAction(f"🔎 Search TS {pill.spec_number} in Spec Search Tab", self)
        act_search.triggered.connect(lambda: self.search_in_spec_search_requested.emit(pill.spec_number, pill.version))
        menu.addAction(act_search)
        menu.addSeparator()

        act_del = QAction(f"🗑️ Delete TS {pill.spec_number} v{pill.version}", self)
        act_del.triggered.connect(lambda: self.delete_version_requested.emit(pill.spec_number, pill.version))
        menu.addAction(act_del)
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

    def is_collapsed(self) -> bool:
        return self._is_collapsed

    def set_collapsed(self, collapsed: bool, emit_signal: bool = True):
        self._is_collapsed = bool(collapsed)
        self.pills_container.setVisible(not self._is_collapsed)
        self.btn_collapse.setText("▸" if self._is_collapsed else "▾")
        self._update_badge_text()
        if emit_signal:
            self.collapsed_changed.emit()

    def _update_badge_text(self):
        checked_count = sum(1 for p in self.pills if p.isChecked())
        total = len(self.pills)
        if 0 < checked_count < total:
            self.count_badge.setText(f"{checked_count}/{total}")
        else:
            self.count_badge.setText(str(total))

    def _update_header_state(self):
        self._updating_checks = True
        checked_count = sum(1 for p in self.pills if p.isChecked())
        if self.pills and checked_count == len(self.pills):
            self.chk_all.setCheckState(Qt.Checked)
        elif checked_count == 0:
            self.chk_all.setCheckState(Qt.Unchecked)
        else:
            self.chk_all.setCheckState(Qt.PartiallyChecked)
        self._update_badge_text()
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
