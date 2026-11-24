"""Per-WG Agenda Item pattern pills widget and individual pill chip."""

from __future__ import annotations

from typing import Dict, List

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.theme import C
from modules.meetings.ui.pills.wg_and_filter_pills import get_wg_colors


class AgendaItemPill(QFrame):
    """Pill chip representing a specific Agenda Item pattern within a Working Group."""

    removed = pyqtSignal(str, str)  # wg, pattern

    def __init__(self, wg: str, pattern: str, parent=None):
        super().__init__(parent)
        self.wg = wg.strip().upper()
        self.pattern = pattern.strip()
        colors = get_wg_colors(self.wg)

        self.setStyleSheet(
            f"""
            QFrame {{
                background-color: {colors['bg_active']};
                border: 1px solid {colors['border_active']};
                border-radius: 10px;
                padding: 1px 4px;
            }}
        """
        )
        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 2, 4, 2)
        layout.setSpacing(4)

        lbl = QLabel(f"<b>{self.pattern}</b>")
        lbl.setStyleSheet(f"color: {colors['fg_active']}; font-size: 11px; background: transparent; border: none;")
        layout.addWidget(lbl)

        btn_close = QPushButton("✕")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setFixedSize(14, 14)
        btn_close.setStyleSheet(
            f"""
            QPushButton {{
                color: {colors['fg_active']};
                background: transparent;
                border: none;
                font-size: 10px;
                font-weight: bold;
                padding: 0;
            }}
            QPushButton:hover {{
                color: {C.DANGER_ACCENT};
            }}
        """
        )
        btn_close.clicked.connect(lambda: self.removed.emit(self.wg, self.pattern))
        layout.addWidget(btn_close)


class AgendaItemPillsWidget(QFrame):
    """
    Per-WG grouped Agenda Item pattern manager.
    Displays a row per Target Working Group with removable pattern pills and '+ Add AI' button.
    """

    rulesChanged = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("AiPillsFrame")
        self.setStyleSheet(
            f"""
            QFrame#AiPillsFrame {{
                background-color: {C.BG_ALT_LIGHT};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 6px;
                padding: 6px;
            }}
        """
        )
        self._rules: Dict[str, List[str]] = {}
        self._active_wgs: List[str] = ["SA2", "RAN2", "CT1"]

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(4, 4, 4, 4)
        self.main_layout.setSpacing(6)

        self.lbl_empty = QLabel("No Agenda Items defined. Use '🔍 Auto-Suggest AIs...' or click '+ Add AI' below.")
        self.lbl_empty.setWordWrap(True)
        self.lbl_empty.setStyleSheet(
            f"color: {C.TEXT_FAINT}; font-size: 11px; font-style: italic; padding: 4px; background: transparent; border: none;"
        )
        self.main_layout.addWidget(self.lbl_empty)

        self._rows_container = QWidget()
        self._rows_layout = QVBoxLayout(self._rows_container)
        self._rows_layout.setContentsMargins(0, 0, 0, 0)
        self._rows_layout.setSpacing(6)
        self.main_layout.addWidget(self._rows_container)

    def set_active_wgs(self, wgs: List[str]):
        """Updates which Working Groups are visible/managed based on Target WGs selection."""
        self._active_wgs = [w.strip().upper() for w in wgs if w.strip()]
        self._rebuild_rows()

    def add_pattern(self, wg: str, pattern: str):
        wg = wg.strip().upper()
        pat = pattern.strip()
        if not wg or not pat:
            return
        if wg not in self._rules:
            self._rules[wg] = []
        if pat not in self._rules[wg]:
            self._rules[wg].append(pat)
            self._rebuild_rows()
            self.rulesChanged.emit()

    def remove_pattern(self, wg: str, pattern: str):
        wg = wg.strip().upper()
        pat = pattern.strip()
        if wg in self._rules and pat in self._rules[wg]:
            self._rules[wg].remove(pat)
            if not self._rules[wg]:
                del self._rules[wg]
            self._rebuild_rows()
            self.rulesChanged.emit()

    def _on_add_clicked(self, wg: str):
        text, ok = QInputDialog.getText(
            self,
            f"Add Agenda Item for {wg}",
            f"Enter Agenda Item pattern for {wg} (e.g. 5.14*, 9.3*):",
        )
        if ok and text.strip():
            for p in text.split(","):
                p_c = p.strip()
                if p_c:
                    self.add_pattern(wg, p_c)

    def _rebuild_rows(self):
        # Clear existing rows immediately so back-to-back rebuilds never overlap visually
        while self._rows_layout.count():
            item = self._rows_layout.takeAt(0)
            if item and item.widget():
                w = item.widget()
                w.hide()
                w.setParent(None)
                w.deleteLater()

        has_any_pills = False
        display_wgs = sorted(list(set(self._active_wgs) | set(self._rules.keys())))

        for wg in display_wgs:
            pats = self._rules.get(wg, [])
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(2, 2, 2, 2)
            row_layout.setSpacing(6)

            colors = get_wg_colors(wg)
            lbl_wg = QLabel(f"<b>{wg}:</b>")
            lbl_wg.setFixedWidth(50)
            lbl_wg.setStyleSheet(
                f"color: {colors['fg_active']}; font-size: 11px; padding-top: 3px; background: transparent; border: none;"
            )
            row_layout.addWidget(lbl_wg, 0, Qt.AlignTop)

            flow_container = QWidget()
            flow_layout = FlowLayout(flow_container, margin=0, h_spacing=4, v_spacing=4)

            for pat in pats:
                has_any_pills = True
                pill = AgendaItemPill(wg, pat, self)
                pill.removed.connect(self.remove_pattern)
                flow_layout.addWidget(pill)

            btn_add = QPushButton("+ Add AI")
            btn_add.setCursor(Qt.PointingHandCursor)
            btn_add.setStyleSheet(
                f"""
                QPushButton {{
                    background-color: {C.BG_CARD};
                    color: {C.TEXT_SECONDARY};
                    border: 1px dashed {C.BORDER_STRONG};
                    border-radius: 10px;
                    padding: 2px 8px;
                    font-size: 10px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {C.BG_CANVAS};
                    color: {C.TEXT_PRIMARY};
                    border-color: {C.TEXT_FAINT};
                }}
            """
            )
            btn_add.clicked.connect(lambda _, w=wg: self._on_add_clicked(w))
            flow_layout.addWidget(btn_add)

            row_layout.addWidget(flow_container, 1)
            self._rows_layout.addWidget(row_widget)

        self.lbl_empty.setVisible(not has_any_pills)
        self.updateGeometry()

    def text(self) -> str:
        """Returns standard format: 'SA2: 5.14*, 5.12*; RAN2: 8.7*; CT1: 6.18*' for backward compatibility."""
        out = []
        for wg in sorted(self._rules.keys()):
            pats = self._rules[wg]
            if pats:
                out.append(f"{wg}: {', '.join(pats)}")
        return "; ".join(out)

    def setText(self, rule_str: str):
        """Parses standard format into internal map and updates pills."""
        self._rules = {}
        if rule_str:
            for chunk in str(rule_str).split(";"):
                chunk = chunk.strip()
                if ":" in chunk:
                    wg_k, pats = chunk.split(":", 1)
                    wg_k = wg_k.strip().upper()
                    p_list = [p.strip() for p in pats.split(",") if p.strip()]
                    if p_list:
                        self._rules[wg_k] = p_list
        self._rebuild_rows()
        self.rulesChanged.emit()

    def patterns_map(self) -> Dict[str, List[str]]:
        return {k: list(v) for k, v in self._rules.items()}
