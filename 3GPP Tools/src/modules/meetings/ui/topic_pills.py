# --- File: src/modules/meetings/ui/topic_pills.py ---
"""
Pill & Tag UI Components for Cross-WG Topic Management.

Provides:
- WgTogglePill: Checkable pill button color-coded by Working Group archetype (SA blue, RAN purple, CT teal).
- TagCloudWidget: Interactive tag cloud editor with removable pill chips and inline entry for WI tokens and keywords.
- AgendaItemPillsWidget: Structured per-WG Agenda Item pattern manager with removable pills and inline adding.
"""

import re
from typing import Dict, List, Optional, Set

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QKeyEvent
from PyQt5.QtWidgets import (
    QWidget, QFrame, QHBoxLayout, QVBoxLayout, QLabel, QPushButton,
    QLineEdit, QInputDialog, QScrollArea, QSizePolicy
)

from core.ui.flow_layout import FlowLayout


def get_wg_colors(wg: str) -> Dict[str, str]:
    """Returns archetype color schemes based on Working Group family."""
    wg_u = str(wg or "").strip().upper()
    if wg_u.startswith("SA"):
        return {
            "bg_active": "#DBEAFE",
            "fg_active": "#1E40AF",
            "border_active": "#93C5FD",
            "bg_inactive": "#F8FAFC",
            "fg_inactive": "#64748B",
            "border_inactive": "#CBD5E1",
            "hover": "#EFF6FF",
        }
    elif wg_u.startswith("RAN"):
        return {
            "bg_active": "#EDE9FE",
            "fg_active": "#5B21B6",
            "border_active": "#C4B5FD",
            "bg_inactive": "#F8FAFC",
            "fg_inactive": "#64748B",
            "border_inactive": "#CBD5E1",
            "hover": "#F5F3FF",
        }
    else:  # CT or others
        return {
            "bg_active": "#CCFBF1",
            "fg_active": "#0F766E",
            "border_active": "#99F6E4",
            "bg_inactive": "#F8FAFC",
            "fg_inactive": "#64748B",
            "border_inactive": "#CBD5E1",
            "hover": "#F0FDFA",
        }


class WgTogglePill(QPushButton):
    """
    Checkable pill button for selecting Target Working Groups.
    Displays '✓ WG' when checked and '+ WG' when unchecked with family-specific colors.
    """
    def __init__(self, wg: str, checked: bool = False, parent=None):
        super().__init__(parent)
        self.wg = str(wg).strip().upper()
        self.colors = get_wg_colors(self.wg)
        self.setCheckable(True)
        self.setChecked(checked)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.setCursor(Qt.PointingHandCursor)
        self._update_appearance()
        self.toggled.connect(self._on_toggled)

    def _on_toggled(self, checked: bool):
        self._update_appearance()

    def setChecked(self, checked: bool):
        super().setChecked(checked)
        self._update_appearance()

    def _update_appearance(self):
        c = self.colors
        if self.isChecked():
            self.setText(f"✓ {self.wg}")
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {c['bg_active']};
                    color: {c['fg_active']};
                    border: 1px solid {c['border_active']};
                    border-radius: 12px;
                    padding: 4px 10px;
                    font-size: 11px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    border-color: {c['fg_active']};
                }}
            """)
        else:
            self.setText(f"+ {self.wg}")
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {c['bg_inactive']};
                    color: {c['fg_inactive']};
                    border: 1px solid {c['border_inactive']};
                    border-radius: 12px;
                    padding: 4px 10px;
                    font-size: 11px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {c['hover']};
                    color: #1E293B;
                    border-color: #94A3B8;
                }}
            """)


class TagPillItem(QFrame):
    """Individual tag pill with a remove button."""
    removed = pyqtSignal(str)

    def __init__(self, text: str, parent=None):
        super().__init__(parent)
        self.tag_text = text
        self.setStyleSheet("""
            QFrame {
                background-color: #F1F5F9;
                border: 1px solid #CBD5E1;
                border-radius: 11px;
                padding: 1px 4px;
            }
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 2, 4, 2)
        layout.setSpacing(4)

        lbl = QLabel(text)
        lbl.setStyleSheet("color: #334155; font-size: 11px; font-weight: 500; background: transparent; border: none;")
        layout.addWidget(lbl)

        btn_close = QPushButton("✕")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setFixedSize(14, 14)
        btn_close.setStyleSheet("""
            QPushButton {
                color: #94A3B8;
                background: transparent;
                border: none;
                font-size: 10px;
                font-weight: bold;
                padding: 0;
            }
            QPushButton:hover {
                color: #EF4444;
            }
        """)
        btn_close.clicked.connect(lambda: self.removed.emit(self.tag_text))
        layout.addWidget(btn_close)


class _TagLineEdit(QLineEdit):
    """Custom QLineEdit that captures Backspace to remove last tag and handles comma separation."""
    backspace_pressed = pyqtSignal()

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key_Backspace and not self.text():
            self.backspace_pressed.emit()
            return
        super().keyPressEvent(event)


class TagCloudWidget(QFrame):
    """
    Modern Tag Cloud input displaying tokens as removable pills with an inline editor.
    Supports comma separation, Enter to add, paste with comma parsing, and backspace removal.
    """
    tagsChanged = pyqtSignal()

    def __init__(self, placeholder: str = "Type term and press Enter or comma...", parent=None):
        super().__init__(parent)
        self.setObjectName("TagCloudFrame")
        self.setStyleSheet("""
            QFrame#TagCloudFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 3px;
            }
            QFrame#TagCloudFrame:focus-within {
                border-color: #3B82F6;
            }
        """)
        self._tags: List[str] = []

        self.flow_layout = FlowLayout(self, margin=2, h_spacing=6, v_spacing=4)

        self.input_edit = _TagLineEdit()
        self.input_edit.setPlaceholderText(placeholder)
        self.input_edit.setStyleSheet("border: none; background: transparent; font-size: 11px; padding: 2px 4px;")
        self.input_edit.returnPressed.connect(self._on_enter_pressed)
        self.input_edit.textChanged.connect(self._on_text_changed)
        self.input_edit.backspace_pressed.connect(self._on_backspace_pressed)
        self.flow_layout.addWidget(self.input_edit)

    def _on_enter_pressed(self):
        text = self.input_edit.text()
        self._ingest_text(text)
        self.input_edit.clear()

    def _on_text_changed(self, text: str):
        if "," in text or ";" in text:
            self._ingest_text(text)
            self.input_edit.clear()

    def _on_backspace_pressed(self):
        if self._tags:
            self.remove_tag(self._tags[-1])

    def _ingest_text(self, text: str):
        parts = re.split(r'[,;]+', text)
        changed = False
        for p in parts:
            p_clean = p.strip()
            if p_clean and p_clean.lower() not in [t.lower() for t in self._tags]:
                self._tags.append(p_clean)
                changed = True
        if changed:
            self._rebuild_pills()
            self.tagsChanged.emit()

    def add_tag(self, tag: str):
        tag_clean = tag.strip()
        if tag_clean and tag_clean.lower() not in [t.lower() for t in self._tags]:
            self._tags.append(tag_clean)
            self._rebuild_pills()
            self.tagsChanged.emit()

    def remove_tag(self, tag: str):
        target = tag.strip().lower()
        self._tags = [t for t in self._tags if t.lower() != target]
        self._rebuild_pills()
        self.tagsChanged.emit()

    def _rebuild_pills(self):
        # Remove all pill widgets from flow layout
        self.flow_layout.clear()

        # Add tag pill widgets
        for t in self._tags:
            pill = TagPillItem(t, self)
            pill.removed.connect(self.remove_tag)
            self.flow_layout.addWidget(pill)

        # Re-add the input edit at the end
        self.flow_layout.addWidget(self.input_edit)
        self.updateGeometry()

    def tags(self) -> List[str]:
        return list(self._tags)

    def text(self) -> str:
        """Returns comma-separated string for backward compatibility."""
        return ", ".join(self._tags)

    def setText(self, text: str):
        """Sets tags from comma-separated or semicolon-separated string."""
        self._tags = []
        if text:
            parts = re.split(r'[,;]+', str(text))
            for p in parts:
                p_clean = p.strip()
                if p_clean and p_clean.lower() not in [t.lower() for t in self._tags]:
                    self._tags.append(p_clean)
        self._rebuild_pills()
        self.tagsChanged.emit()


class AgendaItemPill(QFrame):
    """Pill chip representing a specific Agenda Item pattern within a Working Group."""
    removed = pyqtSignal(str, str)  # wg, pattern

    def __init__(self, wg: str, pattern: str, parent=None):
        super().__init__(parent)
        self.wg = wg.strip().upper()
        self.pattern = pattern.strip()
        colors = get_wg_colors(self.wg)

        self.setStyleSheet(f"""
            QFrame {{
                background-color: {colors['bg_active']};
                border: 1px solid {colors['border_active']};
                border-radius: 10px;
                padding: 1px 4px;
            }}
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 2, 4, 2)
        layout.setSpacing(4)

        lbl = QLabel(f"<b>{self.pattern}</b>")
        lbl.setStyleSheet(f"color: {colors['fg_active']}; font-size: 11px; background: transparent; border: none;")
        layout.addWidget(lbl)

        btn_close = QPushButton("✕")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setFixedSize(14, 14)
        btn_close.setStyleSheet(f"""
            QPushButton {{
                color: {colors['fg_active']};
                background: transparent;
                border: none;
                font-size: 10px;
                font-weight: bold;
                padding: 0;
            }}
            QPushButton:hover {{
                color: #EF4444;
            }}
        """)
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
        self.setStyleSheet("""
            QFrame#AiPillsFrame {
                background-color: #FAFAFA;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 6px;
            }
        """)
        self._rules: Dict[str, List[str]] = {}
        self._active_wgs: List[str] = ["SA2", "RAN2", "CT1"]

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(4, 4, 4, 4)
        self.main_layout.setSpacing(6)

        self.lbl_empty = QLabel("No Agenda Items defined. Use '🔍 Auto-Suggest AIs...' or click '+ Add AI' below.")
        self.lbl_empty.setStyleSheet("color: #94A3B8; font-size: 11px; font-style: italic; padding: 4px;")
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
        # Clear existing rows
        while self._rows_layout.count():
            item = self._rows_layout.takeAt(0)
            if item and item.widget():
                item.widget().deleteLater()

        has_any_pills = False
        # Display rows for all active WGs that either have patterns or are in active_wgs
        display_wgs = sorted(list(set(self._active_wgs) | set(self._rules.keys())))

        for wg in display_wgs:
            pats = self._rules.get(wg, [])
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(2, 2, 2, 2)
            row_layout.setSpacing(6)

            # WG Badge
            colors = get_wg_colors(wg)
            lbl_wg = QLabel(f"<b>{wg}:</b>")
            lbl_wg.setFixedWidth(50)
            lbl_wg.setStyleSheet(f"color: {colors['fg_active']}; font-size: 11px;")
            row_layout.addWidget(lbl_wg)

            # Pills in a FlowLayout
            flow_container = QWidget()
            flow_layout = FlowLayout(flow_container, margin=0, h_spacing=4, v_spacing=4)

            for pat in pats:
                has_any_pills = True
                pill = AgendaItemPill(wg, pat, self)
                pill.removed.connect(self.remove_pattern)
                flow_layout.addWidget(pill)

            btn_add = QPushButton("+ Add AI")
            btn_add.setCursor(Qt.PointingHandCursor)
            btn_add.setStyleSheet("""
                QPushButton {
                    background-color: #FFFFFF;
                    color: #475569;
                    border: 1px dashed #CBD5E1;
                    border-radius: 10px;
                    padding: 2px 8px;
                    font-size: 10px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #F1F5F9;
                    color: #0F172A;
                    border-color: #94A3B8;
                }
            """)
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


class FilterPillButton(QPushButton):
    """
    Clickable, checkable filter pill with title, count badge, and custom color scheme.
    Used for 1-click filtering by Working Group and Decision Category.
    """
    def __init__(self, key: str, label: str, count: Optional[int] = None, colors: Optional[Dict[str, str]] = None, parent=None):
        super().__init__(parent)
        self.key = key
        self.base_label = label
        self._count = count
        self.colors = colors or {
            "bg_active": "#2563EB",
            "fg_active": "#FFFFFF",
            "border_active": "#1D4ED8",
            "bg_inactive": "#FFFFFF",
            "fg_inactive": "#475569",
            "border_inactive": "#CBD5E1",
            "hover": "#F8FAFC",
        }
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self._update_text_and_style()
        self.toggled.connect(self._on_toggled)

    def set_count(self, count: Optional[int]):
        self._count = count
        self._update_text_and_style()

    def _on_toggled(self, checked: bool):
        self._update_text_and_style()

    def _update_text_and_style(self):
        c = self.colors
        cnt_str = f" ({self._count})" if self._count is not None and self._count >= 0 else ""
        self.setText(f"{self.base_label}{cnt_str}")
        if self.isChecked():
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {c['bg_active']};
                    color: {c['fg_active']};
                    border: 1px solid {c['border_active']};
                    border-radius: 11px;
                    padding: 3px 10px;
                    font-size: 11px;
                    font-weight: bold;
                }}
            """)
        else:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {c['bg_inactive']};
                    color: {c['fg_inactive']};
                    border: 1px solid {c['border_inactive']};
                    border-radius: 11px;
                    padding: 3px 10px;
                    font-size: 11px;
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {c['hover']};
                    border-color: #94A3B8;
                    color: #0F172A;
                }}
            """)

