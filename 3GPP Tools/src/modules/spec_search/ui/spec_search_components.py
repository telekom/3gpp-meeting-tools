"""
UI Components: Version Tree Widget and Rich Clause Content Inspector with match excerpts,
term navigation, one-click citation copying, direct document opening, and persistent selection state.
"""

import html
import re
from typing import Any, Dict, List, Optional
from PyQt5.QtCore import QModelIndex, QPoint, QRect, Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QCursor, QFont, QFontMetrics, QPainter, QPen, QTextCursor, QTextDocument
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QCheckBox,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTextBrowser,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.spec_search.core.spec_clause_exporter import copy_clause_to_clipboard
from core.ui.flow_layout import FlowLayout
from modules.spec_search.core.spec_search_db import extract_major_release, parse_version_tuple


# -----------------------------------------------------------------------------
# Inspector Button Styles
# -----------------------------------------------------------------------------
INSPECTOR_BTN_BASE = """
QPushButton {
    font-size: 11px;
    font-weight: 500;
    color: #334155;
    background-color: #F8FAFC;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}
QPushButton:hover {
    background-color: #F1F5F9;
    color: #0F172A;
    border-color: #94A3B8;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

INSPECTOR_BTN_SUCCESS = """
QPushButton {
    font-size: 11px;
    font-weight: 600;
    color: #166534;
    background-color: #F0FDF4;
    border: 1px solid #86EFAC;
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}
QPushButton:hover {
    background-color: #DCFCE7;
    color: #14532D;
    border-color: #4ADE80;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

INSPECTOR_BTN_AI = """
QPushButton {
    font-size: 11px;
    font-weight: 600;
    color: #6D28D9;
    background-color: #F5F3FF;
    border: 1px solid #DDD6FE;
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}
QPushButton:hover {
    background-color: #EDE9FE;
    color: #5B21B6;
    border-color: #C4B5FD;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

INSPECTOR_BTN_WARNING = """
QPushButton {
    font-size: 11px;
    font-weight: 500;
    color: #92400E;
    background-color: #FFFBEB;
    border: 1px solid #FDE68A;
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}
QPushButton:hover {
    background-color: #FEF3C7;
    color: #78350F;
    border-color: #F59E0B;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

INSPECTOR_BTN_INFO = """
QPushButton {
    font-size: 11px;
    font-weight: 500;
    color: #1D4ED8;
    background-color: #EFF6FF;
    border: 1px solid #BFDBFE;
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 22px;
}
QPushButton:hover {
    background-color: #DBEAFE;
    color: #1E40AF;
    border-color: #93C5FD;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""


# -----------------------------------------------------------------------------
# Evolution Matrix Custom Delegate (Sleek Rounded Badges)
# -----------------------------------------------------------------------------
class SpecEvolutionStatusDelegate(QStyledItemDelegate):
    """
    Renders status cells in the Evolution Matrix as sleek, centered, rounded pill badges
    (🟢 Added, ⚡ Post-Cutoff, ✓ Present, or —) without obscuring row selection.
    """

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex):
        col = index.column()

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)

        is_selected = bool(option.state & QStyle.State_Selected)
        is_hover = bool(option.state & QStyle.State_MouseOver)

        # 1. Background fill
        if is_selected:
            bg_brush = QBrush(QColor("#EBF3FC"))
        elif is_hover:
            bg_brush = QBrush(QColor("#F8FAFC"))
        elif index.row() % 2 == 1:
            bg_brush = QBrush(QColor("#F8FAFC"))
        else:
            bg_brush = QBrush(QColor("#FFFFFF"))
        painter.fillRect(option.rect, bg_brush)

        # Subtle bottom separator line
        painter.setPen(QPen(QColor("#F1F5F9"), 1))
        painter.drawLine(option.rect.bottomLeft(), option.rect.bottomRight())

        # 2. Cell Content
        if col == 0:
            val = str(index.data(Qt.DisplayRole) or "")
            text_rect = option.rect.adjusted(8, 0, -8, 0)
            painter.setFont(QFont("Segoe UI", 9, QFont.Bold))
            painter.setPen(QColor("#0284C7" if not is_selected else "#0369A1"))
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, val)
        elif col == 1:
            val = str(index.data(Qt.DisplayRole) or "")
            text_rect = option.rect.adjusted(8, 0, -8, 0)
            painter.setFont(QFont("Segoe UI", 9, QFont.Normal))
            painter.setPen(QColor("#0F172A" if not is_selected else "#1E293B"))
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, val)
        else:
            display_text = str(index.data(Qt.DisplayRole) or "").strip()
            if not display_text or display_text in ("-", "—"):
                painter.setFont(QFont("Segoe UI", 10, QFont.Normal))
                painter.setPen(QColor("#CBD5E1"))
                painter.drawText(option.rect, Qt.AlignCenter, "—")
            else:
                if "Post-Cutoff" in display_text or "⚡" in display_text:
                    pill_bg = QColor("#FEF3C7")
                    pill_border = QColor("#FDE68A")
                    pill_fg = QColor("#B45309")
                    badge_label = "⚡ Post-Cutoff"
                elif "Added" in display_text or "🟢" in display_text:
                    pill_bg = QColor("#DCFCE7")
                    pill_border = QColor("#86EFAC")
                    pill_fg = QColor("#15803D")
                    badge_label = "🟢 Added"
                elif "Present" in display_text or "✓" in display_text:
                    pill_bg = QColor("#F1F5F9")
                    pill_border = QColor("#E2E8F0")
                    pill_fg = QColor("#475569")
                    badge_label = "✓ Present"
                else:
                    pill_bg = QColor("#F8FAFC")
                    pill_border = QColor("#E2E8F0")
                    pill_fg = QColor("#334155")
                    badge_label = display_text

                pill_font = QFont("Segoe UI", 8, QFont.Bold)
                fm = QFontMetrics(pill_font)
                text_w = fm.horizontalAdvance(badge_label)
                pill_h = 20
                pill_w = max(68, min(option.rect.width() - 8, text_w + 14))

                cx = option.rect.center().x()
                cy = option.rect.center().y()
                pill_rect = QRect(cx - pill_w // 2, cy - pill_h // 2, pill_w, pill_h)

                painter.setBrush(QBrush(pill_bg))
                painter.setPen(QPen(pill_border, 1))
                painter.drawRoundedRect(pill_rect, 10, 10)

                painter.setFont(pill_font)
                painter.setPen(pill_fg)
                painter.drawText(pill_rect, Qt.AlignCenter, badge_label)

        painter.restore()


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
        self.setFixedHeight(22)

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
            self.setStyleSheet("""
                QPushButton {
                    background-color: #0284C7;
                    color: #FFFFFF;
                    border: 1px solid #0284C7;
                    border-radius: 11px;
                    padding: 1px 7px;
                    font-size: 11px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #0369A1;
                    border-color: #0369A1;
                }
            """)
        else:
            self.setStyleSheet("""
                QPushButton {
                    background-color: #F8FAFC;
                    color: #64748B;
                    border: 1px solid #CBD5E1;
                    border-radius: 11px;
                    padding: 1px 7px;
                    font-size: 11px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #F1F5F9;
                    color: #0F172A;
                    border-color: #94A3B8;
                }
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
        self.setStyleSheet("""
            QFrame#SpecCard {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
            }
            QFrame#SpecCard:hover {
                border-color: #CBD5E1;
            }
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
        font_spec = QFont()
        font_spec.setBold(True)
        font_spec.setPointSize(9)
        lbl_spec.setFont(font_spec)
        lbl_spec.setStyleSheet("color: #0F172A;")
        hdr.addWidget(lbl_spec)

        clean_title = re.sub(r'^(NR;\s*|LTE;\s*|Universal Mobile Telecommunications System \(UMTS\);\s*)', '', title, flags=re.IGNORECASE).strip()
        parts = [p.strip() for p in clean_title.split(';') if p.strip()]
        display_title = parts[0] if parts else clean_title
        if len(display_title) > 36:
            display_title = display_title[:35].rstrip() + "…"

        self.lbl_title = QLabel(display_title)
        self.lbl_title.setStyleSheet("color: #64748B; font-size: 11px;")
        if title:
            self.lbl_title.setToolTip(f"TS {spec_number}: {title}\n{len(versions)} indexed release(s)")
        hdr.addWidget(self.lbl_title, stretch=1)

        self.badge = QLabel(str(len(versions)))
        self.badge.setStyleSheet("""
            background-color: #F1F5F9;
            color: #475569;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: 1px 6px;
            font-size: 10px;
            font-weight: bold;
        """)
        hdr.addWidget(self.badge)

        self.btn_opt = QToolButton()
        self.btn_opt.setText("⋮")
        self.btn_opt.setCursor(Qt.PointingHandCursor)
        self.btn_opt.setStyleSheet("""
            QToolButton {
                border: none;
                background: transparent;
                color: #64748B;
                font-weight: bold;
                font-size: 12px;
                padding: 0px 2px;
            }
            QToolButton:hover {
                color: #0F172A;
                background-color: #F1F5F9;
                border-radius: 3px;
            }
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
        act_smart = menu.addAction(f"🔄 Check & Ingest Latest Releases for TS {self.spec_number}...")
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


class _TreeCompatChildItem:
    def __init__(self, pill: ReleasePillWidget):
        self.pill = pill

    def text(self, col: int = 0) -> str:
        date_lbl = f" ({self.pill.release_date})" if self.pill.release_date else ""
        return f"v{self.pill.version}{date_lbl}"

    def toolTip(self, col: int = 0) -> str:
        return self.pill.toolTip()

    def data(self, col: int, role: int = Qt.UserRole) -> Any:
        return {
            "type": "version",
            "id": self.pill.version_id,
            "spec_number": self.pill.spec_number,
            "version": self.pill.version,
            "release_date": self.pill.release_date,
        }

    def checkState(self, col: int = 0) -> Qt.CheckState:
        return Qt.Checked if self.pill.isChecked() else Qt.Unchecked

    def setCheckState(self, col: int, state: Qt.CheckState):
        self.pill.setChecked(state == Qt.Checked)


class _TreeCompatSpecItem:
    def __init__(self, card: SpecCardWidget):
        self.card = card

    def text(self, col: int = 0) -> str:
        t = self.card.raw_title
        if t:
            clean = re.sub(r'^(NR;\s*|LTE;\s*|Universal Mobile Telecommunications System \(UMTS\);\s*)', '', t, flags=re.IGNORECASE).strip()
            parts = [p.strip() for p in clean.split(';') if p.strip()]
            display_title = parts[0] if parts else clean
            if len(display_title) > 40:
                display_title = display_title[:39].rstrip() + "…"
            return f"TS {self.card.spec_number}: {display_title}  ({len(self.card.pills)})"
        return f"TS {self.card.spec_number}  ({len(self.card.pills)})"

    def toolTip(self, col: int = 0) -> str:
        if self.card.raw_title:
            return f"TS {self.card.spec_number}: {self.card.raw_title}\n{len(self.card.pills)} indexed release(s)"
        return f"TS {self.card.spec_number}\n{len(self.card.pills)} indexed release(s)"

    def data(self, col: int, role: int = Qt.UserRole) -> Any:
        return {
            "type": "spec",
            "spec_number": self.card.spec_number,
            "title": self.card.raw_title,
        }

    def isHidden(self) -> bool:
        return self.card.isHidden()

    def setHidden(self, val: bool):
        self.card.setHidden(val)

    def setExpanded(self, val: bool):
        pass

    def childCount(self) -> int:
        return len(self.card.pills)

    def child(self, idx: int) -> _TreeCompatChildItem:
        return _TreeCompatChildItem(self.card.pills[idx])

    def checkState(self, col: int = 0) -> Qt.CheckState:
        return self.card.chk_all.checkState()

    def setCheckState(self, col: int, state: Qt.CheckState):
        if state == Qt.Checked:
            self.card.select_all()
        else:
            self.card.deselect_all()


class _TreeCompatRootItem:
    def __init__(self, widget: "SpecSearchVersionCardsWidget"):
        self.widget = widget

    def text(self, col: int = 0) -> str:
        return "All Indexed Specifications"

    def data(self, col: int, role: int = Qt.UserRole) -> Any:
        return {"type": "all"}

    def isHidden(self) -> bool:
        return False

    def setHidden(self, val: bool):
        pass

    def checkState(self, col: int = 0) -> Qt.CheckState:
        total = sum(len(c.pills) for c in self.widget._cards.values())
        checked = len(self.widget.get_selected_version_ids())
        if checked == total and total > 0:
            return Qt.Checked
        elif checked == 0:
            return Qt.Unchecked
        return Qt.PartiallyChecked

    def setCheckState(self, col: int, state: Qt.CheckState):
        if state == Qt.Checked:
            self.widget.select_all()
        else:
            self.widget.deselect_all()


class SpecSearchVersionCardsWidget(QScrollArea):
    """
    Modern Concept B: Release Pills Flow Widget.
    Replaces the deep vertical tree with clean, compact specification cards featuring
    interactive release pills, 1-click status badges, and quick-filtering.
    Includes full backward-compatibility adapters for legacy tree tests and callers.
    """
    selection_changed = pyqtSignal()
    delete_version_requested = pyqtSignal(str, str)
    delete_spec_requested = pyqtSignal(str, int)
    prune_spec_requested = pyqtSignal(str)
    smart_import_requested = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setStyleSheet("""
            QScrollArea {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #F8FAFC;
            }
        """)

        self.container = QWidget()
        self.container.setStyleSheet("background-color: #F8FAFC;")
        self.cards_layout = QVBoxLayout(self.container)
        self.cards_layout.setContentsMargins(4, 4, 4, 4)
        self.cards_layout.setSpacing(6)
        self.cards_layout.addStretch()
        self.setWidget(self.container)

        self._cards: Dict[str, SpecCardWidget] = {}
        self._all_item_compat = _TreeCompatRootItem(self)

    def populate(
        self,
        versions: List[Dict[str, Any]],
        saved_checked: Optional[List[Dict[str, Any]]] = None,
        spec_titles: Optional[Dict[str, str]] = None,
    ):
        self.clear()
        if not versions:
            self.selection_changed.emit()
            return

        saved_tuples = (
            {(str(i.get("spec_number")), str(i.get("version"))) for i in saved_checked if i}
            if saved_checked is not None
            else None
        )

        specs_map: Dict[str, List[Dict[str, Any]]] = {}
        for v in versions:
            specs_map.setdefault(v["spec_number"], []).append(v)

        # Remove stretch before adding cards
        count = self.cards_layout.count()
        if count > 0:
            last = self.cards_layout.itemAt(count - 1)
            if last.spacerItem():
                self.cards_layout.removeItem(last)

        for spec_num in sorted(specs_map.keys()):
            title = ""
            if spec_titles:
                clean_num = re.sub(r'^(TS|TR)\s*', '', spec_num.strip(), flags=re.IGNORECASE)
                title = spec_titles.get(clean_num, "")
                if not title:
                    base_num = clean_num.split('-')[0]
                    title = spec_titles.get(base_num, "")

            card = SpecCardWidget(spec_num, title, specs_map[spec_num], self.container)
            card.selection_changed.connect(self.selection_changed.emit)
            card.delete_version_requested.connect(self.delete_version_requested.emit)
            card.delete_spec_requested.connect(self.delete_spec_requested.emit)
            card.prune_spec_requested.connect(self.prune_spec_requested.emit)
            card.smart_import_requested.connect(self.smart_import_requested.emit)

            # Restore checked states
            for pill in card.pills:
                if saved_tuples is not None:
                    is_chk = (str(spec_num), str(pill.version)) in saved_tuples
                else:
                    is_chk = True
                pill.setChecked(is_chk)
            card._update_header_state()

            self._cards[spec_num] = card
            self.cards_layout.addWidget(card)

        self.cards_layout.addStretch()
        self.selection_changed.emit()

    def clear(self):
        while self.cards_layout.count():
            item = self.cards_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._cards.clear()
        self.cards_layout.addStretch()

    def filter_tree(self, query: str):
        query = query.strip().lower()
        if not query:
            for card in self._cards.values():
                card.setVisible(True)
            return

        for spec_num, card in self._cards.items():
            num_match = query in spec_num.lower()
            title_match = query in card.raw_title.lower()
            pill_match = any(query in p.version.lower() or query in p.text().lower() for p in card.pills)
            card.setVisible(num_match or title_match or pill_match)

    def select_all(self):
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                p.setChecked(True)
            card._updating_checks = False
            card._update_header_state()
        self.selection_changed.emit()

    def deselect_all(self):
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                p.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self.selection_changed.emit()

    def select_latest_per_release(self):
        """Checks the newest version of each major Release for each specification."""
        for card in self._cards.values():
            card._updating_checks = True
            seen_releases = set()
            for pill in card.pills:
                rel = extract_major_release(pill.version)
                if rel not in seen_releases:
                    seen_releases.add(rel)
                    pill.setChecked(True)
                else:
                    pill.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self.selection_changed.emit()

    def select_releases_min(self, min_rel_num: int):
        """Checks only releases with major release number >= min_rel_num (e.g. 18 for Rel-18+)."""
        for card in self._cards.values():
            card._updating_checks = True
            for pill in card.pills:
                rel = extract_major_release(pill.version)
                match = re.search(r'\d+', rel)
                if match and int(match.group()) >= min_rel_num:
                    pill.setChecked(True)
                else:
                    pill.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self.selection_changed.emit()

    def get_selected_version_ids(self) -> List[int]:
        ids = []
        for card in self._cards.values():
            for p in card.pills:
                if p.isChecked() and p.version_id:
                    ids.append(p.version_id)
        return ids

    def get_checked_versions_info(self) -> List[Dict[str, Any]]:
        info = []
        for card in self._cards.values():
            for p in card.pills:
                if p.isChecked():
                    info.append({
                        "spec_number": card.spec_number,
                        "version": p.version,
                    })
        return info

    def get_all_versions_for_spec(self, spec_number: str) -> List[Dict[str, Any]]:
        card = self._cards.get(spec_number)
        if not card:
            return []
        return [{"spec_number": card.spec_number, "version": p.version} for p in card.pills]

    def focus_spec(self, spec_number: str, version: Optional[str] = None) -> bool:
        """Scrolls to and highlights a specific specification card and version."""
        card = self._cards.get(spec_number)
        if not card:
            return False
        card.setVisible(True)
        if version:
            card.set_pill_state(version, True)
        self.ensureWidgetVisible(card)
        return True

    # --- Tree Compatibility Adapters ---
    def topLevelItemCount(self) -> int:
        return 1 + len(self._cards)

    def topLevelItem(self, idx: int):
        if idx == 0:
            return self._all_item_compat
        spec_keys = sorted(self._cards.keys())
        card_idx = idx - 1
        if 0 <= card_idx < len(spec_keys):
            return _TreeCompatSpecItem(self._cards[spec_keys[card_idx]])
        return None

    def setCurrentItem(self, item):
        pass

    def scrollToItem(self, item):
        pass


# Backward compatibility alias
SpecSearchVersionTreeWidget = SpecSearchVersionCardsWidget


class SpecClauseInspector(QGroupBox):
    """
    Renders rich clause content with context excerpts, match navigation, citation copying,
    1-click Word document launching, and deep-linking to the Specifications catalogue.
    Features a clean two-tier header and flexible layout orientation controls.
    """

    compare_clause_requested = pyqtSignal(str, str, str, str)  # spec_number, version, clause_number, clause_title
    open_document_requested = pyqtSignal(str, str, str, str)    # spec_number, version, clause_number, format ("word", "pdf", "folder")
    show_in_specs_tab_requested = pyqtSignal(str, str)          # spec_number, version
    quick_diff_requested = pyqtSignal(str, str, str, str, str)  # spec_number, base_ver, target_ver, clause_number, clause_title
    jump_to_version_requested = pyqtSignal(str, str, str)       # spec_number, target_version, clause_number
    extract_clause_requested = pyqtSignal(str, str, str, str, str, bool)  # spec_number, version, clause_number, clause_title, content, save_as
    export_markdown_requested = pyqtSignal(str, str, str, str, str)       # spec_number, version, clause_number, clause_title, content
    layout_orientation_toggle_requested = pyqtSignal(bool)      # True = side-by-side (horizontal), False = bottom (vertical)
    maximize_toggle_requested = pyqtSignal(bool)                # True = maximized, False = restored
    version_changes_requested = pyqtSignal(str, str)            # spec_number, version
    identify_origin_tdoc_requested = pyqtSignal(str, str, str, str, str)  # spec_number, base_ver, target_ver, clause_number, clause_title

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__("Clause Content Inspector", parent)
        self._current_search_query = ""
        self._last_citation_text = ""
        self._current_spec = ""
        self._current_version = ""
        self._current_clause_num = ""
        self._current_clause_title = ""
        self._current_content = ""
        self._current_release_date = ""
        self._first_added_ver = ""
        self._prev_ver = ""
        self._is_side_by_side = False
        self._is_maximized = False
        self.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #1E293B;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                margin-top: 6px;
                padding-top: 10px;
                background-color: #FFFFFF;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }
        """)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 8)
        layout.setSpacing(5)

        # ---------------------------------------------------------------------
        # Tier 1: Context Title, First Added Chip, Matches & Layout Toggles
        # ---------------------------------------------------------------------
        tier1_bar = QHBoxLayout()
        tier1_bar.setSpacing(6)

        self.title_lbl = QLabel("No clause selected")
        self.title_lbl.setStyleSheet("font-weight: bold; color: #0284C7; font-size: 12px;")
        tier1_bar.addWidget(self.title_lbl, stretch=1)

        # Clickable First Introduced Chip
        self.first_added_badge = QPushButton("")
        self.first_added_badge.setCursor(Qt.PointingHandCursor)
        self.first_added_badge.setToolTip("Click to jump directly to the version where this clause was first introduced")
        self.first_added_badge.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                font-size: 11px;
                background-color: #DCFCE7;
                color: #15803D;
                border: 1px solid #86EFAC;
                padding: 1px 8px;
                border-radius: 10px;
            }
            QPushButton:hover {
                background-color: #BBF7D0;
                border-color: #4ADE80;
            }
        """)
        self.first_added_badge.setVisible(False)
        self.first_added_badge.clicked.connect(self._on_first_added_clicked)
        tier1_bar.addWidget(self.first_added_badge)

        # Match navigation group
        self.match_badge = QLabel("")
        self.match_badge.setStyleSheet("font-weight: bold; font-size: 11px; color: #B45309; background-color: #FEF3C7; border: 1px solid #FDE68A; padding: 2px 6px; border-radius: 4px;")
        self.match_badge.setVisible(False)
        tier1_bar.addWidget(self.match_badge)

        self.btn_prev = QPushButton("◀")
        self.btn_prev.setToolTip("Jump to previous match in clause")
        self.btn_prev.setCursor(Qt.PointingHandCursor)
        self.btn_prev.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { min-width: 20px; font-weight: bold; }")
        self.btn_prev.clicked.connect(self._find_prev)
        self.btn_prev.setEnabled(False)
        tier1_bar.addWidget(self.btn_prev)

        self.btn_next = QPushButton("▶")
        self.btn_next.setToolTip("Jump to next match in clause")
        self.btn_next.setCursor(Qt.PointingHandCursor)
        self.btn_next.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { min-width: 20px; font-weight: bold; }")
        self.btn_next.clicked.connect(self._find_next)
        self.btn_next.setEnabled(False)
        tier1_bar.addWidget(self.btn_next)

        # Layout Orientation Toggle (Bottom Split <-> Right Pane)
        self.btn_layout_toggle = QPushButton("◫ Right Pane")
        self.btn_layout_toggle.setToolTip("Toggle inspector layout: Bottom split vs. Right pane side-by-side")
        self.btn_layout_toggle.setCursor(Qt.PointingHandCursor)
        self.btn_layout_toggle.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { font-weight: 600; padding: 2px 8px; }")
        self.btn_layout_toggle.clicked.connect(self._on_layout_toggle_clicked)
        tier1_bar.addWidget(self.btn_layout_toggle)

        # Maximize / Restore Toggle
        self.btn_maximize = QPushButton("⤢")
        self.btn_maximize.setToolTip("Maximize reading pane to full height")
        self.btn_maximize.setCursor(Qt.PointingHandCursor)
        self.btn_maximize.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { font-weight: bold; min-width: 22px; padding: 2px 6px; }")
        self.btn_maximize.clicked.connect(self._on_maximize_clicked)
        tier1_bar.addWidget(self.btn_maximize)

        layout.addLayout(tier1_bar)

        # ---------------------------------------------------------------------
        # Tier 2: Grouped Action Toolbar (Copy, Extract, Full Doc, Diff, LLM)
        # ---------------------------------------------------------------------
        tier2_bar = QHBoxLayout()
        tier2_bar.setSpacing(6)

        # Group 1: Copy & Extract
        self.btn_copy = QPushButton("📋 Copy ▾")
        self.btn_copy.setToolTip("Copy clause text (Rich HTML for Email/Outlook, Markdown quote block, or Plain Text)")
        self.btn_copy.setCursor(Qt.PointingHandCursor)
        self.btn_copy.setStyleSheet(INSPECTOR_BTN_BASE)
        copy_menu = QMenu(self)
        act_cp_email = copy_menu.addAction("📋 Copy for Email / Outlook (Rich Text)")
        act_cp_email.triggered.connect(lambda: self._copy_formatted("email"))
        act_cp_md = copy_menu.addAction("📝 Copy as Markdown (Quote Block)")
        act_cp_md.triggered.connect(lambda: self._copy_formatted("markdown"))
        act_cp_plain = copy_menu.addAction("📄 Copy Plain Text")
        act_cp_plain.triggered.connect(lambda: self._copy_formatted("text"))
        copy_menu.addSeparator()
        act_cp_cite = copy_menu.addAction("🔖 Copy Citation Only")
        act_cp_cite.triggered.connect(lambda: self._copy_formatted("citation"))
        self.btn_copy.setMenu(copy_menu)
        self.btn_copy.setEnabled(False)
        tier2_bar.addWidget(self.btn_copy)

        self.btn_extract = QPushButton("✂️ Extract ▾")
        self.btn_extract.setToolTip("Extract this clause to a clean Word (.docx) document or export to Markdown (.md)")
        self.btn_extract.setCursor(Qt.PointingHandCursor)
        self.btn_extract.setStyleSheet(INSPECTOR_BTN_SUCCESS)
        extract_menu = QMenu(self)
        act_ext_word = extract_menu.addAction("✂️ Extract Clause in Word (.docx)")
        act_ext_word.triggered.connect(lambda: self._on_extract_clicked(save_as=False))
        act_ext_save = extract_menu.addAction("💾 Save Clause as Word (.docx)...")
        act_ext_save.triggered.connect(lambda: self._on_extract_clicked(save_as=True))
        extract_menu.addSeparator()
        act_ext_md = extract_menu.addAction("📘 Export Clause to Markdown (.md)...")
        act_ext_md.triggered.connect(self._on_export_markdown_clicked)
        self.btn_extract.setMenu(extract_menu)
        self.btn_extract.setEnabled(False)
        tier2_bar.addWidget(self.btn_extract)

        # Separator line
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet("color: #E2E8F0;")
        tier2_bar.addWidget(sep1)

        # Group 2: Full Document & Specs Catalogue
        self.btn_open_doc = QPushButton("📖 Full Doc ▾")
        self.btn_open_doc.setToolTip("Open complete specification document in Microsoft Word or PDF")
        self.btn_open_doc.setCursor(Qt.PointingHandCursor)
        self.btn_open_doc.setStyleSheet(INSPECTOR_BTN_BASE)
        doc_menu = QMenu(self)
        act_word = doc_menu.addAction("📝 Open Full Specification in Word (.docx / .doc)")
        act_word.triggered.connect(lambda: self._on_open_doc_clicked("word"))
        act_pdf = doc_menu.addAction("📕 Open Converted PDF (.pdf)")
        act_pdf.triggered.connect(lambda: self._on_open_doc_clicked("pdf"))
        doc_menu.addSeparator()
        act_folder = doc_menu.addAction("📂 Open Containing Folder in Explorer")
        act_folder.triggered.connect(lambda: self._on_open_doc_clicked("folder"))
        self.btn_open_doc.setMenu(doc_menu)
        self.btn_open_doc.setEnabled(False)
        tier2_bar.addWidget(self.btn_open_doc)

        self.btn_show_in_specs = QPushButton("🔍 Specs Tab")
        self.btn_show_in_specs.setToolTip("Switch to Specifications Tab with this spec and version selected")
        self.btn_show_in_specs.setCursor(Qt.PointingHandCursor)
        self.btn_show_in_specs.setStyleSheet(INSPECTOR_BTN_INFO)
        self.btn_show_in_specs.clicked.connect(self._on_show_in_specs_clicked)
        self.btn_show_in_specs.setEnabled(False)
        tier2_bar.addWidget(self.btn_show_in_specs)

        # Separator line
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet("color: #E2E8F0;")
        tier2_bar.addWidget(sep2)

        # Group 3: Comparison & LLM Analysis
        self.btn_quick_diff = QPushButton("🔍 Diff Prev")
        self.btn_quick_diff.setToolTip("Compare this clause with the immediately preceding release")
        self.btn_quick_diff.setCursor(Qt.PointingHandCursor)
        self.btn_quick_diff.setStyleSheet(INSPECTOR_BTN_WARNING)
        self.btn_quick_diff.clicked.connect(self._on_quick_diff_clicked)
        self.btn_quick_diff.setEnabled(False)
        tier2_bar.addWidget(self.btn_quick_diff)

        self.btn_llm_diff = QPushButton("🤖 LLM Diff")
        self.btn_llm_diff.setToolTip("Compare this clause with another version and generate an LLM analysis prompt with hierarchical context.")
        self.btn_llm_diff.setCursor(Qt.PointingHandCursor)
        self.btn_llm_diff.setStyleSheet(INSPECTOR_BTN_AI)
        self.btn_llm_diff.clicked.connect(self._on_llm_diff_clicked)
        self.btn_llm_diff.setEnabled(False)
        tier2_bar.addWidget(self.btn_llm_diff)

        self.btn_version_changes = QPushButton("📜 Version Changes")
        self.btn_version_changes.setToolTip("View all integrated CRs, p-CRs, and TDocs for this version, and directly open the source contribution documents.")
        self.btn_version_changes.setCursor(Qt.PointingHandCursor)
        self.btn_version_changes.setStyleSheet(INSPECTOR_BTN_BASE)
        self.btn_version_changes.clicked.connect(self._on_version_changes_clicked)
        self.btn_version_changes.setEnabled(False)
        tier2_bar.addWidget(self.btn_version_changes)

        self.btn_origin_tdoc = QPushButton("🎯 Origin TDoc")
        self.btn_origin_tdoc.setToolTip("Identify which specific CR or p-CR introduced or modified this clause across releases.")
        self.btn_origin_tdoc.setCursor(Qt.PointingHandCursor)
        self.btn_origin_tdoc.setStyleSheet(INSPECTOR_BTN_SUCCESS)
        self.btn_origin_tdoc.clicked.connect(self._on_origin_tdoc_clicked)
        self.btn_origin_tdoc.setEnabled(False)
        tier2_bar.addWidget(self.btn_origin_tdoc)

        tier2_bar.addStretch()
        layout.addLayout(tier2_bar)

        # Document Browser
        self.browser = QTextBrowser()
        self.browser.setReadOnly(True)
        self.browser.setOpenExternalLinks(True)
        self.browser.setPlaceholderText("Select a cell in the evolution matrix above to inspect the full clause content and match context...")
        self.browser.setStyleSheet("""
            QTextBrowser {
                border: 1px solid #E2E8F0;
                border-radius: 4px;
                background-color: #FFFFFF;
                padding: 6px;
            }
        """)
        layout.addWidget(self.browser)

    def set_side_by_side(self, side_by_side: bool):
        """Sets the side-by-side orientation state externally."""
        self._is_side_by_side = side_by_side
        self._update_layout_toggle_btn()

    def _on_layout_toggle_clicked(self):
        self._is_side_by_side = not self._is_side_by_side
        self._update_layout_toggle_btn()
        self.layout_orientation_toggle_requested.emit(self._is_side_by_side)

    def _update_layout_toggle_btn(self):
        if self._is_side_by_side:
            self.btn_layout_toggle.setText("⬒ Bottom")
            self.btn_layout_toggle.setToolTip("Switch inspector back to bottom split layout")
        else:
            self.btn_layout_toggle.setText("◫ Right Pane")
            self.btn_layout_toggle.setToolTip("Switch inspector to right-pane side-by-side layout")

    def _on_maximize_clicked(self):
        self._is_maximized = not self._is_maximized
        self._update_maximize_btn()
        self.maximize_toggle_requested.emit(self._is_maximized)

    def _update_maximize_btn(self):
        if self._is_maximized:
            self.btn_maximize.setText("⤡")
            self.btn_maximize.setToolTip("Restore normal inspector height")
        else:
            self.btn_maximize.setText("⤢")
            self.btn_maximize.setToolTip("Maximize reading pane to full height")

    def display_clause(
        self,
        clause_number: str,
        clause_title: str,
        spec_number: str,
        version: str,
        content: str,
        release_date: Optional[str] = None,
        search_query: str = "",
        first_added_ver: Optional[str] = None,
        prev_ver: Optional[str] = None,
        spec_title: Optional[str] = None,
    ):
        self._current_spec = spec_number
        self._current_version = version
        self._current_clause_num = clause_number
        self._current_clause_title = clause_title
        self._current_content = content
        self._current_release_date = release_date or ""
        self._current_search_query = search_query.strip()
        self._first_added_ver = first_added_ver or ""
        self._prev_ver = prev_ver or ""
        self._current_spec_title = spec_title or ""

        date_str = f" ({release_date})" if release_date else ""
        if spec_title:
            header_title = f"TS {spec_number} ({spec_title}) v{version}{date_str} — Clause {clause_number}: {clause_title}"
            self.title_lbl.setToolTip(f"TS {spec_number}: {spec_title}")
        else:
            header_title = f"TS {spec_number} v{version}{date_str} — Clause {clause_number}: {clause_title}"
            self.title_lbl.setToolTip("")
        self.title_lbl.setText(header_title)

        self._last_citation_text = f"3GPP TS {spec_number} v{version}{date_str}, Clause {clause_number} ({clause_title}):\n\n{content}"
        self.btn_copy.setEnabled(True)
        self.btn_extract.setEnabled(True)
        self.btn_llm_diff.setEnabled(True)
        self.btn_open_doc.setEnabled(True)
        self.btn_show_in_specs.setEnabled(True)
        self.btn_version_changes.setEnabled(True)
        self.btn_version_changes.setText(f"📜 v{version} Changes")
        self.btn_version_changes.setToolTip(f"View all CRs, p-CRs, and TDocs integrated in TS {spec_number} v{version}, and directly open the actual TDoc documents.")
        self.btn_origin_tdoc.setEnabled(True)
        self.btn_origin_tdoc.setToolTip(f"Identify which CR or p-CR added or modified Clause {clause_number} in v{version}")

        # Update Quick Diff button
        if self._prev_ver:
            self.btn_quick_diff.setEnabled(True)
            self.btn_quick_diff.setToolTip(f"Compare clause {clause_number} between preceding v{self._prev_ver} and current v{version}")
        else:
            self.btn_quick_diff.setEnabled(False)
            self.btn_quick_diff.setToolTip("No preceding release indexed to compare against")

        # Update First Introduced chip
        if self._first_added_ver:
            clean_first = self._first_added_ver.lstrip("v")
            clean_curr = version.lstrip("v")
            if clean_first == clean_curr:
                self.first_added_badge.setText(f"🟢 First Added in v{clean_first}")
            else:
                self.first_added_badge.setText(f"🟢 Introduced in v{clean_first} (Click to View)")
            self.first_added_badge.setVisible(True)
        else:
            self.first_added_badge.setVisible(False)

        match_count = 0
        if self._current_search_query:
            match_count = len(re.findall(re.escape(self._current_search_query), content, re.IGNORECASE))

        if match_count > 0:
            self.match_badge.setText(f"🎯 {match_count} Match(es)")
            self.match_badge.setVisible(True)
            self.btn_prev.setEnabled(True)
            self.btn_next.setEnabled(True)
        else:
            self.match_badge.setVisible(False)
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)

        excerpt_html = ""
        if self._current_search_query and match_count > 0:
            excerpt_html = self._generate_match_excerpt(content, self._current_search_query, clause_title)

        # Deduplicate redundant first line if it duplicates the clause title banner
        cleaned_body = content.strip()
        lines = cleaned_body.splitlines()
        if lines:
            first_line = lines[0].strip()
            norm_header = f"{clause_number} {clause_title}".strip().lower()
            if (
                first_line.lower() == norm_header
                or first_line.lower() == clause_title.lower()
                or first_line.lower() == clause_number.lower()
            ):
                cleaned_body = "\n".join(lines[1:]).strip()

        escaped_content = html.escape(cleaned_body)
        if self._current_search_query:
            pattern = re.compile(re.escape(html.escape(self._current_search_query)), re.IGNORECASE)
            escaped_content = pattern.sub(
                lambda m: f'<mark style="background-color: #FEF08A; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: #0F172A;">{m.group(0)}</mark>',
                escaped_content,
            )

        formatted_html = f"""
        <html>
        <body style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; line-height: 1.6; color: #1E293B; margin: 4px;">
            {excerpt_html}
            <div style="background-color: #F1F5F9; border-left: 4px solid #0284C7; border-radius: 4px; padding: 6px 12px; margin: 8px 0 10px 0;">
                <span style="color: #0369A1; font-weight: bold; font-size: 13px;">Clause {clause_number}: {html.escape(clause_title)}</span>
            </div>
            <pre style="white-space: pre-wrap; font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; color: #334155; line-height: 1.6;">{escaped_content}</pre>
        </body>
        </html>
        """
        self.browser.setHtml(formatted_html)

        if self._current_search_query and match_count > 0:
            self._find_next()

    def _on_first_added_clicked(self):
        if self._current_spec and self._first_added_ver:
            self.jump_to_version_requested.emit(
                self._current_spec,
                self._first_added_ver,
                self._current_clause_num,
            )

    def _on_open_doc_clicked(self, fmt="word"):
        if self._current_spec and self._current_version:
            self.open_document_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                fmt,
            )

    def _on_show_in_specs_clicked(self):
        if self._current_spec:
            self.show_in_specs_tab_requested.emit(
                self._current_spec,
                self._current_version,
            )

    def _on_quick_diff_clicked(self):
        if self._current_spec and self._prev_ver and self._current_version:
            self.quick_diff_requested.emit(
                self._current_spec,
                self._prev_ver,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
            )

    def _on_llm_diff_clicked(self):
        if self._current_spec and self._current_clause_num:
            self.compare_clause_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
            )

    def _generate_match_excerpt(self, full_text: str, query: str, clause_title: str = "") -> str:
        paragraphs = [p.strip() for p in full_text.split("\n") if p.strip()]
        matched_paras = [p for p in paragraphs if query.lower() in p.lower()]

        # Filter out paragraphs that simply echo the clause title or clause number + title
        body_matches = [
            p for p in matched_paras
            if p.lower() != clause_title.lower() and not p.lower().endswith(clause_title.lower())
        ]

        if body_matches:
            target_excerpt = body_matches[0]
            escaped_excerpt = html.escape(target_excerpt)
            pattern = re.compile(re.escape(html.escape(query)), re.IGNORECASE)
            highlighted_excerpt = pattern.sub(
                lambda m: f'<mark style="background-color: #FEF08A; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: #0F172A;">{m.group(0)}</mark>',
                escaped_excerpt,
            )
            return f"""
            <div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-left: 4px solid #0284C7; border-radius: 6px; padding: 8px 12px; margin-bottom: 10px;">
                <div style="font-weight: 600; color: #0369A1; font-size: 11px; margin-bottom: 3px;">💡 Key Match Excerpt & Context:</div>
                <div style="color: #1E293B; font-size: 12px; line-height: 1.55;">{highlighted_excerpt}</div>
            </div>
            """
        elif clause_title and query.lower() in clause_title.lower():
            escaped_title = html.escape(clause_title)
            pattern = re.compile(re.escape(html.escape(query)), re.IGNORECASE)
            highlighted_title = pattern.sub(
                lambda m: f'<mark style="background-color: #FEF08A; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: #0F172A;">{m.group(0)}</mark>',
                escaped_title,
            )
            return f"""
            <div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-left: 4px solid #0284C7; border-radius: 6px; padding: 8px 12px; margin-bottom: 10px;">
                <div style="font-weight: 600; color: #0369A1; font-size: 11px; margin-bottom: 3px;">💡 Match Location:</div>
                <div style="color: #1E293B; font-size: 12px; line-height: 1.55;">Query matches the <b>Clause Title</b>: {highlighted_title}</div>
            </div>
            """
        elif matched_paras:
            escaped_excerpt = html.escape(matched_paras[0])
            pattern = re.compile(re.escape(html.escape(query)), re.IGNORECASE)
            highlighted_excerpt = pattern.sub(
                lambda m: f'<mark style="background-color: #FEF08A; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: #0F172A;">{m.group(0)}</mark>',
                escaped_excerpt,
            )
            return f"""
            <div style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-left: 4px solid #0284C7; border-radius: 6px; padding: 8px 12px; margin-bottom: 10px;">
                <div style="font-weight: 600; color: #0369A1; font-size: 11px; margin-bottom: 3px;">💡 Key Match Excerpt & Context:</div>
                <div style="color: #1E293B; font-size: 12px; line-height: 1.55;">{highlighted_excerpt}</div>
            </div>
            """
        return ""

    def _find_next(self):
        if not self._current_search_query:
            return
        found = self.browser.find(self._current_search_query)
        if not found:
            self.browser.moveCursor(QTextCursor.Start)
            self.browser.find(self._current_search_query)

    def _find_prev(self):
        if not self._current_search_query:
            return
        found = self.browser.find(self._current_search_query, QTextDocument.FindBackward)
        if not found:
            self.browser.moveCursor(QTextCursor.End)
            self.browser.find(self._current_search_query, QTextDocument.FindBackward)

    def _copy_formatted(self, mode: str = "email"):
        if not self._current_spec or not self._current_clause_num or not self._current_content:
            return
        status = copy_clause_to_clipboard(
            self._current_spec,
            self._current_version,
            self._current_clause_num,
            self._current_clause_title,
            self._current_content,
            self._current_release_date,
            mode=mode,
        )
        self.btn_copy.setText("✅ Copied!")
        QTimer.singleShot(1500, lambda: self.btn_copy.setText("📋 Copy ▾"))

    def _on_extract_clicked(self, save_as: bool = False):
        if self._current_spec and self._current_version and self._current_clause_num:
            self.extract_clause_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
                self._current_content,
                save_as,
            )

    def _on_export_markdown_clicked(self):
        if self._current_spec and self._current_version and self._current_clause_num:
            self.export_markdown_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
                self._current_content,
            )

    def _copy_citation(self):
        self._copy_formatted("citation")

    def _on_version_changes_clicked(self):
        if self._current_spec and self._current_version:
            self.version_changes_requested.emit(self._current_spec, self._current_version)

    def _on_origin_tdoc_clicked(self):
        if self._current_spec and self._current_version and self._current_clause_num:
            self.identify_origin_tdoc_requested.emit(
                self._current_spec,
                self._prev_ver or "",
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
            )

    def clear_display(self):
        self.title_lbl.setText("No clause selected")
        self.match_badge.setVisible(False)
        self.first_added_badge.setVisible(False)
        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)
        self.btn_copy.setEnabled(False)
        self.btn_extract.setEnabled(False)
        self.btn_open_doc.setEnabled(False)
        self.btn_show_in_specs.setEnabled(False)
        self.btn_quick_diff.setEnabled(False)
        self.btn_llm_diff.setEnabled(False)
        self.btn_version_changes.setEnabled(False)
        self.btn_version_changes.setText("📜 Version Changes")
        self.btn_origin_tdoc.setEnabled(False)
        self.browser.clear()
        self._current_search_query = ""
        self._last_citation_text = ""
        self._current_spec = ""
        self._current_version = ""
        self._current_clause_num = ""
        self._current_clause_title = ""
        self._current_content = ""
        self._current_release_date = ""
        self._first_added_ver = ""
        self._prev_ver = ""
