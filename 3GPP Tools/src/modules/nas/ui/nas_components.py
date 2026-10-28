import html
import re
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import QPoint, QRect, QRectF, QSize, Qt, QUrl, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QFont, QFontMetrics, QPainter, QPainterPath, QPen
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QStackedWidget,
    QStyle,
    QStyledItemDelegate,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import get_wg_colors
from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
from modules.nas.ui.nas_dialogs import get_spec_metadata, get_spec_title, get_spec_tooltip


class NASMessageItemDelegate(QStyledItemDelegate):
    """
    Renders protocol messages and information elements with high visual polish:
    - Left: Message / IE Name in bold typeface
    - Right: Color-coded Working Group / Specification chip and Clause pill
    """

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)

    def sizeHint(self, option, index) -> QSize:
        return QSize(option.rect.width(), 32)

    def paint(self, painter: QPainter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)

        rect = option.rect
        is_selected = bool(option.state & QStyle.State_Selected)
        is_hovered = bool(option.state & QStyle.State_MouseOver)

        # 1. Background rendering
        if is_selected:
            bg_brush = QBrush(QColor("#EBF3FC"))
            border_pen = QPen(QColor("#1E5C99"), 1.2)
        elif is_hovered:
            bg_brush = QBrush(QColor("#F8FAFC"))
            border_pen = QPen(QColor("#E2E8F0"), 1)
        else:
            bg_brush = QBrush(QColor("#FFFFFF"))
            border_pen = QPen(QColor("#F1F5F9"), 0.8)

        bg_path = QPainterPath()
        bg_rect = QRectF(rect.x() + 2, rect.y() + 1, rect.width() - 4, rect.height() - 2)
        bg_path.addRoundedRect(bg_rect, 4.0, 4.0)
        painter.fillPath(bg_path, bg_brush)
        painter.setPen(border_pen)
        painter.drawPath(bg_path)

        # Active selection accent marker on left border
        if is_selected:
            accent_path = QPainterPath()
            accent_rect = QRectF(rect.x() + 2, rect.y() + 2, 3.5, rect.height() - 4)
            accent_path.addRoundedRect(accent_rect, 1.5, 1.5)
            painter.fillPath(accent_path, QBrush(QColor("#1E5C99")))

        # 2. Extract item metadata
        msg_name = str(index.data(Qt.DisplayRole) or "")
        spec_num = str(index.data(Qt.UserRole + 2) or "").replace("TS", "").strip()
        clause = str(index.data(Qt.UserRole + 3) or "").strip()

        meta = get_spec_metadata(spec_num) if spec_num else {}
        wg = meta.get("wg", "")
        acronym = meta.get("acronym", "")

        right_cursor = rect.right() - 8

        # 3. Draw Right-aligned Clause Pill
        if clause and clause != msg_name:
            clause_font = QFont(option.font)
            clause_font.setPointSize(9)
            clause_font.setBold(True)
            clause_fm = QFontMetrics(clause_font)
            clause_text = f"Cl. {clause}"
            pill_w = clause_fm.horizontalAdvance(clause_text) + 12
            pill_h = 18
            pill_y = rect.y() + (rect.height() - pill_h) // 2
            pill_x = right_cursor - pill_w

            pill_path = QPainterPath()
            pill_path.addRoundedRect(QRectF(pill_x, pill_y, pill_w, pill_h), 9.0, 9.0)
            painter.fillPath(pill_path, QBrush(QColor("#F1F5F9")))
            painter.setPen(QPen(QColor("#CBD5E1"), 1))
            painter.drawPath(pill_path)

            painter.setFont(clause_font)
            painter.setPen(QColor("#475569"))
            painter.drawText(
                QRect(pill_x, pill_y, pill_w, pill_h),
                Qt.AlignCenter,
                clause_text,
            )
            right_cursor = pill_x - 6

        # 4. Draw Right-aligned Specification / WG Chip
        if spec_num:
            wg_colors = get_wg_colors(wg) if wg else {"bg": "#F8FAFC", "text": "#334155", "border": "#CBD5E1"}
            spec_font = QFont(option.font)
            spec_font.setPointSize(9)
            spec_font.setBold(True)
            spec_fm = QFontMetrics(spec_font)

            spec_label = f"{wg} · {spec_num}" if wg and wg != "3GPP" else f"TS {spec_num}"
            chip_w = spec_fm.horizontalAdvance(spec_label) + 10
            chip_h = 18
            chip_y = rect.y() + (rect.height() - chip_h) // 2
            chip_x = right_cursor - chip_w

            if chip_x > rect.x() + 100:  # Only draw if enough horizontal room
                chip_path = QPainterPath()
                chip_path.addRoundedRect(QRectF(chip_x, chip_y, chip_w, chip_h), 3.0, 3.0)
                painter.fillPath(chip_path, QBrush(QColor(wg_colors.get("bg", "#F8FAFC"))))
                painter.setPen(QPen(QColor(wg_colors.get("border", "#CBD5E1")), 1))
                painter.drawPath(chip_path)

                painter.setFont(spec_font)
                painter.setPen(QColor(wg_colors.get("text", "#334155")))
                painter.drawText(
                    QRect(chip_x, chip_y, chip_w, chip_h),
                    Qt.AlignCenter,
                    spec_label,
                )
                right_cursor = chip_x - 6

        # 5. Draw Left-aligned Message Name
        name_font = QFont(option.font)
        name_font.setPointSize(10)
        name_font.setBold(is_selected)
        painter.setFont(name_font)
        painter.setPen(QColor("#0F172A") if not is_selected else QColor("#1E5C99"))

        left_x = rect.x() + (10 if not is_selected else 12)
        avail_w = max(40, right_cursor - left_x)
        text_rect = QRect(left_x, rect.y(), avail_w, rect.height())

        elided_text = QFontMetrics(name_font).elidedText(msg_name, Qt.ElideRight, avail_w)
        painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, elided_text)

        painter.restore()


class NASVersionTreeWidget(QTreeWidget):
    """Hierarchical specification and version management tree with tri-state selection."""

    selection_changed = pyqtSignal()
    delete_version_requested = pyqtSignal(str, str)
    delete_spec_requested = pyqtSignal(str, int)
    structure_changed = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._updating_checks: bool = False

        self.setHeaderHidden(True)
        self.setAlternatingRowColors(True)
        self.setIndentation(14)
        self.setStyleSheet("""
            QTreeWidget {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #FFFFFF;
                padding: 2px;
                outline: none;
            }
            QTreeWidget::item {
                padding: 3px 2px;
                min-height: 22px;
                border-radius: 3px;
            }
            QTreeWidget::item:hover {
                background-color: #F8FAFC;
            }
            QTreeWidget::item:selected {
                background-color: #EBF3FC;
                color: #0F172A;
            }
        """)

        self.itemChanged.connect(self._on_item_changed)
        self.itemExpanded.connect(lambda _: self.structure_changed.emit())
        self.itemCollapsed.connect(lambda _: self.structure_changed.emit())
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._on_context_menu)

    def select_all(self):
        """Checks all specifications and versions."""
        all_item = self.topLevelItem(0)
        if all_item:
            all_item.setCheckState(0, Qt.Checked)

    def select_none(self):
        """Unchecks all specifications and versions."""
        all_item = self.topLevelItem(0)
        if all_item:
            all_item.setCheckState(0, Qt.Unchecked)

    def select_latest_only(self):
        """Selects only the highest (latest) release version for each indexed specification."""
        self._updating_checks = True
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            child_count = spec_item.childCount()
            if child_count > 0:
                spec_item.child(0).setCheckState(0, Qt.Checked)
                for c_idx in range(1, child_count):
                    spec_item.child(c_idx).setCheckState(0, Qt.Unchecked)
        self._update_parent_states()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def populate(
            self,
            versions: List[Dict[str, Any]],
            saved_checked: Optional[List[Dict[str, Any]]] = None,
            saved_collapsed: Optional[Set[str]] = None,
    ):
        self._updating_checks = True
        self.clear()

        if not versions:
            self._updating_checks = False
            self.selection_changed.emit()
            return

        saved_tuples: Optional[Set[Tuple[str, str]]] = None
        if saved_checked is not None:
            saved_tuples = {(item.get("spec_number"), item.get("version")) for item in saved_checked}

        if saved_collapsed is None:
            saved_collapsed = set()

        base_font = self.font()
        bold_font = QFont(base_font)
        bold_font.setBold(True)

        text_color_header = QBrush(QColor("#0F172A"))
        text_color_item = QBrush(QColor("#334155"))

        # Master toggle item
        all_item = QTreeWidgetItem(self)
        all_item.setText(0, "All Specifications")
        all_item.setData(0, Qt.UserRole, {"type": "all"})
        all_item.setFlags(all_item.flags() | Qt.ItemIsUserCheckable)
        all_item.setFont(0, bold_font)
        all_item.setForeground(0, text_color_header)
        all_item.setCheckState(0, Qt.Checked)

        specs_map: Dict[str, List[Dict[str, Any]]] = {}
        for v in versions:
            specs_map.setdefault(v["spec_number"], []).append(v)

        for spec_num in sorted(specs_map.keys()):
            spec_versions = specs_map[spec_num]
            meta = get_spec_metadata(spec_num)

            spec_item = QTreeWidgetItem(self)
            display_title = f"[{meta['wg']}] TS {spec_num} — {meta['acronym']}"
            spec_item.setText(0, display_title)
            spec_item.setToolTip(0, get_spec_tooltip(spec_num))
            spec_item.setData(0, Qt.UserRole, {"type": "spec", "spec_number": spec_num})
            spec_item.setFlags(spec_item.flags() | Qt.ItemIsUserCheckable)
            spec_item.setFont(0, bold_font)
            spec_item.setForeground(0, text_color_header)
            spec_item.setCheckState(0, Qt.Checked)

            sorted_v_list = sorted(spec_versions, key=lambda x: parse_version_tuple(x["version"]), reverse=True)

            for v in sorted_v_list:
                clean_v = v["version"].lstrip("vV")
                parts = clean_v.split(".")
                major_rel = f"Rel-{parts[0]} · " if parts and parts[0].isdigit() else ""

                child = QTreeWidgetItem(spec_item)
                child.setText(0, f"{major_rel}v{v['version']}")
                child.setToolTip(
                    0,
                    f"TS {v['spec_number']} v{v['version']} ({meta['title']})\n"
                    f"3GPP Release: {major_rel.rstrip(' ·') if major_rel else 'Unknown'}\n"
                    f"(Right-click to delete)"
                )
                child.setData(
                    0,
                    Qt.UserRole,
                    {
                        "type": "version",
                        "id": v["id"],
                        "spec_number": v["spec_number"],
                        "version": v["version"],
                    },
                )
                child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                child.setFont(0, base_font)
                child.setForeground(0, text_color_item)

                if saved_tuples is not None:
                    is_checked = (v["spec_number"], v["version"]) in saved_tuples
                    child.setCheckState(0, Qt.Checked if is_checked else Qt.Unchecked)
                else:
                    child.setCheckState(0, Qt.Checked)

            spec_item.setExpanded(spec_num not in saved_collapsed)

        self._update_parent_states()
        self._updating_checks = False
        self.selection_changed.emit()

    def get_selected_version_ids(self) -> List[int]:
        selected_ids = []
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            for c_idx in range(spec_item.childCount()):
                child = spec_item.child(c_idx)
                if child.checkState(0) == Qt.Checked:
                    data = child.data(0, Qt.UserRole)
                    if data and "id" in data:
                        selected_ids.append(data["id"])
        return selected_ids

    def get_checked_versions_info(self) -> List[Dict[str, Any]]:
        checked = []
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            for c_idx in range(spec_item.childCount()):
                child = spec_item.child(c_idx)
                if child.checkState(0) == Qt.Checked:
                    c_data = child.data(0, Qt.UserRole) or {}
                    if c_data.get("type") == "version":
                        checked.append({
                            "spec_number": c_data.get("spec_number"),
                            "version": c_data.get("version"),
                        })
        return checked

    def get_all_versions_for_spec(self, spec_number: str) -> List[Dict[str, Any]]:
        results = []
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            s_data = spec_item.data(0, Qt.UserRole) or {}
            if s_data.get("spec_number") == spec_number:
                for c_idx in range(spec_item.childCount()):
                    child = spec_item.child(c_idx)
                    c_data = child.data(0, Qt.UserRole) or {}
                    if c_data.get("type") == "version":
                        results.append({
                            "spec_number": c_data.get("spec_number"),
                            "version": c_data.get("version"),
                        })
        return results

    def get_collapsed_spec_numbers(self) -> List[str]:
        collapsed = []
        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            s_data = spec_item.data(0, Qt.UserRole) or {}
            spec_num = s_data.get("spec_number", "")
            if not spec_item.isExpanded() and spec_num:
                collapsed.append(spec_num)
        return collapsed

    def _on_item_changed(self, item: QTreeWidgetItem, column: int):
        if self._updating_checks:
            return

        self._updating_checks = True
        state = item.checkState(0)
        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")

        if item_type == "all":
            target_state = Qt.Checked if state == Qt.Checked else Qt.Unchecked
            for s_idx in range(1, self.topLevelItemCount()):
                spec_item = self.topLevelItem(s_idx)
                spec_item.setCheckState(0, target_state)
                for c_idx in range(spec_item.childCount()):
                    child = spec_item.child(c_idx)
                    child.setCheckState(0, target_state)
        elif item_type == "spec":
            target_state = Qt.Checked if state == Qt.Checked else Qt.Unchecked
            for c_idx in range(item.childCount()):
                child = item.child(c_idx)
                child.setCheckState(0, target_state)
            self._update_parent_states()
        elif item_type == "version":
            self._update_parent_states()

        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def _update_parent_states(self):
        total_version_count = 0
        total_checked_count = 0

        for s_idx in range(1, self.topLevelItemCount()):
            spec_item = self.topLevelItem(s_idx)
            child_count = spec_item.childCount()
            if child_count == 0:
                continue

            checked_children = sum(
                1 for c_idx in range(child_count)
                if spec_item.child(c_idx).checkState(0) == Qt.Checked
            )

            total_version_count += child_count
            total_checked_count += checked_children

            if checked_children == child_count:
                spec_item.setCheckState(0, Qt.Checked)
            elif checked_children == 0:
                spec_item.setCheckState(0, Qt.Unchecked)
            else:
                spec_item.setCheckState(0, Qt.PartiallyChecked)

        all_item = self.topLevelItem(0)
        if all_item:
            if total_version_count > 0 and total_checked_count == total_version_count:
                all_item.setCheckState(0, Qt.Checked)
            elif total_checked_count == 0:
                all_item.setCheckState(0, Qt.Unchecked)
            else:
                all_item.setCheckState(0, Qt.PartiallyChecked)

    def _on_context_menu(self, pos):
        item = self.itemAt(pos)
        if not item:
            return

        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                font-size: 11px;
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                padding: 4px;
            }
            QMenu::item {
                padding: 4px 15px;
            }
            QMenu::item:selected {
                background-color: #FEE2E2;
                color: #B91C1C;
            }
        """)

        if item_type == "version":
            spec_num = data.get("spec_number", "")
            version = data.get("version", "")
            act_delete = QAction(f"🗑️ Delete TS {spec_num} v{version}", self)
            act_delete.triggered.connect(lambda: self.delete_version_requested.emit(spec_num, version))
            menu.addAction(act_delete)
            menu.exec_(self.viewport().mapToGlobal(pos))

        elif item_type == "spec":
            spec_num = data.get("spec_number", "")
            child_count = item.childCount()
            act_delete_spec = QAction(f"🗑️ Delete all versions of TS {spec_num} ({child_count} versions)", self)
            act_delete_spec.triggered.connect(lambda: self.delete_spec_requested.emit(spec_num, child_count))
            menu.addAction(act_delete_spec)
            menu.exec_(self.viewport().mapToGlobal(pos))


RE_NAS_CROSS_REF = re.compile(
    r"(?:see\s+)?(?:subclause|clause|sub-clause)\s+([0-9A-Za-z\.]+)\s+in\s+(?:3GPP\s+)?TS\s+(24\.(?:301|501))",
    re.IGNORECASE,
)


class NASInspectorWidget(QFrame):
    """Modern Card Frame rendering Clause 9 diagrams, ASN.1 syntax, and field descriptions."""

    jump_to_message_requested = pyqtSignal(str)
    filter_by_ie_requested = pyqtSignal(str)
    import_spec_requested = pyqtSignal(str, int)  # (spec_number, major_version)

    def __init__(self, db: Optional[NASDatabase] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.db = db
        self._current_clause_defs: Dict[str, Dict[str, Any]] = {}
        self._current_ie_clause: Optional[str] = None
        self._current_ie_name: Optional[str] = None
        self._current_ie_spec: Optional[str] = None
        self._current_containing_msgs: List[Dict[str, Any]] = []
        self._updating_combo: bool = False

        self.setObjectName("cardFrame")
        self.setStyleSheet("""
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
            }
        """)

        self._setup_ui()

    def set_database(self, db: NASDatabase):
        self.db = db

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 1. Header Toolbar
        header_frame = QFrame()
        header_frame.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border-bottom: 1px solid #E2E8F0;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 4px 8px;
            }
        """)
        insp_header = QHBoxLayout(header_frame)
        insp_header.setContentsMargins(6, 4, 6, 4)
        insp_header.setSpacing(8)

        card_title = QLabel("🔍 Field & Structure Inspector:")
        card_title.setStyleSheet("font-weight: bold; color: #1E293B; font-size: 11px;")
        insp_header.addWidget(card_title)

        self.title_lbl = QLabel("No Selection")
        self.title_lbl.setStyleSheet("""
            font-weight: bold;
            color: #0284C7;
            font-size: 11px;
            background-color: #E0F2FE;
            padding: 2px 8px;
            border-radius: 4px;
            border: 1px solid #BAE6FD;
        """)
        insp_header.addWidget(self.title_lbl)

        insp_header.addStretch()

        self.usage_btn = QPushButton("Used in: 0 msgs ▾")
        self.usage_btn.setVisible(False)
        self.usage_btn.setCursor(Qt.PointingHandCursor)
        self.usage_btn.setToolTip("View other messages that contain this Information Element")
        self.usage_btn.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: bold;
                color: #0369A1;
                background-color: #E0F2FE;
                border: 1px solid #BAE6FD;
                border-radius: 4px;
                padding: 2px 8px;
            }
            QPushButton:hover {
                background-color: #BAE6FD;
                border-color: #0284C7;
            }
        """)
        self.usage_btn.clicked.connect(self._show_usage_menu)
        insp_header.addWidget(self.usage_btn)

        self.version_combo = QComboBox()
        self.version_combo.setToolTip("Switch specification release for this definition")
        self.version_combo.setStyleSheet("""
            QComboBox {
                font-weight: bold;
                font-size: 11px;
                padding: 2px 8px;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #FFFFFF;
                min-width: 120px;
            }
            QComboBox:hover {
                border-color: #0284C7;
            }
        """)
        self.version_combo.currentIndexChanged.connect(self._on_version_changed)
        insp_header.addWidget(self.version_combo)

        self.copy_btn = QPushButton("📋 Copy Definition")
        self.copy_btn.setCursor(Qt.PointingHandCursor)
        self.copy_btn.setToolTip("Copy raw definition text or ASN.1 syntax to clipboard")
        self.copy_btn.setEnabled(False)
        self.copy_btn.setStyleSheet("""
            QPushButton {
                font-size: 11px;
                font-weight: 500;
                color: #334155;
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 2px 8px;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                color: #0F172A;
                border-color: #94A3B8;
            }
            QPushButton:disabled {
                color: #94A3B8;
                background-color: #F8FAFC;
                border-color: #E2E8F0;
            }
        """)
        self.copy_btn.clicked.connect(self._copy_current_text)
        insp_header.addWidget(self.copy_btn)

        layout.addWidget(header_frame)

        # 2. Content Stack: Empty State (0) vs Text Browser (1)
        self.stack = QStackedWidget()

        # Empty State Placeholder
        self.empty_widget = QWidget()
        empty_layout = QVBoxLayout(self.empty_widget)
        empty_layout.setAlignment(Qt.AlignCenter)
        empty_layout.setSpacing(6)

        empty_icon = QLabel("🔍")
        empty_icon.setStyleSheet("font-size: 26px; color: #94A3B8;")
        empty_icon.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_icon)

        empty_title = QLabel("Information Element & Field Inspector")
        empty_title.setStyleSheet("font-weight: bold; font-size: 13px; color: #334155;")
        empty_title.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_title)

        empty_sub = QLabel(
            "Click on any Information Element or field in the Evolution Matrix above\n"
            "to inspect its bit-level octet layout, ASN.1 syntax, and field descriptions."
        )
        empty_sub.setStyleSheet("font-size: 11px; color: #64748B; line-height: 1.4;")
        empty_sub.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(empty_sub)

        self.stack.addWidget(self.empty_widget)

        # Text Browser View
        self.text_area = QTextBrowser()
        self.text_area.setReadOnly(True)
        self.text_area.setOpenLinks(False)
        self.text_area.setOpenExternalLinks(False)
        self.text_area.anchorClicked.connect(self._on_anchor_clicked)
        self.text_area.setStyleSheet("""
            QTextBrowser {
                border: none;
                padding: 8px;
                background-color: #FFFFFF;
            }
        """)
        self.stack.addWidget(self.text_area)

        layout.addWidget(self.stack)

    def _copy_current_text(self):
        text = self.text_area.toPlainText().strip()
        if text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            prev_label = self.copy_btn.text()
            self.copy_btn.setText("✓ Copied!")
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(1500, lambda: self.copy_btn.setText(prev_label))

    def _on_anchor_clicked(self, url: QUrl):
        url_str = url.toString()
        if url_str.startswith("action:import:"):
            parts = url_str.replace("action:import:", "").split(":")
            spec_num = parts[0]
            major_ver = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
            self.import_spec_requested.emit(spec_num, major_ver)

    def display_definitions(
            self,
            clause: str,
            ie_name: str,
            spec_number: Optional[str],
            defs: List[Dict[str, Any]],
            containing_msgs: List[Dict[str, Any]],
            target_version_hint: Optional[str] = None,
            fallback_type_ref: str = "",
            field_description: str = "",
            **kwargs,
    ):
        self._current_ie_clause = clause
        self._current_ie_name = ie_name
        self._current_ie_spec = spec_number
        self._current_containing_msgs = containing_msgs

        meta = get_spec_metadata(spec_number) if spec_number else {}
        acronym_str = f" [{meta.get('acronym', f'TS {spec_number}')}]" if spec_number else ""

        if defs:
            resolved_name = defs[0]["ie_name"]
            self.title_lbl.setText(f"{resolved_name}{acronym_str}")
            self._current_clause_defs = {f"{d['spec_number']} v{d['version']}": d for d in defs}

            if containing_msgs:
                num_msgs = len(containing_msgs)
                self.usage_btn.setText(f"Used in: {num_msgs} msg{'s' if num_msgs != 1 else ''} ▾")
                self.usage_btn.setEnabled(True)
            else:
                self.usage_btn.setText("Used in: ... ▾")
                self.usage_btn.setEnabled(False)
            self.usage_btn.setVisible(True)

            self._updating_combo = True
            self.version_combo.clear()
            for d in defs:
                clean_v = str(d.get("version", "")).lstrip("vV")
                parts = clean_v.split(".")
                rel_str = f"Rel-{parts[0]} " if parts and parts[0].isdigit() else ""
                key_name = f"TS {d['spec_number']} {rel_str}v{d['version']}"
                self.version_combo.addItem(key_name, f"{d['spec_number']} v{d['version']}")

            target_version_key: Optional[str] = None
            if target_version_hint:
                clean_hint = target_version_hint.replace("TS ", "").strip()
                if clean_hint in self._current_clause_defs:
                    target_version_key = clean_hint
                else:
                    for k in self._current_clause_defs:
                        if target_version_hint in k or k.endswith(f"v{target_version_hint}"):
                            target_version_key = k
                            break

            if not target_version_key and defs:
                target_version_key = f"{defs[0]['spec_number']} v{defs[0]['version']}"

            for idx in range(self.version_combo.count()):
                if self.version_combo.itemData(idx) == target_version_key:
                    self.version_combo.setCurrentIndex(idx)
                    break

            self._updating_combo = False
            selected_def = self._current_clause_defs.get(target_version_key) or defs[0]
            self._render_selected_definition(selected_def)
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(True)
            return

        self.usage_btn.setVisible(False)
        self.version_combo.clear()
        self._current_clause_defs.clear()

        self.title_lbl.setText(f"{ie_name}{acronym_str}")

        is_inline_type = any(
            fallback_type_ref.strip().upper().startswith(kw)
            for kw in ("ENUMERATED", "INTEGER", "BOOLEAN", "BIT STRING", "OCTET STRING", "NULL", "CHOICE", "SEQUENCE")
        )

        if is_inline_type or fallback_type_ref:
            html_output = self._render_inline_type_html(
                ie_name=ie_name,
                type_ref=fallback_type_ref,
                field_description=field_description,
                spec_number=spec_number,
            )
            self.text_area.setHtml(html_output)
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(True)
        else:
            self.text_area.setPlainText(
                f"Type / Reference: {fallback_type_ref}\n(No structure definition found in database)"
            )
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(bool(fallback_type_ref))

    def _render_selected_definition(self, definition: Dict[str, Any]):
        base_html = definition.get("raw_description", "")
        current_spec = definition.get("spec_number", "")
        current_ver_str = definition.get("version", "")

        # Check for 4G <-> 5G cross-specification references
        if self.db and current_spec in ("24.501", "24.301"):
            match = RE_NAS_CROSS_REF.search(base_html)
            if match:
                target_clause = match.group(1).rstrip(".[] ")
                target_spec = match.group(2).strip()

                if target_spec != current_spec:
                    ver_tuple = parse_version_tuple(current_ver_str)
                    major_ver = (ver_tuple[0][1] if isinstance(ver_tuple[0], tuple) else ver_tuple[0]) if ver_tuple else 15

                    cross_def = self.db.get_cross_referenced_ie_definition(
                        target_spec=target_spec,
                        target_clause=target_clause,
                        alt_name=definition.get("ie_name", ""),
                        major_version=major_ver,
                    )

                    if cross_def and cross_def.get("raw_description"):
                        cross_body = cross_def.get("raw_description", "").strip()
                        cross_html = (
                            f'<div style="margin-top: 14px; border: 1.5px solid #0284C7; background-color: #F0F9FF; border-radius: 6px; padding: 12px;">'
                            f'<div style="margin-bottom: 10px; border-bottom: 1px solid #BAE6FD; padding-bottom: 6px;">'
                            f'<span style="font-size: 12px; font-weight: bold; color: #0369A1;">'
                            f'🔗 Cross-Referenced from TS {target_spec} v{cross_def["version"]} (Clause {cross_def.get("clause", target_clause)})'
                            f'</span>'
                            f'<span style="font-size: 10px; color: #0284C7; margin-left: 8px; background-color: #E0F2FE; padding: 2px 6px; border-radius: 3px; border: 1px solid #7DD3FC;">'
                            f'Rel-{major_ver} latest'
                            f'</span>'
                            f'</div>'
                            f'<div style="background-color: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 4px; padding: 10px;">'
                            f'{cross_body}'
                            f'</div>'
                            f'</div>'
                        )
                        base_html += cross_html
                    else:
                        missing_html = (
                            f'<div style="margin-top: 14px; border: 1px solid #FCD34D; background-color: #FFFBEB; border-radius: 6px; padding: 10px; border-left: 4px solid #F59E0B; font-family: Segoe UI, sans-serif;">'
                            f'<div style="font-weight: bold; color: #92400E; font-size: 12px; margin-bottom: 4px;">'
                            f'ℹ️ Cross-Specification Reference: TS {target_spec} (Clause {target_clause})'
                            f'</div>'
                            f'<div style="color: #78350F; font-size: 11px; line-height: 1.5;">'
                            f'This Information Element references <b>TS {target_spec} (Release {major_ver})</b>, which is not currently present in your local database.<br/>'
                            f'<p style="margin-top: 6px; margin-bottom: 2px;">'
                            f'<a href="action:import:{target_spec}:{major_ver}" style="display: inline-block; background-color: #0284C7; color: #FFFFFF; padding: 4px 12px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 11px;">'
                            f'📥 Import TS {target_spec} (Rel-{major_ver}) from 3GPP Archive'
                            f'</a>'
                            f'</p>'
                            f'</div>'
                            f'</div>'
                        )
                        base_html += missing_html

        self.text_area.setHtml(base_html)

    def _on_version_changed(self, index: int):
        if self._updating_combo or index < 0:
            return
        version_key = self.version_combo.itemData(index)
        if version_key in self._current_clause_defs:
            self._render_selected_definition(self._current_clause_defs[version_key])

    @staticmethod
    def _render_inline_type_html(
        ie_name: str,
        type_ref: str,
        field_description: str = "",
        spec_number: Optional[str] = None,
    ) -> str:
        meta = get_spec_metadata(spec_number) if spec_number else {}
        spec_title = f" ({meta.get('acronym', f'TS {spec_number}')})" if spec_number else ""
        header_html = (
            f'<h2 style="color: #0369A1; margin-top: 2px; margin-bottom: 8px; '
            f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
            f'{html.escape(ie_name)}{spec_title}</h2>'
        )

        clean_type_ref = re.sub(r"\s+", " ", type_ref).strip()
        formatted_asn1 = f"{ie_name} ::= {clean_type_ref}"
        if "ENUMERATED" in clean_type_ref:
            enum_match = re.search(r"ENUMERATED\s*\{([^}]+)\}", clean_type_ref, re.IGNORECASE)
            if enum_match:
                inner_vals = [v.strip() for v in enum_match.group(1).split(",") if v.strip()]
                indented_vals = ",\n    ".join(inner_vals)
                formatted_asn1 = f"{ie_name} ::= ENUMERATED {{\n    {indented_vals}\n}}"

        asn1_html = (
            f'<div style="margin-bottom: 10px;">'
            f'<div style="font-size: 11px; font-weight: bold; color: #475569; margin-bottom: 3px;">Inline ASN.1 Definition:</div>'
            f'<pre style="background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 8px; '
            f'border-radius: 4px; font-family: Consolas, monospace; font-size: 11px; color: #0F172A; '
            f'white-space: pre-wrap; margin: 0;">{html.escape(formatted_asn1)}</pre>'
            f'</div>'
        )

        desc_html = ""
        if field_description and field_description.strip():
            desc_html = (
                f'<div style="margin-top: 8px;">'
                f'<div style="font-size: 11px; font-weight: bold; color: #475569; margin-bottom: 4px;">'
                f'Field Description & Semantics:</div>'
                f'<div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-left: 3px solid #0284C7; '
                f'padding: 8px; border-radius: 3px; font-family: Segoe UI, sans-serif; font-size: 11px; color: #334155; line-height: 1.4;">'
                f'{html.escape(field_description)}'
                f'</div>'
                f'</div>'
            )

        return f"{header_html}{asn1_html}{desc_html}"

    def clear_display(self):
        self.title_lbl.setText("No Selection")
        self.usage_btn.setVisible(False)
        self.version_combo.clear()
        self.text_area.clear()
        self._current_clause_defs.clear()
        self.copy_btn.setEnabled(False)
        self.stack.setCurrentIndex(0)

    def _show_usage_menu(self):
        if not self._current_ie_clause:
            return

        clause = self._current_ie_clause
        name = self._current_ie_name or ""
        spec_num = self._current_ie_spec

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                font-size: 11px;
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                padding: 4px;
            }
            QMenu::item {
                padding: 4px 20px 4px 10px;
                border-radius: 3px;
            }
            QMenu::item:selected {
                background-color: #E0F2FE;
                color: #0369A1;
            }
            QMenu::separator {
                height: 1px;
                background-color: #E2E8F0;
                margin: 4px 0;
            }
        """)

        header_title = f"Messages referencing {name or clause}"
        if spec_num:
            meta = get_spec_metadata(spec_num)
            header_title += f" [{meta['acronym']}]"
        header_action = QAction(f"{header_title}:", self)
        header_action.setEnabled(False)
        menu.addAction(header_action)
        menu.addSeparator()

        if self._current_containing_msgs:
            for m in self._current_containing_msgs:
                msg_name = m["message_name"]
                clause_ref = m["clause"]
                m_spec = m.get("spec_number", "")
                spec_tag = f" [{m_spec}]" if m_spec and not spec_num else ""
                action = QAction(f"{msg_name} ({clause_ref}){spec_tag}", self)
                action.triggered.connect(lambda checked, mn=msg_name: self.jump_to_message_requested.emit(mn))
                menu.addAction(action)
        else:
            none_act = QAction("No messages found in active releases", self)
            none_act.setEnabled(False)
            menu.addAction(none_act)

        menu.addSeparator()
        filter_action = QAction(f"🔍 Filter message list by '{name or clause}'", self)
        filter_action.triggered.connect(lambda: self.filter_by_ie_requested.emit(name or clause))
        menu.addAction(filter_action)

        menu.exec_(self.usage_btn.mapToGlobal(self.usage_btn.rect().bottomLeft()))

    def set_containing_messages(self, containing_msgs: List[Dict[str, Any]]):
        self._current_containing_msgs = containing_msgs
        num_msgs = len(containing_msgs)
        self.usage_btn.setText(f"Used in: {num_msgs} msg{'s' if num_msgs != 1 else ''} ▾")
        self.usage_btn.setVisible(True)
        self.usage_btn.setEnabled(True)
