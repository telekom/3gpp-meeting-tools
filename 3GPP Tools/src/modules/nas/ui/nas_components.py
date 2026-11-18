import html
import re
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import QPoint, QRect, QRectF, QSize, Qt, QUrl, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QCursor, QFont, QFontMetrics, QPainter, QPainterPath, QPen
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
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

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scaled_font, scale_px
from core.ui.ui_components import get_wg_colors
from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
from modules.nas.core.nas_exporter import copy_definition_to_clipboard
from modules.nas.ui.nas_dialogs import get_spec_metadata, get_spec_title, get_spec_tooltip
from core.ui.theme import C, T


def _extract_major_release(version: str) -> str:
    """Extract major release label (e.g. 'Rel-20') from a 3GPP version string."""
    clean_v = str(version or "").lstrip("vV").strip()
    first = clean_v.split(".")[0] if clean_v else ""
    if first.isdigit():
        return f"Rel-{first}"
    return "Rel-?"


class NASMessageItemDelegate(QStyledItemDelegate):
    """
    Renders protocol messages and information elements in a clean, scannable 2-line layout:
    - Line 1 (Top): Message / IE Name (regular weight unselected; semibold + accent when selected)
    - Line 2 (Bottom): Compact WG chip + muted 'Spec · Acronym' subtitle, and right-aligned Clause
    """

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)

    def sizeHint(self, option, index) -> QSize:
        return QSize(option.rect.width(), scale_px(42, minimum=42))

    def paint(self, painter: QPainter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)

        rect = option.rect
        is_selected = bool(option.state & QStyle.State_Selected)
        is_hovered = bool(option.state & QStyle.State_MouseOver)

        # 1. Background rendering
        if is_selected:
            bg_brush = QBrush(QColor(C.ACCENT_SOFT))
            border_pen = QPen(QColor(C.ACCENT_RING), 1.0)
        elif is_hovered:
            bg_brush = QBrush(QColor(C.BG_SURFACE))
            border_pen = QPen(QColor(C.BORDER_SUBTLE), 1.0)
        else:
            bg_brush = QBrush(QColor(C.BG_CARD))
            border_pen = None

        bg_rect = QRectF(rect.x() + 2, rect.y() + 1, rect.width() - 4, rect.height() - 2)
        bg_path = QPainterPath()
        bg_path.addRoundedRect(bg_rect, 4.0, 4.0)
        painter.fillPath(bg_path, bg_brush)
        if border_pen is not None:
            painter.setPen(border_pen)
            painter.drawPath(bg_path)
        else:
            painter.setPen(QPen(QColor(C.BORDER_ROW), 1.0))
            painter.drawLine(rect.x() + 8, rect.bottom(), rect.right() - 8, rect.bottom())

        # Active selection accent marker on left border
        if is_selected:
            accent_path = QPainterPath()
            accent_rect = QRectF(rect.x() + 2, rect.y() + 3, 3.0, rect.height() - 6)
            accent_path.addRoundedRect(accent_rect, 1.5, 1.5)
            painter.fillPath(accent_path, QBrush(QColor(C.ACCENT_PRIMARY)))

        # 2. Extract item metadata
        msg_name = str(index.data(Qt.DisplayRole) or "")
        spec_num = str(index.data(Qt.UserRole + 2) or "").replace("TS", "").strip()
        clause = str(index.data(Qt.UserRole + 3) or "").strip()

        meta = get_spec_metadata(spec_num) if spec_num else {}
        wg = meta.get("wg", "")
        acronym = str(index.data(Qt.UserRole + 4) or meta.get("acronym", "")).strip()

        left_x = rect.x() + (10 if not is_selected else 12)
        right_limit = rect.right() - 8

        # 3. Line 1: Message / IE Name (regular when unselected; semibold when selected)
        name_font = scaled_font(option.font, T.SIZE_TABLE)
        name_font.setWeight(QFont.DemiBold if is_selected else QFont.Normal)
        painter.setFont(name_font)
        if is_selected:
            painter.setPen(QColor(C.ACCENT_PRIMARY))
        elif is_hovered:
            painter.setPen(QColor(C.TEXT_PRIMARY))
        else:
            painter.setPen(QColor(C.TEXT_BODY))

        line1_h = max(scale_px(17), (rect.height() - 6) // 2)
        line1_rect = QRect(left_x, rect.y() + 3, max(40, right_limit - left_x), line1_h)
        elided_text = QFontMetrics(name_font).elidedText(msg_name, Qt.ElideRight, line1_rect.width())
        painter.drawText(line1_rect, Qt.AlignLeft | Qt.AlignVCenter, elided_text)

        # 4. Line 2: Compact WG chip + clean 'Spec · Acronym' subtitle, Right-aligned Clause
        chip_y = rect.y() + 3 + line1_h + 1
        chip_h = max(scale_px(15), min(scale_px(16), rect.bottom() - chip_y - 3))

        sub_font = scaled_font(option.font, T.SIZE_CAPTION)
        sub_font.setWeight(QFont.Normal)
        sub_fm = QFontMetrics(sub_font)

        clause_left_edge = right_limit
        if clause and clause != msg_name:
            clause_text = f"Cl. {clause}"
            clause_w = sub_fm.horizontalAdvance(clause_text) + 4
            clause_x = right_limit - clause_w
            if clause_x > left_x + 40:
                painter.setFont(sub_font)
                painter.setPen(QColor(C.TEXT_MUTED))
                painter.drawText(QRect(clause_x, chip_y, clause_w, chip_h), Qt.AlignRight | Qt.AlignVCenter, clause_text)
                clause_left_edge = clause_x - 6

        cur_x = left_x
        if wg:
            wg_colors = get_wg_colors(wg)
            wg_font = scaled_font(option.font, 10)
            wg_font.setWeight(QFont.DemiBold)
            wg_fm = QFontMetrics(wg_font)
            wg_w = wg_fm.horizontalAdvance(wg) + 8
            if cur_x + wg_w <= clause_left_edge:
                wg_path = QPainterPath()
                wg_path.addRoundedRect(QRectF(cur_x, chip_y + 0.5, wg_w, chip_h - 1), 3.0, 3.0)
                painter.fillPath(wg_path, QBrush(QColor(wg_colors.get("bg", C.BG_SURFACE))))
                painter.setPen(QPen(QColor(wg_colors.get("border", C.BORDER_STRONG)), 0.8))
                painter.drawPath(wg_path)

                painter.setFont(wg_font)
                painter.setPen(QColor(wg_colors.get("text", C.TEXT_BODY)))
                painter.drawText(QRect(cur_x, chip_y, wg_w, chip_h), Qt.AlignCenter, wg)
                cur_x += wg_w + 6

        subtitle_parts = []
        if spec_num:
            subtitle_parts.append(spec_num)
        if acronym and acronym != f"TS {spec_num}":
            subtitle_parts.append(acronym)
        subtitle_str = " · ".join(subtitle_parts)

        if subtitle_str:
            avail_sub = clause_left_edge - cur_x
            if avail_sub > 24:
                sub_elided = sub_fm.elidedText(subtitle_str, Qt.ElideRight, avail_sub)
                painter.setFont(sub_font)
                painter.setPen(QColor(C.TEXT_SECONDARY))
                painter.drawText(
                    QRect(cur_x, chip_y, avail_sub, chip_h),
                    Qt.AlignLeft | Qt.AlignVCenter,
                    sub_elided,
                )

        painter.restore()


class NASMatrixTreeDelegate(QStyledItemDelegate):
    """Renders crisp geometric tree guide lines for nested fields in the Evolution Matrix."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)

    def paint(self, painter: QPainter, option, index):
        guide_info = index.data(Qt.UserRole + 1)
        depth = index.data(Qt.UserRole)
        if guide_info is None or depth is None or depth <= 0:
            super().paint(painter, option, index)
            return

        ancestor_continues, is_last_sibling, raw_label = guide_info

        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        opt.text = ""

        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.CE_ItemViewItem, opt, painter, opt.widget)

        painter.save()
        rect = opt.rect
        pad_left = scale_px(6, minimum=4)
        indent_w = scale_px(16, minimum=14)
        mid_y = rect.center().y()

        painter.setPen(QPen(QColor(C.BORDER_STRONG), 1.0))

        for lvl_idx, continues in enumerate(ancestor_continues):
            if continues:
                cx = rect.left() + pad_left + lvl_idx * indent_w + indent_w // 2
                painter.drawLine(cx, rect.top(), cx, rect.bottom())

        branch_cx = rect.left() + pad_left + (depth - 1) * indent_w + indent_w // 2
        branch_end_y = mid_y if is_last_sibling else rect.bottom()
        painter.drawLine(branch_cx, rect.top(), branch_cx, branch_end_y)
        painter.drawLine(branch_cx, mid_y, branch_cx + indent_w // 2 + 2, mid_y)

        text_x = rect.left() + pad_left + depth * indent_w + 4
        text_rect = QRect(text_x, rect.top(), max(20, rect.right() - text_x - 4), rect.height())
        painter.setFont(opt.font)
        is_selected = bool(opt.state & QStyle.State_Selected)
        painter.setPen(QColor(C.TEXT_PRIMARY) if is_selected else QColor(C.TEXT_BODY))
        elided = QFontMetrics(opt.font).elidedText(str(raw_label), Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, elided)
        painter.restore()


class NASVersionTreeWidget(QTreeWidget):
    """Hierarchical specification and version management tree with tri-state selection."""

    selection_changed = pyqtSignal()
    reload_version_requested = pyqtSignal(str, str)
    reload_spec_requested = pyqtSignal(str, int)
    delete_version_requested = pyqtSignal(str, str)
    delete_spec_requested = pyqtSignal(str, int)
    structure_changed = pyqtSignal()
    open_spec_requested = pyqtSignal(str, str, str, str)  # spec_number, version, clause, format
    show_in_specs_requested = pyqtSignal(str, str)         # spec_number, version
    search_in_spec_search_requested = pyqtSignal(str, str) # spec_number, version

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._updating_checks: bool = False

        self.setHeaderHidden(True)
        self.setAlternatingRowColors(True)
        self.setIndentation(14)
        self.setStyleSheet(f"""
            QTreeWidget {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background-color: {C.BG_CARD};
                padding: 2px;
                outline: none;
            }}
            QTreeWidget::item {{
                padding: 3px 2px;
                min-height: 22px;
                border-radius: 3px;
            }}
            QTreeWidget::item:hover {{
                background-color: {C.BG_SURFACE};
            }}
            QTreeWidget::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.TEXT_PRIMARY};
            }}
        """)

        self.itemChanged.connect(self._on_item_changed)
        self.itemDoubleClicked.connect(self._on_item_double_clicked)
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

        text_color_header = QBrush(QColor(C.TEXT_PRIMARY))
        text_color_item = QBrush(QColor(C.TEXT_BODY))

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

    def _on_item_double_clicked(self, item: QTreeWidgetItem, column: int):
        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")
        if item_type == "version":
            spec_num = data.get("spec_number", "")
            version = data.get("version", "")
            if spec_num and version:
                self.open_spec_requested.emit(spec_num, version, "", "word")
        elif item_type == "spec":
            spec_num = data.get("spec_number", "")
            if spec_num:
                versions = self.get_all_versions_for_spec(spec_num)
                latest_v = versions[0]["version"] if versions else ""
                self.open_spec_requested.emit(spec_num, latest_v, "", "word")

    def _on_context_menu(self, pos):
        item = self.itemAt(pos)
        if not item:
            return

        data = item.data(0, Qt.UserRole) or {}
        item_type = data.get("type")

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

        if item_type == "version":
            spec_num = data.get("spec_number", "")
            version = data.get("version", "")

            act_word = QAction(f"📝 Open TS {spec_num} v{version} (Word)", self)
            act_word.triggered.connect(lambda: self.open_spec_requested.emit(spec_num, version, "", "word"))
            menu.addAction(act_word)

            act_pdf = QAction(f"📕 Open TS {spec_num} v{version} (PDF)", self)
            act_pdf.triggered.connect(lambda: self.open_spec_requested.emit(spec_num, version, "", "pdf"))
            menu.addAction(act_pdf)

            act_folder = QAction("📂 Open Containing Folder in Explorer", self)
            act_folder.triggered.connect(lambda: self.open_spec_requested.emit(spec_num, version, "", "folder"))
            menu.addAction(act_folder)

            menu.addSeparator()

            act_reload = QAction(f"🔄 Reload TS {spec_num} v{version}", self)
            act_reload.triggered.connect(lambda: self.reload_version_requested.emit(spec_num, version))
            menu.addAction(act_reload)

            menu.addSeparator()

            act_specs = QAction(f"📚 Show TS {spec_num} v{version} in Specifications Tab", self)
            act_specs.triggered.connect(lambda: self.show_in_specs_requested.emit(spec_num, version))
            menu.addAction(act_specs)

            act_search = QAction(f"🔎 Search TS {spec_num} in Spec Search Tab", self)
            act_search.triggered.connect(lambda: self.search_in_spec_search_requested.emit(spec_num, version))
            menu.addAction(act_search)

            menu.addSeparator()

            act_delete = QAction(f"🗑️ Delete TS {spec_num} v{version}", self)
            act_delete.triggered.connect(lambda: self.delete_version_requested.emit(spec_num, version))
            menu.addAction(act_delete)
            menu.exec_(self.viewport().mapToGlobal(pos))

        elif item_type == "spec":
            spec_num = data.get("spec_number", "")
            child_count = item.childCount()
            all_vers = self.get_all_versions_for_spec(spec_num)
            latest_v = all_vers[0]["version"] if all_vers else ""

            act_word = QAction(f"📝 Open Latest TS {spec_num} (v{latest_v}) in Word", self)
            act_word.triggered.connect(lambda: self.open_spec_requested.emit(spec_num, latest_v, "", "word"))
            menu.addAction(act_word)

            act_folder = QAction("📂 Open Local Folder in Explorer", self)
            act_folder.triggered.connect(lambda: self.open_spec_requested.emit(spec_num, latest_v, "", "folder"))
            menu.addAction(act_folder)

            menu.addSeparator()

            act_reload_spec = QAction(f"🔄 Reload all versions of TS {spec_num} ({child_count} versions)", self)
            act_reload_spec.triggered.connect(lambda: self.reload_spec_requested.emit(spec_num, child_count))
            menu.addAction(act_reload_spec)

            menu.addSeparator()

            act_specs = QAction(f"📚 Show TS {spec_num} in Specifications Tab", self)
            act_specs.triggered.connect(lambda: self.show_in_specs_requested.emit(spec_num, latest_v))
            menu.addAction(act_specs)

            act_search = QAction(f"🔎 Search TS {spec_num} in Spec Search Tab", self)
            act_search.triggered.connect(lambda: self.search_in_spec_search_requested.emit(spec_num, latest_v))
            menu.addAction(act_search)

            menu.addSeparator()

            act_delete_spec = QAction(f"🗑️ Delete all versions of TS {spec_num} ({child_count} versions)", self)
            act_delete_spec.triggered.connect(lambda: self.delete_spec_requested.emit(spec_num, child_count))
            menu.addAction(act_delete_spec)
            menu.exec_(self.viewport().mapToGlobal(pos))


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
        self.setFixedHeight(scale_px(20, minimum=18))

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
        lbl_spec.setFont(scaled_font(base_px=T.SIZE_TABLE, bold=True))
        lbl_spec.setStyleSheet(f"color: {C.TEXT_PRIMARY};")
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


class _NASCardCompatChild:
    def __init__(self, pill: NASReleasePillWidget):
        self.pill = pill

    def text(self, col: int = 0) -> str:
        rel = _extract_major_release(self.pill.version)
        return f"{rel} · v{self.pill.version}"

    def checkState(self, col: int = 0) -> Qt.CheckState:
        return Qt.Checked if self.pill.isChecked() else Qt.Unchecked


class _NASCardCompatSpec:
    def __init__(self, card: NASSpecCardWidget):
        self.card = card

    def text(self, col: int = 0) -> str:
        wg = self.card.meta.get("wg", "3GPP")
        acr = self.card.meta.get("acronym", "")
        return f"[{wg}] TS {self.card.spec_number} — {acr}".strip()

    def childCount(self) -> int:
        return len(self.card.pills)

    def child(self, idx: int) -> _NASCardCompatChild:
        return _NASCardCompatChild(self.card.pills[idx])


class NASVersionCardsWidget(QScrollArea):
    """
    Modern Spec-Search-aligned release pills selector for Protocols (`NASTab`).
    Replaces the deep vertical tree with compact specification cards and wrapping release pills.
    """

    selection_changed = pyqtSignal()
    reload_version_requested = pyqtSignal(str, str)
    reload_spec_requested = pyqtSignal(str, int)
    delete_version_requested = pyqtSignal(str, str)
    delete_spec_requested = pyqtSignal(str, int)
    structure_changed = pyqtSignal()
    open_spec_requested = pyqtSignal(str, str, str, str)
    show_in_specs_requested = pyqtSignal(str, str)
    search_in_spec_search_requested = pyqtSignal(str, str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._updating_checks: bool = False
        self._cards: Dict[str, NASSpecCardWidget] = {}

        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setStyleSheet(f"""
            QScrollArea {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background-color: {C.BG_SURFACE};
            }}
        """)

        self.container = QWidget()
        self.container.setObjectName("nasCardsContainer")
        self.container.setStyleSheet(
            f"QWidget#nasCardsContainer {{ background-color: {C.BG_SURFACE}; }}"
        )
        self.cards_layout = QVBoxLayout(self.container)
        self.cards_layout.setContentsMargins(4, 4, 4, 4)
        self.cards_layout.setSpacing(6)
        self.cards_layout.addStretch()
        self.setWidget(self.container)

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
            saved_tuples = {(item.get("spec_number"), item.get("version")) for item in saved_checked if item}

        specs_map: Dict[str, List[Dict[str, Any]]] = {}
        for v in versions:
            specs_map.setdefault(v["spec_number"], []).append(v)

        # Remove trailing stretch before adding cards
        count = self.cards_layout.count()
        if count > 0 and self.cards_layout.itemAt(count - 1).spacerItem():
            self.cards_layout.removeItem(self.cards_layout.itemAt(count - 1))

        for spec_num in sorted(specs_map.keys()):
            card = NASSpecCardWidget(spec_num, specs_map[spec_num], self.container)
            card.selection_changed.connect(self._on_card_selection_changed)
            card.collapsed_changed.connect(self.structure_changed.emit)
            card.reload_version_requested.connect(self.reload_version_requested.emit)
            card.reload_spec_requested.connect(self.reload_spec_requested.emit)
            card.delete_version_requested.connect(self.delete_version_requested.emit)
            card.delete_spec_requested.connect(self.delete_spec_requested.emit)
            card.open_spec_requested.connect(self.open_spec_requested.emit)
            card.show_in_specs_requested.connect(self.show_in_specs_requested.emit)
            card.search_in_spec_search_requested.connect(self.search_in_spec_search_requested.emit)

            card._updating_checks = True
            for pill in card.pills:
                if saved_tuples is not None:
                    is_chk = (str(spec_num), str(pill.version)) in saved_tuples
                else:
                    is_chk = True
                pill.setChecked(is_chk)
            card._updating_checks = False
            card._update_header_state()
            if saved_collapsed and spec_num in saved_collapsed:
                card.set_collapsed(True, emit_signal=False)

            self._cards[spec_num] = card
            self.cards_layout.addWidget(card)

        self.cards_layout.addStretch()
        self._updating_checks = False
        self.selection_changed.emit()

    def clear(self):
        while self.cards_layout.count():
            item = self.cards_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._cards.clear()
        self.cards_layout.addStretch()

    def _on_card_selection_changed(self):
        if self._updating_checks:
            return
        self.selection_changed.emit()
        self.structure_changed.emit()

    def filter_tree(self, query: str):
        q = str(query or "").strip().lower()
        if not q:
            for card in self._cards.values():
                card.setVisible(True)
            return
        for spec_num, card in self._cards.items():
            hay = " ".join([
                spec_num,
                str(card.meta.get("wg", "")),
                str(card.meta.get("acronym", "")),
                str(card.meta.get("title", "")),
                " ".join(p.text() for p in card.pills),
            ]).lower()
            card.setVisible(all(tok in hay for tok in q.split()))

    def select_all(self):
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                p.setChecked(True)
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def deselect_all(self):
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                p.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def select_none(self):
        self.deselect_all()

    def select_latest_only(self):
        """Selects only the single newest version for each indexed protocol specification."""
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            for idx, p in enumerate(card.pills):
                p.setChecked(idx == 0)
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def select_latest_per_release(self):
        """Selects the newest version of each major Release per specification for evolution diffs."""
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            seen = set()
            for p in card.pills:
                rel = _extract_major_release(p.version)
                if rel not in seen:
                    seen.add(rel)
                    p.setChecked(True)
                else:
                    p.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def select_releases_min(self, min_rel_num: int):
        """Selects only releases with major release >= min_rel_num (e.g. Rel-18+)."""
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                rel = _extract_major_release(p.version)
                m = re.search(r"\d+", rel)
                p.setChecked(bool(m and int(m.group()) >= min_rel_num))
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def get_selected_version_ids(self) -> List[int]:
        ids: List[int] = []
        for card in self._cards.values():
            for p in card.pills:
                if p.isChecked() and p.version_id is not None:
                    ids.append(p.version_id)
        return ids

    def get_checked_versions_info(self) -> List[Dict[str, Any]]:
        checked: List[Dict[str, Any]] = []
        for spec_num in sorted(self._cards.keys()):
            card = self._cards[spec_num]
            for p in card.pills:
                if p.isChecked():
                    checked.append({"spec_number": spec_num, "version": p.version})
        return checked

    def get_all_versions_for_spec(self, spec_number: str) -> List[Dict[str, Any]]:
        card = self._cards.get(spec_number)
        if not card:
            return []
        return [{"spec_number": card.spec_number, "version": p.version} for p in card.pills]

    def get_collapsed_spec_numbers(self) -> List[str]:
        return [spec_num for spec_num, card in self._cards.items() if card.is_collapsed()]

    def topLevelItemCount(self) -> int:
        return 1 + len(self._cards)

    def topLevelItem(self, idx: int):
        keys = sorted(self._cards.keys())
        card_idx = idx - 1
        if 0 <= card_idx < len(keys):
            return _NASCardCompatSpec(self._cards[keys[card_idx]])
        return None


RE_NAS_CROSS_REF = re.compile(
    r"(?:see\s+)?(?:subclause|clause|sub-clause)\s+([0-9A-Za-z\.]+)\s+in\s+(?:3GPP\s+)?TS\s+(24\.(?:301|501|526|588|555|578|514))",
    re.IGNORECASE,
)
RE_EXPLICIT_SPEC_REF = re.compile(
    r"^TS\s+((?:24|25|29|36|37|38)\.[0-9]{3})\s+([0-9A-Za-z\.]+)$",
    re.IGNORECASE,
)


class NASInspectorWidget(QFrame):
    """Modern Card Frame rendering Clause 9 diagrams, ASN.1 syntax, and field descriptions."""

    jump_to_message_requested = pyqtSignal(str)
    filter_by_ie_requested = pyqtSignal(str)
    import_spec_requested = pyqtSignal(str, int)  # (spec_number, major_version)
    open_spec_requested = pyqtSignal(str, str, str, str)  # (spec_number, version, clause_number, doc_format)
    show_in_specs_tab_requested = pyqtSignal(str, str)     # (spec_number, version)
    search_in_spec_search_requested = pyqtSignal(str, str) # (spec_number, version)

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

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 1. Header Toolbar
        header_frame = QFrame()
        header_frame.setObjectName("inspHeaderFrame")
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
        insp_header.setSpacing(8)

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
        self.stack.addWidget(self.text_area)

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

        prev_label = self.copy_btn.text()
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
            full_badge = f"{resolved_name}{acronym_str}"
            elided_badge = self.title_lbl.fontMetrics().elidedText(
                full_badge, Qt.ElideRight, max(140, self.title_lbl.maximumWidth() - 18)
            )
            self.title_lbl.setText(elided_badge)
            self.title_lbl.setToolTip(full_badge)
            self._current_clause_defs = {f"{d['spec_number']} v{d['version']}": d for d in defs}

            if containing_msgs:
                num_msgs = len(containing_msgs)
                self.usage_btn.setText(f"Used in: {num_msgs} msg{'s' if num_msgs != 1 else ''} ▾")
                self.usage_btn.setEnabled(True)
                self.usage_btn.setVisible(True)
            else:
                self.usage_btn.setEnabled(False)
                self.usage_btn.setVisible(False)

            self._updating_combo = True
            self.version_combo.clear()
            for d in defs:
                clean_v = str(d.get("version", "")).lstrip("vV")
                parts = clean_v.split(".")
                rel_str = f"· R{parts[0]} " if parts and parts[0].isdigit() else ""
                key_name = f"TS {d['spec_number']} {rel_str}(v{d['version']})" if rel_str else f"TS {d['spec_number']} v{d['version']}"
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
            self._update_open_spec_menu()
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

        match_explicit = RE_EXPLICIT_SPEC_REF.match(fallback_type_ref.strip())
        if match_explicit:
            ext_spec = match_explicit.group(1)
            ext_clause = match_explicit.group(2)
            ext_meta = get_spec_metadata(ext_spec)
            html_output = (
                f'<h2 style="color: {C.ACCENT_SKY_DARK}; margin-top: 2px; margin-bottom: 8px; '
                f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid {C.BORDER_SUBTLE}; padding-bottom: 4px;">'
                f'{html.escape(ie_name)} (TS {html.escape(ext_spec)} Clause {html.escape(ext_clause)})</h2>'
                f'<div style="margin-top: 10px; border: 1px solid {C.WARNING_BORDER_STRONG}; background-color: {C.WARNING_SOFT}; border-radius: 6px; padding: 12px; border-left: 4px solid {C.WARNING_ACCENT}; font-family: Segoe UI, sans-serif;">'
                f'<div style="font-weight: bold; color: {C.WARNING_DARK}; font-size: 12px; margin-bottom: 4px;">'
                f'ℹ️ Cross-Specification Reference: TS {html.escape(ext_spec)} — {html.escape(ext_meta.get("title", ""))} (Clause {html.escape(ext_clause)})'
                f'</div>'
                f'<div style="color: {C.WARNING_BROWN_DEEP}; font-size: 11px; line-height: 1.5;">'
                f'This UE policy structure is defined in <b>3GPP TS {html.escape(ext_spec)} Clause {html.escape(ext_clause)}</b>, which is not currently present in your local database.<br/>'
                f'<p style="margin-top: 8px; margin-bottom: 2px;">'
                f'<a href="action:import:{ext_spec}:0" style="display: inline-block; background-color: {C.ACCENT_SKY}; color: {C.BG_CARD}; padding: 4px 12px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 11px;">'
                f'📥 Import TS {html.escape(ext_spec)} from 3GPP Archive'
                f'</a>'
                f'</p>'
                f'</div>'
                f'</div>'
            )
            self._current_html_content = html_output
            self.text_area.setHtml(html_output)
            self._update_open_spec_menu()
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(True)
        elif is_inline_type or fallback_type_ref:
            html_output = self._render_inline_type_html(
                ie_name=ie_name,
                type_ref=fallback_type_ref,
                field_description=field_description,
                spec_number=spec_number,
            )
            self._current_html_content = html_output
            self.text_area.setHtml(html_output)
            self._update_open_spec_menu()
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(True)
        else:
            fb_text = f"Type / Reference: {fallback_type_ref}\n(No structure definition found in database)"
            self._current_html_content = fb_text
            self.text_area.setPlainText(fb_text)
            self._update_open_spec_menu()
            self.stack.setCurrentIndex(1)
            self.copy_btn.setEnabled(bool(fallback_type_ref))

    def _render_selected_definition(self, definition: Dict[str, Any]):
        base_html = definition.get("raw_description", "")
        current_spec = definition.get("spec_number", "")
        current_ver_str = definition.get("version", "")
        current_clause = definition.get("clause", "")

        # Check for 4G <-> 5G and UE Policy (24.501 -> 24.526, 24.588, 24.555, 24.514, 24.578) cross-specification references
        if self.db and current_spec in ("24.501", "24.301"):
            cross_targets: List[Tuple[str, str]] = []
            match = RE_NAS_CROSS_REF.search(base_html)
            if match:
                target_clause = match.group(1).rstrip(".[] ")
                target_spec = match.group(2).strip()
                if target_spec != current_spec:
                    cross_targets.append((target_spec, target_clause))

            if current_spec == "24.501" and current_clause in ("D.6.2", "D.6.8"):
                policy_clause_map = {
                    "24.526": "5.2",
                    "24.588": "5.2",
                    "24.555": "5.2",
                    "24.514": "12.2",
                    "24.578": "5.2",
                }
                for p_spec, p_clause in policy_clause_map.items():
                    if p_spec in base_html and not any(t[0] == p_spec for t in cross_targets):
                        cross_targets.append((p_spec, p_clause))

            ver_tuple = parse_version_tuple(current_ver_str)
            major_ver = (ver_tuple[0][1] if isinstance(ver_tuple[0], tuple) else ver_tuple[0]) if ver_tuple else 15

            for target_spec, target_clause in cross_targets:
                cross_def = self.db.get_cross_referenced_ie_definition(
                    target_spec=target_spec,
                    target_clause=target_clause,
                    alt_name=definition.get("ie_name", ""),
                    major_version=major_ver,
                )
                if not cross_def:
                    cross_def = self.db.get_cross_referenced_ie_definition(
                        target_spec=target_spec,
                        target_clause=target_clause,
                        alt_name="",
                        major_version=None,
                    )

                if cross_def and cross_def.get("raw_description"):
                    cross_body = cross_def.get("raw_description", "").strip()
                    c_ver = cross_def.get("version", "")
                    c_cl = cross_def.get("clause", target_clause)
                    cross_html = (
                        f'<div style="margin-top: 14px; border: 1.5px solid {C.ACCENT_SKY}; background-color: {C.SKY_SURFACE}; border-radius: 6px; padding: 12px;">'
                        f'<div style="margin-bottom: 10px; border-bottom: 1px solid {C.SKY_BORDER}; padding-bottom: 6px;">'
                        f'<span style="font-size: 12px; font-weight: bold; color: {C.ACCENT_SKY_DARK};">'
                        f'<a href="action:open_spec:{target_spec}:{c_ver}:{c_cl}" style="color: {C.ACCENT_SKY_DARK}; text-decoration: underline;" title="Click to open TS {target_spec} v{c_ver} in Word">'
                        f'🔗 Cross-Referenced from TS {target_spec} v{c_ver} (Clause {c_cl}) [Open Word]'
                        f'</a>'
                        f'</span>'
                        f'<span style="font-size: 10px; color: {C.ACCENT_SKY}; margin-left: 8px; background-color: {C.SKY_SOFT}; padding: 2px 6px; border-radius: 3px; border: 1px solid {C.SKY_BORDER_STRONG};">'
                        f'v{c_ver}'
                        f'</span>'
                        f'</div>'
                        f'<div style="background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 10px;">'
                        f'{cross_body}'
                        f'</div>'
                        f'</div>'
                    )
                    base_html += cross_html
                else:
                    missing_html = (
                        f'<div style="margin-top: 14px; border: 1px solid {C.WARNING_BORDER_STRONG}; background-color: {C.WARNING_SOFT}; border-radius: 6px; padding: 10px; border-left: 4px solid {C.WARNING_ACCENT}; font-family: Segoe UI, sans-serif;">'
                        f'<div style="font-weight: bold; color: {C.WARNING_DARK}; font-size: 12px; margin-bottom: 4px;">'
                        f'ℹ️ Cross-Specification Reference: TS {target_spec} (Clause {target_clause})'
                        f'</div>'
                        f'<div style="color: {C.WARNING_BROWN_DEEP}; font-size: 11px; line-height: 1.5;">'
                        f'This Information Element references <b>TS {target_spec} (Release {major_ver})</b>, which is not currently present in your local database.<br/>'
                        f'<p style="margin-top: 6px; margin-bottom: 2px;">'
                        f'<a href="action:import:{target_spec}:{major_ver}" style="display: inline-block; background-color: {C.ACCENT_SKY}; color: {C.BG_CARD}; padding: 4px 12px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 11px;">'
                        f'📥 Import TS {target_spec} (Rel-{major_ver}) from 3GPP Archive'
                        f'</a>'
                        f'</p>'
                        f'</div>'
                        f'</div>'
                    )
                    base_html += missing_html

        self._current_html_content = base_html
        self.text_area.setHtml(base_html)

    def _on_version_changed(self, index: int):
        if self._updating_combo or index < 0:
            return
        version_key = self.version_combo.itemData(index)
        if version_key in self._current_clause_defs:
            self._render_selected_definition(self._current_clause_defs[version_key])
            self._update_open_spec_menu()

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
            f'<h2 style="color: {C.ACCENT_SKY_DARK}; margin-top: 2px; margin-bottom: 8px; '
            f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid {C.BORDER_SUBTLE}; padding-bottom: 4px;">'
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
            f'<div style="font-size: 11px; font-weight: bold; color: {C.TEXT_SECONDARY}; margin-bottom: 3px;">Inline ASN.1 Definition:</div>'
            f'<pre style="background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; padding: 8px; '
            f'border-radius: 4px; font-family: Consolas, monospace; font-size: 11px; color: {C.TEXT_PRIMARY}; '
            f'white-space: pre-wrap; margin: 0;">{html.escape(formatted_asn1)}</pre>'
            f'</div>'
        )

        desc_html = ""
        if field_description and field_description.strip():
            desc_html = (
                f'<div style="margin-top: 8px;">'
                f'<div style="font-size: 11px; font-weight: bold; color: {C.TEXT_SECONDARY}; margin-bottom: 4px;">'
                f'Field Description & Semantics:</div>'
                f'<div style="background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-left: 3px solid {C.ACCENT_SKY}; '
                f'padding: 8px; border-radius: 3px; font-family: Segoe UI, sans-serif; font-size: 11px; color: {C.TEXT_BODY}; line-height: 1.4;">'
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
        self._current_html_content = ""
        self._current_clause_defs.clear()
        self.copy_btn.setEnabled(False)
        self.open_spec_btn.setEnabled(False)
        self.open_spec_btn.setText("📖 Open Spec")
        self.open_spec_btn.setMenu(None)
        self.stack.setCurrentIndex(0)

    def _show_usage_menu(self):
        if not self._current_ie_clause:
            return

        clause = self._current_ie_clause
        name = self._current_ie_name or ""
        spec_num = self._current_ie_spec

        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                font-size: 11px;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                padding: 4px;
            }}
            QMenu::item {{
                padding: 4px 20px 4px 10px;
                border-radius: 3px;
            }}
            QMenu::item:selected {{
                background-color: {C.SKY_SOFT};
                color: {C.ACCENT_SKY_DARK};
            }}
            QMenu::separator {{
                height: 1px;
                background-color: {C.BORDER_SUBTLE};
                margin: 4px 0;
            }}
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
        if num_msgs > 0:
            self.usage_btn.setText(f"Used in: {num_msgs} msg{'s' if num_msgs != 1 else ''} ▾")
            self.usage_btn.setVisible(True)
            self.usage_btn.setEnabled(True)
        else:
            self.usage_btn.setVisible(False)
            self.usage_btn.setEnabled(False)
