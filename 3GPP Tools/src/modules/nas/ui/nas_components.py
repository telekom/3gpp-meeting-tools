from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QAction,
    QComboBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.nas.core.nas_db import parse_version_tuple
from modules.nas.ui.nas_dialogs import get_spec_title


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
                padding: 1px;
            }
            QTreeWidget::item {
                padding: 1px 0px;
                min-height: 18px;
            }
            QTreeWidget::item:hover {
                background-color: #F1F5F9;
            }
            QTreeWidget::item:selected {
                background-color: #E2E8F0;
                color: #0F172A;
            }
        """)

        self.itemChanged.connect(self._on_item_changed)
        self.itemExpanded.connect(lambda _: self.structure_changed.emit())
        self.itemCollapsed.connect(lambda _: self.structure_changed.emit())
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._on_context_menu)

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
            spec_item = QTreeWidgetItem(self)
            spec_item.setText(0, get_spec_title(spec_num))
            spec_item.setData(0, Qt.UserRole, {"type": "spec", "spec_number": spec_num})
            spec_item.setFlags(spec_item.flags() | Qt.ItemIsUserCheckable)
            spec_item.setFont(0, bold_font)
            spec_item.setForeground(0, text_color_header)
            spec_item.setCheckState(0, Qt.Checked)

            sorted_v_list = sorted(spec_versions, key=lambda x: parse_version_tuple(x["version"]), reverse=True)

            for v in sorted_v_list:
                child = QTreeWidgetItem(spec_item)
                child.setText(0, f"v{v['version']}")
                child.setToolTip(0, f"TS {v['spec_number']} v{v['version']}\n(Right-click to delete)")
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


class NASInspectorWidget(QGroupBox):
    """Renders Clause 9 diagrams, ASN.1 syntax blocks, and field description tables."""

    jump_to_message_requested = pyqtSignal(str)
    filter_by_ie_requested = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__("Structure && Field Descriptions Inspector", parent)
        self._current_clause_defs: Dict[str, Dict[str, Any]] = {}
        self._current_ie_clause: Optional[str] = None
        self._current_ie_name: Optional[str] = None
        self._current_ie_spec: Optional[str] = None
        self._current_containing_msgs: List[Dict[str, Any]] = []
        self._updating_combo: bool = False

        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        insp_header = QHBoxLayout()
        insp_header.setContentsMargins(0, 0, 0, 4)

        self.title_lbl = QLabel("No Information Element selected")
        self.title_lbl.setStyleSheet("font-weight: bold; color: #0284C7; font-size: 12px;")
        insp_header.addWidget(self.title_lbl)

        insp_header.addStretch()

        self.usage_btn = QPushButton("Used in: 0 messages ▾")
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
                padding: 2px 8px;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #F8FAFC;
                min-width: 110px;
            }
            QComboBox:hover {
                border-color: #0284C7;
                background-color: #FFFFFF;
            }
        """)
        self.version_combo.currentIndexChanged.connect(self._on_version_changed)
        insp_header.addWidget(self.version_combo)
        layout.addLayout(insp_header)

        self.text_area = QTextEdit()
        self.text_area.setReadOnly(True)
        self.text_area.setPlaceholderText(
            "Click on an Information Element above to inspect its details and field descriptions..."
        )
        layout.addWidget(self.text_area)

    def display_definitions(
            self,
            clause: str,
            ie_name: str,
            spec_number: Optional[str],
            defs: List[Dict[str, Any]],
            containing_msgs: List[Dict[str, Any]],
            target_version_hint: Optional[str] = None,
            fallback_type_ref: str = "",
    ):
        self._current_ie_clause = clause
        self._current_ie_name = ie_name
        self._current_ie_spec = spec_number
        self._current_containing_msgs = containing_msgs

        if not defs:
            self.title_lbl.setText(f"{ie_name} ({fallback_type_ref})")
            self.usage_btn.setVisible(False)
            self.version_combo.clear()
            self._current_clause_defs.clear()
            self.text_area.setPlainText(
                f"Type / Reference: {fallback_type_ref}\n(No structure definition found in database)"
            )
            return

        resolved_name = defs[0]["ie_name"]
        spec_badge = f" (TS {spec_number})" if spec_number else ""
        self.title_lbl.setText(f"{resolved_name}{spec_badge}")
        self._current_clause_defs = {f"{d['spec_number']} v{d['version']}": d for d in defs}

        num_msgs = len(containing_msgs)
        self.usage_btn.setText(f"Used in: {num_msgs} message{'s' if num_msgs != 1 else ''} ▾")
        self.usage_btn.setVisible(True)

        self._updating_combo = True
        self.version_combo.clear()
        for d in defs:
            key_name = f"TS {d['spec_number']} v{d['version']}"
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
        self.text_area.setHtml(selected_def["raw_description"])

    def clear_display(self):
        self.title_lbl.setText("No Information Element selected")
        self.usage_btn.setVisible(False)
        self.version_combo.clear()
        self.text_area.clear()
        self._current_clause_defs.clear()

    def _on_version_changed(self, index: int):
        if self._updating_combo or index < 0:
            return
        version_key = self.version_combo.itemData(index)
        if version_key in self._current_clause_defs:
            self.text_area.setHtml(self._current_clause_defs[version_key]["raw_description"])

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
            header_title += f" [TS {spec_num}]"
        header_action = QAction(f"{header_title}:", self)
        header_action.setEnabled(False)
        menu.addAction(header_action)
        menu.addSeparator()

        if self._current_containing_msgs:
            for m in self._current_containing_msgs:
                msg_name = m["message_name"]
                clause_ref = m["clause"]
                m_spec = m.get("spec_number", "")
                spec_tag = f" [TS {m_spec}]" if m_spec and not spec_num else ""
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