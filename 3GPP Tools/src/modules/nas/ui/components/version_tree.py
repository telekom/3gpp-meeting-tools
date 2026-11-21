"""Legacy collapsible tree selector for imported protocol specifications and versions."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import QAction, QMenu, QTreeWidget, QTreeWidgetItem, QWidget

from core.ui.theme import C
from modules.nas.core.nas_db import parse_version_tuple
from modules.nas.ui.nas_dialogs import get_spec_metadata, get_spec_tooltip


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
