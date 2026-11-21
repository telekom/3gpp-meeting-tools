"""Dual WI/AI tree population, filtering, and synchronized selection mixin."""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QTreeWidget, QTreeWidgetItem

from modules.meetings.ui.dialogs.pre_meeting_export.styles import (
    STYLE_TOGGLE_ACTIVE,
    STYLE_TOGGLE_INACTIVE,
    natural_ai_sort_key,
)


class PreMeetingTreeSelectionMixin:
    """Manages synchronized WI->AI and AI->WI checkable selection trees."""

    def _populate_tree_wi(self):
        """Populates Page 0: Work Item -> Agenda Items."""
        self.tree_wi.blockSignals(True)
        try:
            self.tree_wi.clear()
            for wi in self.meeting_wis:
                ai_dict = self.wi_to_ais.get(wi, {})
                n_ais = len(ai_dict)
                total_docs = sum(ai_dict.values())

                parent_item = QTreeWidgetItem(self.tree_wi)
                parent_item.setText(0, f"{wi}  ({n_ais} AIs, {total_docs} TDocs)")
                parent_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                parent_item.setCheckState(0, Qt.Unchecked)
                parent_item.setData(0, Qt.UserRole, wi)

                sorted_ais = sorted(ai_dict.keys(), key=natural_ai_sort_key)
                for ai in sorted_ais:
                    docs_count = ai_dict[ai]
                    desc = self.ai_descriptions.get(ai, "")
                    desc_str = f": {desc}" if desc else ""
                    child_label = f"AI {ai}{desc_str} ({docs_count} docs)"

                    child_item = QTreeWidgetItem(parent_item)
                    child_item.setText(0, child_label)
                    child_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                    child_item.setCheckState(0, Qt.Unchecked)
                    child_item.setData(0, Qt.UserRole, (wi, ai))
        finally:
            self.tree_wi.blockSignals(False)

    def _populate_tree_ai(self):
        """Populates Page 1: Agenda Item -> Work Items."""
        self.tree_ai.blockSignals(True)
        try:
            self.tree_ai.clear()
            for ai in self.meeting_ais:
                wi_dict = self.ai_to_wis.get(ai, {})
                n_wis = len(wi_dict)
                total_docs = sum(wi_dict.values())
                desc = self.ai_descriptions.get(ai, "")
                desc_str = f" · {desc}" if desc else ""

                parent_item = QTreeWidgetItem(self.tree_ai)
                parent_item.setText(0, f"AI {ai}{desc_str}  ({n_wis} WIs, {total_docs} TDocs)")
                parent_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                parent_item.setCheckState(0, Qt.Unchecked)
                parent_item.setData(0, Qt.UserRole, ai)

                sorted_wis = sorted(wi_dict.keys(), key=lambda x: (x == "General", x.lower()))
                for wi in sorted_wis:
                    docs_count = wi_dict[wi]
                    child_label = f"{wi} ({docs_count} docs)"

                    child_item = QTreeWidgetItem(parent_item)
                    child_item.setText(0, child_label)
                    child_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                    child_item.setCheckState(0, Qt.Unchecked)
                    child_item.setData(0, Qt.UserRole, (wi, ai))
        finally:
            self.tree_ai.blockSignals(False)

    def _on_tree_item_changed(self, tree: QTreeWidget, item: QTreeWidgetItem, column: int):
        if self._updating_tree:
            return

        if not self.rb_custom.isChecked():
            self.rb_custom.setChecked(True)

        self._updating_tree = True
        try:
            is_parent = (item.parent() is None)
            new_state = item.checkState(0)

            if is_parent:
                target_state = Qt.Checked if new_state == Qt.Checked else Qt.Unchecked
                item.setCheckState(0, target_state)
                for i in range(item.childCount()):
                    child = item.child(i)
                    child.setCheckState(0, target_state)
                    pair = child.data(0, Qt.UserRole)
                    if isinstance(pair, tuple):
                        if target_state == Qt.Checked:
                            self.selected_wi_ais.add(pair)
                        else:
                            self.selected_wi_ais.discard(pair)
            else:
                pair = item.data(0, Qt.UserRole)
                if isinstance(pair, tuple):
                    if new_state == Qt.Checked:
                        self.selected_wi_ais.add(pair)
                    else:
                        self.selected_wi_ais.discard(pair)
                self._update_parent_check_state(item.parent())
        finally:
            self._updating_tree = False

        self._update_estimates()

    def _update_parent_check_state(self, parent: QTreeWidgetItem):
        if not parent or parent.childCount() == 0:
            return
        n = parent.childCount()
        n_checked = sum(1 for i in range(n) if parent.child(i).checkState(0) == Qt.Checked)
        n_partial = sum(1 for i in range(n) if parent.child(i).checkState(0) == Qt.PartiallyChecked)

        if n_checked == n:
            parent.setCheckState(0, Qt.Checked)
        elif n_checked == 0 and n_partial == 0:
            parent.setCheckState(0, Qt.Unchecked)
        else:
            parent.setCheckState(0, Qt.PartiallyChecked)

    def _sync_tree_from_selection(self, tree: QTreeWidget):
        tree.blockSignals(True)
        self._updating_tree = True
        try:
            root = tree.invisibleRootItem()
            for i in range(root.childCount()):
                parent = root.child(i)
                for j in range(parent.childCount()):
                    child = parent.child(j)
                    pair = child.data(0, Qt.UserRole)
                    if isinstance(pair, tuple):
                        child.setCheckState(0, Qt.Checked if pair in self.selected_wi_ais else Qt.Unchecked)
                self._update_parent_check_state(parent)
        finally:
            self._updating_tree = False
            tree.blockSignals(False)

    def _expand_checked_parents(self, tree: QTreeWidget):
        root = tree.invisibleRootItem()
        for i in range(root.childCount()):
            parent = root.child(i)
            parent.setExpanded(parent.checkState(0) in (Qt.Checked, Qt.PartiallyChecked))

    def _switch_custom_view(self, mode: str):
        if mode == "wi":
            self.btn_view_wi.setStyleSheet(STYLE_TOGGLE_ACTIVE)
            self.btn_view_ai.setStyleSheet(STYLE_TOGGLE_INACTIVE)
            self._sync_tree_from_selection(self.tree_wi)
            self.tree_stack.setCurrentWidget(self.tree_wi)
            self._filter_tree(self.tree_wi, self.custom_search.text())
        else:
            self.btn_view_wi.setStyleSheet(STYLE_TOGGLE_INACTIVE)
            self.btn_view_ai.setStyleSheet(STYLE_TOGGLE_ACTIVE)
            self._sync_tree_from_selection(self.tree_ai)
            self.tree_stack.setCurrentWidget(self.tree_ai)
            self._filter_tree(self.tree_ai, self.custom_search.text())

    def _filter_custom_tree(self, text: str):
        active_tree = self.tree_stack.currentWidget()
        if isinstance(active_tree, QTreeWidget):
            self._filter_tree(active_tree, text)

    def _filter_tree(self, tree: QTreeWidget, query: str):
        q = query.strip().lower()
        root = tree.invisibleRootItem()
        for i in range(root.childCount()):
            parent = root.child(i)
            parent_matches = bool(q and q in parent.text(0).lower())
            any_child_matched = False

            for j in range(parent.childCount()):
                child = parent.child(j)
                child_matches = bool(not q or q in child.text(0).lower() or parent_matches)
                child.setHidden(not child_matches)
                if child_matches:
                    any_child_matched = True

            if not q:
                parent.setHidden(False)
            else:
                parent.setHidden(not (parent_matches or any_child_matched))
                if any_child_matched:
                    parent.setExpanded(True)

    def _expand_active_tree(self):
        active_tree = self.tree_stack.currentWidget()
        if isinstance(active_tree, QTreeWidget):
            active_tree.expandAll()

    def _collapse_active_tree(self):
        active_tree = self.tree_stack.currentWidget()
        if isinstance(active_tree, QTreeWidget):
            active_tree.collapseAll()

    def _select_all_custom(self):
        if not self.rb_custom.isChecked():
            self.rb_custom.setChecked(True)
        for wi in self.meeting_wis:
            for ai in self.wi_to_ais.get(wi, {}):
                self.selected_wi_ais.add((wi, ai))
        self._sync_tree_from_selection(self.tree_wi)
        self._sync_tree_from_selection(self.tree_ai)
        self._update_estimates()

    def _clear_all_custom(self):
        if not self.rb_custom.isChecked():
            self.rb_custom.setChecked(True)
        self.selected_wi_ais.clear()
        self._sync_tree_from_selection(self.tree_wi)
        self._sync_tree_from_selection(self.tree_ai)
        self._update_estimates()
