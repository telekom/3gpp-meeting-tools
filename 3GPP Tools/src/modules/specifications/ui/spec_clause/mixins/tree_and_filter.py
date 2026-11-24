"""Document preparation, hierarchical heading tree population, and selection filtering mixin."""

from __future__ import annotations

from pathlib import Path
from typing import List

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMessageBox, QTreeWidgetItem

from core.ui.theme import C
from core.ui.ui_components import BADGE_STYLE_PRIMARY
from modules.specifications.ui.threads import SpecExcerptPrepThread


class SpecClauseTreeMixin:
    """Manages SpecExcerptPrepThread callbacks, tree population, search filtering, and check cascades."""

    def _start_preparation(self):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Starting specification preparation...")

        self._prep_thread = SpecExcerptPrepThread(
            spec_num=self.spec_num,
            version=self.version,
            fname=self.fname,
            url=self.url,
            spec_dl_dir=self.spec_dl_dir,
        )
        self._prep_thread.progress.connect(self._on_prep_progress)
        self._prep_thread.prepared.connect(self._on_prep_finished)
        self._prep_thread.error.connect(self._on_prep_error)
        self._prep_thread.start()

    def _on_prep_progress(self, pct: int, msg: str):
        self.progress_bar.setValue(pct)
        self.status_label.setText(msg)

    def _on_prep_finished(self, ordered_paths: list, headings: list):
        self._ordered_paths = list(ordered_paths)
        self._headings = list(headings)

        parts_count = len(self._ordered_paths)
        parts_text = f"{parts_count} part" if parts_count == 1 else f"{parts_count} parts"
        self.source_summary_badge.setText(f"📄 {parts_text} ({len(self._headings)} clauses)")
        self.source_summary_badge.setStyleSheet(BADGE_STYLE_PRIMARY)

        self.progress_bar.setValue(100)
        self.status_label.setText("Ready: Select clauses and click Extract.")
        self._populate_tree()

    def _on_prep_error(self, err_msg: str):
        self.progress_bar.setVisible(False)
        self.status_label.setText("⚠️ Preparation failed.")
        self.status_label.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")
        self.source_summary_badge.setText("Failed")
        self.source_summary_badge.setStyleSheet(f"color: {C.DANGER_ALT};")
        QMessageBox.critical(
            self,
            "Preparation Failed",
            f"Could not prepare specification for extraction:\n{err_msg}",
        )

    def _populate_tree(self):
        self._is_updating_checks = True
        self.clause_tree.clear()

        by_source = {}
        for h in self._headings:
            by_source.setdefault(h.source_path, []).append(h)

        multi_part = len(self._ordered_paths) > 1

        for source_path in self._ordered_paths:
            source_name = Path(source_path).name
            if multi_part:
                source_item = QTreeWidgetItem([f"📄 {source_name}", source_name])
                source_item.setFlags(source_item.flags() & ~Qt.ItemIsUserCheckable)
                font = source_item.font(0)
                font.setBold(True)
                source_item.setFont(0, font)
                self.clause_tree.addTopLevelItem(source_item)
                parent_root = source_item
            else:
                parent_root = None

            parents = {}
            for h in by_source.get(source_path, []):
                item = QTreeWidgetItem([h.text, source_name])
                item.setData(0, Qt.UserRole, h.id)
                item.setCheckState(0, Qt.Unchecked)

                parent = None
                for level in range(h.level - 1, 0, -1):
                    if level in parents:
                        parent = parents[level]
                        break

                if parent is not None:
                    parent.addChild(item)
                elif parent_root is not None:
                    parent_root.addChild(item)
                else:
                    self.clause_tree.addTopLevelItem(item)

                parents[h.level] = item
                for level in list(parents):
                    if level > h.level:
                        del parents[level]

        self.clause_tree.expandToDepth(1 if multi_part else 0)
        self._is_updating_checks = False
        self._update_extract_button()

    def _on_tree_item_changed(self, item: QTreeWidgetItem, column: int):
        if self._is_updating_checks or column != 0:
            return

        self._is_updating_checks = True
        try:
            state = item.checkState(0)

            # Cascade down to all child items
            def set_child_states(parent_item):
                for i in range(parent_item.childCount()):
                    child = parent_item.child(i)
                    if child.data(0, Qt.UserRole):
                        child.setCheckState(0, state)
                    set_child_states(child)

            set_child_states(item)
        finally:
            self._is_updating_checks = False

        self._update_extract_button()

    def _filter_tree(self, text: str):
        query = text.strip().lower()
        matches = 0

        def visit(item: QTreeWidgetItem) -> bool:
            nonlocal matches
            own_match = not query or query in item.text(0).lower()
            if query and own_match and item.data(0, Qt.UserRole):
                matches += 1
            child_match = False
            for i in range(item.childCount()):
                child_match = visit(item.child(i)) or child_match

            is_visible = own_match or child_match
            item.setHidden(not is_visible)
            if query and child_match:
                item.setExpanded(True)
            return is_visible

        for i in range(self.clause_tree.topLevelItemCount()):
            visit(self.clause_tree.topLevelItem(i))

        if hasattr(self, "lbl_filter_matches"):
            if query:
                self.lbl_filter_matches.setText(f"{matches} match(es)")
            else:
                self.lbl_filter_matches.setText("")

    def _clear_selection(self):
        self._is_updating_checks = True
        try:

            def clear_node(item: QTreeWidgetItem):
                if item.data(0, Qt.UserRole):
                    item.setCheckState(0, Qt.Unchecked)
                for i in range(item.childCount()):
                    clear_node(item.child(i))

            for i in range(self.clause_tree.topLevelItemCount()):
                clear_node(self.clause_tree.topLevelItem(i))
        finally:
            self._is_updating_checks = False

        self._update_extract_button()

    def _get_selected_ids(self) -> List[str]:
        selected = []

        def collect(item: QTreeWidgetItem):
            h_id = item.data(0, Qt.UserRole)
            if h_id and item.checkState(0) == Qt.Checked:
                selected.append(h_id)
            for i in range(item.childCount()):
                collect(item.child(i))

        for i in range(self.clause_tree.topLevelItemCount()):
            collect(self.clause_tree.topLevelItem(i))

        return selected

    def _update_extract_button(self):
        selected_ids = self._get_selected_ids()
        count = len(selected_ids)
        self.btn_extract.setEnabled(count > 0)
        self.btn_cr_pcr.setEnabled(count > 0)
        self.btn_clear.setEnabled(count > 0)
        self.btn_cr_pcr.setText(
            f"📝 Draft CR / pCR ({count})..." if count > 0 else "📝 Draft CR / pCR..."
        )
        if count > 0:
            self.btn_extract.setText(f"✂️ Extract Selected Clauses ({count})...")
            if hasattr(self, "lbl_selection_count"):
                self.lbl_selection_count.setText(f"<b>{count}</b> clause(s) selected")
                self.lbl_selection_count.setStyleSheet(
                    f"color: {C.ACCENT_SKY}; font-size: 12px; font-weight: 500;"
                )
        else:
            self.btn_extract.setText("✂️ Extract Selected Clauses...")
            if hasattr(self, "lbl_selection_count"):
                self.lbl_selection_count.setText("No clauses selected")
                self.lbl_selection_count.setStyleSheet(
                    f"color: {C.TEXT_MUTED}; font-size: 12px; font-weight: 500;"
                )
