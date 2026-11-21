"""Spec preparation, clause tree filtering, change-block grouping, and naming preview mixin."""

from pathlib import Path
from typing import List

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMessageBox, QTreeWidgetItem

from core.ui.ui_components import BADGE_STYLE_MUTED, BADGE_STYLE_PRIMARY
from modules.specifications.ui.threads import SpecExcerptPrepThread
from modules.word_tools.core.cr_pcr_generator import (
    extract_clause_summary,
    group_selected_headings_into_blocks,
)
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.word_excerpt_naming import generate_cr_pcr_export_name
from core.ui.theme import C


class CrPcrTreeNamingMixin:
    """Manages clause tree checks, search filtering, change-block summary, and export naming."""

    def _on_mode_changed(self):
        is_cr = self.radio_cr.isChecked()
        self.current_mode = "cr" if is_cr else "pcr"
        self.mode_stack.setCurrentIndex(1 if is_cr else 0)
        self._update_change_block_summary()
        self._update_preview_filename(force=True)

    def _save_default_company(self, text: str):
        comp = text.strip()
        if comp:
            cfg = WordConfig.load()
            cfg["default_source_company"] = comp
            WordConfig.save(cfg)

    def _start_spec_preparation(self):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Preparing specification document for extraction...")
        self.status_badge.setText("⏳ Preparing Spec...")
        self.status_badge.setStyleSheet(BADGE_STYLE_MUTED)

        self._prep_thread = SpecExcerptPrepThread(
            spec_num=self.spec_num,
            version=self.version,
            fname=self.fname,
            url=self.url,
            spec_dl_dir=self.spec_dl_dir,
        )
        self._prep_thread.progress.connect(lambda pct, msg: (self.progress_bar.setValue(pct), self.status_label.setText(msg)))
        self._prep_thread.prepared.connect(self._on_prep_finished)
        self._prep_thread.error.connect(self._on_prep_error)
        self._prep_thread.start()

    def _on_prep_finished(self, ordered_paths: list, headings: list):
        self._ordered_paths = list(ordered_paths)
        self._headings = list(headings)

        parts_count = len(self._ordered_paths)
        parts_text = f"{parts_count} part" if parts_count == 1 else f"{parts_count} parts"
        self.status_badge.setText(f"📄 {parts_text} ({len(self._headings)} clauses)")
        self.status_badge.setStyleSheet(BADGE_STYLE_PRIMARY)

        self.progress_bar.setValue(100)
        self.status_label.setText("Ready: Select clauses to generate contribution.")
        self._populate_tree()

    def _on_prep_error(self, err_msg: str):
        self.progress_bar.setVisible(False)
        self.status_label.setText("⚠️ Specification preparation failed.")
        self.status_badge.setText("Failed")
        self.status_badge.setStyleSheet(f"color: {C.DANGER_ALT};")
        QMessageBox.critical(self, "Preparation Error", f"Could not prepare specification:\n{err_msg}")

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
        self._update_change_block_summary()

    def _on_tree_item_changed(self, item: QTreeWidgetItem, column: int):
        if self._is_updating_checks or column != 0:
            return

        self._is_updating_checks = True
        try:
            state = item.checkState(0)

            def set_child_states(parent_item):
                for i in range(parent_item.childCount()):
                    child = parent_item.child(i)
                    if child.data(0, Qt.UserRole):
                        child.setCheckState(0, state)
                    set_child_states(child)

            set_child_states(item)
        finally:
            self._is_updating_checks = False

        self._update_change_block_summary()

    def _filter_tree(self, text: str):
        query = text.strip().lower()

        def visit(item: QTreeWidgetItem) -> bool:
            own_match = not query or query in item.text(0).lower()
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

        self._update_change_block_summary()

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

    def _update_change_block_summary(self):
        selected_ids = self._get_selected_ids()
        count = len(selected_ids)
        self.btn_generate.setEnabled(count > 0)

        if count == 0:
            self.block_summary_label.setText("No clauses selected.")
            self.btn_generate.setText("🚀 Generate Contribution Document...")
            return

        blocks = group_selected_headings_into_blocks(self._headings, selected_ids, self._all_new_map)
        summary_str = extract_clause_summary(blocks)

        block_desc = []
        for b in blocks:
            lbl = "First change" if b.index == 0 else "Next change"
            clauses = ", ".join(h.text[:25] for h in b.headings[:2])
            if len(b.headings) > 2:
                clauses += f" (+{len(b.headings) - 2} more)"
            block_desc.append(f"<b>{lbl}:</b> {clauses}")

        info_text = f"<b>{len(blocks)} Change Section(s)</b> ({summary_str})<br>" + " | ".join(block_desc)
        self.block_summary_label.setText(info_text)

        btn_type = "pCR" if self.current_mode == "pcr" else "CR"
        self.btn_generate.setText(f"🚀 Generate {btn_type} ({len(blocks)} Section{'s' if len(blocks) != 1 else ''})...")
        self._update_preview_filename()

    # -----------------------------------------------------------------
    # Differentiated Export Naming Helpers
    # -----------------------------------------------------------------

    def _on_naming_scheme_changed(self):
        scheme = self.naming_combo.currentData()
        if scheme:
            WordConfig.set_cr_naming_scheme(scheme)
        self._custom_filename_active = False
        self._update_preview_filename(force=True)

    def _on_filename_text_edited(self, text: str):
        self._custom_filename_active = bool(text.strip())

    def _reset_filename_to_auto(self):
        self._custom_filename_active = False
        self._update_preview_filename(force=True)

    def _compute_auto_filename(self) -> str:
        scheme = self.naming_combo.currentData() if hasattr(self, "naming_combo") else "cr_standard"
        is_cr = self.current_mode == "cr"
        tdoc_id = self.input_tdoc_id.text().strip() if hasattr(self, "input_tdoc_id") else ""
        cr_num = self.input_cr_num.text().strip() if is_cr and hasattr(self, "input_cr_num") else ""
        cr_rev = self.input_cr_rev.text().strip() if is_cr and hasattr(self, "input_cr_rev") else ""
        work_item = self.input_work_item.text().strip() if hasattr(self, "input_work_item") else ""
        title = self.input_title.text().strip() if hasattr(self, "input_title") else ""
        selected_ids = self._get_selected_ids() if hasattr(self, "clause_tree") else []

        return generate_cr_pcr_export_name(
            mode=self.current_mode,
            spec_num=self.spec_num,
            spec_type=self.spec_type,
            version=self.version,
            tdoc_id=tdoc_id,
            cr_num=cr_num,
            cr_rev=cr_rev,
            work_item=work_item,
            title=title,
            headings=self._headings,
            selected_ids=selected_ids,
            scheme=scheme or "cr_standard",
            ext=".docx",
        )

    def _update_preview_filename(self, force: bool = False):
        if not hasattr(self, "filename_edit"):
            return
        if self._custom_filename_active and not force:
            return
        auto_name = self._compute_auto_filename()
        self.filename_edit.setText(auto_name)
