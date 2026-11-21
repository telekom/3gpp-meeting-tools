"""Filter persistence, dismissible tags, TSG scope, and star/follow tracking mixin."""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QPushButton

from core.ui.ui_components import DISMISSIBLE_TAG_STYLE


class WorkItemsFiltersTrackingMixin:
    """Filter state synchronization, active tag chips, and tracking toggles."""

    def _save_filters(self):
        filters = {
            "search": self.search_input.text().strip(),
            "tsg": self._current_tsg,
            "releases": self.release_combo.getCheckedItems(),
            "wgs": self.wg_combo.getCheckedItems(),
            "status": self.status_combo.currentData() or "active",
            "starred_only": self.starred_only,
            "following_only": self.following_only,
        }
        self.settings.save_filters(filters)

    def _load_filters(self):
        filters = self.settings.get_filters()
        self.search_input.blockSignals(True)
        self.release_combo.blockSignals(True)
        self.wg_combo.blockSignals(True)
        self.status_combo.blockSignals(True)
        try:
            if not filters:
                idx = self.status_combo.findData("active")
                if idx >= 0:
                    self.status_combo.setCurrentIndex(idx)
                return

            if "search" in filters:
                self.search_input.setText(filters["search"])

            if "tsg" in filters:
                self._current_tsg = filters["tsg"]

            if "releases" in filters:
                self._apply_checked_items(self.release_combo, filters["releases"])

            if "wgs" in filters:
                self._apply_checked_items(self.wg_combo, filters["wgs"])

            if "status" in filters:
                idx = self.status_combo.findData(filters["status"])
                if idx >= 0:
                    self.status_combo.setCurrentIndex(idx)
            else:
                idx = self.status_combo.findData("active")
                if idx >= 0:
                    self.status_combo.setCurrentIndex(idx)

            if "starred_only" in filters:
                self.starred_only = bool(filters["starred_only"])
            if "following_only" in filters:
                self.following_only = bool(filters["following_only"])
        finally:
            self.search_input.blockSignals(False)
            self.release_combo.blockSignals(False)
            self.wg_combo.blockSignals(False)
            self.status_combo.blockSignals(False)

    def _apply_checked_items(self, combo, items_to_check):
        model = combo.model()
        row_count = model.rowCount()
        if row_count == 0:
            return

        model.blockSignals(True)
        try:
            if not items_to_check:
                for row in range(row_count):
                    if hasattr(model, 'item'):
                        item = model.item(row)
                        if item:
                            item.setCheckState(Qt.Checked)
                    else:
                        index = model.index(row, 0)
                        model.setData(index, Qt.Checked, Qt.CheckStateRole)
            else:
                items_to_check_stripped = {str(x).strip() for x in items_to_check}
                all_checked = True
                for row in range(1, row_count):
                    if hasattr(model, 'item'):
                        item = model.item(row)
                        if item:
                            user_val = str(item.data(Qt.UserRole) or item.text()).strip()
                            checked = (user_val in items_to_check_stripped or item.text().strip() in items_to_check_stripped)
                            item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
                            if not checked:
                                all_checked = False
                    else:
                        index = model.index(row, 0)
                        text = str(model.data(index, Qt.DisplayRole)).strip()
                        checked = text in items_to_check_stripped
                        model.setData(index, Qt.Checked if checked else Qt.Unchecked, Qt.CheckStateRole)
                        if not checked:
                            all_checked = False

                if hasattr(model, 'item'):
                    root_item = model.item(0)
                    if root_item:
                        root_item.setCheckState(Qt.Checked if all_checked else Qt.Unchecked)
                else:
                    root_index = model.index(0, 0)
                    model.setData(root_index, Qt.Checked if all_checked else Qt.Unchecked, Qt.CheckStateRole)
        finally:
            model.blockSignals(False)

        if hasattr(combo, 'updateText'):
            combo.updateText()
        elif hasattr(combo, 'update'):
            combo.update()
        elif hasattr(combo, 'repaint'):
            combo.repaint()

    def _on_tsg_clicked(self, tsg: str):
        if tsg == "All":
            self._current_tsg = ""
        else:
            self._current_tsg = tsg

        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _update_active_tags(self):
        """Re-render active filter tags dynamically with 1-click removal."""
        tags = []

        query = self.search_input.text().strip()
        if query:
            disp_q = query if len(query) <= 18 else f"{query[:15]}..."
            tags.append((f'🔍 "{disp_q}" ✕', lambda: self.search_input.clear()))

        if self._current_tsg:
            tags.append((f'TSG: {self._current_tsg} ✕', lambda: self._on_tsg_clicked("All")))

        if hasattr(self.wg_combo, "isFiltered") and self.wg_combo.isFiltered():
            selected_wgs = self.wg_combo.getCheckedItems()
            if selected_wgs:
                w_str = ", ".join(selected_wgs)
                disp_w = w_str if len(w_str) <= 18 else f"{w_str[:15]}..."
                tags.append((f'WG: {disp_w} ✕', lambda: (self._apply_checked_items(self.wg_combo, []), self.refresh_table(True))))

        if hasattr(self.release_combo, "isFiltered") and self.release_combo.isFiltered():
            selected_rels = self.release_combo.getCheckedItems()
            if selected_rels:
                r_str = ", ".join(selected_rels)
                disp_r = r_str if len(r_str) <= 18 else f"{r_str[:15]}..."
                tags.append((f'Rel: {disp_r} ✕', lambda: (self._apply_checked_items(self.release_combo, []), self.refresh_table(True))))

        cur_status = self.status_combo.currentData()
        if cur_status and cur_status != "active":
            status_name = "Finished" if cur_status == "finished" else "All Statuses"
            tags.append((f'Status: {status_name} ✕', lambda: self.status_combo.setCurrentIndex(0)))

        if self.following_only:
            tags.append(('🎯 Following ✕', self._toggle_following_filter))

        if self.starred_only:
            tags.append(('⭐ Starred ✕', self._toggle_starred_filter))

        tag_labels = [label for label, _ in tags]
        if getattr(self, "_rendered_tag_labels", None) == tag_labels:
            return
        self._rendered_tag_labels = tag_labels

        while self.active_tags_container_layout.count():
            item = self.active_tags_container_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        has_tags = len(tags) > 0
        if hasattr(self, "clear_all_btn"):
            self.clear_all_btn.setVisible(has_tags)

        for label, callback in tags:
            tag_btn = QPushButton(label)
            tag_btn.setCursor(Qt.PointingHandCursor)
            tag_btn.setToolTip("Click to remove this filter constraint")
            tag_btn.setStyleSheet(DISMISSIBLE_TAG_STYLE)
            tag_btn.clicked.connect(callback)
            self.active_tags_container_layout.addWidget(tag_btn)

    def _toggle_starred_filter(self):
        self.starred_only = not self.starred_only
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _toggle_following_filter(self):
        self.following_only = not self.following_only
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _update_chip_styles(self):
        starred_count = len(self.starred)
        following_count = len(self.following)
        tracked_count = len(self.starred | self.following)

        self.starred_chip.setText(f"⭐ Starred ({starred_count})")
        self.following_chip.setText(f"🎯 Following ({following_count})")
        if hasattr(self, "sync_tracked_btn"):
            self.sync_tracked_btn.setText(f"🎯 Update Tracked ({tracked_count})")
            self.sync_tracked_btn.setEnabled(tracked_count > 0)

        self.starred_chip.setChecked(self.starred_only)
        self.following_chip.setChecked(self.following_only)

        if hasattr(self, "tsg_buttons"):
            for tsg_val, btn in self.tsg_buttons.items():
                btn.setChecked((tsg_val == "All" and not self._current_tsg) or (tsg_val == self._current_tsg))

    def _get_active_filter_codes(self):
        """Returns list of WI codes if filter chips are active, or None if no tracking filter is active."""
        if not self.starred_only and not self.following_only:
            return None
        codes = set()
        if self.starred_only:
            codes |= self.starred
        if self.following_only:
            codes |= self.following
        return sorted(list(codes))

    def _toggle_star(self, wi_code: str):
        code = str(wi_code).strip()
        if not code:
            return
        if code in self.starred:
            self.starred.remove(code)
        else:
            self.starred.add(code)
        self.settings.save_tracking_sets(self.starred, self.following)
        self.table_model.set_tracking_sets(self.starred, self.following)
        self._update_chip_styles()

        if self.starred_only:
            self.refresh_table(reset_pagination=True)
        else:
            self.table_model.refresh_row_by_code(code)

    def _toggle_follow(self, wi_code: str):
        code = str(wi_code).strip()
        if not code:
            return
        if code in self.following:
            self.following.remove(code)
        else:
            self.following.add(code)
        self.settings.save_tracking_sets(self.starred, self.following)
        self.table_model.set_tracking_sets(self.starred, self.following)
        self._update_chip_styles()

        if self.following_only:
            self.refresh_table(reset_pagination=True)
        else:
            self.table_model.refresh_row_by_code(code)

    def _batch_set_starred(self, wi_codes: list, add: bool):
        for c in wi_codes:
            code = str(c).strip()
            if not code:
                continue
            if add:
                self.starred.add(code)
            else:
                self.starred.discard(code)
        self.settings.save_tracking_sets(self.starred, self.following)
        self.table_model.set_tracking_sets(self.starred, self.following)
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _batch_set_following(self, wi_codes: list, add: bool):
        for c in wi_codes:
            code = str(c).strip()
            if not code:
                continue
            if add:
                self.following.add(code)
            else:
                self.following.discard(code)
        self.settings.save_tracking_sets(self.starred, self.following)
        self.table_model.set_tracking_sets(self.starred, self.following)
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def _clear_all_filters(self):
        self.search_input.clear()
        self._current_tsg = ""
        self._apply_checked_items(self.release_combo, [])
        self._apply_checked_items(self.wg_combo, [])
        self.status_combo.setCurrentIndex(0)
        self.starred_only = False
        self.following_only = False
        self._update_chip_styles()
        self._update_active_tags()
        self.refresh_table(reset_pagination=True)
