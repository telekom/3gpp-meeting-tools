"""TSG/WG/date filter persistence, table refresh, caching probes, and focus navigation."""

from pathlib import Path

from PyQt5.QtCore import QDate, Qt

from core.ui.theme import C


class MeetingsTabFiltersDataMixin:
    """Handles filter serialization, TSG/date presets, DB search, and table refresh."""

    def _save_filters(self):
        filters = {
            "wg": self.wg_filter.getCheckedItems(),
            "adhoc": self.adhoc_filter.currentText(),
            "type": self.type_filter.currentText(),
            "search": self.search_input.text().strip(),
            "enable_dates": self.enable_dates_cb.isChecked(),
            "date_from": self.date_from.date().toString("yyyy-MM-dd"),
            "date_to": self.date_to.date().toString("yyyy-MM-dd"),
            "date_preset": self.date_preset_combo.currentText(),
            "tsg": getattr(self, "_current_tsg", ""),
        }
        self.settings.save_filters(filters)

    def _load_filters(self):
        filters = self.settings.get_filters()
        if not filters:
            return

        self.wg_filter.blockSignals(True)
        self.adhoc_filter.blockSignals(True)
        self.type_filter.blockSignals(True)
        self.search_input.blockSignals(True)
        self.enable_dates_cb.blockSignals(True)
        self.date_from.blockSignals(True)
        self.date_to.blockSignals(True)
        self.date_preset_combo.blockSignals(True)

        if "tsg" in filters:
            self._current_tsg = filters["tsg"]
            self._update_tsg_button_styles()

        if "wg" in filters:
            saved_wgs = filters["wg"]
            if isinstance(saved_wgs, list):
                all_checked = True
                for i in range(1, self.wg_filter.model().rowCount()):
                    item = self.wg_filter.model().item(i)
                    if item.text() in saved_wgs:
                        item.setCheckState(Qt.Checked)
                    else:
                        item.setCheckState(Qt.Unchecked)
                        all_checked = False
                root_item = self.wg_filter.model().item(0)
                if root_item is not None and self.wg_filter.model().rowCount() > 1:
                    root_item.setCheckState(Qt.Checked if all_checked else Qt.Unchecked)
                self.wg_filter.updateText()

        if "adhoc" in filters:
            idx = self.adhoc_filter.findText(filters["adhoc"])
            if idx >= 0:
                self.adhoc_filter.setCurrentIndex(idx)

        if "type" in filters:
            idx = self.type_filter.findText(filters["type"])
            if idx >= 0:
                self.type_filter.setCurrentIndex(idx)

        if "search" in filters:
            self.search_input.setText(filters["search"])

        if "date_preset" in filters:
            preset_name = filters["date_preset"]
            idx = self.date_preset_combo.findText(preset_name)
            if idx >= 0:
                self.date_preset_combo.setCurrentIndex(idx)
                is_custom = (preset_name == "Custom Range...")
                self.date_from.setVisible(is_custom)
                self.arrow_lbl.setVisible(is_custom)
                self.date_to.setVisible(is_custom)

        if "enable_dates" in filters:
            use_dates = filters["enable_dates"] and self.date_preset_combo.currentText() not in (
                "All Dates",
                "Unknown Dates",
            )
            self.enable_dates_cb.setChecked(use_dates)
            self._toggle_date_inputs(use_dates)

        if "date_from" in filters:
            d = QDate.fromString(filters["date_from"], "yyyy-MM-dd")
            if d.isValid():
                self.date_from.setDate(d)

        if "date_to" in filters:
            d = QDate.fromString(filters["date_to"], "yyyy-MM-dd")
            if d.isValid():
                self.date_to.setDate(d)

        self.wg_filter.blockSignals(False)
        self.adhoc_filter.blockSignals(False)
        self.type_filter.blockSignals(False)
        self.search_input.blockSignals(False)
        self.enable_dates_cb.blockSignals(False)
        self.date_from.blockSignals(False)
        self.date_to.blockSignals(False)
        self.date_preset_combo.blockSignals(False)

    def _on_tsg_button_clicked(self, tsg_val: str):
        if getattr(self, "_current_tsg", "") == tsg_val and tsg_val != "":
            self._current_tsg = ""
        else:
            self._current_tsg = tsg_val

        self._update_tsg_button_styles()

        model = self.wg_filter.model()
        self.wg_filter.blockSignals(True)
        try:
            for i in range(1, model.rowCount()):
                item = model.item(i)
                wg_name = item.text().upper()
                if not self._current_tsg:
                    item.setCheckState(Qt.Unchecked)
                elif self._current_tsg == "SA" and wg_name.startswith("SA"):
                    item.setCheckState(Qt.Checked)
                elif self._current_tsg == "RAN" and (wg_name.startswith("RAN") or wg_name.startswith("GERAN")):
                    item.setCheckState(Qt.Checked)
                elif self._current_tsg == "CT" and (wg_name.startswith("CT") or wg_name.startswith("CN")):
                    item.setCheckState(Qt.Checked)
                else:
                    item.setCheckState(Qt.Unchecked)
            self.wg_filter.updateText()
        finally:
            self.wg_filter.blockSignals(False)

        self.refresh_table()

    def _update_tsg_button_styles(self):
        curr = getattr(self, "_current_tsg", "")
        for t_val, btn in getattr(self, "tsg_buttons", {}).items():
            btn.setChecked(curr == t_val)

    def _on_date_preset_changed(self):
        preset = self.date_preset_combo.currentText()
        today = QDate.currentDate()
        self.enable_dates_cb.blockSignals(True)
        self.date_from.blockSignals(True)
        self.date_to.blockSignals(True)

        is_custom = (preset == "Custom Range...")
        self.date_from.setVisible(is_custom)
        self.arrow_lbl.setVisible(is_custom)
        self.date_to.setVisible(is_custom)

        try:
            if preset == "All Dates":
                self.enable_dates_cb.setChecked(False)
                self.date_from.setEnabled(False)
                self.date_to.setEnabled(False)
            elif preset == "Upcoming":
                self.enable_dates_cb.setChecked(True)
                self.date_from.setDate(today)
                self.date_to.setDate(today.addYears(3))
                self.date_from.setEnabled(False)
                self.date_to.setEnabled(False)
            elif preset == "This Year":
                self.enable_dates_cb.setChecked(True)
                self.date_from.setDate(QDate(today.year(), 1, 1))
                self.date_to.setDate(QDate(today.year(), 12, 31))
                self.date_from.setEnabled(False)
                self.date_to.setEnabled(False)
            elif preset == "Past 12 Months":
                self.enable_dates_cb.setChecked(True)
                self.date_from.setDate(today.addYears(-1))
                self.date_to.setDate(today)
                self.date_from.setEnabled(False)
                self.date_to.setEnabled(False)
            elif preset == "Custom Range...":
                self.enable_dates_cb.setChecked(True)
                self.date_from.setEnabled(True)
                self.date_to.setEnabled(True)
            elif preset == "Unknown Dates":
                self.enable_dates_cb.setChecked(False)
                self.date_from.setEnabled(False)
                self.date_to.setEnabled(False)
        finally:
            self.enable_dates_cb.blockSignals(False)
            self.date_from.blockSignals(False)
            self.date_to.blockSignals(False)

        self.refresh_table()

    def _reset_filters(self):
        self._current_tsg = ""
        self._update_tsg_button_styles()
        self.search_input.clear()
        self.adhoc_filter.setCurrentIndex(0)
        self.type_filter.setCurrentIndex(0)
        self.date_preset_combo.setCurrentIndex(0)
        self.date_from.setVisible(False)
        self.arrow_lbl.setVisible(False)
        self.date_to.setVisible(False)

        model = self.wg_filter.model()
        self.wg_filter.blockSignals(True)
        try:
            for i in range(model.rowCount()):
                item = model.item(i)
                if item is not None:
                    item.setCheckState(Qt.Checked)
            self.wg_filter.updateText()
        finally:
            self.wg_filter.blockSignals(False)

        self.refresh_table()

    def ensure_loaded(self):
        """Ensure initial meetings table state is loaded (safe idempotent hook)."""
        if self.table_model.rowCount() == 0:
            self.refresh_table()

    def focus_meeting(self, query_or_tdoc: str):
        """Focus a meeting or TDoc, broadening active WG/date filters if needed."""
        import re

        clean = (query_or_tdoc or "").strip()
        if not clean:
            return

        if re.match(r"^[A-Z0-9]{2,4}-\d{6,8}$", clean, re.IGNORECASE):
            self.global_tdoc_input.setText(clean)
            if hasattr(self, "search_controller") and hasattr(
                self.search_controller, "on_tdoc_input_changed"
            ):
                self.search_controller.on_tdoc_input_changed(clean)
            owning = self.db.find_meeting_by_tdoc(clean)
            if owning and owning.get("meeting_number"):
                clean = str(owning["meeting_number"])

        self.search_input.blockSignals(True)
        self.search_input.setText(clean)
        self.search_input.blockSignals(False)
        self.refresh_table()

        if self.table_model.rowCount() == 0:
            self.enable_dates_cb.blockSignals(True)
            self.enable_dates_cb.setChecked(False)
            self.enable_dates_cb.blockSignals(False)
            self._reset_filters()
            self.search_input.blockSignals(True)
            self.search_input.setText(clean)
            self.search_input.blockSignals(False)
            self.refresh_table()

        if self.table_model.rowCount() == 0:
            unk_idx = self.date_preset_combo.findText("Unknown Dates")
            if unk_idx >= 0:
                self.date_preset_combo.setCurrentIndex(unk_idx)
                if self.table_model.rowCount() == 0:
                    self.date_preset_combo.setCurrentIndex(0)

        if self.table_model.rowCount() > 0:
            self.table.selectRow(0)
            self.table.scrollTo(self.table_model.index(0, 0))

    def _toggle_date_inputs(self, checked):
        self.date_from.setEnabled(checked)
        self.date_to.setEnabled(checked)

    def _populate_filters(self):
        selected_wgs = set(self.wg_filter.getCheckedItems())
        wgs = self.db.get_working_groups()

        self.wg_filter.blockSignals(True)
        try:
            self.wg_filter.updateItems(wgs)
            model = self.wg_filter.model()
            for i in range(1, model.rowCount()):
                item = model.item(i)
                item.setCheckState(Qt.Checked if item.text() in selected_wgs else Qt.Unchecked)
            self.wg_filter.updateText()
        finally:
            self.wg_filter.blockSignals(False)

    def refresh_table(self):
        self.save_filters_timer.start()

        selected_wgs = self.wg_filter.getCheckedItems()
        preset = self.date_preset_combo.currentText() if hasattr(self, "date_preset_combo") else "All Dates"
        date_mode = "unknown" if preset == "Unknown Dates" else "dated"
        date_from = self.date_from.date().toString("yyyy-MM-dd") if self.enable_dates_cb.isChecked() else None
        date_to = self.date_to.date().toString("yyyy-MM-dd") if self.enable_dates_cb.isChecked() else None
        adhoc_val = self.adhoc_filter.currentText()
        type_val = self.type_filter.currentText()
        search_val = self.search_input.text().strip()

        data = self.db.search_meetings(
            wg_name=selected_wgs, search_term=search_val,
            location=None, date_from=date_from, date_to=date_to,
            adhoc_filter=adhoc_val, type_filter=type_val,
            date_mode=date_mode,
        )

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.settings.cache_dir
        cache_path = Path(current_cache)
        cached_folders = set(f.name for f in cache_path.iterdir() if f.is_dir()) if cache_path.exists() else set()

        cached_count = 0
        live_count = 0
        today = QDate.currentDate()

        for row in data:
            mtg_id = row.get("mtg_id")
            s = row.get("start_date", "")
            e = row.get("end_date", "")
            if s and e:
                try:
                    d1 = QDate.fromString(s, "yyyy-MM-dd")
                    d2 = QDate.fromString(e, "yyyy-MM-dd")
                    if d1.isValid() and d2.isValid() and d1 <= today <= d2:
                        live_count += 1
                except Exception:
                    pass

            if not mtg_id:
                row['tdoc_btn_status'] = 'na'
                continue

            folder_name = row.get("folder_name") or row.get("meeting_number", "")
            filepath = None

            if folder_name in cached_folders:
                agenda_dir = cache_path / folder_name / "Agenda"
                if agenda_dir.exists():
                    filepath = next((f for f in agenda_dir.iterdir() if (
                                "tdoc_list_meeting_" in f.name.lower() or "tdocs_list_" in f.name.lower()) and f.name.endswith(
                        ".xlsx")), None)
                    if not filepath:
                        fallback = agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"
                        filepath = fallback if fallback.exists() else None

            if filepath:
                row['tdoc_btn_status'] = 'open'
                row['tdoc_filepath'] = str(filepath)
                cached_count += 1
            else:
                row['tdoc_btn_status'] = 'get'

        self.table_model.update_data(data)

        # Update status bar summary
        total_count = len(data)
        summary_parts = [f"Showing <b>{total_count}</b> meeting{'s' if total_count != 1 else ''}"]
        if cached_count > 0:
            summary_parts.append(f"<b>{cached_count}</b> cached")
        if live_count > 0:
            summary_parts.append(f"<span style='color:{C.EMERALD_PRIMARY}; font-weight:bold;'>🟢 {live_count} live</span>")
        if date_mode != "unknown" and hasattr(self.db, "count_unknown_date_meetings"):
            unknown_count = self.db.count_unknown_date_meetings(
                wg_name=selected_wgs,
                search_term=search_val,
                adhoc_filter=adhoc_val,
                type_filter=type_val,
            )
            if unknown_count > 0:
                summary_parts.append(
                    f"<span style='color:{C.TEXT_MUTED};'>{unknown_count} undated (Unknown Dates)</span>"
                )

        self.status_lbl.setText(" · ".join(summary_parts))
