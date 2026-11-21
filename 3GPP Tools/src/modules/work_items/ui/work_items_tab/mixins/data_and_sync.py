"""Database pagination, deep-link focusing, and portal scraper sync mixin."""

import logging

from PyQt5.QtWidgets import QMessageBox

from core.ui.toast import notify_info, notify_success
from modules.work_items.core.wi_scraper import (
    TargetedWIScraperThread,
    WorkItemsScraperThread,
)

logger = logging.getLogger("modules.work_items.ui.ui_tabs")


class WorkItemsDataSyncMixin:
    """Incremental chunk loading, DB search, deep linking, and WI scraping."""

    def _sync_tracked_wis(self):
        tracked = sorted(list(self.starred | self.following))
        if not tracked:
            notify_info(
                self,
                "No Tracked Work Items",
                "Star or Follow Work Items first using the Track column or right-click menu.",
            )
            return
        self._start_targeted_sync(tracked)

    def ensure_loaded(self):
        """Ensure Work Items table is populated."""
        if self.table_model.rowCount() == 0:
            self.refresh_table(reset_pagination=True)

    def focus_work_item(self, query: str):
        """Search and select a Work Item, broadening status/tracking filters if hidden."""
        clean = (query or "").strip()
        if not clean:
            return

        self.search_timer.stop()
        self.search_input.blockSignals(True)
        self.search_input.setText(clean)
        self.search_input.blockSignals(False)
        self.refresh_table(reset_pagination=True)

        if self.table_model.rowCount() == 0:
            self._clear_all_filters()
            all_idx = self.status_combo.findData("all")
            if all_idx >= 0:
                self.status_combo.blockSignals(True)
                self.status_combo.setCurrentIndex(all_idx)
                self.status_combo.blockSignals(False)
            self.search_input.blockSignals(True)
            self.search_input.setText(clean)
            self.search_input.blockSignals(False)
            self.refresh_table(reset_pagination=True)

        if self.table_model.rowCount() > 0:
            self.table.selectRow(0)
            self.table.scrollTo(self.table_model.index(0, 0))

    def _on_scroll(self, value):
        max_val = self.table.verticalScrollBar().maximum()
        if value >= max_val - 5 and not self._is_loading_chunk:
            if len(self._loaded_data) < self._total_count:
                self._load_next_chunk()

    def _log_and_forward_action(self, text, action):
        self.global_action_requested.emit(text, action)

    def _populate_filters(self):
        options = self.db.get_filter_options()

        self.release_combo.blockSignals(True)
        self.release_combo.updateItems(options.get('releases', []))
        self.release_combo.blockSignals(False)

        self.wg_combo.blockSignals(True)
        self.wg_combo.updateItems(options.get('groups', []))
        self.wg_combo.blockSignals(False)

    def _resolve_effective_wgs(self) -> list:
        if hasattr(self.wg_combo, "isFiltered") and self.wg_combo.isFiltered():
            return self.wg_combo.getCheckedItems()
        elif self._current_tsg:
            all_wgs = self.db.get_filter_options().get('groups', [])
            if self._current_tsg == "SA":
                return [g for g in all_wgs if g.startswith("SA")]
            elif self._current_tsg == "RAN":
                return [g for g in all_wgs if g.startswith("RAN") or g.startswith("GERAN")]
            elif self._current_tsg == "CT":
                return [g for g in all_wgs if g.startswith("CT") or g.startswith("CN")]
        return []

    def refresh_table(self, reset_pagination: bool = True):
        self.save_filters_timer.start()

        if reset_pagination:
            self._current_offset = 0
            self._loaded_data = []

        search_term = self.search_input.text().strip()
        selected_releases = self.release_combo.getCheckedItems() if (hasattr(self.release_combo, "isFiltered") and self.release_combo.isFiltered()) else []
        selected_wgs = self._resolve_effective_wgs()
        selected_status = self.status_combo.currentData() or "active"
        filter_codes = self._get_active_filter_codes()

        self._total_count = self.db.count_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status,
            wi_codes=filter_codes
        )

        chunk = self.db.search_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status,
            limit=self._chunk_size,
            offset=self._current_offset,
            wi_codes=filter_codes
        )

        self._loaded_data = chunk
        self.table_model.set_search_term(search_term)
        self.table_model.update_data(self._loaded_data)
        self._update_count_label()
        self._update_chip_styles()
        self._update_active_tags()

    def _load_next_chunk(self):
        self._is_loading_chunk = True
        self._current_offset += self._chunk_size

        search_term = self.search_input.text().strip()
        selected_releases = self.release_combo.getCheckedItems() if (hasattr(self.release_combo, "isFiltered") and self.release_combo.isFiltered()) else []
        selected_wgs = self._resolve_effective_wgs()
        selected_status = self.status_combo.currentData() or "active"
        filter_codes = self._get_active_filter_codes()

        chunk = self.db.search_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status,
            limit=self._chunk_size,
            offset=self._current_offset,
            wi_codes=filter_codes
        )

        if chunk:
            self.table_model.append_data(chunk)
            self._loaded_data.extend(chunk)
            self._update_count_label()

        self._is_loading_chunk = False

    def _update_count_label(self):
        loaded = len(self._loaded_data)
        query = self.search_input.text().strip()
        exact_count = 0
        if query:
            exact_count = sum(1 for r in self._loaded_data if str(r.get("acronym") or "").strip().casefold() == query.casefold())

        prefix = f'{self._total_count} matches for "{query}"' if query else f"{self._total_count} Work Items"
        if query and exact_count:
            prefix += f" · Exact acronym: {exact_count}"
        if self._total_count > loaded:
            self.count_label.setText(f"{prefix} · showing {loaded} (scroll for more)")
        else:
            self.count_label.setText(prefix)

    def _start_sync(self):
        logger.info("🔄 [Work Items] Starting full Work Items synchronization.")
        self.sync_btn.setEnabled(False)
        self.sync_btn.setText("⏳ Syncing...")
        if hasattr(self, "add_wi_btn"):
            self.add_wi_btn.setEnabled(False)
        if hasattr(self, "sync_tracked_btn"):
            self.sync_tracked_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.scraper_thread = WorkItemsScraperThread(self.db_path, self)
        self.scraper_thread.progress.connect(self._update_progress)
        self.scraper_thread.finished_sync.connect(self._on_sync_finished)
        self.scraper_thread.start()

    def _update_progress(self, current: int, total: int, msg: str):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_lbl.setText(msg)

    def _on_sync_finished(self, success: bool, msg: str):
        if success:
            logger.info(f"✅ [Work Items] {msg}")
        else:
            logger.error(f"❌ [Work Items] {msg}")

        self.sync_btn.setEnabled(True)
        self.sync_btn.setText("🔄 Sync 3GPP WIs")
        if hasattr(self, "add_wi_btn"):
            self.add_wi_btn.setEnabled(True)
        if hasattr(self, "sync_tracked_btn"):
            self.sync_tracked_btn.setEnabled(len(self.starred | self.following) > 0)
        self.progress_bar.setVisible(False)
        self.status_lbl.setText("")

        prev_releases = self.release_combo.getCheckedItems()
        prev_wgs = self.wg_combo.getCheckedItems()
        self._populate_filters()
        if prev_releases:
            self.release_combo.blockSignals(True)
            self._apply_checked_items(self.release_combo, prev_releases)
            self.release_combo.blockSignals(False)
        if prev_wgs:
            self.wg_combo.blockSignals(True)
            self._apply_checked_items(self.wg_combo, prev_wgs)
            self.wg_combo.blockSignals(False)
        self.refresh_table(reset_pagination=True)

        if success:
            notify_success(self, "Sync Complete", msg)
        else:
            QMessageBox.warning(self, "Sync Failed", msg)

    def _start_targeted_sync(self, wi_codes: list):
        logger.info(f"🔄 [Work Items] Starting targeted update for {len(wi_codes)} WI(s): {wi_codes}")
        self.sync_btn.setEnabled(False)
        self.sync_btn.setText("⏳ Updating...")
        if hasattr(self, "add_wi_btn"):
            self.add_wi_btn.setEnabled(False)
        if hasattr(self, "sync_tracked_btn"):
            self.sync_tracked_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.targeted_thread = TargetedWIScraperThread(self.db_path, wi_codes, self)
        self.targeted_thread.progress.connect(self._update_progress)
        self.targeted_thread.finished_sync.connect(self._on_sync_finished)
        self.targeted_thread.start()
