"""Settings persistence, favorites, WG scoping, and inline filter pill mixin."""

from __future__ import annotations

import json
import logging

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMenu, QPushButton

from core.ui.ui_components import DISMISSIBLE_TAG_STYLE, WG_DESCRIPTIONS
from ..constants import (
    ACTIVE_WG_ORDER,
    FULL_WG_ORDER,
    USEFUL_CT,
    USEFUL_RAN,
    USEFUL_SA,
    USEFUL_WGS,
)

logger = logging.getLogger(__name__)


class SpecsTabFiltersMixin:
    """Manages persisted filter preferences, favorites, WG scope toggles, and active filter chips."""

    def _load_settings(self):
        if self.config_file.exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.download_dir = data.get("download_dir", self.default_dl_dir)
                    self.saved_search = data.get("search_query", "")
                    self.saved_version = data.get("version_query", "")
                    self.saved_downloaded_only = data.get("downloaded_only", False)
                    self.favorites = set(data.get("favorites", []))
                    self.favorites_only = data.get("favorites_only", False)
                    self.copy_link_mode = data.get("copy_link_mode", "both")
                    self._show_all_wgs = data.get("show_all_wgs", False)
                    saved_filters = data.get("table_filters", {})
                    if isinstance(saved_filters, dict):
                        self.table_filters.update(saved_filters)
            except Exception as e:
                logger.error(f"Error loading specs_config.json: {e}", exc_info=True)
        else:
            self.download_dir = self.default_dl_dir
            self.copy_link_mode = "both"

    def _save_settings(self):
        try:
            data = {
                "download_dir": self.download_dir,
                "search_query": self.spec_search_input.text().strip(),
                "version_query": self.version_search_input.text().strip(),
                "downloaded_only": self.downloaded_only_cb.isChecked(),
                "favorites": sorted(list(self.favorites)),
                "favorites_only": self.favorites_only,
                "copy_link_mode": getattr(self, "copy_link_mode", "both"),
                "show_all_wgs": getattr(self, "_show_all_wgs", False),
                "table_filters": self.table_filters,
            }
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            logger.error(f"Error saving specs_config.json: {e}", exc_info=True)

    def _toggle_favorite(self, spec_num: str):
        """Toggle favorite state for a single specification."""
        if spec_num in self.favorites:
            self.favorites.remove(spec_num)
        else:
            self.favorites.add(spec_num)

        self._save_settings()
        self._update_chip_styles()
        self._update_active_tags()
        self.refresh_table()

    def _batch_set_favorites(self, spec_nums: list, add: bool):
        """Batch add or remove specifications from favorites."""
        for num in spec_nums:
            if add:
                self.favorites.add(num)
            else:
                self.favorites.discard(num)

        self._save_settings()
        self._update_chip_styles()
        self._update_active_tags()
        self.refresh_table()

    def _toggle_favorites_filter(self, checked: bool = False):
        """Toggle the favorites-only filter mode."""
        self.favorites_only = self.fav_chip.isChecked()
        self.save_settings_timer.start()
        self._update_chip_styles()
        self._update_active_tags()
        self.filter_timer.start()

    def _toggle_downloaded_filter(self, checked: bool = False):
        """Toggle the downloaded-only filter mode."""
        self.save_settings_timer.start()
        self._update_active_tags()
        self.filter_timer.start()

    def _on_search_changed(self):
        self.search_timer.start()
        self.save_settings_timer.start()

    def _sync_ui_to_filters(self):
        """Synchronize UI controls with internal self.table_filters, favorites, and downloaded states."""
        # 1. Align TSG and WG
        grp = self.table_filters.get("group", "")
        grp_list = [g.strip() for g in grp.split(",") if g.strip()] if grp else []

        if any(g not in USEFUL_WGS for g in grp_list) and not self._show_all_wgs:
            self._set_wg_scope(True)

        if len(grp_list) == 1 and grp_list[0] in ("SA", "RAN", "CT"):
            self._current_tsg = grp_list[0]
            self._populate_wg_combo()
            self.wg_combo.setCheckedItems([])
            self.table_filters["group"] = ""
        elif grp_list:
            tsgs = set()
            for g in grp_list:
                if g.startswith("SA"):
                    tsgs.add("SA")
                elif g.startswith("RAN") or g.startswith("GERAN"):
                    tsgs.add("RAN")
                elif g.startswith("CT") or g.startswith("CN"):
                    tsgs.add("CT")
            if len(tsgs) == 1:
                self._current_tsg = list(tsgs)[0]
            else:
                self._current_tsg = ""
            self._populate_wg_combo()
            self.wg_combo.setCheckedItems(grp_list)
        else:
            self._current_tsg = ""
            self._populate_wg_combo()
            self.wg_combo.setCheckedItems([])

        # 2. Align Series
        ser = self.table_filters.get("series", "")
        if ser:
            ser_list = [s.strip() for s in ser.split(",") if s.strip()]
            self.series_combo.setCheckedItems(ser_list)
        else:
            self.series_combo.setCheckedItems([])

        # 3. Align Tech
        tech = self.table_filters.get("tech", "")
        if tech and tech != "Any":
            tech_list = [t.strip() for t in tech.split(",") if t.strip()]
            self.tech_combo.setCheckedItems(tech_list)
        else:
            self.tech_combo.setCheckedItems([])

        # 4. Align Type
        stype = self.table_filters.get("spec_type", "Any")
        for t_val, btn in self.type_buttons.items():
            btn.setChecked(t_val == stype)

        # 5. Align Downloaded & Favorites
        self.downloaded_only_btn.setChecked(self.saved_downloaded_only)
        self.fav_chip.setChecked(self.favorites_only)

        self._populate_tag_filter_combo()
        self._update_chip_styles()
        self._update_active_tags()

    def _populate_wg_combo(self):
        all_groups = self.filter_options.get("groups", [])
        clean_groups = [g for g in all_groups if g and g != "-"]

        if not self._show_all_wgs:
            # --- ACTIVE / USEFUL WORKING GROUPS ONLY ---
            if self._current_tsg == "SA":
                target_groups = [g for g in clean_groups if g in USEFUL_SA]
            elif self._current_tsg == "RAN":
                target_groups = [g for g in clean_groups if g in USEFUL_RAN]
            elif self._current_tsg == "CT":
                target_groups = [g for g in clean_groups if g in USEFUL_CT]
            else:
                target_groups = [g for g in clean_groups if g in USEFUL_WGS]

            target_groups.sort(key=lambda g: ACTIVE_WG_ORDER.index(g) if g in ACTIVE_WG_ORDER else 999)
        else:
            # --- FULL LIST (INCLUDING LEGACY & DISBANDED GROUPS) ---
            if self._current_tsg == "SA":
                target_groups = [g for g in clean_groups if g.startswith("SA")]
            elif self._current_tsg == "RAN":
                target_groups = [g for g in clean_groups if g.startswith("RAN") or g.startswith("GERAN")]
            elif self._current_tsg == "CT":
                target_groups = [g for g in clean_groups if g.startswith("CT") or g.startswith("CN") or g in ("T2", "T3")]
            else:
                target_groups = [g for g in clean_groups]

            # In Full mode, active groups appear first (SA -> RAN -> CT), followed by legacy groups grouped by domain
            target_groups.sort(key=lambda g: FULL_WG_ORDER.index(g) if g in FULL_WG_ORDER else 999)

        if hasattr(self, "wg_scope_btn"):
            self.wg_scope_btn.blockSignals(True)
            self.wg_scope_btn.setChecked(self._show_all_wgs)
            if self._show_all_wgs:
                self.wg_scope_btn.setText(f"Full List ({len(target_groups)})")
                self.wg_scope_btn.setToolTip(
                    f"Currently showing full list of {len(target_groups)} Working Groups (including legacy/disbanded groups).\n"
                    "Click to limit to active/useful groups only."
                )
            else:
                self.wg_scope_btn.setText(f"Active WGs ({len(target_groups)})")
                self.wg_scope_btn.setToolTip(
                    f"Currently showing {len(target_groups)} active/useful 3GPP Working Groups (SA, RAN, CT).\n"
                    "Click to extend to full list including legacy groups (CN, GERAN, SMG, etc.)."
                )
            self.wg_scope_btn.blockSignals(False)

        self.wg_combo.blockSignals(True)
        self.wg_combo.updateItems(target_groups, item_tooltips=WG_DESCRIPTIONS)
        self.wg_combo.blockSignals(False)

    def _toggle_wg_scope(self):
        self._set_wg_scope(self.wg_scope_btn.isChecked())

    def _set_wg_scope(self, show_all: bool):
        self._show_all_wgs = show_all
        self._populate_wg_combo()
        self.save_settings_timer.start()

    def _show_wg_context_menu(self, pos):
        menu = QMenu(self)
        act_active = menu.addAction("⚡ Show Active / Useful Working Groups Only")
        act_active.setCheckable(True)
        act_active.setChecked(not self._show_all_wgs)
        act_full = menu.addAction("🌐 Show Full List (Include Legacy && Disbanded Groups)")
        act_full.setCheckable(True)
        act_full.setChecked(self._show_all_wgs)

        action = menu.exec_(self.wg_combo.mapToGlobal(pos))
        if action == act_active and self._show_all_wgs:
            self._set_wg_scope(False)
        elif action == act_full and not self._show_all_wgs:
            self._set_wg_scope(True)

    def _on_tsg_button_clicked(self, tsg_val: str):
        if self._current_tsg == tsg_val and tsg_val != "":
            self._current_tsg = ""
        else:
            self._current_tsg = tsg_val

        self._populate_wg_combo()
        self.wg_combo.setCheckedItems([])
        self.table_filters["group"] = ""
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _on_chip_clicked(self, tsg_val: str):
        """Backward compatibility wrapper for TSG chip clicks."""
        self._on_tsg_button_clicked(tsg_val)

    def _on_wg_filter_changed(self, checked_items: list):
        if self.wg_combo.isAllChecked() or not checked_items:
            self.table_filters["group"] = ""
        else:
            self.table_filters["group"] = ",".join(checked_items)

        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _on_wg_combo_changed(self):
        """Backward compatibility alias for wg combo changes."""
        self._on_wg_filter_changed(self.wg_combo.getCheckedItems())

    def _on_wg_badge_clicked(self, wg: str):
        """Filters the specification list by the clicked Working Group badge."""
        if not wg or wg == "-":
            return
        tsg = ""
        wg_upper = wg.upper()
        if wg_upper.startswith("SA"):
            tsg = "SA"
        elif wg_upper.startswith("RAN") or wg_upper.startswith("GERAN"):
            tsg = "RAN"
        elif wg_upper.startswith("CT") or wg_upper.startswith("CN"):
            tsg = "CT"

        if wg not in USEFUL_WGS and not self._show_all_wgs:
            self._set_wg_scope(True)

        if self._current_tsg != tsg and tsg != "":
            self._current_tsg = tsg
            self._populate_wg_combo()
            self._update_chip_styles()

        self.wg_combo.setCheckedItems([wg])
        self.table_filters["group"] = wg
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _on_series_filter_changed(self, checked_items: list):
        if self.series_combo.isAllChecked() or not checked_items:
            self.table_filters["series"] = ""
        else:
            self.table_filters["series"] = ",".join(checked_items)

        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _on_tech_filter_changed(self, checked_items: list):
        if self.tech_combo.isAllChecked() or not checked_items:
            self.table_filters["tech"] = ""
        else:
            self.table_filters["tech"] = ",".join(checked_items)

        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _on_tech_combo_changed(self):
        """Backward compatibility alias for tech combo changes."""
        self._on_tech_filter_changed(self.tech_combo.getCheckedItems())

    def _on_type_button_clicked(self, type_val: str):
        self.table_filters["spec_type"] = type_val
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _populate_tag_filter_combo(self):
        if not hasattr(self, "tag_filter_combo"):
            return
        current = self.tag_filter_combo.currentData() or ""
        self.tag_filter_combo.blockSignals(True)
        self.tag_filter_combo.clear()
        self.tag_filter_combo.addItem("All Tags", "")
        if self.db:
            unique_tags = self.db.get_all_unique_tags()
            for t in unique_tags:
                self.tag_filter_combo.addItem(f"🏷️ {t}", t)
        if current:
            idx = self.tag_filter_combo.findData(current)
            if idx >= 0:
                self.tag_filter_combo.setCurrentIndex(idx)
            else:
                self.tag_filter_combo.setCurrentIndex(0)
                self.tag_filter = None
        else:
            self.tag_filter_combo.setCurrentIndex(0)
        self.tag_filter_combo.blockSignals(False)

    def _on_tag_filter_changed(self):
        data = self.tag_filter_combo.currentData() or ""
        self.tag_filter = data if data else None
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _reset_tag_filter(self):
        self.tag_filter_combo.setCurrentIndex(0)
        self.tag_filter = None
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _update_chip_styles(self):
        fav_count = len(self.favorites)
        self.fav_chip.setText(f"⭐ Favorites ({fav_count})")
        self.fav_chip.setChecked(self.favorites_only)

        for tsg_val, btn in self.tsg_buttons.items():
            btn.setChecked(tsg_val == self._current_tsg)

        current_type = self.table_filters.get("spec_type", "Any")
        for t_val, btn in self.type_buttons.items():
            btn.setChecked(t_val == current_type)

    def _update_active_tags(self):
        """Re-render active filter tags dynamically with 1-click removal."""
        tags = []

        sq = self.spec_search_input.text().strip()
        if sq:
            disp_sq = sq if len(sq) <= 18 else f"{sq[:15]}..."
            tags.append((f'🔍 "{disp_sq}" ✕', lambda: (self.spec_search_input.clear(), self._on_search_changed())))

        vq = self.version_search_input.text().strip()
        if vq:
            tags.append((f'Ver: {vq} ✕', lambda: (self.version_search_input.clear(), self._on_search_changed())))

        if self.downloaded_only_btn.isChecked():
            tags.append(('💾 Cached ✕', lambda: self.downloaded_only_btn.setChecked(False)))

        if self.favorites_only:
            tags.append(('⭐ Favs ✕', lambda: self.fav_chip.setChecked(False)))

        if self.tag_filter:
            disp_tag = self.tag_filter if len(self.tag_filter) <= 15 else f"{self.tag_filter[:12]}..."
            tags.append((f'🏷️ {disp_tag} ✕', self._reset_tag_filter))

        if self._current_tsg:
            tags.append((f'TSG: {self._current_tsg} ✕', self._reset_tsg_filter))

        grp = self.table_filters.get("group", "")
        if grp:
            disp_grp = grp if len(grp) <= 15 else f"{grp[:12]}..."
            tags.append((f'WG: {disp_grp} ✕', self._reset_group_filter))

        ser = self.table_filters.get("series", "")
        if ser:
            disp_ser = ser if len(ser) <= 15 else f"{ser[:12]}..."
            tags.append((f'Series: {disp_ser} ✕', self._reset_series_filter))

        tech = self.table_filters.get("tech", "")
        if tech and tech != "Any":
            disp_tech = tech if len(tech) <= 15 else f"{tech[:12]}..."
            tags.append((f'Tech: {disp_tech} ✕', self._reset_tech_filter))

        stype = self.table_filters.get("spec_type", "Any")
        if stype and stype != "Any":
            tags.append((f'Type: {stype} ✕', self._reset_type_filter))

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
        self.active_tags_label.setVisible(has_tags)
        self.clear_all_btn.setVisible(has_tags)

        for label, callback in tags:
            tag_btn = QPushButton(label)
            tag_btn.setCursor(Qt.PointingHandCursor)
            tag_btn.setToolTip("Click to remove this filter constraint")
            tag_btn.setStyleSheet(DISMISSIBLE_TAG_STYLE)
            tag_btn.clicked.connect(callback)
            self.active_tags_container_layout.addWidget(tag_btn)

    def _reset_tsg_filter(self):
        self._current_tsg = ""
        self._populate_wg_combo()
        self.wg_combo.setCheckedItems([])
        self.table_filters["group"] = ""
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _reset_group_filter(self):
        self.wg_combo.setCheckedItems([])
        self.table_filters["group"] = ""
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _reset_series_filter(self):
        self.series_combo.setCheckedItems([])
        self.table_filters["series"] = ""
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _reset_tech_filter(self):
        self.tech_combo.setCheckedItems([])
        self.table_filters["tech"] = ""
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _reset_type_filter(self):
        self.table_filters["spec_type"] = "Any"
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _clear_all_filters(self):
        self._rendered_tag_labels = None
        self.spec_search_input.blockSignals(True)
        self.version_search_input.blockSignals(True)
        self.downloaded_only_btn.blockSignals(True)
        self.fav_chip.blockSignals(True)
        self.wg_combo.blockSignals(True)
        self.tech_combo.blockSignals(True)
        if hasattr(self, "tag_filter_combo"):
            self.tag_filter_combo.blockSignals(True)

        self.spec_search_input.clear()
        self.version_search_input.clear()
        self.downloaded_only_btn.setChecked(False)
        self.favorites_only = False
        self.fav_chip.setChecked(False)
        self.tag_filter = None
        if hasattr(self, "tag_filter_combo"):
            self.tag_filter_combo.setCurrentIndex(0)

        self._current_tsg = ""
        self._populate_wg_combo()
        self.wg_combo.setCheckedItems([])

        self.series_combo.setCheckedItems([])
        self.tech_combo.setCheckedItems([])

        for t_val, btn in self.type_buttons.items():
            btn.setChecked(t_val == "Any")

        self.table_filters = {"series": "", "tech": "", "group": "", "spec_type": "Any"}

        self.spec_search_input.blockSignals(False)
        self.version_search_input.blockSignals(False)
        self.downloaded_only_btn.blockSignals(False)
        self.fav_chip.blockSignals(False)
        self.wg_combo.blockSignals(False)
        self.tech_combo.blockSignals(False)
        if hasattr(self, "tag_filter_combo"):
            self.tag_filter_combo.blockSignals(False)

        self._update_chip_styles()
        self._update_active_tags()
        self._save_settings()
        self.refresh_table()

    def _clear_table_filters(self):
        self._clear_all_filters()

    def _update_filter_ui(self):
        self._update_active_tags()
