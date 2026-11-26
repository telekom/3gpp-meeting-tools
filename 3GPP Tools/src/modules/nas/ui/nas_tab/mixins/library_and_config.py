"""Protocol library selection, config persistence, ribbon KPI, and message banner mixin."""

from __future__ import annotations

import json
import logging
import re
from typing import Any, Dict, List

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QAction

from modules.nas.core.nas_exporter import copy_matrix_to_clipboard
from modules.nas.ui.nas_ingest_controller import (
    delete_single_version,
    delete_spec_group,
    reload_checked_versions,
    reload_single_version,
    reload_spec_group,
)
from ..styles import RE_CTX_LCH, RE_CTX_RLC, RE_CTX_SRB, RE_CTX_TRAILING_CAPTION

logger = logging.getLogger(__name__)


class NASTabLibraryAndConfigMixin:
    """Handles version cards refresh, config save/load, context banners, and matrix copy actions."""

    def _on_select_all_specs(self):
        self.version_tree.select_all()

    def _on_select_latest_specs(self):
        self.version_tree.select_latest_only()

    def _on_select_none_specs(self):
        self.version_tree.select_none()

    def _copy_matrix_formatted(self, mode: str = "word"):
        """Exports the active evolution matrix to the clipboard in Word, Markdown, or TSV format."""
        model = self.matrix_table.model()
        if not model or model.rowCount() == 0:
            return

        visible_cols = [
            c for c in range(model.columnCount())
            if not self.matrix_table.isColumnHidden(c)
        ]
        if not visible_cols:
            visible_cols = list(range(model.columnCount()))

        headers = [
            str(model.headerData(c, Qt.Horizontal, Qt.DisplayRole) or "")
            for c in visible_cols
        ]
        rows = [
            [
                str(model.data(model.index(r, c), Qt.DisplayRole) or "")
                for c in visible_cols
            ]
            for r in range(model.rowCount())
        ]

        # Extract diff colors if available
        diff_colors = []
        if hasattr(model, "_bg_brush_matrix"):
            for r in range(model.rowCount()):
                row_colors = []
                for c in visible_cols:
                    brush = (
                        model._bg_brush_matrix[r][c]
                        if r < len(model._bg_brush_matrix) and c < len(model._bg_brush_matrix[r])
                        else None
                    )
                    if brush and hasattr(brush, "color"):
                        row_colors.append(brush.color().name())
                    else:
                        row_colors.append(None)
                diff_colors.append(row_colors)

        title = self.matrix_title.text().strip() if hasattr(self, "matrix_title") else ""
        if not title:
            title = self.current_selected_message_name or self.current_selected_ie_key or "Protocol Evolution Matrix"

        copy_matrix_to_clipboard(
            title=title,
            headers=headers,
            rows=rows,
            diff_colors=diff_colors,
            mode=mode,
        )

        self.copy_matrix_btn.setText("✓ Copied!")
        QTimer.singleShot(1500, lambda: self.copy_matrix_btn.setText("📋 Copy Table ▾"))

    def _copy_matrix_tsv(self):
        """Legacy fallback calling TSV clipboard copy."""
        self._copy_matrix_formatted(mode="tsv")

    def _update_stats_ribbon(self):
        """Updates live KPI stats badges in the toolbar ribbon."""
        try:
            versions = self.db.get_imported_versions()
            unique_specs = len(set(v.get("spec_number") for v in versions if v.get("spec_number")))
            total_versions = len(versions)

            with self.db._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(DISTINCT message_name) FROM nas_messages")
                total_msgs = cursor.fetchone()[0] or 0
                cursor.execute("SELECT COUNT(DISTINCT ie_name) FROM ie_definitions")
                total_ies = cursor.fetchone()[0] or 0

            self.stats_specs_lbl.setText(f"📦 {unique_specs} Spec{'s' if unique_specs != 1 else ''}")
            self.stats_vers_lbl.setText(f"🏷️ {total_versions} Version{'s' if total_versions != 1 else ''}")
            self.stats_msgs_lbl.setText(f"📨 {total_msgs} Messages · {total_ies} IEs")
            if hasattr(self, "tree_stats_badge"):
                self.tree_stats_badge.setText(str(unique_specs))
                self.tree_stats_badge.setToolTip(
                    f"Indexed Protocols: {unique_specs} Spec(s) · {total_versions} Version(s)\n"
                    f"Library Contents: {total_msgs} Messages · {total_ies} Information Elements"
                )
        except Exception as e:
            logger.debug(f"Could not update stats ribbon: {e}")

    # -------------------------------------------------------------------------
    # Browse Mode & Configuration Persistence
    # -------------------------------------------------------------------------

    def _set_browse_mode(self, mode: str, trigger_populate: bool = True):
        mode = "ies" if mode == "ies" else "messages"
        self._browse_mode = mode

        self.mode_messages_btn.blockSignals(True)
        self.mode_ies_btn.blockSignals(True)
        self.mode_messages_btn.setChecked(mode == "messages")
        self.mode_ies_btn.setChecked(mode == "ies")
        self.mode_messages_btn.blockSignals(False)
        self.mode_ies_btn.blockSignals(False)

        self._update_search_placeholders()

        if trigger_populate:
            self._save_config()
            self._populate_messages()

    def _update_search_placeholders(self):
        checked = self.deep_search_btn.isChecked()
        if self._browse_mode == "ies":
            self.msg_search.setPlaceholderText("Filter IE / Type name, acronym, or Clause (e.g. UE parameters, 9.11.3)...")
            self.ie_search.setPlaceholderText(
                "Filter IE by keyword or Description..."
                if checked
                else "Additional IE / Clause keyword filter..."
            )
        else:
            self.msg_search.setPlaceholderText("Filter by name, acronym, or spec (e.g. SIB1, PFCP, 24.501)...")
            self.ie_search.setPlaceholderText(
                "Filter by Field or Description (e.g. emergency)..."
                if checked
                else "Filter by IE / Field (e.g. RadioBearer, CellGroup)..."
            )

    def _on_filter_by_ie_requested(self, term: str):
        if self._browse_mode != "messages":
            self._set_browse_mode("messages", trigger_populate=False)
        self.ie_search.setText(term)

    def _load_config(self) -> Dict[str, Any]:
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logging.warning(f"Could not load Protocol config from {self.config_path}: {e}")
        return {}

    def _save_config(self):
        if self._loading_config or self.version_tree._updating_checks:
            return

        try:
            config_data = {
                "checked_versions": self.version_tree.get_checked_versions_info(),
                "browse_mode": self._browse_mode,
                "msg_filter": self.msg_search.text(),
                "ie_filter": self.ie_search.text(),
                "search_descriptions": self.deep_search_btn.isChecked(),
                "selected_message": self.current_selected_message_name or "",
                "selected_ie": self.current_selected_ie_key or "",
                "collapsed_specs": self.version_tree.get_collapsed_spec_numbers(),
            }
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            logging.warning(f"Could not save Protocol config to {self.config_path}: {e}")

    # -------------------------------------------------------------------------
    # Search Handlers
    # -------------------------------------------------------------------------

    def _on_deep_search_toggled(self, checked: bool):
        self._update_search_placeholders()
        self.deep_search_btn.setText("📖 Desc: ON" if checked else "📖 Desc")
        self._save_config()
        self._populate_messages()

    def _on_ie_search_context_menu(self, pos):
        menu = self.ie_search.createStandardContextMenu()
        menu.addSeparator()

        act_toggle = QAction("📖 Include Descriptions in Search", self)
        act_toggle.setCheckable(True)
        act_toggle.setChecked(self.deep_search_btn.isChecked())
        act_toggle.toggled.connect(self.deep_search_btn.setChecked)
        menu.addAction(act_toggle)
        menu.exec_(self.ie_search.mapToGlobal(pos))

    # -------------------------------------------------------------------------
    # Version Tree Management
    # -------------------------------------------------------------------------

    def refresh_versions(self):
        versions = self.db.get_imported_versions()
        saved_config = self._load_config()

        self.version_tree.populate(
            versions=versions,
            saved_checked=saved_config.get("checked_versions"),
            saved_collapsed=set(saved_config.get("collapsed_specs", [])),
        )

        self._update_stats_ribbon()

        if not self._initialized_filters:
            self._loading_config = True
            if "browse_mode" in saved_config:
                self._set_browse_mode(str(saved_config["browse_mode"]), trigger_populate=False)
            if "msg_filter" in saved_config:
                self.msg_search.setText(saved_config["msg_filter"])
            if "ie_filter" in saved_config:
                self.ie_search.setText(saved_config["ie_filter"])
            if "search_descriptions" in saved_config:
                self.deep_search_btn.setChecked(bool(saved_config["search_descriptions"]))
            if "selected_message" in saved_config and saved_config["selected_message"]:
                self.current_selected_message_name = saved_config["selected_message"]
            if "selected_ie" in saved_config and saved_config["selected_ie"]:
                self.current_selected_ie_key = saved_config["selected_ie"]
            self._loading_config = False
            self._initialized_filters = True
        if hasattr(self, "layout_controller"):
            self.layout_controller.update_specs_summary(self.version_tree.get_checked_versions_info())

    def _on_version_selection_changed(self):
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        if hasattr(self, "layout_controller"):
            self.layout_controller.update_specs_summary(self.version_tree.get_checked_versions_info())
        self._save_config()
        self._populate_messages()

        current_msg_item = self.msg_list.currentItem()
        if current_msg_item:
            self._on_message_clicked(current_msg_item)
        else:
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)
            self._hide_message_context_banner()

    def _hide_message_context_banner(self):
        if hasattr(self, "layout_controller"):
            self.layout_controller.set_banner_active(False)
        if hasattr(self, "msg_context_banner"):
            self.msg_context_banner.setVisible(False)
        if hasattr(self, "msg_direction_badge"):
            self.msg_direction_badge.setVisible(False)
            self.msg_direction_badge.setText("")
        if hasattr(self, "msg_meta_badge"):
            self.msg_meta_badge.setVisible(False)
            self.msg_meta_badge.setText("")
        if hasattr(self, "msg_description_lbl"):
            self.msg_description_lbl.setVisible(False)
            self.msg_description_lbl.setText("")

    def _update_message_context_banner(self, description: str, direction: str):
        desc_clean = (description or "").strip()
        dir_clean = (direction or "").strip()

        # Extract RRC radio-bearer metadata into clean chips
        meta_chips: List[str] = []
        if desc_clean:
            for pattern, label in (
                (RE_CTX_SRB, "SRB"),
                (RE_CTX_RLC, "RLC"),
                (RE_CTX_LCH, "Ch"),
            ):
                m = pattern.search(desc_clean)
                if m:
                    val = m.group(1).strip()
                    if val:
                        meta_chips.append(f"{label}: {val}")
                    desc_clean = (desc_clean[:m.start()] + " " + desc_clean[m.end():]).strip()

            desc_clean = re.sub(r"\s+", " ", desc_clean).strip()
            desc_clean = RE_CTX_TRAILING_CAPTION.sub("", desc_clean).strip()

        if not desc_clean and not dir_clean and not meta_chips:
            self._hide_message_context_banner()
            return

        if dir_clean:
            self.msg_direction_badge.setText(f"⇄ Direction: {dir_clean}")
            self.msg_direction_badge.setVisible(True)
        else:
            self.msg_direction_badge.setText("")
            self.msg_direction_badge.setVisible(False)

        if hasattr(self, "msg_meta_badge"):
            if meta_chips:
                self.msg_meta_badge.setText(" · ".join(meta_chips))
                self.msg_meta_badge.setVisible(True)
            else:
                self.msg_meta_badge.setText("")
                self.msg_meta_badge.setVisible(False)

        if desc_clean:
            self.msg_description_lbl.setText(desc_clean)
            self.msg_description_lbl.setVisible(True)
        else:
            self.msg_description_lbl.setText("")
            self.msg_description_lbl.setVisible(False)

        if hasattr(self, "layout_controller"):
            self.layout_controller.set_banner_active(True)
        if getattr(self, "layout_controller", None) and self.layout_controller.current_mode() == "inspector_only":
            self.msg_context_banner.setVisible(False)
        else:
            self.msg_context_banner.setVisible(True)

    def _reload_single_version(self, spec_number: str, version: str):
        reload_single_version(self, spec_number, version)

    def _reload_spec_group(self, spec_number: str, child_count: int):
        reload_spec_group(self, spec_number, child_count)

    def _on_reload_checked_clicked(self):
        reload_checked_versions(self)

    def _delete_single_version(self, spec_number: str, version: str):
        delete_single_version(self, spec_number, version)

    def _delete_spec_group(self, spec_number: str, child_count: int):
        delete_spec_group(self, spec_number, child_count)

    def _on_search_timer_timeout(self):
        self._save_config()
        self._populate_messages()
