"""Message/IE list loading, selection, diff filtering, and evolution matrix mixin."""

from __future__ import annotations

import html
import re
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QHeaderView, QListWidgetItem

from core.ui.font_scale import scale_px
from core.ui.theme import C
from modules.nas.ui.nas_dialogs import get_spec_metadata
from modules.nas.ui.nas_matrix_controller import handle_table_cell_click
from modules.nas.ui.nas_models import NASEvolutionMatrixModel
from modules.nas.ui.nas_workers import NASMessageListWorker, ReverseLookupWorker
from ..styles import RE_CLAUSE_FROM_REF, RE_SPEC_FROM_COL


class NASTabMessagesAndMatrixMixin:
    """Populates message/IE lists, handles selection/diffing, and rebuilds the matrix model."""

    def _populate_messages(self):
        if not self.selected_version_ids:
            self.msg_list.clear()
            self.msg_count_lbl.setText("0 items")
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)
            self.interface_combo.setVisible(False)
            if hasattr(self, "depth_combo"):
                self.depth_combo.setVisible(False)
            self._hide_message_context_banner()
            default_prompt = (
                "Select an Information Element to View Evolution & Definition"
                if self._browse_mode == "ies"
                else "Select a Message to View Evolution Matrix"
            )
            self.matrix_title.setText(default_prompt)
            self.inspector.clear_display()
            return

        ie_query = self.ie_search.text().strip()
        if self._browse_mode == "ies":
            top_query = self.msg_search.text().strip()
            combined_query = f"{top_query} {ie_query}".strip()
        else:
            combined_query = ie_query

        search_desc = self.deep_search_btn.isChecked()

        self._msg_query_req_id += 1
        req_id = self._msg_query_req_id

        worker = NASMessageListWorker(
            db=self.db,
            ie_query=combined_query,
            version_ids=list(self.selected_version_ids),
            search_desc=search_desc,
            request_id=req_id,
            browse_mode=self._browse_mode,
        )
        self._msg_query_workers[req_id] = worker
        worker.results_ready.connect(self._on_message_list_ready)
        worker.finished.connect(lambda rid=req_id: self._release_msg_query_worker(rid))
        worker.start()

    def _release_msg_query_worker(self, req_id: int):
        """Release a completed message-query worker without touching running workers."""
        worker = self._msg_query_workers.pop(req_id, None)
        if worker is not None:
            worker.deleteLater()

    def _on_message_list_ready(self, items: List[Dict[str, Any]], req_id: int):
        if req_id != self._msg_query_req_id:
            return

        msg_query = self.msg_search.text().strip().lower()
        msg_tokens = [t for t in msg_query.split() if t]

        self.msg_list.blockSignals(True)
        self.msg_list.clear()

        target_item = None
        visible_count = 0

        if self._browse_mode == "ies":
            target_key = self.current_selected_ie_key
            for d in items:
                ie_name = d.get("ie_name", "")
                clause = d.get("clause", "")
                spec_num = str(d.get("spec_number", "")).replace("TS", "").strip()
                meta = get_spec_metadata(spec_num) if spec_num else {}

                ie_key = f"{ie_name}::{clause}"

                item = QListWidgetItem(ie_name)
                item.setData(Qt.UserRole, ie_key)
                item.setData(Qt.UserRole + 1, d)
                item.setData(Qt.UserRole + 2, spec_num)
                item.setData(Qt.UserRole + 3, clause)
                item.setData(Qt.UserRole + 4, meta.get("acronym", ""))
                item.setData(Qt.UserRole + 5, meta.get("wg", ""))

                tooltip_lines = [f"Information Element: {ie_name}"]
                if clause:
                    tooltip_lines.append(f"Clause / Type: {clause}")
                if spec_num:
                    tooltip_lines.append(f"Specification: TS {spec_num} ({meta.get('title', '')})")
                    tooltip_lines.append(f"Working Group: {meta.get('wg', '')} • Domain: {meta.get('tech', '')}")
                item.setToolTip("\n".join(tooltip_lines))

                self.msg_list.addItem(item)

                search_target = f"{ie_name} {clause} {spec_num} {meta.get('acronym', '')} {meta.get('title', '')}".lower()
                if msg_tokens and not all(tok in search_target for tok in msg_tokens):
                    item.setHidden(True)
                else:
                    visible_count += 1

                if target_key and ie_key == target_key and not item.isHidden():
                    target_item = item
        else:
            target_msg_name = self.current_selected_message_name
            for m in items:
                msg_name = m["message_name"]
                clause = m.get("clause", "")
                msg_desc = (m.get("description") or "").strip()
                msg_dir = (m.get("direction") or "").strip()
                spec_num = str(m.get("spec_number", "")).replace("TS", "").strip()
                meta = get_spec_metadata(spec_num) if spec_num else {}

                item = QListWidgetItem(msg_name)
                item.setData(Qt.UserRole, msg_name)
                item.setData(Qt.UserRole + 1, m)
                item.setData(Qt.UserRole + 2, spec_num)
                item.setData(Qt.UserRole + 3, clause)
                item.setData(Qt.UserRole + 4, meta.get("acronym", ""))
                item.setData(Qt.UserRole + 5, meta.get("wg", ""))

                tooltip_lines = [f"Message: {msg_name}"]
                if clause:
                    tooltip_lines.append(f"Clause: {clause}")
                if msg_dir:
                    tooltip_lines.append(f"Direction: {msg_dir}")
                if msg_desc:
                    tooltip_lines.append(f"Description: {msg_desc}")
                if spec_num:
                    tooltip_lines.append(f"Specification: TS {spec_num} ({meta.get('title', '')})")
                    tooltip_lines.append(f"Working Group: {meta.get('wg', '')} • Domain: {meta.get('tech', '')}")
                item.setToolTip("\n".join(tooltip_lines))

                self.msg_list.addItem(item)

                extra_search = f" {msg_desc} {msg_dir}" if self.deep_search_btn.isChecked() else f" {msg_dir}"
                search_target = f"{msg_name} {clause} {spec_num} {meta.get('acronym', '')} {meta.get('title', '')}{extra_search}".lower()
                if msg_tokens and not all(tok in search_target for tok in msg_tokens):
                    item.setHidden(True)
                else:
                    visible_count += 1

                if target_msg_name and msg_name == target_msg_name and not item.isHidden():
                    target_item = item

        self.msg_list.blockSignals(False)

        # Update counter label
        total_count = self.msg_list.count()
        if visible_count == total_count:
            self.msg_count_lbl.setText(f"{total_count} items")
        else:
            self.msg_count_lbl.setText(f"{visible_count} / {total_count} items")

        default_prompt = (
            "Select an Information Element to View Evolution & Definition"
            if self._browse_mode == "ies"
            else "Select a Message to View Evolution Matrix"
        )

        if target_item:
            self.msg_list.setCurrentItem(target_item)
            self._on_message_clicked(target_item)
        elif self.msg_list.count() == 1 and (msg_tokens or self.ie_search.text().strip()):
            only_item = self.msg_list.item(0)
            if not only_item.isHidden():
                self.msg_list.setCurrentItem(only_item)
                self._on_message_clicked(only_item)
        elif self.msg_list.count() > 0:
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)
            self.interface_combo.setVisible(False)
            if hasattr(self, "depth_combo"):
                self.depth_combo.setVisible(False)
            self._hide_message_context_banner()
            self.matrix_title.setText(default_prompt)
            self._current_message_clause = None
            self._current_message_spec = None
            self._update_open_spec_button()
        else:
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)
            self.interface_combo.setVisible(False)
            if hasattr(self, "depth_combo"):
                self.depth_combo.setVisible(False)
            self._hide_message_context_banner()
            self.matrix_title.setText(default_prompt)
            self.inspector.clear_display()
            self._current_message_clause = None
            self._current_message_spec = None
            self._update_open_spec_button()

    def _on_message_clicked(self, item: QListWidgetItem):
        if self._browse_mode == "ies":
            self._hide_message_context_banner()
            ie_data = item.data(Qt.UserRole + 1) or {}
            ie_name = ie_data.get("ie_name", "")
            clause = ie_data.get("clause", "") or item.data(Qt.UserRole + 3) or ""
            self.current_selected_ie_key = item.data(Qt.UserRole)
            self._current_message_clause = clause
            self._save_config()

            df = self.db.get_ie_definition_evolution_df(
                clause=clause,
                ie_name=ie_name,
                version_ids=self.selected_version_ids,
            )

            if not df.empty and "spec_number" in df.columns:
                unique_specs = sorted(df["spec_number"].dropna().unique())
                self._current_message_spec = str(unique_specs[0]) if len(unique_specs) == 1 else None
            else:
                self._current_message_spec = None

            meta = get_spec_metadata(self._current_message_spec) if self._current_message_spec else {}
            spec_badge = f" [{meta.get('acronym', f'TS {self._current_message_spec}')}]" if self._current_message_spec else ""
            clause_suffix = f" [Clause {clause}]" if clause and clause != ie_name else ""
            self.matrix_title.setText(f"Information Element{spec_badge}: {ie_name}{clause_suffix}")
            self.interface_combo.setVisible(False)
            if hasattr(self, "depth_combo"):
                self.depth_combo.setVisible(False)

            self._current_matrix_df = df
            self._rebuild_matrix_model(ignore_ie_filter=True)
            self._update_open_spec_button()

            # Immediately display the full Clause 9 / ASN.1 definition in the Inspector below
            defs = self.db.get_ie_definitions_by_clause(
                clause=clause,
                alt_name=ie_name,
                spec_number=self._current_message_spec,
                version_ids=self.selected_version_ids,
            )
            self.inspector.display_definitions(
                clause=clause,
                ie_name=ie_name,
                spec_number=self._current_message_spec,
                defs=defs,
                containing_msgs=[],
                fallback_type_ref=f"{ie_name} {clause}".strip(),
            )
            if defs:
                self._start_reverse_lookup(
                    clause=clause,
                    ie_name=ie_name,
                    spec_num=self._current_message_spec,
                )
            return

        msg_name = item.data(Qt.UserRole)
        m_data = item.data(Qt.UserRole + 1) or {}
        clause = m_data.get("clause", "") or item.data(Qt.UserRole + 3) or ""
        self.current_selected_message_name = msg_name
        self._current_message_clause = clause
        self._save_config()
        self._save_config()

        ie_query = self.ie_search.text().strip()
        search_desc = self.deep_search_btn.isChecked()
        desc_label = " (incl. Desc)" if search_desc and ie_query else ""
        title_suffix = f" (Filtered by IE{desc_label}: '{ie_query}')" if ie_query else ""

        need_descriptions = search_desc and bool(ie_query)

        df = self.db.get_message_evolution_df(
            message_name=msg_name,
            version_ids=self.selected_version_ids,
            include_descriptions=need_descriptions,
        )

        if not df.empty and "spec_number" in df.columns:
            unique_specs = sorted(df["spec_number"].dropna().unique())
            self._current_message_spec = str(unique_specs[0]) if len(unique_specs) == 1 else None
        else:
            self._current_message_spec = None

        meta = get_spec_metadata(self._current_message_spec) if self._current_message_spec else {}
        spec_badge = f" [{meta.get('acronym', f'TS {self._current_message_spec}')}]" if self._current_message_spec else ""
        self.matrix_title.setText(f"Message{spec_badge}: {msg_name}{title_suffix}")

        has_appl = bool(
            "applicability" in df.columns
            and df["applicability"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
            .any()
        )
        self.interface_combo.blockSignals(True)
        self.interface_combo.setVisible(has_appl)
        if has_appl:
            current_choice = self.interface_combo.currentText()
            unique_ifaces = set()
            for raw_val in df["applicability"].dropna():
                for part in str(raw_val).split(","):
                    token = part.strip()
                    if token and token.upper() not in ("ALL INTERFACES", "ALL"):
                        unique_ifaces.add(token)

            self.interface_combo.clear()
            self.interface_combo.addItem("All Interfaces")
            for iface in sorted(unique_ifaces):
                self.interface_combo.addItem(iface)

            self.interface_combo.setMaxVisibleItems(max(20, len(unique_ifaces) + 2))

            idx = self.interface_combo.findText(current_choice)
            self.interface_combo.setCurrentIndex(idx if idx >= 0 else 0)
        else:
            self.interface_combo.setCurrentIndex(0)
        self.interface_combo.blockSignals(False)

        has_nested = bool(
            "depth" in df.columns
            and df["depth"].fillna(0).astype(int).gt(0).any()
        )
        if hasattr(self, "depth_combo"):
            self.depth_combo.blockSignals(True)
            self.depth_combo.setVisible(has_nested)
            if not has_nested:
                self.depth_combo.setCurrentIndex(0)
            self.depth_combo.blockSignals(False)

        self._current_matrix_df = df
        self._rebuild_matrix_model()
        self._update_open_spec_button()

        # Update Message Context Banner (Direction + Description)
        ctx = self.db.get_message_context_info(
            message_name=msg_name,
            clause=clause,
            version_ids=self.selected_version_ids,
        )
        msg_desc = ctx.get("description", "") or str(m_data.get("description", "") or "").strip()
        msg_dir = ctx.get("direction", "") or str(m_data.get("direction", "") or "").strip()
        eff_clause = ctx.get("clause", "") or clause
        self._update_message_context_banner(description=msg_desc, direction=msg_dir)

        # Immediately populate the Field & Structure Inspector with the message definition
        msg_defs = self.db.get_ie_definitions_by_clause(
            clause=eff_clause,
            alt_name=msg_name,
            spec_number=self._current_message_spec,
            version_ids=self.selected_version_ids,
        )
        if not msg_defs:
            spec_num = self._current_message_spec or str(m_data.get("spec_number", "") or "")
            ver_hint = ""
            if not df.empty and "version" in df.columns:
                vers = [str(v) for v in df["version"].dropna().unique() if str(v).strip()]
                if vers:
                    ver_hint = vers[-1]
            synth_html = (
                f'<h2 style="color: {C.ACCENT_SKY}; margin-top: 4px; margin-bottom: 6px; '
                f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid {C.BORDER_SUBTLE}; padding-bottom: 4px;">'
                f'{html.escape(msg_name)}{f" (Clause {html.escape(eff_clause)})" if eff_clause else ""}</h2>'
            )
            if msg_dir:
                synth_html += (
                    f'<p style="margin: 4px 0; line-height: 1.4; color: {C.ACCENT_SKY_DARK}; font-weight: bold; font-size: 11px;">'
                    f'Direction: {html.escape(msg_dir)}</p>'
                )
            if msg_desc:
                synth_html += (
                    f'<p style="margin: 4px 0; line-height: 1.4; color: {C.TEXT_BODY}; font-size: 11px;">'
                    f'{html.escape(msg_desc)}</p>'
                )
            msg_defs = [{
                "clause": eff_clause,
                "ie_name": msg_name,
                "raw_description": synth_html,
                "structure_table": "[]",
                "spec_number": spec_num,
                "version": ver_hint,
                "release": "",
            }]
        self.inspector.display_definitions(
            clause=eff_clause,
            ie_name=msg_name,
            spec_number=self._current_message_spec,
            defs=msg_defs,
            containing_msgs=[],
            fallback_type_ref=f"{msg_name} {eff_clause}".strip(),
            message_name=msg_name,
        )

    def _get_active_diff_filter(self) -> set:
        active = set()
        if getattr(self, "btn_diff_added", None) and self.btn_diff_added.isChecked():
            active.add("added")
        if getattr(self, "btn_diff_modified", None) and self.btn_diff_modified.isChecked():
            active.add("modified")
        if getattr(self, "btn_diff_removed", None) and self.btn_diff_removed.isChecked():
            active.add("removed")
        return active

    def _update_diff_filter_badges(self, counts: Optional[Dict[str, int]] = None):
        counts = counts or {"added": 0, "modified": 0, "removed": 0}
        specs = (
            ("btn_diff_added", "🟢 Added", counts.get("added", 0)),
            ("btn_diff_modified", "🟡 Modified", counts.get("modified", 0)),
            ("btn_diff_removed", "🔴 Removed", counts.get("removed", 0)),
        )
        for attr, base_lbl, cnt in specs:
            btn = getattr(self, attr, None)
            if btn is not None:
                btn.setText(f"{base_lbl} ({cnt})" if cnt > 0 else base_lbl)

    def _rebuild_matrix_model(self, ignore_ie_filter: bool = False):
        if not hasattr(self, "_current_matrix_df") or self._current_matrix_df.empty:
            self.matrix_table.setModel(None)
            self.matrix_stack.setCurrentIndex(0)
            self.copy_matrix_btn.setEnabled(False)
            self._update_diff_filter_badges(None)
            return

        df = self._current_matrix_df
        ie_query = "" if (ignore_ie_filter or self._browse_mode == "ies") else self.ie_search.text().strip()
        search_desc = self.deep_search_btn.isChecked()
        active_if = self.interface_combo.currentText() if self.interface_combo.isVisible() else None
        max_depth = (
            self.depth_combo.currentData()
            if hasattr(self, "depth_combo") and self.depth_combo.isVisible()
            else None
        )

        model = NASEvolutionMatrixModel(
            df,
            ie_filter=ie_query,
            search_descriptions=(search_desc and bool(ie_query)),
            interface_filter=active_if,
            diff_filter=self._get_active_diff_filter(),
            max_depth=max_depth,
        )
        self.matrix_table.setModel(model)
        self._update_diff_filter_badges(getattr(model, "diff_counts", None))
        has_rows = model.rowCount() > 0
        self.matrix_stack.setCurrentIndex(1 if has_rows else 0)
        self.copy_matrix_btn.setEnabled(has_rows)
        if hasattr(self, "layout_controller"):
            self.layout_controller.apply_auto_compact_for_matrix(self._browse_mode, model.rowCount())

        h_header = self.matrix_table.horizontalHeader()
        if model.columnCount() > 0:
            has_any_iei = any(
                str(model.data(model.index(r, 0), Qt.DisplayRole) or "").strip() not in ("", "-")
                for r in range(model.rowCount())
            )
            self.matrix_table.setColumnHidden(0, not has_any_iei)
            h_header.setSectionResizeMode(0, QHeaderView.Interactive)
            self.matrix_table.setColumnWidth(0, max(75, scale_px(75)))

            h_header.setSectionResizeMode(1, QHeaderView.Stretch)
            h_header.setSectionResizeMode(2, QHeaderView.Interactive)
            self.matrix_table.setColumnWidth(2, max(280, scale_px(280)))

            for col in range(3, model.columnCount()):
                h_header.setSectionResizeMode(col, QHeaderView.Interactive)
                self.matrix_table.setColumnWidth(col, max(130, scale_px(130)))

    def _start_reverse_lookup(self, clause: str, ie_name: str, spec_num: Optional[str]):
        self._reverse_lookup_request_id += 1
        current_req_id = self._reverse_lookup_request_id

        worker = ReverseLookupWorker(
            db=self.db,
            clause=clause,
            ie_name=ie_name,
            spec_number=spec_num,
            version_ids=list(self.selected_version_ids),
            request_id=current_req_id,
        )
        self._reverse_lookup_workers[current_req_id] = worker
        worker.results_ready.connect(self._on_reverse_lookup_finished)
        worker.finished.connect(
            lambda rid=current_req_id: self._release_reverse_lookup_worker(rid)
        )
        worker.start()

    def _on_table_cell_clicked(self, index):
        handle_table_cell_click(self, index)

    def _on_table_cell_double_clicked(self, index):
        """Double-clicking a cell in the Evolution Matrix opens the corresponding specification in Word."""
        model = self.matrix_table.model()
        if not model or not index.isValid():
            return

        row, col = index.row(), index.column()
        ie_name = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip().lstrip("└─ ")
        type_ref = str(model.data(model.index(row, 2), Qt.DisplayRole) or "").strip()

        match = RE_CLAUSE_FROM_REF.search(type_ref)
        clause = match.group(1).strip() if match else ie_name
        spec_num = getattr(self, "_current_message_spec", None)

        target_version = None
        if col >= 3 and hasattr(model, "_get_visible_column_name"):
            col_name = str(model._get_visible_column_name(col))
            match_spec = RE_SPEC_FROM_COL.search(col_name)
            if match_spec:
                spec_num = match_spec.group(0)
            m_ver = re.search(r"v?([0-9]+\.[0-9]+\.[0-9]+)", col_name)
            if m_ver:
                target_version = m_ver.group(1)

        if spec_num:
            self.open_specification(
                spec_number=spec_num,
                version=target_version,
                clause_number=clause,
                doc_format="word",
            )
