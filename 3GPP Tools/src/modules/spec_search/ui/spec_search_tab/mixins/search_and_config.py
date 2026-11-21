"""Config persistence, FTS5 query execution, and Evolution Matrix population mixin."""

from __future__ import annotations

import html
import json
import logging
import re

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import QHeaderView, QTableView

from core.ui.font_scale import register_scaled_row_height, scale_px
from core.ui.theme import C
from modules.spec_search.core.spec_search_threads import SpecSearchQueryWorker
from modules.spec_search.ui.spec_search_components import SpecEvolutionStatusDelegate
from modules.spec_search.ui.spec_search_models import SpecEvolutionMatrixModel


class SpecSearchTabSearchMixin:
    """Manages filter persistence, inspector layout toggles, async search, and table selection."""

    def _save_config(self):
        """Saves active search text, clause filter, cutoff dates, and checked specification versions."""
        try:
            checked_versions = self.version_tree.get_checked_versions_info()
            config_data = {
                "search_query": self.search_input.text(),
                "clause_filter": self.clause_filter_input.text(),
                "enable_cutoff": self.chk_filing_date.isChecked(),
                "cutoff_date": self.cutoff_date_edit.date().toString("yyyy-MM-dd"),
                "only_post_cutoff": self.chk_post_date_only.isChecked(),
                "auto_replace_older_versions": getattr(self, "auto_replace_older_versions", True),
                "checked_versions": checked_versions,
                "inspector_side_by_side": (self.right_splitter.orientation() == Qt.Horizontal),
            }
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            logging.error(f"Failed to save spec search config: {e}")

    def _load_config(self) -> dict:
        """Loads saved specification search settings from JSON."""
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logging.error(f"Failed to read spec search config: {e}")
        return {}

    def _load_config_and_refresh(self):
        """Restores UI filter inputs, populates the tree with saved checkboxes, and runs search."""
        config = self._load_config()
        self.auto_replace_older_versions = config.get("auto_replace_older_versions", True) if config else True
        if hasattr(self, "act_auto_replace"):
            self.act_auto_replace.setChecked(self.auto_replace_older_versions)
        versions = self.db.get_imported_versions()

        saved_checked = config.get("checked_versions", None) if config else None
        self.version_tree.populate(versions, saved_checked=saved_checked, spec_titles=self._spec_titles)
        self.selected_version_ids = self.version_tree.get_selected_version_ids()

        if config:
            # Block signals while restoring to avoid redundant triggers
            self.search_input.blockSignals(True)
            self.clause_filter_input.blockSignals(True)
            self.chk_filing_date.blockSignals(True)
            self.cutoff_date_edit.blockSignals(True)
            self.chk_post_date_only.blockSignals(True)

            self.search_input.setText(config.get("search_query", ""))
            self.clause_filter_input.setText(config.get("clause_filter", ""))

            enable_cutoff = config.get("enable_cutoff", False)
            self.chk_filing_date.setChecked(enable_cutoff)
            self.cutoff_date_edit.setEnabled(enable_cutoff)
            self.chk_post_date_only.setEnabled(enable_cutoff)

            saved_date = config.get("cutoff_date", "")
            if saved_date:
                d = QDate.fromString(saved_date, "yyyy-MM-dd")
                if d.isValid():
                    self.cutoff_date_edit.setDate(d)

            self.chk_post_date_only.setChecked(config.get("only_post_cutoff", False))

            self.search_input.blockSignals(False)
            self.clause_filter_input.blockSignals(False)
            self.chk_filing_date.blockSignals(False)
            self.cutoff_date_edit.blockSignals(False)
            self.chk_post_date_only.blockSignals(False)

            if config.get("inspector_side_by_side", False):
                self.inspector.set_side_by_side(True)
                self.right_splitter.setOrientation(Qt.Horizontal)
                self.right_splitter.setSizes([550, 450])

        self._update_db_stats()

        # Trigger search if a saved query exists
        if self.search_input.text().strip():
            self._execute_search()

    def _toggle_inspector_layout(self, side_by_side: bool):
        if side_by_side:
            self.right_splitter.setOrientation(Qt.Horizontal)
            self.right_splitter.setSizes([550, 450])
        else:
            self.right_splitter.setOrientation(Qt.Vertical)
            self.right_splitter.setSizes([380, 260])
        self._save_config()

    def _toggle_inspector_maximize(self, is_maximized: bool):
        if is_maximized:
            self._pre_maximize_sizes = self.right_splitter.sizes()
            self.right_splitter.setSizes([0, 1000])
        else:
            if hasattr(self, "_pre_maximize_sizes") and self._pre_maximize_sizes and any(self._pre_maximize_sizes):
                self.right_splitter.setSizes(self._pre_maximize_sizes)
            else:
                if self.right_splitter.orientation() == Qt.Horizontal:
                    self.right_splitter.setSizes([550, 450])
                else:
                    self.right_splitter.setSizes([380, 260])

    def _update_db_stats(self):
        try:
            versions = self.db.get_imported_versions()
            specs = set(v.get("spec_number") for v in versions if v.get("spec_number"))
            num_specs = len(specs)
            num_vers = len(versions)
            if hasattr(self, "tree_stats_badge"):
                self.tree_stats_badge.setText(f"{num_specs}")
                self.tree_stats_badge.setToolTip(f"{num_specs} Specification(s), {num_vers} Release(s) Indexed")
            if hasattr(self, "db_stats_label") and self.db_stats_label != self.tree_stats_badge:
                self.db_stats_label.setText(f"📦 {num_specs} Specs · {num_vers} Releases Indexed")
        except Exception:
            if hasattr(self, "tree_stats_badge"):
                self.tree_stats_badge.setText("0")

    def _on_filter_changed(self):
        self._search_timer.start()
        self._save_config_timer.start()

    def _on_cutoff_date_filter_toggled(self, checked: bool):
        self.cutoff_date_edit.setEnabled(checked)
        self.chk_post_date_only.setEnabled(checked)
        self._on_filter_changed()

    def _on_version_selection_changed(self):
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        self._search_timer.start()
        self._save_config_timer.start()

    # -------------------------------------------------------------------------
    # Search Execution
    # -------------------------------------------------------------------------

    def _execute_search(self):
        query = self.search_input.text().strip()
        clause_filter = self.clause_filter_input.text().strip()

        if not self.selected_version_ids or not query:
            self.spec_results_tabs.clear()
            self.matrix_title.setText(f"<span style='color: {C.TEXT_MUTED}; font-size: 12px;'>💡 Type a query above to trace clause evolution across releases.</span>")
            self.inspector.clear_display()
            return

        self._current_req_id += 1
        req_id = self._current_req_id

        for w in list(self._active_query_workers.values()):
            if w.isRunning():
                w.requestInterruption()

        worker = SpecSearchQueryWorker(
            db=self.db,
            query_str=query,
            version_ids=self.selected_version_ids,
            clause_filter=clause_filter,
            request_id=req_id,
        )
        self._query_worker = worker
        self._active_query_workers[req_id] = worker
        worker.results_ready.connect(self._on_search_results_ready)
        worker.finished.connect(lambda r=req_id: self._release_query_worker(r))
        worker.start()

    def _release_query_worker(self, req_id: int):
        worker = self._active_query_workers.pop(req_id, None)
        if worker:
            worker.deleteLater()

    def _on_search_results_ready(self, df, req_id: int):
        if req_id != self._current_req_id:
            return

        query = self.search_input.text().strip()
        self.spec_results_tabs.clear()

        if df.empty:
            self.matrix_title.setText(f"<span style='color: {C.TEXT_MUTED}; font-size: 12px;'>No matching clauses found for '<b>{html.escape(query)}</b>' across selected releases.</span>")
            self.inspector.clear_display()
            return

        cutoff_date = self.cutoff_date_edit.date().toString("yyyy-MM-dd") if self.chk_filing_date.isChecked() else None
        only_post_date = self.chk_post_date_only.isChecked() if self.chk_filing_date.isChecked() else False

        total_clauses_count = 0
        specs_with_hits = 0

        grouped_specs = sorted(df["spec_number"].unique())

        for spec_num in grouped_specs:
            spec_df = df[df["spec_number"] == spec_num].copy()

            model = SpecEvolutionMatrixModel(
                raw_df=spec_df,
                cutoff_date=cutoff_date,
                only_added_after_cutoff=only_post_date,
            )

            if model.rowCount() == 0:
                continue

            specs_with_hits += 1
            total_clauses_count += model.rowCount()

            table = QTableView()
            table.setAlternatingRowColors(False)
            table.setSelectionBehavior(QTableView.SelectRows)
            table.setSelectionMode(QTableView.SingleSelection)
            register_scaled_row_height(table, 26)
            table.verticalHeader().setVisible(False)
            table.setItemDelegate(SpecEvolutionStatusDelegate(table))
            table.setModel(model)
            table.setStyleSheet(f"""
                QTableView {{
                    background-color: {C.BG_CARD};
                    border: 1px solid {C.BORDER_STRONG};
                    border-radius: 4px;
                    gridline-color: transparent;
                    font-size: 12px;
                    selection-background-color: {C.ACCENT_SOFT};
                    selection-color: {C.TEXT_PRIMARY};
                }}
                QHeaderView::section {{
                    background-color: {C.BG_SURFACE};
                    color: {C.TEXT_SECONDARY};
                    font-weight: bold;
                    font-size: 11px;
                    border: none;
                    border-bottom: 1px solid {C.BORDER_STRONG};
                    border-right: 1px solid {C.BORDER_SUBTLE};
                    padding: 5px 8px;
                }}
            """)

            h_header = table.horizontalHeader()
            h_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            h_header.setSectionResizeMode(1, QHeaderView.Stretch)
            for c in range(2, model.columnCount()):
                h_header.setSectionResizeMode(c, QHeaderView.Interactive)
                table.setColumnWidth(c, max(115, scale_px(115)))

            table.clicked.connect(lambda idx, s=spec_num, t=table: self._on_tab_table_clicked(idx, s, t))
            table.doubleClicked.connect(lambda idx, s=spec_num, t=table: self._on_table_double_clicked(idx, s, t))
            table.setContextMenuPolicy(Qt.CustomContextMenu)
            table.customContextMenuRequested.connect(lambda pos, t=table, s=spec_num: self._show_matrix_context_menu(pos, t, s))

            short_name = self.get_spec_short_name(spec_num)
            tab_title = f"TS {spec_num} · {short_name} ({model.rowCount()})" if short_name else f"TS {spec_num} ({model.rowCount()})"
            tab_idx = self.spec_results_tabs.addTab(table, tab_title)
            full_title = self.get_spec_title(spec_num)
            if full_title:
                self.spec_results_tabs.setTabToolTip(tab_idx, f"TS {spec_num}: {full_title} ({model.rowCount()} clauses)")
            else:
                self.spec_results_tabs.setTabToolTip(tab_idx, f"TS {spec_num} ({model.rowCount()} clauses)")

        if specs_with_hits == 0:
            self.matrix_title.setText(f"<span style='color: {C.TEXT_MUTED}; font-size: 12px;'>No matching clauses found for '<b>{html.escape(query)}</b>' matching active filters.</span>")
            self.inspector.clear_display()
            return

        date_badge = (
            f" <span style='background-color: {C.WARNING_SOFT_ALT}; color: {C.WARNING}; font-weight: bold; padding: 2px 7px; border-radius: 4px; border: 1px solid {C.WARNING_BORDER}; font-size: 11px;'>Cutoff ≥ {cutoff_date}</span>"
            if cutoff_date and only_post_date
            else ""
        )
        self.matrix_title.setText(
            f"<span style='color: {C.TEXT_SECONDARY}; font-size: 11px; font-weight: 600;'>Results:</span> "
            f"<span style='background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_SKY}; font-weight: bold; padding: 2px 8px; border-radius: 4px; border: 1px solid {C.INFO_BORDER_ALT}; font-size: 11px;'>🔍 {html.escape(query)}</span> "
            f"<span style='background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS}; font-weight: bold; padding: 2px 8px; border-radius: 4px; border: 1px solid {C.SUCCESS_BORDER_MEDIUM}; font-size: 11px;'>{total_clauses_count} Clauses</span> "
            f"<span style='background-color: {C.BG_CANVAS}; color: {C.TEXT_BODY}; font-weight: bold; padding: 2px 8px; border-radius: 4px; border: 1px solid {C.BORDER_STRONG}; font-size: 11px;'>{specs_with_hits} Spec(s)</span>"
            f"{date_badge}"
        )

    def _on_tab_table_clicked(self, index, spec_num: str, table: QTableView):
        model = table.model()
        if not model:
            return

        row, col = index.row(), index.column()
        clause_num = str(model.data(model.index(row, 0), Qt.DisplayRole) or "").strip()
        clause_title = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip()
        query = self.search_input.text().strip()

        # Resolve target version header
        if col >= 2:
            ver_header = str(model._version_cols[col - 2])
        else:
            ver_header = str(model._version_cols[0])

        ver_match = re.search(r"v([0-9\.]+)", ver_header)
        version = ver_match.group(1) if ver_match else ver_header.lstrip("v")

        # Resolve introductory version & preceding version
        first_added_ver_header = model.get_first_added_version(row)
        first_added_ver = None
        if first_added_ver_header:
            m = re.search(r"v([0-9\.]+)", first_added_ver_header)
            first_added_ver = m.group(1) if m else first_added_ver_header.lstrip("v")

        prev_ver_header = model.get_preceding_version(ver_header)
        prev_ver = None
        if prev_ver_header:
            m = re.search(r"v([0-9\.]+)", prev_ver_header)
            prev_ver = m.group(1) if m else prev_ver_header.lstrip("v")

        # 1. Fetch by direct SQLite primary key (clause_pk)
        clause_data = None
        target_pk = model.get_clause_pk(row, ver_header)
        if target_pk is None:
            target_pk = model.get_row_first_pk(row)

        if target_pk is not None:
            clause_data = self.db.get_clause_content(target_pk)

        # 2. Fallback to longest content query by spec/version/clause
        if not clause_data:
            clause_data = self.db.get_clause_content_by_spec_ver(spec_num, version, clause_num)

        if clause_data:
            spec_title = self.get_spec_short_name(spec_num) or self.get_spec_title(spec_num)
            self.inspector.display_clause(
                clause_number=clause_num,
                clause_title=clause_title,
                spec_number=clause_data.get("spec_number", spec_num),
                version=clause_data.get("version", version),
                release_date=clause_data.get("release_date"),
                content=clause_data.get("content", ""),
                search_query=query,
                first_added_ver=first_added_ver,
                prev_ver=prev_ver,
                spec_title=spec_title,
            )
