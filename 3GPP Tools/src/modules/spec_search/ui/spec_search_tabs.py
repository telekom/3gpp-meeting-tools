"""
Main Tab for Specification Substring Search, Release Evolution Tracking, and Cutoff Date Analysis.
Visualizes multi-specification search results in dedicated per-specification tabs with automatic
persistence of active filters and selected specification versions.
"""

import json
import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional
from PyQt5.QtCore import QDate, Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QCheckBox,
    QDateEdit,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core.utils.paths import get_project_root
from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.spec_search.core.spec_search_threads import (
    SpecSearchImportThread,
    SpecSearchQueryWorker,
    SpecSearchWipeWorker,
    is_change_mark_file,
)
from modules.spec_search.ui.spec_search_components import SpecClauseInspector, SpecSearchVersionTreeWidget
from modules.spec_search.ui.spec_search_dialogs import SpecSearchVersionSelectDialog
from modules.spec_search.ui.spec_search_models import SpecEvolutionMatrixModel
from modules.specifications.core.database import SpecsDatabase


class SpecSearchTab(QWidget):
    """Main Search Tab: Tree, Query Inputs, Cutoff Date Filter, Tabbed Evolution Matrix, and Inspector."""

    log_msg = pyqtSignal(str, int)

    def __init__(self, search_db_path: Path, specs_db_path: Optional[Path] = None):
        super().__init__()
        self.search_db_path = Path(search_db_path)
        self.specs_db_path = Path(specs_db_path) if specs_db_path else None
        self.config_path = self.search_db_path.parent / "spec_search_config.json"

        try:
            settings = MeetingsSettings()
            self.cache_dir = Path(settings.cache_dir).parent / "specs"
        except Exception:
            self.cache_dir = Path.home() / "3GPP_Delegate_Helper" / "specs"

        self.db = SpecSearchDatabase(self.search_db_path)
        self.specs_db = SpecsDatabase(self.specs_db_path) if self.specs_db_path and self.specs_db_path.exists() else None

        self._query_worker: Optional[SpecSearchQueryWorker] = None
        self._current_req_id: int = 0
        self.selected_version_ids: List[int] = []

        # Debounce timer for search execution
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(300)
        self._search_timer.timeout.connect(self._execute_search)

        # Debounce timer for saving filter/selection state
        self._save_config_timer = QTimer(self)
        self._save_config_timer.setSingleShot(True)
        self._save_config_timer.setInterval(500)
        self._save_config_timer.timeout.connect(self._save_config)

        self._setup_ui()
        self._load_config_and_refresh()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # ---------------------------------------------------------------------
        # Toolbar
        # ---------------------------------------------------------------------
        toolbar = QHBoxLayout()
        self.fetch_btn = QPushButton("📥 Import from Specs DB")
        self.fetch_btn.clicked.connect(self._on_fetch_clicked)
        toolbar.addWidget(self.fetch_btn)

        self.import_local_btn = QPushButton("📁 Import Local .docx")
        self.import_local_btn.clicked.connect(self._on_import_local_clicked)
        toolbar.addWidget(self.import_local_btn)

        self.clear_ver_btn = QPushButton("🗑️ Clear Version")
        self.clear_ver_btn.clicked.connect(self._on_clear_version_clicked)
        toolbar.addWidget(self.clear_ver_btn)

        self.wipe_db_btn = QPushButton("⚠️ Wipe DB")
        self.wipe_db_btn.setStyleSheet("color: #D32F2F; font-weight: bold;")
        self.wipe_db_btn.clicked.connect(self._on_wipe_db_clicked)
        toolbar.addWidget(self.wipe_db_btn)

        toolbar.addStretch()
        layout.addLayout(toolbar)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # ---------------------------------------------------------------------
        # Splitter Layout (Left: Version Tree, Right: Tabs & Inspector)
        # ---------------------------------------------------------------------
        main_splitter = QSplitter(Qt.Horizontal)

        # Left Panel: Versions Tree
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        tree_group = QGroupBox("Indexed Specifications && Releases")
        tree_layout = QVBoxLayout(tree_group)
        self.version_tree = SpecSearchVersionTreeWidget()
        self.version_tree.selection_changed.connect(self._on_version_selection_changed)
        self.version_tree.delete_version_requested.connect(self._delete_single_version)
        self.version_tree.delete_spec_requested.connect(self._delete_spec_group)
        tree_layout.addWidget(self.version_tree)
        left_layout.addWidget(tree_group)
        main_splitter.addWidget(left_widget)

        # Right Panel: Search Controls, Tabbed Evolution Matrix, Inspector
        right_splitter = QSplitter(Qt.Vertical)

        matrix_widget = QWidget()
        matrix_layout = QVBoxLayout(matrix_widget)
        matrix_layout.setContentsMargins(0, 0, 0, 0)

        # Search substring & Clause Filter
        search_bar_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search exact substring or phrase (e.g. 'slice replacement', 'emergency', 'PDU session')...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._on_filter_changed)
        search_bar_layout.addWidget(self.search_input)

        self.clause_filter_input = QLineEdit()
        self.clause_filter_input.setPlaceholderText("Filter clause (e.g. 5.2, 8.1)...")
        self.clause_filter_input.setFixedWidth(160)
        self.clause_filter_input.setClearButtonEnabled(True)
        self.clause_filter_input.textChanged.connect(self._on_filter_changed)
        search_bar_layout.addWidget(self.clause_filter_input)
        matrix_layout.addLayout(search_bar_layout)

        # Cutoff Date Filter Bar
        cutoff_date_bar = QHBoxLayout()
        self.chk_filing_date = QCheckBox("🎯 Date Cutoff:")
        self.chk_filing_date.toggled.connect(self._on_cutoff_date_filter_toggled)
        cutoff_date_bar.addWidget(self.chk_filing_date)

        self.cutoff_date_edit = QDateEdit()
        self.cutoff_date_edit.setCalendarPopup(True)
        self.cutoff_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.cutoff_date_edit.setDate(QDate.currentDate().addYears(-3))
        self.cutoff_date_edit.setEnabled(False)
        self.cutoff_date_edit.dateChanged.connect(self._on_filter_changed)
        cutoff_date_bar.addWidget(self.cutoff_date_edit)

        self.chk_post_date_only = QCheckBox("Show Only Post-Cutoff Additions")
        self.chk_post_date_only.setToolTip("Hides clauses where matching text was already present prior to the selected date.")
        self.chk_post_date_only.setEnabled(False)
        self.chk_post_date_only.toggled.connect(self._on_filter_changed)
        cutoff_date_bar.addWidget(self.chk_post_date_only)

        cutoff_date_bar.addStretch()
        matrix_layout.addLayout(cutoff_date_bar)

        self.matrix_title = QLabel("Type a query above to see Release Evolution Matrix")
        self.matrix_title.setStyleSheet("font-weight: bold; font-size: 12px; color: #1E293B; margin-top: 4px;")
        matrix_layout.addWidget(self.matrix_title)

        # Tabbed Container for Per-Specification Matrices
        self.spec_results_tabs = QTabWidget()
        self.spec_results_tabs.setDocumentMode(True)
        matrix_layout.addWidget(self.spec_results_tabs)
        right_splitter.addWidget(matrix_widget)

        self.inspector = SpecClauseInspector()
        right_splitter.addWidget(self.inspector)

        right_splitter.setSizes([420, 240])
        main_splitter.addWidget(right_splitter)
        main_splitter.setSizes([260, 690])
        layout.addWidget(main_splitter)

    # -------------------------------------------------------------------------
    # Configuration Persistence
    # -------------------------------------------------------------------------

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
                "checked_versions": checked_versions,
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
        versions = self.db.get_imported_versions()

        saved_checked = config.get("checked_versions", None) if config else None
        self.version_tree.populate(versions, saved_checked=saved_checked)
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

        # Trigger search if a saved query exists
        if self.search_input.text().strip():
            self._execute_search()

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
            self.matrix_title.setText("Type a query above to see Release Evolution Matrix")
            self.inspector.clear_display()
            return

        self._current_req_id += 1
        req_id = self._current_req_id

        if self._query_worker and self._query_worker.isRunning():
            self._query_worker.terminate()
            self._query_worker.wait()

        self._query_worker = SpecSearchQueryWorker(
            db=self.db,
            query_str=query,
            version_ids=self.selected_version_ids,
            clause_filter=clause_filter,
            request_id=req_id,
        )
        self._query_worker.results_ready.connect(self._on_search_results_ready)
        self._query_worker.start()

    def _on_search_results_ready(self, df, req_id: int):
        if req_id != self._current_req_id:
            return

        query = self.search_input.text().strip()
        self.spec_results_tabs.clear()

        if df.empty:
            self.matrix_title.setText(f"No matching clauses found for '{query}' across selected releases.")
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
            table.setAlternatingRowColors(True)
            table.setSelectionBehavior(QTableView.SelectRows)
            table.setSelectionMode(QTableView.SingleSelection)
            table.verticalHeader().setDefaultSectionSize(24)
            table.verticalHeader().setVisible(False)
            table.setModel(model)

            h_header = table.horizontalHeader()
            h_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            h_header.setSectionResizeMode(1, QHeaderView.Stretch)
            for c in range(2, model.columnCount()):
                h_header.setSectionResizeMode(c, QHeaderView.Interactive)
                table.setColumnWidth(c, 110)

            table.clicked.connect(lambda idx, s=spec_num, t=table: self._on_tab_table_clicked(idx, s, t))

            tab_title = f"TS {spec_num} ({model.rowCount()})"
            self.spec_results_tabs.addTab(table, tab_title)

        if specs_with_hits == 0:
            self.matrix_title.setText(f"No matching clauses found for '{query}' matching the active date cutoff.")
            self.inspector.clear_display()
            return

        date_suffix = f" (Cutoff: ≥ {cutoff_date})" if cutoff_date and only_post_date else ""
        self.matrix_title.setText(
            f"Found matches in {total_clauses_count} clause(s) across {specs_with_hits} specification(s) for query: '{query}'{date_suffix}"
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
            self.inspector.display_clause(
                clause_number=clause_num,
                clause_title=clause_title,
                spec_number=clause_data.get("spec_number", spec_num),
                version=clause_data.get("version", version),
                release_date=clause_data.get("release_date"),
                content=clause_data.get("content", ""),
                search_query=query,
            )

    # -------------------------------------------------------------------------
    # Ingestion Actions
    # -------------------------------------------------------------------------

    def refresh_versions(self):
        versions = self.db.get_imported_versions()
        saved_checked = self.version_tree.get_checked_versions_info()
        if not saved_checked:
            cfg = self._load_config()
            saved_checked = cfg.get("checked_versions") if cfg else None

        self.version_tree.populate(versions, saved_checked=saved_checked)
        self.selected_version_ids = self.version_tree.get_selected_version_ids()

    def _on_fetch_clicked(self):
        if not self.specs_db:
            QMessageBox.warning(self, "Specs DB Unavailable", "The specifications database (3gpp_data.db) is not configured.")
            return

        dialog = SpecSearchVersionSelectDialog(self.specs_db, self.db, self.cache_dir, self)
        if dialog.exec_() == dialog.Accepted and dialog.selected_files_info:
            self._start_batch_ingestion(dialog.selected_files_info)

    def _on_import_local_clicked(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Specification Document(s) (.docx)", "", "Word Files (*.docx)"
        )
        if not paths:
            return

        clean_paths = [p for p in paths if not is_change_mark_file(p)]
        if len(clean_paths) < len(paths):
            skipped = len(paths) - len(clean_paths)
            self.log_msg.emit(f"ℹ️ Skipped {skipped} revision mark (-rm) file(s).", logging.INFO)

        grouped: Dict[str, List[Path]] = {}
        for fp in clean_paths:
            p = Path(fp)
            base_key = re.sub(r"_\d+_.*$", "", p.stem)
            base_key = re.sub(r"[-_]cl$", "", base_key, flags=re.IGNORECASE)
            grouped.setdefault(base_key, []).append(p)

        tasks = []
        for base_key, p_list in grouped.items():
            dispatcher = ProtocolDocxDispatcher(p_list)
            spec_num = dispatcher.extract_spec_number()
            version = dispatcher.extract_version_from_filename()
            tasks.append({
                "spec_number": spec_num,
                "version": version,
                "filename": f"{base_key}.docx",
                "file_url": "",
                "release_date": "",
                "local_docx_paths": p_list,
            })

        self._start_batch_ingestion(tasks)

    def _start_batch_ingestion(self, tasks: List[Dict[str, Any]]):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.fetch_btn.setEnabled(False)
        self.import_local_btn.setEnabled(False)

        self.thread = SpecSearchImportThread(
            search_db_path=self.search_db_path,
            tasks=tasks,
            cache_dir=self.cache_dir,
        )
        self.thread.progress.connect(self._on_import_progress)
        self.thread.finished_success.connect(self._on_import_success)
        self.thread.error.connect(self._on_import_error)
        self.thread.start()

    def _on_import_progress(self, msg: str, val: int):
        self.progress_bar.setValue(val)
        self.log_msg.emit(msg, logging.INFO)

    def _on_import_success(self, specs_count: int, clauses_count: int):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        self.log_msg.emit(f"✅ Indexed {specs_count} specification(s) ({clauses_count} total clauses).", logging.INFO)
        self.refresh_versions()
        self._save_config()

    def _on_import_error(self, err: str):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        QMessageBox.critical(self, "Indexing Error", err)
        self.log_msg.emit(f"❌ Indexing error: {err}", logging.ERROR)

    def _delete_single_version(self, spec_number: str, version: str):
        if QMessageBox.question(self, "Confirm Delete", f"Delete TS {spec_number} v{version}?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.clear_version(spec_number, version)
            self.refresh_versions()
            self._save_config()
            self._execute_search()

    def _delete_spec_group(self, spec_number: str, count: int):
        if QMessageBox.question(self, "Confirm Delete", f"Delete all {count} version(s) of TS {spec_number}?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for item in self.version_tree.get_checked_versions_info():
                if item.get("spec_number") == spec_number:
                    self.db.clear_version(spec_number, item["version"])
            self.refresh_versions()
            self._save_config()
            self._execute_search()

    def _on_clear_version_clicked(self):
        checked = self.version_tree.get_checked_versions_info()
        if not checked:
            QMessageBox.warning(self, "Select Version", "Please check at least one version to clear.")
            return
        if QMessageBox.question(self, "Confirm Delete", f"Delete {len(checked)} checked version(s)?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for item in checked:
                self.db.clear_version(item["spec_number"], item["version"])
            self.refresh_versions()
            self._save_config()
            self._execute_search()

    def _on_wipe_db_clicked(self):
        if QMessageBox.critical(
            self,
            "Confirm Wipe",
            "Delete all indexed specification clauses and full-text indexes?",
            QMessageBox.Yes | QMessageBox.No,
        ) != QMessageBox.Yes:
            return

        if self._query_worker and self._query_worker.isRunning():
            self._query_worker.terminate()
            self._query_worker.wait()

        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.wipe_db_btn.setEnabled(False)
        self.fetch_btn.setEnabled(False)
        self.import_local_btn.setEnabled(False)
        self.clear_ver_btn.setEnabled(False)

        self._wipe_worker = SpecSearchWipeWorker(self.db)
        self._wipe_worker.finished_success.connect(self._on_wipe_db_success)
        self._wipe_worker.error.connect(self._on_wipe_db_error)
        self._wipe_worker.start()

    def _on_wipe_db_success(self):
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.wipe_db_btn.setEnabled(True)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        self.clear_ver_btn.setEnabled(True)

        self.refresh_versions()
        self.spec_results_tabs.clear()
        self.inspector.clear_display()
        self.matrix_title.setText("Type a query above to see Release Evolution Matrix")
        self._save_config()
        self.log_msg.emit("🧹 Specification Search DB wiped successfully.", logging.INFO)

    def _on_wipe_db_error(self, err: str):
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.wipe_db_btn.setEnabled(True)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        self.clear_ver_btn.setEnabled(True)

        QMessageBox.critical(self, "Wipe Error", f"Failed to wipe search database: {err}")
        self.log_msg.emit(f"❌ Wipe error: {err}", logging.ERROR)