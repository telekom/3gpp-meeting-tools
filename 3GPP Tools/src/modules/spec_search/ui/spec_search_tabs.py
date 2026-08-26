"""
Main Tab for Specification Substring Search, Release Evolution Tracking, and Cutoff Date Analysis.
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
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core.utils.paths import get_project_root
from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.spec_search.core.spec_search_threads import SpecSearchImportThread, SpecSearchQueryWorker
from modules.spec_search.ui.spec_search_components import SpecClauseInspector, SpecSearchVersionTreeWidget
from modules.spec_search.ui.spec_search_dialogs import SpecSearchVersionSelectDialog
from modules.spec_search.ui.spec_search_models import SpecEvolutionMatrixModel
from modules.specifications.core.database import SpecsDatabase


class SpecSearchTab(QWidget):
    """Main Search Tab: Tree, Query Inputs, Cutoff Date Filter, Evolution Matrix, and Inspector."""

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

        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(300)
        self._search_timer.timeout.connect(self._execute_search)

        self._setup_ui()
        self.refresh_versions()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Toolbar
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

        # Splitter Layout
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

        # Right Panel: Search Controls, Matrix Table, Inspector
        right_splitter = QSplitter(Qt.Vertical)

        matrix_widget = QWidget()
        matrix_layout = QVBoxLayout(matrix_widget)
        matrix_layout.setContentsMargins(0, 0, 0, 0)

        # Row 1: Search substring & Clause Filter
        search_bar_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search exact substring or phrase (e.g. 'initial registration', 'emergency', 'PDU session')...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(lambda: self._search_timer.start())
        search_bar_layout.addWidget(self.search_input)

        self.clause_filter_input = QLineEdit()
        self.clause_filter_input.setPlaceholderText("Filter clause (e.g. 5.2, 8.1)...")
        self.clause_filter_input.setFixedWidth(160)
        self.clause_filter_input.setClearButtonEnabled(True)
        self.clause_filter_input.textChanged.connect(lambda: self._search_timer.start())
        search_bar_layout.addWidget(self.clause_filter_input)
        matrix_layout.addLayout(search_bar_layout)

        # Row 2: Cutoff Date Bar
        cutoff_date_bar = QHBoxLayout()
        self.chk_cutoff_date = QCheckBox("🎯 Date Cutoff:")
        self.chk_cutoff_date.toggled.connect(self._on_cutoff_date_filter_toggled)
        cutoff_date_bar.addWidget(self.chk_cutoff_date)

        self.cutoff_date_edit = QDateEdit()
        self.cutoff_date_edit.setCalendarPopup(True)
        self.cutoff_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.cutoff_date_edit.setDate(QDate.currentDate().addYears(-3))
        self.cutoff_date_edit.setEnabled(False)
        self.cutoff_date_edit.dateChanged.connect(lambda: self._search_timer.start())
        cutoff_date_bar.addWidget(self.cutoff_date_edit)

        self.chk_post_date_only = QCheckBox("Show Only Post-Cutoff Additions")
        self.chk_post_date_only.setToolTip("Hides clauses where matching text was already present prior to the selected date.")
        self.chk_post_date_only.setEnabled(False)
        self.chk_post_date_only.toggled.connect(lambda: self._search_timer.start())
        cutoff_date_bar.addWidget(self.chk_post_date_only)

        cutoff_date_bar.addStretch()
        matrix_layout.addLayout(cutoff_date_bar)

        self.matrix_title = QLabel("Type a query above to see Release Evolution Matrix")
        self.matrix_title.setStyleSheet("font-weight: bold; font-size: 12px; color: #1E293B; margin-top: 4px;")
        matrix_layout.addWidget(self.matrix_title)

        self.matrix_table = QTableView()
        self.matrix_table.setAlternatingRowColors(True)
        self.matrix_table.setSelectionBehavior(QTableView.SelectRows)
        self.matrix_table.setSelectionMode(QTableView.SingleSelection)
        self.matrix_table.verticalHeader().setDefaultSectionSize(24)
        self.matrix_table.verticalHeader().setVisible(False)
        self.matrix_table.clicked.connect(self._on_table_clicked)
        matrix_layout.addWidget(self.matrix_table)
        right_splitter.addWidget(matrix_widget)

        self.inspector = SpecClauseInspector()
        right_splitter.addWidget(self.inspector)

        right_splitter.setSizes([400, 250])
        main_splitter.addWidget(right_splitter)
        main_splitter.setSizes([260, 690])
        layout.addWidget(main_splitter)

    def _on_cutoff_date_filter_toggled(self, checked: bool):
        self.cutoff_date_edit.setEnabled(checked)
        self.chk_post_date_only.setEnabled(checked)
        self._search_timer.start()

    # -------------------------------------------------------------------------
    # Search Execution
    # -------------------------------------------------------------------------

    def _on_version_selection_changed(self):
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        self._execute_search()

    def _execute_search(self):
        query = self.search_input.text().strip()
        clause_filter = self.clause_filter_input.text().strip()

        if not self.selected_version_ids or not query:
            self.matrix_table.setModel(None)
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
        if df.empty:
            self.matrix_table.setModel(None)
            self.matrix_title.setText(f"No matching clauses found for '{query}' across selected releases.")
            self.inspector.clear_display()
            return

        cutoff_date = self.cutoff_date_edit.date().toString("yyyy-MM-dd") if self.chk_cutoff_date.isChecked() else None
        only_post_date = self.chk_post_date_only.isChecked() if self.chk_cutoff_date.isChecked() else False

        model = SpecEvolutionMatrixModel(
            raw_df=df,
            cutoff_date=cutoff_date,
            only_added_after_cutoff=only_post_date,
        )
        self.matrix_table.setModel(model)

        h_header = self.matrix_table.horizontalHeader()
        h_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        h_header.setSectionResizeMode(1, QHeaderView.Stretch)
        for c in range(2, model.columnCount()):
            h_header.setSectionResizeMode(c, QHeaderView.Interactive)
            self.matrix_table.setColumnWidth(c, 120)

        date_suffix = f" (Post-Date Cutoff: ≥ {cutoff_date})" if cutoff_date and only_post_date else ""
        self.matrix_title.setText(f"Found matches in {len(model._pivot_df)} clause(s) for query: '{query}'{date_suffix}")

    def _on_table_clicked(self, index):
        model = self.matrix_table.model()
        if not model:
            return

        row, col = index.row(), index.column()
        clause_num = str(model.data(model.index(row, 0), Qt.DisplayRole) or "").strip()
        clause_title = str(model.data(model.index(row, 1), Qt.DisplayRole) or "").strip()
        query = self.search_input.text().strip()

        # Resolve target version
        if col >= 2:
            ver_header = str(model._version_cols[col - 2])
        else:
            ver_header = str(model._version_cols[0])

        spec_match = re.search(r"TS\s+([0-9\.]+)", ver_header)
        ver_match = re.search(r"v([0-9\.]+)", ver_header)

        spec_num = spec_match.group(1) if spec_match else ""
        version = ver_match.group(1) if ver_match else ver_header.lstrip("v")

        # Fetch clause content
        clause_data = self.db.get_clause_content_by_spec_ver(spec_num, version, clause_num) if spec_num else None
        if not clause_data:
            df = getattr(model, "_raw_df", None)
            if df is not None and not df.empty:
                matches = df[(df["clause_number"] == clause_num) & (df["version"] == version)]
                if not matches.empty:
                    pk = int(matches.iloc[0]["clause_pk"])
                    clause_data = self.db.get_clause_content(pk)

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
        self.version_tree.populate(versions)
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

        from modules.spec_search.core.spec_search_threads import is_change_mark_file

        # Filter out revision marked (-rm) files
        clean_paths = [p for p in paths if not is_change_mark_file(p)]
        if len(clean_paths) < len(paths):
            skipped = len(paths) - len(clean_paths)
            self.log_msg.emit(f"ℹ️ Skipped {skipped} revision mark (-rm) file(s).", logging.INFO)

        grouped: Dict[str, List[Path]] = {}
        for fp in clean_paths:
            p = Path(fp)
            base_key = re.sub(r"_\d+_.*$", "", p.stem)
            # Remove clean suffix if present so grouping matches the base version
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
            self._execute_search()

    def _delete_spec_group(self, spec_number: str, count: int):
        if QMessageBox.question(self, "Confirm Delete", f"Delete all {count} version(s) of TS {spec_number}?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for item in self.version_tree.get_checked_versions_info():
                if item.get("spec_number") == spec_number:
                    self.db.clear_version(spec_number, item["version"])
            self.refresh_versions()
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
            self._execute_search()

    def _on_wipe_db_clicked(self):
        if QMessageBox.critical(self, "Confirm Wipe", "Delete all indexed specification clauses?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.wipe_database()
            self.refresh_versions()
            self.matrix_table.setModel(None)
            self.inspector.clear_display()
            self.log_msg.emit("🧹 Specification Search DB wiped.", logging.INFO)