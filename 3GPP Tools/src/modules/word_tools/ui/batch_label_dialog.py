# --- File: src/modules/word_tools/ui/batch_label_dialog.py ---
"""
Batch Sensitivity Label Dialog.
Allows delegates to batch-apply configured Microsoft Purview / MIP sensitivity labels
(e.g., OFFEN, INTERN, VERTRAULICH) across an arbitrary set of Word documents or an entire folder.
"""
import os
import platform
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import QThread, Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SEARCH_INPUT_STYLE,
)
from modules.word_tools.core.sensitivity_label import (
    batch_set_sensitivity_label,
    get_document_sensitivity_label,
)
from modules.word_tools.core.word_config import WordConfig


class BatchSensitivityLabelThread(QThread):
    """Background worker thread to batch apply sensitivity labels without freezing the UI."""

    progress_signal = pyqtSignal(int, int, str, str)  # current, total, filename, status
    finished_signal = pyqtSignal(dict)  # summary dict
    error_signal = pyqtSignal(str)

    def __init__(
        self,
        folder_path: Optional[str] = None,
        file_paths: Optional[List[str]] = None,
        label_id: str = "",
        label_name: str = "",
        overwrite_existing: bool = False,
        recursive: bool = False,
        parent=None,
    ):
        super().__init__(parent)
        self.folder_path = folder_path
        self.file_paths = file_paths
        self.label_id = label_id
        self.label_name = label_name
        self.overwrite_existing = overwrite_existing
        self.recursive = recursive

    def run(self):
        try:
            def _cb(curr, tot, fname, status):
                self.progress_signal.emit(curr, tot, fname, status)

            summary = batch_set_sensitivity_label(
                folder_path=self.folder_path,
                file_paths=self.file_paths,
                label_id=self.label_id,
                label_name=self.label_name,
                overwrite_existing=self.overwrite_existing,
                recursive=self.recursive,
                progress_callback=_cb,
            )
            self.finished_signal.emit(summary)
        except Exception as e:
            self.error_signal.emit(str(e))


class BatchSensitivityLabelDialog(QDialog):
    """UI Dialog for applying configured sensitivity labels in bulk to a folder or selected documents."""

    def __init__(self, parent=None, initial_folder: Optional[str] = None, initial_files: Optional[List[str]] = None):
        super().__init__(parent)
        self.setWindowTitle("🛡️ Batch Sensitivity Label Applicator")
        self.setMinimumWidth(680)
        self.setMinimumHeight(540)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setAcceptDrops(True)
        self._thread: Optional[BatchSensitivityLabelThread] = None
        self._selected_files: List[Path] = []
        self._setup_ui()

        if initial_folder and Path(initial_folder).exists():
            self.edit_folder.setText(str(Path(initial_folder).resolve()))
        if initial_files:
            self._add_file_paths([Path(f) for f in initial_files])
            self.tabs.setCurrentIndex(1)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        # Header description
        desc_lbl = QLabel(
            "<b>Batch apply enterprise sensitivity labels</b> across a set of Word documents or an entire folder "
            "to comply with corporate DLP policies and prevent email or save restrictions."
        )
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet("color: #334155; font-size: 12px; margin-bottom: 2px;")
        main_layout.addWidget(desc_lbl)

        # Dual Mode Tabs: Folder vs Selected Documents
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                background-color: #FFFFFF;
                top: -1px;
            }
            QTabBar::tab {
                background: #F1F5F9;
                border: 1px solid #CBD5E1;
                padding: 6px 16px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                color: #475569;
                font-size: 12px;
            }
            QTabBar::tab:selected {
                background: #FFFFFF;
                border-bottom-color: #FFFFFF;
                color: #0F172A;
                font-weight: bold;
            }
        """)

        # Tab 1: Folder Mode
        tab_folder = QWidget()
        folder_layout = QVBoxLayout(tab_folder)
        folder_layout.setContentsMargins(12, 12, 12, 12)
        folder_layout.setSpacing(10)

        f_row = QHBoxLayout()
        self.edit_folder = QLineEdit()
        self.edit_folder.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_folder.setPlaceholderText("Select folder containing Word documents...")
        f_row.addWidget(self.edit_folder, 1)

        self.btn_browse_folder = QPushButton("📂 Browse Folder...")
        self.btn_browse_folder.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_browse_folder.clicked.connect(self._browse_folder)
        f_row.addWidget(self.btn_browse_folder)
        folder_layout.addLayout(f_row)

        self.chk_recursive = QCheckBox("Include subfolders recursively")
        self.chk_recursive.setChecked(False)
        folder_layout.addWidget(self.chk_recursive)
        folder_layout.addStretch()
        self.tabs.addTab(tab_folder, "📁 Entire Folder")

        # Tab 2: Document Selection Mode
        tab_files = QWidget()
        files_layout = QVBoxLayout(tab_files)
        files_layout.setContentsMargins(12, 12, 12, 12)
        files_layout.setSpacing(8)

        btn_files_bar = QHBoxLayout()
        self.btn_add_files = QPushButton("➕ Add Documents...")
        self.btn_add_files.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_add_files.clicked.connect(self._browse_files)
        btn_files_bar.addWidget(self.btn_add_files)

        self.btn_remove_file = QPushButton("🗑️ Remove")
        self.btn_remove_file.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_remove_file.clicked.connect(self._remove_selected_files)
        btn_files_bar.addWidget(self.btn_remove_file)

        self.btn_clear_files = QPushButton("Clear All")
        self.btn_clear_files.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_clear_files.clicked.connect(self._clear_files)
        btn_files_bar.addWidget(self.btn_clear_files)

        self.lbl_file_count = QLabel("0 documents selected (drag & drop files here)")
        self.lbl_file_count.setStyleSheet("font-size: 11px; color: #64748B; margin-left: 8px;")
        btn_files_bar.addWidget(self.lbl_file_count, 1)
        files_layout.addLayout(btn_files_bar)

        self.tbl_files = QTableWidget()
        self.tbl_files.setColumnCount(3)
        self.tbl_files.setHorizontalHeaderLabels(["Document Name", "Size", "Current Classification"])
        self.tbl_files.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tbl_files.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tbl_files.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl_files.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_files.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl_files.setStyleSheet("""
            QTableWidget {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                background-color: #F8FAFC;
                font-size: 11px;
            }
            QHeaderView::section {
                background-color: #F1F5F9;
                padding: 4px 8px;
                font-weight: bold;
                border: none;
                border-bottom: 1px solid #CBD5E1;
            }
        """)
        files_layout.addWidget(self.tbl_files)
        self.tabs.addTab(tab_files, "📄 Specific Documents")

        main_layout.addWidget(self.tabs)

        # Settings Card (Target Label & Options)
        cfg_card = QFrame()
        cfg_card.setStyleSheet(CARD_FRAME_STYLE)
        cfg_layout = QFormLayout(cfg_card)
        cfg_layout.setContentsMargins(12, 10, 12, 10)
        cfg_layout.setSpacing(8)

        self.combo_label = QComboBox()
        self.combo_label.setStyleSheet(SEARCH_INPUT_STYLE)
        self._populate_labels()
        cfg_layout.addRow("<b>Target Sensitivity Level:</b>", self.combo_label)

        self.chk_unlabeled_only = QCheckBox("Only apply to unlabeled documents (recommended - skips already classified files)")
        self.chk_unlabeled_only.setChecked(True)
        cfg_layout.addRow("", self.chk_unlabeled_only)
        main_layout.addWidget(cfg_card)

        # Progress Section
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("Ready.")
        self.lbl_status.setStyleSheet("font-size: 11px; color: #64748B;")
        main_layout.addWidget(self.lbl_status)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumHeight(110)
        self.log_view.setStyleSheet("font-family: Consolas, monospace; font-size: 11px; background-color: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 4px;")
        self.log_view.setVisible(False)
        main_layout.addWidget(self.log_view)

        # Bottom Button Row
        btn_row = QHBoxLayout()
        self.btn_open_folder = QPushButton("📂 Open Folder")
        self.btn_open_folder.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_folder.clicked.connect(self._open_target_folder)
        self.btn_open_folder.setVisible(False)
        btn_row.addWidget(self.btn_open_folder)

        btn_row.addStretch()

        self.btn_cancel = QPushButton("Close")
        self.btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(self.btn_cancel)

        self.btn_apply = QPushButton("🚀 Apply Sensitivity Labels")
        self.btn_apply.setObjectName("primaryBtn")
        self.btn_apply.clicked.connect(self._start_batch_labeling)
        btn_row.addWidget(self.btn_apply)

        main_layout.addLayout(btn_row)

    def _populate_labels(self):
        self.combo_label.clear()
        labels = WordConfig.get_sensitivity_labels()
        active = WordConfig.get_active_sensitivity_label()
        active_id = active.get("id", "")

        sel_idx = 0
        for idx, lbl in enumerate(labels):
            name = lbl.get("name", "Unknown")
            lid = lbl.get("id", "")
            desc = lbl.get("description", "")
            label_text = f"{name} ({desc})" if desc else name
            self.combo_label.addItem(label_text, (lid, name))
            if lid == active_id:
                sel_idx = idx

        self.combo_label.setCurrentIndex(sel_idx)

    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder for Sensitivity Labeling", self.edit_folder.text())
        if folder:
            self.edit_folder.setText(folder)

    def _browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Word Documents to Label",
            "",
            "Word Documents (*.docx *.doc *.docm);;All Files (*.*)",
        )
        if files:
            self._add_file_paths([Path(f) for f in files])

    def _add_file_paths(self, paths: List[Path]):
        existing_set = {p.resolve() for p in self._selected_files}
        for p in paths:
            rp = p.resolve()
            if rp.is_file() and rp.suffix.lower() in {".docx", ".doc", ".docm"} and rp not in existing_set:
                self._selected_files.append(rp)
                existing_set.add(rp)
        self._refresh_file_table()

    def _remove_selected_files(self):
        selected_rows = sorted({idx.row() for idx in self.tbl_files.selectedIndexes()}, reverse=True)
        for r in selected_rows:
            if 0 <= r < len(self._selected_files):
                self._selected_files.pop(r)
        self._refresh_file_table()

    def _clear_files(self):
        self._selected_files.clear()
        self._refresh_file_table()

    def _refresh_file_table(self):
        self.tbl_files.setRowCount(len(self._selected_files))
        for row, p in enumerate(self._selected_files):
            # File name
            it_name = QTableWidgetItem(p.name)
            it_name.setToolTip(str(p))
            self.tbl_files.setItem(row, 0, it_name)

            # Size
            size_str = ""
            if p.exists():
                sz_kb = p.stat().st_size / 1024
                size_str = f"{sz_kb:.1f} KB" if sz_kb < 1024 else f"{sz_kb/1024:.2f} MB"
            self.tbl_files.setItem(row, 1, QTableWidgetItem(size_str))

            # Inspect sensitivity label headlessly
            lbl = get_document_sensitivity_label(p)
            if lbl and lbl.get("name"):
                cls_text = f"🟢 {lbl['name']}"
            else:
                cls_text = "⚪ Unclassified"
            it_cls = QTableWidgetItem(cls_text)
            self.tbl_files.setItem(row, 2, it_cls)

        count = len(self._selected_files)
        self.lbl_file_count.setText(f"{count} document(s) queued for labeling")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        dropped_paths = [Path(u.toLocalFile()) for u in urls if u.isLocalFile()]
        new_files: List[Path] = []
        for p in dropped_paths:
            if p.is_dir():
                # If a folder was dropped, set folder input or scan
                self.edit_folder.setText(str(p.resolve()))
                self.tabs.setCurrentIndex(0)
                return
            elif p.is_file() and p.suffix.lower() in {".docx", ".doc", ".docm"}:
                new_files.append(p)
        if new_files:
            self._add_file_paths(new_files)
            self.tabs.setCurrentIndex(1)
        event.acceptProposedAction()

    def _open_target_folder(self):
        target = ""
        if self.tabs.currentIndex() == 0:
            target = self.edit_folder.text().strip()
        elif self._selected_files:
            target = str(self._selected_files[0].parent)

        if target and Path(target).exists() and hasattr(os, "startfile"):
            os.startfile(target)

    def _start_batch_labeling(self):
        if platform.system() != "Windows":
            QMessageBox.warning(self, "Windows Required", "Microsoft Word COM automation is only available on Windows.")
            return

        data = self.combo_label.currentData()
        if not data:
            QMessageBox.warning(self, "No Label Selected", "Please select a sensitivity label to apply.")
            return

        label_id, label_name = data
        overwrite = not self.chk_unlabeled_only.isChecked()
        mode_idx = self.tabs.currentIndex()

        target_folder = None
        target_files = None

        if mode_idx == 0:
            folder = self.edit_folder.text().strip()
            if not folder or not Path(folder).is_dir():
                QMessageBox.warning(self, "Invalid Folder", "Please select a valid existing directory.")
                return
            target_folder = folder
            recursive = self.chk_recursive.isChecked()
        else:
            if not self._selected_files:
                QMessageBox.warning(self, "No Documents", "Please add at least one Word document to process.")
                return
            target_files = [str(p) for p in self._selected_files]
            recursive = False

        self.btn_apply.setEnabled(False)
        self.btn_browse_folder.setEnabled(False)
        self.btn_add_files.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_view.setVisible(True)
        self.log_view.clear()

        desc_target = f"Folder: {target_folder}" if target_folder else f"{len(target_files)} specific document(s)"
        self.log_view.append(f"Starting batch labeling on {desc_target}")
        self.log_view.append(f"Target Label: {label_name} (GUID: {label_id})")
        self.lbl_status.setText("Scanning documents...")

        self._thread = BatchSensitivityLabelThread(
            folder_path=target_folder,
            file_paths=target_files,
            label_id=label_id,
            label_name=label_name,
            overwrite_existing=overwrite,
            recursive=recursive,
            parent=self,
        )
        self._thread.progress_signal.connect(self._on_progress)
        self._thread.finished_signal.connect(self._on_finished)
        self._thread.error_signal.connect(self._on_error)
        self._thread.start()

    def _on_progress(self, current: int, total: int, filename: str, status: str):
        if total > 0:
            pct = int((current / total) * 100)
            self.progress_bar.setValue(pct)
        self.lbl_status.setText(f"Processing ({current}/{total}): {filename}")
        self.log_view.append(f"[{current}/{total}] {filename}: {status}")

    def _on_finished(self, summary: dict):
        self.progress_bar.setValue(100)
        self.btn_apply.setEnabled(True)
        self.btn_browse_folder.setEnabled(True)
        self.btn_add_files.setEnabled(True)
        self.btn_open_folder.setVisible(True)

        tot = summary.get("total", 0)
        lab = summary.get("labeled", 0)
        skp = summary.get("skipped", 0)
        err = summary.get("errors", 0)

        status_msg = f"Completed: {tot} scanned, {lab} newly labeled, {skp} skipped, {err} errors."
        self.lbl_status.setText(status_msg)
        self.log_view.append(f"\n✅ {status_msg}")

        # Refresh file table classifications if in document mode
        if self.tabs.currentIndex() == 1:
            self._refresh_file_table()

        QMessageBox.information(
            self,
            "Batch Sensitivity Labeling Complete",
            f"<b>Batch Labeling Finished!</b><br><br>"
            f"• <b>Total documents processed:</b> {tot}<br>"
            f"• <b>Successfully labeled:</b> {lab}<br>"
            f"• <b>Skipped (already classified):</b> {skp}<br>"
            f"• <b>Errors encountered:</b> {err}",
        )

    def _on_error(self, err_msg: str):
        self.btn_apply.setEnabled(True)
        self.btn_browse_folder.setEnabled(True)
        self.btn_add_files.setEnabled(True)
        self.lbl_status.setText(f"Failed: {err_msg}")
        self.log_view.append(f"\n❌ Error: {err_msg}")
        QMessageBox.critical(self, "Batch Labeling Error", f"An error occurred during batch labeling:\n\n{err_msg}")
