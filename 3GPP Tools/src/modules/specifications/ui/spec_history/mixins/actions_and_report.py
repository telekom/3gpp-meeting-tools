"""Clipboard export, TDoc opening, local re-scan, and dossier report generation mixin."""

import logging
from typing import Any, Dict

from PyQt5.QtCore import QPoint, Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QMenu,
    QMessageBox,
    QProgressDialog,
    QTableWidgetItem,
    QToolTip,
)

from modules.specifications.core.spec_history_extractor import (
    SpecHistoryExtractThread,
    SpecHistoryExtractor,
)
from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
from modules.specifications.core.spec_tdoc_report import SpecTDocReportThread
from modules.specifications.ui.spec_tdoc_report_dialog import SpecTDocReportDialog

logger = logging.getLogger(__name__)


class _HistoryActionsMixin:
    """Handles table context actions, TDoc clipboard copy, re-scan, and PDF/Markdown dossier launch."""

    def _copy_selected_entry_tdocs(self):
        """Copies all TDocs referenced in the currently selected row."""
        row_idx = self.table.currentRow()
        if row_idx < 0 or row_idx >= len(self.filtered_records):
            return
        self._copy_entry_tdocs(self.filtered_records[row_idx])

    def _open_selected_tdoc(self):
        """Opens the primary or first TDoc of the currently selected row."""
        row_idx = self.table.currentRow()
        if row_idx < 0 or row_idx >= len(self.filtered_records):
            return
        rec = self.filtered_records[row_idx]
        primary = rec.get("source_tdoc", "")
        if (not primary or primary == "-") and rec.get("tdocs"):
            primary = rec["tdocs"][0]["raw_tdoc"]
        if primary and primary != "-":
            self._open_tdoc(primary, rec.get("meeting", ""))

    def _on_row_double_clicked(self, item: QTableWidgetItem):
        self._open_selected_tdoc()

    def _open_tdoc(self, tdoc_id: str, meeting_name: str = ""):
        """Dispatches direct opening of the target TDoc Word/ZIP file."""
        SpecTDocOpener.open_tdoc(tdoc_id, meeting_name, parent=self)

    def _show_context_menu(self, pos: QPoint):
        """Context menu on table rows."""
        row_idx = self.table.currentRow()
        if row_idx < 0 or row_idx >= len(self.filtered_records):
            return

        rec = self.filtered_records[row_idx]
        primary_tdoc = rec.get("source_tdoc", "")
        if not primary_tdoc and rec.get("tdocs"):
            primary_tdoc = rec["tdocs"][0]["raw_tdoc"]

        menu = QMenu(self)
        if primary_tdoc:
            act_open = menu.addAction(f"📄 Open {primary_tdoc}")
            act_open.triggered.connect(lambda: self._open_tdoc(primary_tdoc, rec.get("meeting", "")))

            act_jump = menu.addAction(f"🚀 Jump to {primary_tdoc} in Meetings Tab")
            act_jump.triggered.connect(lambda: SpecTDocOpener.jump_to_meetings_tab(primary_tdoc))

            menu.addSeparator()
            act_copy = menu.addAction(f"📋 Copy Citation ({primary_tdoc})")
            act_copy.triggered.connect(
                lambda: SpecTDocOpener.copy_tdoc_citation(
                    primary_tdoc, self.spec_number, rec.get("version", ""), rec.get("meeting", ""), rec.get("subject", ""), parent=self
                )
            )

        act_copy_all = menu.addAction("📋 Copy All TDocs for this Entry")
        act_copy_all.triggered.connect(lambda: self._copy_entry_tdocs(rec))

        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def _copy_entry_tdocs(self, rec: Dict[str, Any]):
        """Copies all TDoc tokens from a single entry."""
        from modules.meetings.ui.tdocs_menus import _copy_to_clipboard

        tdocs = [t["raw_tdoc"] for t in rec.get("tdocs", [])]
        primary = rec.get("source_tdoc", "")
        if primary and primary not in tdocs:
            tdocs.insert(0, primary)

        plain = ", ".join(tdocs)
        _copy_to_clipboard(plain, label=f"{len(tdocs)} TDoc(s)", parent=self)

    def _copy_tdocs_list(self):
        """Copies all TDocs for the active version filter as a comma-separated list."""
        from modules.meetings.ui.tdocs_menus import _copy_to_clipboard

        all_tokens = []
        for r in self.filtered_records:
            primary = r.get("source_tdoc", "")
            if primary and primary != "-" and primary not in all_tokens:
                all_tokens.append(primary)
            for t in r.get("tdocs", []):
                tok = t.get("raw_tdoc", "")
                if tok and tok not in all_tokens:
                    all_tokens.append(tok)

        if not all_tokens:
            QToolTip.showText(QCursor.pos(), "No TDocs to copy.", self)
            return

        plain = ", ".join(all_tokens)
        _copy_to_clipboard(plain, label=f"{len(all_tokens)} TDoc(s)", parent=self)

    def _open_spec_document(self):
        """Locates and opens the local specification document."""
        from modules.word_tools.core.document_opener import open_document

        target = SpecHistoryExtractor.resolve_spec_document(
            self.spec_number, self.version_combo.currentData() or ""
        )
        if target and target.exists():
            open_document(target)
        else:
            QToolTip.showText(QCursor.pos(), f"No local document found for {self.spec_number}", self)

    def _on_rescan_clicked(self):
        """Executes background re-extraction from local document."""
        self.rescan_btn.setEnabled(False)
        self.rescan_btn.setText("⏳ Scanning...")

        self.extract_thread = SpecHistoryExtractThread(
            spec_number=self.spec_number,
            db=self.history_db,
            version=self.version_combo.currentData() or ""
        )

        def _on_success(s_num, count):
            self.rescan_btn.setEnabled(True)
            self.rescan_btn.setText("🔄 Re-scan from Local Doc")
            self._load_data(auto_extract_if_empty=False)
            QToolTip.showText(QCursor.pos(), f"✅ Updated: {count} records indexed.", self)

        def _on_error(err):
            self.rescan_btn.setEnabled(True)
            self.rescan_btn.setText("🔄 Re-scan from Local Doc")
            QToolTip.showText(QCursor.pos(), f"❌ Scan failed: {err}", self)

        self.extract_thread.finished_success.connect(_on_success)
        self.extract_thread.error.connect(_on_error)
        self.extract_thread.start()

    def _generate_tdoc_report(self, scope: str = "version"):
        """
        Gathers TDocs referenced in the specified scope, resolves their technical metadata
        from meeting spreadsheets in background, and presents the interactive dossier dialog.
        """
        tdoc_set = set()
        default_meeting = ""

        if scope == "selection":
            selected_rows = self.table.selectionModel().selectedRows()
            if not selected_rows:
                QMessageBox.information(self, "No Rows Selected", "Please select one or more rows in the Change History table first.")
                return
            for idx in selected_rows:
                row_idx = idx.row()
                item = self.table.item(row_idx, 0)
                if item:
                    r = item.data(Qt.UserRole)
                    if r:
                        if r.get("source_tdoc"):
                            tdoc_set.add(r["source_tdoc"])
                        for t in r.get("tdocs", []):
                            tdoc_set.add(t.get("raw_tdoc") or t.get("tdoc"))
                        if not default_meeting and r.get("meeting"):
                            default_meeting = r["meeting"]
        elif scope == "all":
            for r in self.all_records:
                if r.get("source_tdoc"):
                    tdoc_set.add(r["source_tdoc"])
                for t in r.get("tdocs", []):
                    tdoc_set.add(t.get("raw_tdoc") or t.get("tdoc"))
                if not default_meeting and r.get("meeting"):
                    default_meeting = r["meeting"]
        else:  # "version"
            target_ver = self.version_combo.currentData() or ""
            target_records = [r for r in self.all_records if not target_ver or r.get("version") == target_ver]
            for r in target_records:
                if r.get("source_tdoc"):
                    tdoc_set.add(r["source_tdoc"])
                for t in r.get("tdocs", []):
                    tdoc_set.add(t.get("raw_tdoc") or t.get("tdoc"))
                if not default_meeting and r.get("meeting"):
                    default_meeting = r["meeting"]

        clean_tdocs = sorted([t.strip() for t in tdoc_set if t and t.strip() != "-"])
        if not clean_tdocs:
            QMessageBox.information(
                self,
                "No TDocs Found",
                f"No TDocs or p-CRs are referenced in the selected scope for TS {self.spec_number}."
            )
            return

        active_ver = self.version_combo.currentData() or (self.filtered_records[0].get("version", "") if self.filtered_records else "")
        ver_label = active_ver if scope != "all" else "all"

        logger.info(f"📑 [TDoc Report] Starting dossier generation for TS {self.spec_number} v{ver_label} ({len(clean_tdocs)} TDocs)...")

        # Progress Dialog
        progress_dlg = QProgressDialog(
            f"Resolving metadata for {len(clean_tdocs)} TDocs from meeting agendas...",
            "Cancel",
            0,
            100,
            self
        )
        progress_dlg.setWindowTitle(f"Generating Technical Dossier: TS {self.spec_number}")
        progress_dlg.setMinimumDuration(150)
        progress_dlg.setValue(5)

        thread = SpecTDocReportThread(
            spec_number=self.spec_number,
            version=ver_label,
            tdoc_list=clean_tdocs,
            spec_title=self.title_lbl.text(),
            default_meeting=default_meeting,
            parent=self
        )

        def _on_prog(pct: int, msg: str):
            if not progress_dlg.wasCanceled():
                progress_dlg.setValue(pct)
                progress_dlg.setLabelText(msg)

        def _on_finish(results: list, md_report: str):
            progress_dlg.close()
            logger.info(f"✅ [TDoc Report] Generated dossier with {len(results)} TDocs for TS {self.spec_number} v{ver_label}.")
            report_dlg = SpecTDocReportDialog(
                spec_number=self.spec_number,
                version=ver_label,
                tdoc_entries=results,
                markdown_report=md_report,
                spec_title=self.title_lbl.text(),
                parent=self._parent_ref or self
            )
            self._report_dialog = report_dlg
            report_dlg.show()
            report_dlg.raise_()
            report_dlg.activateWindow()

        def _on_err(err_msg: str):
            progress_dlg.close()
            logger.error(f"❌ [TDoc Report] Error generating dossier: {err_msg}")
            QMessageBox.critical(self, "Report Generation Error", f"Failed to generate TDoc report:\n{err_msg}")

        thread.progress.connect(_on_prog)
        thread.finished.connect(_on_finish)
        thread.error.connect(_on_err)
        progress_dlg.canceled.connect(thread.requestInterruption)

        self._active_report_thread = thread
        thread.start()
