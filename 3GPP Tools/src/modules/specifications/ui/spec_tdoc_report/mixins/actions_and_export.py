"""TDoc opener, context menu, and Markdown/CSV export actions mixin."""

from __future__ import annotations

import logging
from pathlib import Path

from PyQt5.QtCore import QPoint, Qt, QTimer
from PyQt5.QtWidgets import QApplication, QFileDialog, QMenu, QMessageBox, QTableWidgetItem

from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener
from modules.specifications.core.spec_tdoc_report import SpecTDocReportGenerator

logger = logging.getLogger(__name__)


class SpecTDocReportActionsMixin:
    """Implements double-click/menu TDoc launching and clipboard/file dossier exports."""

    def _on_table_double_clicked(self, item: QTableWidgetItem):
        """Double clicking any cell in a row opens the TDoc in Word."""
        row_idx = item.row()
        tdoc_item = self.table.item(row_idx, 0)
        if tdoc_item:
            r = tdoc_item.data(Qt.UserRole)
            if r and r.get("tdoc"):
                logger.info(f"📂 [TDoc Report] Double-click launch for TDoc {r['tdoc']}")
                SpecTDocOpener.open_tdoc(r["tdoc"], meeting_hint=r.get("meeting", ""), parent=self)

    def _open_selected_tdoc(self):
        """Opens the currently selected TDoc document."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return
        row_idx = selected_rows[0].row()
        item = self.table.item(row_idx, 0)
        if item:
            r = item.data(Qt.UserRole)
            if r and r.get("tdoc"):
                SpecTDocOpener.open_tdoc(r["tdoc"], meeting_hint=r.get("meeting", ""), parent=self)

    def _show_context_menu(self, pos: QPoint):
        """Right-click context menu for table rows."""
        item = self.table.itemAt(pos)
        if not item:
            return

        row_idx = item.row()
        tdoc_item = self.table.item(row_idx, 0)
        if not tdoc_item:
            return

        r = tdoc_item.data(Qt.UserRole)
        if not r or not r.get("tdoc"):
            return

        tdoc = r["tdoc"]
        meeting = r.get("meeting", "")
        title = r.get("title", "")

        menu = QMenu(self)

        act_open = menu.addAction(f"📂 Open TDoc Document ({tdoc}.docx)")
        act_open.triggered.connect(
            lambda: SpecTDocOpener.open_tdoc(tdoc, meeting_hint=meeting, parent=self)
        )

        act_portal = menu.addAction(f"🌐 Open {tdoc} in 3GU Web Portal")
        base_id = tdoc.split("r")[0].split("R")[0]
        act_portal.triggered.connect(lambda: SpecTDocOpener.open_tdoc_in_portal(base_id))

        act_jump = menu.addAction(f"🚀 Jump to {tdoc} in Meetings Tab")
        act_jump.triggered.connect(
            lambda: SpecTDocOpener.jump_to_meetings_tab(tdoc, main_window=self._parent_ref)
        )

        menu.addSeparator()

        act_citation = menu.addAction(f"📋 Copy Citation for {tdoc}")
        act_citation.triggered.connect(
            lambda: SpecTDocOpener.copy_tdoc_citation(
                tdoc_id=tdoc,
                spec_number=self.spec_number,
                version=self.version,
                meeting=meeting,
                subject=title,
                parent=self,
            )
        )

        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def _copy_markdown(self):
        """Copies the Markdown dossier report to clipboard."""
        if not self.markdown_report:
            self.markdown_report = SpecTDocReportGenerator.generate_markdown_report(
                self.spec_number, self.version, self.all_entries, spec_title=self.spec_title
            )
        QApplication.clipboard().setText(self.markdown_report)
        self.btn_copy_md.setText("✅ Copied Dossier!")
        logger.info(
            f"📋 [TDoc Report] Copied Markdown dossier for TS {self.spec_number} v{self.version} to clipboard."
        )
        QTimer.singleShot(1500, lambda: self.btn_copy_md.setText("📋 Copy Markdown Dossier"))

    def _export_markdown_file(self):
        """Saves the Markdown dossier report to disk."""
        if not self.markdown_report:
            self.markdown_report = SpecTDocReportGenerator.generate_markdown_report(
                self.spec_number, self.version, self.all_entries, spec_title=self.spec_title
            )

        default_name = f"TS_{self.spec_number}_v{self.version}_TDoc_Dossier.md"
        dest_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save TDoc Dossier Report",
            default_name,
            "Markdown Files (*.md);;All Files (*)",
        )
        if dest_path:
            try:
                Path(dest_path).write_text(self.markdown_report, encoding="utf-8")
                logger.info(f"💾 [TDoc Report] Saved Markdown dossier to: {dest_path}")
                QMessageBox.information(
                    self, "Export Complete", f"Successfully exported dossier to:\n{dest_path}"
                )
            except Exception as e:
                logger.error(f"Error saving markdown dossier: {e}", exc_info=True)
                QMessageBox.critical(self, "Export Error", f"Failed to save dossier:\n{e}")

    def _export_csv_file(self):
        """Exports the entries to a CSV file."""
        csv_text = SpecTDocReportGenerator.generate_csv_report(self.filtered_entries)
        default_name = f"TS_{self.spec_number}_v{self.version}_TDocs.csv"
        dest_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save TDoc CSV Report",
            default_name,
            "CSV Files (*.csv);;All Files (*)",
        )
        if dest_path:
            try:
                Path(dest_path).write_text(csv_text, encoding="utf-8")
                logger.info(f"📊 [TDoc Report] Saved CSV report to: {dest_path}")
                QMessageBox.information(
                    self, "Export Complete", f"Successfully exported CSV to:\n{dest_path}"
                )
            except Exception as e:
                logger.error(f"Error saving CSV: {e}", exc_info=True)
                QMessageBox.critical(self, "Export Error", f"Failed to save CSV:\n{e}")
