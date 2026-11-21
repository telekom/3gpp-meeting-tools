"""Interactive maintenance actions (inspect, single/batch compact, wipe, folder)."""

from __future__ import annotations

import logging
from pathlib import Path

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import QApplication, QMessageBox

from core.ui.database_inspector import DatabaseInspectorDialog
from core.utils.paths import get_project_root

from .storage_ops import format_file_size

logger = logging.getLogger(__name__)


class _MaintenanceActionsMixin:
    """User-triggered database inspection, compaction, and wipe handlers."""

    def _open_db_folder(self):
        """Opens the directory containing SQLite databases in Windows File Explorer."""
        root = get_project_root()
        db_dir = root / "db"
        if not db_dir.exists():
            db_dir = root
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(db_dir)))

    def _inspect_database(self, db_path: Path, display_name: str):
        """Open the reusable read-only SQLite inspector as a non-blocking modeless window."""
        if not db_path.exists():
            QMessageBox.warning(self, "Database Missing", f"Database does not exist:\n{db_path}")
            return

        key = str(db_path.resolve())
        dialog = self._inspector_dialogs.get(key)
        if dialog is None or not dialog.isVisible():
            dialog = DatabaseInspectorDialog(db_path, display_name, self)
            dialog.setWindowModality(Qt.NonModal)
            self._inspector_dialogs[key] = dialog
            dialog.show()
        else:
            dialog.raise_()
            dialog.activateWindow()

    def _wipe_database(self, db_path: Path, row_idx: int):
        """Completely wipes and reinitializes a wipeable auxiliary database (e.g. Change History)."""
        reply = QMessageBox.warning(
            self,
            "Confirm Database Wipe",
            f"Are you sure you want to completely wipe:\n{db_path.name}?\n\n"
            "All indexed specification change history and TDoc cross-references will be cleared.\n"
            "This will NOT affect your specifications catalogue or meeting data.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        from modules.specifications.core.spec_history_db import SpecHistoryDatabase

        success = SpecHistoryDatabase(db_path).wipe_database()
        if success:
            QMessageBox.information(
                self,
                "Wipe Complete",
                f"Successfully cleared and reinitialized {db_path.name}.",
            )
        else:
            QMessageBox.critical(self, "Wipe Error", f"Failed to wipe {db_path.name}.")

        self._refresh_databases()

    def _vacuum_single(self, db_path: Path, row_idx: int):
        """Compacts an individual database and updates table status with real-time feedback."""
        self.status_lbl.setText(f"Compacting {db_path.name}...")
        self.status_lbl.setVisible(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(50)
        QApplication.processEvents()

        success, before, after, msg = self._vacuum_database_file(db_path)

        self.progress_bar.setValue(100)
        if success:
            saved_bytes = max(0, before - after)
            saved_str = (
                f" (reclaimed {format_file_size(saved_bytes)})"
                if saved_bytes > 0
                else " (already compact)"
            )
            logger.info(f"✅ Database compacted: {db_path.name}{saved_str}")
            self.status_lbl.setText(f"Compacted {db_path.name}{saved_str}")
        else:
            logger.error(f"❌ Compaction error on {db_path.name}: {msg}")
            QMessageBox.critical(
                self,
                "Compaction Error",
                f"Failed to compact {db_path.name}:\n{msg}",
            )

        self._refresh_databases()
        self.progress_bar.setVisible(False)
        self.status_lbl.setVisible(False)

    def _vacuum_all(self):
        """Compacts all existing tracked databases in sequence with real-time progress."""
        db_entries = self._get_tracked_databases()
        existing_entries = [e for e in db_entries if e["path"].exists()]
        if not existing_entries:
            return

        total_count = len(existing_entries)
        total_saved = 0
        compacted_count = 0

        self.vacuum_all_btn.setEnabled(False)
        self.vacuum_all_btn.setText("⏳ Compacting...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_lbl.setVisible(True)
        QApplication.processEvents()

        for idx, entry in enumerate(existing_entries):
            db_path = entry["path"]
            self.status_lbl.setText(f"Compacting {db_path.name} ({idx + 1}/{total_count})...")
            self.progress_bar.setValue(int((idx / total_count) * 100))
            QApplication.processEvents()

            success, before, after, msg = self._vacuum_database_file(db_path)
            if success:
                compacted_count += 1
                total_saved += max(0, before - after)
            else:
                logger.error(f"❌ Failed to compact {db_path.name}: {msg}")

        self.progress_bar.setValue(100)
        self.status_lbl.setText(f"Compaction complete. Reclaimed {format_file_size(total_saved)}.")
        self.vacuum_all_btn.setText("🧹 Compact All Databases")
        self.vacuum_all_btn.setEnabled(True)
        self._refresh_databases()

        summary_msg = (
            f"Compacted {compacted_count} database(s).\n"
            f"Total disk space reclaimed: {format_file_size(total_saved)}"
        )
        logger.info(f"✅ {summary_msg.replace(chr(10), ' ')}")
        QMessageBox.information(self, "Compaction Complete", summary_msg)

        self.progress_bar.setVisible(False)
        self.status_lbl.setVisible(False)
