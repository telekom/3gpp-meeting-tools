"""Storage table population and executive KPI update mixin."""

from __future__ import annotations

from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C

from .storage_ops import format_file_size
from .styles import (
    CELL_SUB_STYLE,
    CELL_TITLE_STYLE,
    COMPACT_BTN_STYLE,
    INSPECT_BTN_STYLE,
    WIPE_BTN_STYLE,
)


class _MaintenanceTableMixin:
    """Populates the database metrics table and refreshes summary KPI cards."""

    def _refresh_databases(self):
        """Populates the table and updates executive KPI statistics."""
        db_entries = self._get_tracked_databases()
        self.table.setRowCount(0)

        base_font = self.table.font()
        bold_font = QFont(base_font)
        bold_font.setBold(True)

        total_storage = 0
        total_reclaimable = 0
        total_tables = 0
        active_wal_count = 0
        active_wal_size = 0
        active_db_count = 0

        for row_idx, entry in enumerate(db_entries):
            db_path: Path = entry["path"]
            self.table.insertRow(row_idx)

            metrics = self._query_db_metrics(db_path)

            if metrics["exists"]:
                active_db_count += 1
                total_storage += metrics["size"] + metrics["wal_size"]
                total_reclaimable += metrics["reclaimable"]
                total_tables += metrics["tables"]
                if metrics["wal_size"] > 0:
                    active_wal_count += 1
                    active_wal_size += metrics["wal_size"]

            # ----------------------------------------------------
            # Col 0: Database & Purpose (Two-tier custom widget)
            # ----------------------------------------------------
            cell_widget = QWidget()
            cell_layout = QVBoxLayout(cell_widget)
            cell_layout.setContentsMargins(10, 4, 10, 4)
            cell_layout.setSpacing(2)

            lbl_title = QLabel(f"{entry.get('icon', '🗄️')}  {entry['name']}")
            lbl_title.setStyleSheet(CELL_TITLE_STYLE)

            sub_text = f"{db_path.name}   •   {entry.get('subtitle', '')}"
            lbl_sub = QLabel(sub_text)
            lbl_sub.setStyleSheet(CELL_SUB_STYLE)

            cell_layout.addWidget(lbl_title)
            cell_layout.addWidget(lbl_sub)
            cell_widget.setToolTip(str(db_path))
            self.table.setCellWidget(row_idx, 0, cell_widget)

            # Background item for selection support
            bg_item = QTableWidgetItem()
            self.table.setItem(row_idx, 0, bg_item)

            # ----------------------------------------------------
            # Col 1: Tables Count
            # ----------------------------------------------------
            tbl_text = f"{metrics['tables']} tbl" if metrics["exists"] else "—"
            tbl_item = QTableWidgetItem(tbl_text)
            tbl_item.setTextAlignment(Qt.AlignCenter)
            tbl_item.setFont(bold_font)
            tbl_item.setForeground(
                QBrush(QColor(C.TEXT_BODY) if metrics["exists"] else QColor(C.TEXT_FAINT))
            )
            tbl_item.setToolTip(
                f"User tables in database: {metrics['tables']}" if metrics["exists"] else "Not created"
            )
            self.table.setItem(row_idx, 1, tbl_item)

            # ----------------------------------------------------
            # Col 2: File Size
            # ----------------------------------------------------
            size_text = format_file_size(metrics["size"]) if metrics["exists"] else "Not Created"
            size_item = QTableWidgetItem(size_text)
            size_item.setTextAlignment(Qt.AlignCenter)
            size_item.setFont(bold_font)
            size_item.setForeground(
                QBrush(QColor(C.TEXT_PRIMARY) if metrics["exists"] else QColor(C.TEXT_FAINT))
            )
            self.table.setItem(row_idx, 2, size_item)

            # ----------------------------------------------------
            # Col 3: Reclaimable Space (Freelist calculation)
            # ----------------------------------------------------
            if not metrics["exists"]:
                rec_item = QTableWidgetItem("—")
                rec_item.setForeground(QBrush(QColor(C.TEXT_FAINT)))
            elif metrics["reclaimable"] > 0:
                rec_item = QTableWidgetItem(f"⚡ {format_file_size(metrics['reclaimable'])}")
                rec_item.setFont(bold_font)
                rec_item.setForeground(QBrush(QColor(C.ACCENT_SKY)))
                rec_item.setToolTip(
                    f"Freelist pages can be reclaimed via Compact:\n"
                    f"{format_file_size(metrics['reclaimable'])} ({metrics['freelist_pages']} unused pages)"
                )
            else:
                rec_item = QTableWidgetItem("✓ Optimal")
                rec_item.setFont(bold_font)
                rec_item.setForeground(QBrush(QColor(C.SUCCESS_DARK)))
                rec_item.setToolTip("Database is defragmented with 0 unused freelist pages.")
            rec_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, 3, rec_item)

            # ----------------------------------------------------
            # Col 4: WAL Status
            # ----------------------------------------------------
            if not metrics["exists"]:
                wal_item = QTableWidgetItem("⚪ Inactive")
                wal_item.setForeground(QBrush(QColor(C.TEXT_FAINT)))
            elif metrics["wal_size"] > 0:
                wal_item = QTableWidgetItem(f"🟡 {format_file_size(metrics['wal_size'])}")
                wal_item.setFont(bold_font)
                wal_item.setForeground(QBrush(QColor(C.WARNING)))
                wal_item.setToolTip(
                    f"Active Write-Ahead Log: {format_file_size(metrics['wal_size'])}\n"
                    f"Will be checkpointed and truncated during compaction."
                )
            else:
                wal_item = QTableWidgetItem("🟢 Clean")
                wal_item.setFont(bold_font)
                wal_item.setForeground(QBrush(QColor(C.SUCCESS_DARK)))
                wal_item.setToolTip("Write-Ahead Log is fully flushed and clean.")
            wal_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, 4, wal_item)

            # ----------------------------------------------------
            # Col 5: Actions (Inspect, Compact, Wipe)
            # Dynamic sizing with generous padding for high-DPI
            # ----------------------------------------------------
            if metrics["exists"]:
                action_widget = QWidget()
                action_layout = QHBoxLayout(action_widget)
                action_layout.setContentsMargins(4, 4, 4, 4)
                action_layout.setSpacing(6)

                btn_inspect = QPushButton("🔍 Inspect")
                btn_inspect.setFixedHeight(28)
                btn_inspect.setStyleSheet(INSPECT_BTN_STYLE)
                btn_inspect.setToolTip("Browse tables, schemas, and diagnostics in read-only mode.")
                btn_inspect.clicked.connect(
                    lambda _, p=db_path, n=entry["name"]: self._inspect_database(p, n)
                )
                action_layout.addWidget(btn_inspect)

                btn_compact = QPushButton("⚡ Compact")
                btn_compact.setFixedHeight(28)
                btn_compact.setStyleSheet(COMPACT_BTN_STYLE)
                btn_compact.setToolTip("Run PRAGMA wal_checkpoint(TRUNCATE) and VACUUM on this database.")
                btn_compact.clicked.connect(lambda _, p=db_path, r=row_idx: self._vacuum_single(p, r))
                action_layout.addWidget(btn_compact)

                if entry.get("can_wipe"):
                    btn_wipe = QPushButton("🗑️ Wipe")
                    btn_wipe.setFixedHeight(28)
                    btn_wipe.setStyleSheet(WIPE_BTN_STYLE)
                    btn_wipe.setToolTip(
                        "Completely wipe and clear all indexed change history and TDoc cross-references."
                    )
                    btn_wipe.clicked.connect(lambda _, p=db_path, r=row_idx: self._wipe_database(p, r))
                    action_layout.addWidget(btn_wipe)
                else:
                    action_layout.addStretch()

                self.table.setCellWidget(row_idx, 5, action_widget)
            else:
                lbl_none = QLabel("—")
                lbl_none.setFont(base_font)
                lbl_none.setAlignment(Qt.AlignCenter)
                self.table.setCellWidget(row_idx, 5, lbl_none)

        # ----------------------------------------------------
        # 5. UPDATE EXECUTIVE KPI CARDS
        # ----------------------------------------------------
        self.kpi_storage.set_data(
            format_file_size(total_storage),
            f"across {active_db_count} active database(s)",
            value_color=C.TEXT_PRIMARY,
        )

        if total_reclaimable > 0:
            self.kpi_reclaimable.set_data(
                f"~{format_file_size(total_reclaimable)}",
                "freelist pages ready to reclaim",
                value_color=C.ACCENT_SKY,
            )
        else:
            self.kpi_reclaimable.set_data(
                "✓ Optimal",
                "0 B unused freelist pages",
                value_color=C.SUCCESS_DARK,
            )

        if active_wal_count > 0:
            self.kpi_wal.set_data(
                f"{active_wal_count} Active",
                f"{format_file_size(active_wal_size)} pending checkpoint",
                value_color=C.WARNING,
            )
        else:
            self.kpi_wal.set_data(
                "All Clean",
                "All WAL logs fully checkpointed",
                value_color=C.SUCCESS_DARK,
            )

        self.kpi_tables.set_data(
            f"{total_tables} Tables",
            "across all active schemas",
            value_color=C.TEXT_PRIMARY,
        )
