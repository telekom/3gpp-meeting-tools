"""Superseded specification version review and pruning dialog."""

from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from modules.spec_search.core.spec_search_db import SpecSearchDatabase


class SpecSearchPruneDialog(QDialog):
    """
    Dialog to review and prune older specification versions superseded by newer releases
    within the same 3GPP major release.
    """

    def __init__(
        self,
        db: SpecSearchDatabase,
        superseded_versions: List[Dict[str, Any]],
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.db = db
        self.superseded_versions = superseded_versions
        self.pruned_count = 0
        self.setWindowTitle("🧹 Prune Superseded Specification Versions")
        self.resize(780, 500)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # Header Info Banner
        info_banner = QFrame()
        info_banner.setStyleSheet(f"""
            QFrame {{
                background-color: {C.INFO_SOFT};
                border: 1px solid {C.INFO_BORDER};
                border-radius: 6px;
                padding: 8px;
            }}
        """)
        ib_layout = QVBoxLayout(info_banner)
        ib_layout.setContentsMargins(6, 6, 6, 6)
        ib_layout.setSpacing(3)

        title_lbl = QLabel(f"<b>Found {len(self.superseded_versions)} Superseded Specification Version(s)</b>")
        title_lbl.setStyleSheet(f"color: {C.INFO_NAVY}; font-size: 12px;")
        desc_lbl = QLabel(
            "These older minor releases belong to the same 3GPP Release as a newer indexed version.\n"
            "Pruning them will remove their clause records and FTS indices, freeing database space and cleaning the evolution matrix."
        )
        desc_lbl.setStyleSheet(f"color: {C.INFO_DARK}; font-size: 11px;")
        desc_lbl.setWordWrap(True)
        ib_layout.addWidget(title_lbl)
        ib_layout.addWidget(desc_lbl)
        layout.addWidget(info_banner)

        # Quick selection bar
        sel_bar = QHBoxLayout()
        btn_sel_all = QPushButton("Select All")
        btn_sel_all.clicked.connect(self._select_all)
        sel_bar.addWidget(btn_sel_all)

        btn_desel_all = QPushButton("Deselect All")
        btn_desel_all.clicked.connect(self._deselect_all)
        sel_bar.addWidget(btn_desel_all)

        sel_bar.addStretch()
        self.lbl_selected = QLabel(f"Selected: {len(self.superseded_versions)} version(s)")
        self.lbl_selected.setStyleSheet(f"font-weight: bold; color: {C.ACCENT_SKY}; font-size: 11px;")
        sel_bar.addWidget(self.lbl_selected)
        layout.addLayout(sel_bar)

        # Table of superseded versions
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Prune", "Specification", "Superseded Version", "Release", "Active Replacement", "Clauses"
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)

        self._populate_table()
        layout.addWidget(self.table)

        # Action Buttons
        btn_box = QHBoxLayout()
        btn_box.addStretch()

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(btn_cancel)

        self.btn_prune = QPushButton(f"🧹 Prune Selected ({len(self.superseded_versions)})")
        self.btn_prune.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.DANGER_ALT};
                color: {C.BG_CARD};
                border: none;
                border-radius: 4px;
                padding: 6px 14px;
                font-weight: bold;
                font-size: 11px;
            }}
            QPushButton:hover {{
                background-color: {C.DANGER};
            }}
            QPushButton:disabled {{
                background-color: {C.BORDER_STRONG};
                color: {C.TEXT_FAINT};
            }}
        """)
        self.btn_prune.clicked.connect(self._on_prune_confirmed)
        btn_box.addWidget(self.btn_prune)
        layout.addLayout(btn_box)

    def _populate_table(self):
        self.table.setRowCount(len(self.superseded_versions))
        for r, item in enumerate(self.superseded_versions):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            chk.setCheckState(Qt.Checked)
            self.table.setItem(r, 0, chk)

            s_item = QTableWidgetItem(f"TS {item['spec_number']}")
            s_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 1, s_item)

            v_item = QTableWidgetItem(f"v{item['version']}")
            v_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 2, v_item)

            rel_item = QTableWidgetItem(str(item.get("release", "")))
            rel_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(r, 3, rel_item)

            sup_item = QTableWidgetItem(f"v{item.get('superseded_by', '')} (Newer)")
            sup_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            sup_item.setForeground(QColor(C.SUCCESS))
            self.table.setItem(r, 4, sup_item)

            c_item = QTableWidgetItem(f"{item.get('total_clauses', 0):,}")
            c_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            c_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(r, 5, c_item)

        self.table.itemChanged.connect(self._update_selected_count)

    def _update_selected_count(self):
        selected = 0
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item and item.checkState() == Qt.Checked:
                selected += 1
        self.lbl_selected.setText(f"Selected: {selected} version(s)")
        self.btn_prune.setText(f"🧹 Prune Selected ({selected})")
        self.btn_prune.setEnabled(selected > 0)

    def _select_all(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item:
                item.setCheckState(Qt.Checked)
        self.table.blockSignals(False)
        self._update_selected_count()

    def _deselect_all(self):
        self.table.blockSignals(True)
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item:
                item.setCheckState(Qt.Unchecked)
        self.table.blockSignals(False)
        self._update_selected_count()

    def _on_prune_confirmed(self):
        target_ids = []
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item and item.checkState() == Qt.Checked:
                target_ids.append(self.superseded_versions[r]["id"])

        if not target_ids:
            return

        confirm = QMessageBox.question(
            self,
            "Confirm Pruning",
            f"Are you sure you want to permanently delete {len(target_ids)} superseded version(s) from the search index?\nThis action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm == QMessageBox.Yes:
            self.pruned_count = self.db.prune_superseded_versions(target_ids)
            self.accept()
