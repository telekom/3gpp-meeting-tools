# --- File: src/modules/specifications/ui/dialogs.py ---
import logging
import os
import re
import webbrowser
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
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
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BADGE_STYLE_INFO,
    BADGE_STYLE_MUTED,
    BADGE_STYLE_PRIMARY,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SEARCH_INPUT_STYLE,
    TABLE_STYLE_CLEAN,
)

from modules.specifications.core.database import SpecsDatabase
from modules.specifications.core.scraper import fetch_metadata_from_dynareport


class SpecsConfigDialog(QDialog):
    """Configuration dialog for setting and verifying the specifications download directory."""

    def __init__(self, current_path: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Specifications Settings")
        self.setModal(True)
        self.resize(520, 160)
        self.selected_path = current_path

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        info_label = QLabel(
            "Specify the local directory where downloaded 3GPP specifications, "
            "ZIP archives, and converted documents are stored:"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #4A5568; font-size: 12px;")
        layout.addWidget(info_label)

        path_layout = QHBoxLayout()
        path_layout.setSpacing(6)

        self.path_input = QLineEdit(current_path)
        self.path_input.setPlaceholderText("Select folder...")

        browse_btn = QPushButton("📂 Browse...")
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.clicked.connect(self._browse)

        open_btn = QPushButton("↗️ Open")
        open_btn.setCursor(Qt.PointingHandCursor)
        open_btn.setToolTip("Open this directory in Windows Explorer")
        open_btn.clicked.connect(self._open_folder)

        path_layout.addWidget(self.path_input)
        path_layout.addWidget(browse_btn)
        path_layout.addWidget(open_btn)
        layout.addLayout(path_layout)

        layout.addStretch()

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.clicked.connect(self.reject)

        save_btn = QPushButton("💾 Save Settings")
        save_btn.setCursor(Qt.PointingHandCursor)
        save_btn.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                background-color: #0066CC;
                color: white;
                padding: 6px 16px;
                border-radius: 4px;
                border: 1px solid #0055AA;
            }
            QPushButton:hover {
                background-color: #0052A3;
            }
        """)
        save_btn.clicked.connect(self._save_and_accept)

        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def _browse(self):
        new_dir = QFileDialog.getExistingDirectory(
            self, "Select Download Directory", self.path_input.text().strip()
        )
        if new_dir:
            self.path_input.setText(new_dir)

    def _open_folder(self):
        p = Path(self.path_input.text().strip())
        if not p.exists():
            try:
                p.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                QMessageBox.warning(self, "Directory Error", f"Could not create directory:\n{e}")
                return
        try:
            os.startfile(str(p))
        except Exception as e:
            QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")

    def _save_and_accept(self):
        self.selected_path = self.path_input.text().strip()
        self.accept()

    def get_download_path(self) -> str:
        return self.selected_path


class NumericTableWidgetItem(QTableWidgetItem):
    """Table widget item that sorts numerically rather than alphabetically."""

    def __lt__(self, other):
        try:
            return int(self.text()) < int(other.text())
        except (ValueError, TypeError):
            return super().__lt__(other)


class SpecInfoDialog(QDialog):
    """Modernized Specification Details Dialog with dedicated Work Items table and search filter."""

    def __init__(self, details: dict, parent=None):
        super().__init__(parent)
        self.details = details
        spec_num = details.get("number", "Unknown")
        spec_type = details.get("type", "TS")
        title = details.get("title", "No Title Available")
        self.related_wis = details.get("related_wis", [])

        self.setWindowTitle(f"Specification Details: {spec_num}")
        self.setMinimumWidth(640)

        # Adaptive initial sizing: taller when a WI table needs to be displayed
        if self.related_wis:
            self.resize(760, 660)
            self.setMinimumHeight(500)
        else:
            self.resize(640, 400)

        # Rely on shared card styling and standard Qt dialog background
        self.setStyleSheet(f"""
            QDialog {{ background-color: #F8F9FA; }}
            {CARD_FRAME_STYLE}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # 1. Header Card
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(6)

        title_row = QHBoxLayout()
        type_badge = QLabel(f"<b>{spec_type}</b>")
        type_badge.setStyleSheet(BADGE_STYLE_INFO)

        number_label = QLabel(f"<b>{spec_num}</b>")
        number_label.setStyleSheet("font-size: 17px; color: #1A202C; font-weight: bold;")
        number_label.setTextInteractionFlags(Qt.TextSelectableByMouse)

        title_row.addWidget(type_badge)
        title_row.addWidget(number_label)
        title_row.addStretch()
        header_layout.addLayout(title_row)

        desc_label = QLabel(title)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #4A5568; font-size: 13px; line-height: 1.4;")
        desc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        header_layout.addWidget(desc_label)

        layout.addWidget(header_card)

        # 2. Details & Links Card
        details_card = QFrame()
        details_card.setObjectName("cardFrame")
        form = QFormLayout(details_card)
        form.setContentsMargins(14, 12, 14, 12)
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignRight)

        clean_number = spec_num.split("-")[0].replace(".", "").strip()
        dynareport_url = (
            f"https://www.3gpp.org/DynaReport/{clean_number}.htm" if clean_number else ""
        )
        ftp_url = details.get("url", "")

        primary_group = details.get("primary_group") or "-"
        sec_groups = details.get("secondary_groups") or "-"
        tech = details.get("radio_technology") or details.get("radio_tech") or "-"
        init_rel = details.get("initial_release") or "-"

        self._add_row(form, "Primary Group", primary_group)
        self._add_row(form, "Secondary Groups", sec_groups)
        self._add_row(form, "Radio Technology", tech)
        self._add_row(form, "Initial Release", init_rel)

        if ftp_url:
            ftp_label = QLabel(
                f'<a href="{ftp_url}" style="color: #1E5C99; text-decoration: none;">{ftp_url}</a>'
            )
            ftp_label.setOpenExternalLinks(True)
            ftp_label.setTextInteractionFlags(Qt.TextBrowserInteraction | Qt.TextSelectableByMouse)
            form.addRow(self._make_key_label("FTP Archive:"), ftp_label)

        if dynareport_url:
            dyna_label = QLabel(
                f'<a href="{dynareport_url}" style="color: #1E5C99; text-decoration: none;">'
                f"Open 3GPP Portal Report ({clean_number}.htm) ↗</a>"
            )
            dyna_label.setOpenExternalLinks(True)
            dyna_label.setTextInteractionFlags(Qt.TextBrowserInteraction | Qt.TextSelectableByMouse)
            form.addRow(self._make_key_label("DynaReport:"), dyna_label)

        if not self.related_wis:
            self._add_row(form, "Related WIs", "None")

        excluded_keys = {
            "id",
            "series_id",
            "number",
            "type",
            "title",
            "url",
            "primary_group",
            "secondary_groups",
            "radio_technology",
            "radio_tech",
            "initial_release",
            "related_wis",
        }
        for key, value in details.items():
            if key not in excluded_keys and value:
                display_key = key.replace("_", " ").title()
                self._add_row(form, display_key, str(value))

        layout.addWidget(details_card)

        # 3. Dedicated Work Items Card (Rendered only when WIs exist)
        if self.related_wis:
            wi_card = self._build_work_items_card()
            layout.addWidget(wi_card, stretch=1)

        # 4. Action Buttons Footer
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        if dynareport_url:
            dynareport_btn = QPushButton("🌐 Open DynaReport")
            dynareport_btn.setObjectName("primaryBtn")
            dynareport_btn.setCursor(Qt.PointingHandCursor)
            dynareport_btn.clicked.connect(lambda: webbrowser.open(dynareport_url))
            btn_layout.addWidget(dynareport_btn)

        if ftp_url:
            ftp_btn = QPushButton("📂 Open FTP Archive")
            ftp_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            ftp_btn.setCursor(Qt.PointingHandCursor)
            ftp_btn.clicked.connect(lambda: webbrowser.open(ftp_url))
            btn_layout.addWidget(ftp_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _build_work_items_card(self) -> QFrame:
        """Constructs the dedicated Work Items section with instant search and sortable table."""
        card = QFrame()
        card.setObjectName("cardFrame")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 14, 12)
        card_layout.setSpacing(8)

        # Header bar: Section Title, Count Badge, and Search Filter
        header_bar = QHBoxLayout()
        header_bar.setSpacing(8)

        title_lbl = QLabel("<b>Related Work Items</b>")
        title_lbl.setStyleSheet("font-size: 13px; color: #2D3748;")
        header_bar.addWidget(title_lbl)

        self.wi_count_badge = QLabel(f"{len(self.related_wis)} WIs")
        self.wi_count_badge.setStyleSheet(BADGE_STYLE_MUTED)
        header_bar.addWidget(self.wi_count_badge)
        header_bar.addStretch()

        self.wi_filter_input = QLineEdit()
        self.wi_filter_input.setPlaceholderText("🔍 Filter by acronym, code, or name...")
        self.wi_filter_input.setClearButtonEnabled(True)
        self.wi_filter_input.setFixedWidth(260)
        self.wi_filter_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.wi_filter_input.textChanged.connect(self._filter_wis)
        header_bar.addWidget(self.wi_filter_input)
        card_layout.addLayout(header_bar)

        # Work Items Table
        self.wi_table = QTableWidget()
        self.wi_table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.wi_table.setColumnCount(5)
        self.wi_table.setHorizontalHeaderLabels(
            ["Role", "Acronym", "Code", "Work Item Description", "Link"]
        )

        header = self.wi_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        self.wi_table.setColumnWidth(0, 95)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        self.wi_table.setColumnWidth(1, 140)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        self.wi_table.setColumnWidth(2, 70)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        self.wi_table.setColumnWidth(4, 65)

        self.wi_table.verticalHeader().setVisible(False)
        self.wi_table.verticalHeader().setDefaultSectionSize(32)
        self.wi_table.setAlternatingRowColors(True)
        self.wi_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.wi_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.wi_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.wi_table.cellDoubleClicked.connect(self._on_wi_double_clicked)
        self.wi_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.wi_table.customContextMenuRequested.connect(self._show_wi_context_menu)

        # Sort: Primary WIs first, then alphabetical by Acronym
        sorted_wis = sorted(
            self.related_wis,
            key=lambda w: (
                0 if bool(w.get("is_primary")) else 1,
                str(w.get("acronym", "")).lower(),
                str(w.get("wi_code", "")),
            ),
        )

        self.wi_table.setSortingEnabled(False)
        self.wi_table.setRowCount(len(sorted_wis))

        for row, wi in enumerate(sorted_wis):
            code = str(wi.get("wi_code", "")).strip()
            acronym = str(wi.get("acronym", "")).strip()
            name = str(wi.get("name", "")).strip()
            is_primary = bool(wi.get("is_primary", False))
            portal_url = (
                f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={code}"
                if code
                else ""
            )

            # Column 0: Role Badge (⭐ Primary vs Secondary)
            role_item = QTableWidgetItem("0" if is_primary else "1")
            self.wi_table.setItem(row, 0, role_item)

            badge_widget = QWidget()
            badge_layout = QHBoxLayout(badge_widget)
            badge_layout.setContentsMargins(4, 2, 4, 2)
            badge_layout.setAlignment(Qt.AlignCenter)

            role_badge = QLabel("⭐ Primary" if is_primary else "Secondary")
            role_badge.setStyleSheet(BADGE_STYLE_PRIMARY if is_primary else BADGE_STYLE_MUTED)
            badge_layout.addWidget(role_badge)
            self.wi_table.setCellWidget(row, 0, badge_widget)

            # Column 1: Acronym
            acronym_item = QTableWidgetItem(acronym if acronym else "-")
            if is_primary:
                acronym_font = acronym_item.font()
                acronym_font.setBold(True)
                acronym_item.setFont(acronym_font)
            acronym_item.setToolTip(f"{name}\nDouble-click to view on 3GPP Portal")
            self.wi_table.setItem(row, 1, acronym_item)

            # Column 2: Code (Numerically sortable)
            code_item = NumericTableWidgetItem(code)
            code_item.setTextAlignment(Qt.AlignCenter)
            code_item.setToolTip("Double-click to view on 3GPP Portal")
            self.wi_table.setItem(row, 2, code_item)

            # Column 3: Description
            name_item = QTableWidgetItem(name if name else "(No description)")
            name_item.setToolTip(f"{name}\nDouble-click to view on 3GPP Portal")
            self.wi_table.setItem(row, 3, name_item)

            # Column 4: Text Hyperlink (Clean link without button boxes)
            if portal_url:
                link_lbl = QLabel(
                    f'<a href="{portal_url}" style="color: #1E5C99; text-decoration: none; font-weight: 600;">Open ↗</a>'
                )
                link_lbl.setOpenExternalLinks(True)
                link_lbl.setAlignment(Qt.AlignCenter)
                link_lbl.setToolTip(f"Open Work Item #{code} on 3GPP Portal")
                self.wi_table.setCellWidget(row, 4, link_lbl)
            else:
                empty_item = QTableWidgetItem("-")
                empty_item.setTextAlignment(Qt.AlignCenter)
                self.wi_table.setItem(row, 4, empty_item)

        self.wi_table.setSortingEnabled(True)
        card_layout.addWidget(self.wi_table)
        return card

    def _open_wi_url(self, wi_code: str):
        if not wi_code:
            return
        url = f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}"
        webbrowser.open(url)

    def _on_wi_double_clicked(self, row: int, column: int):
        code_item = self.wi_table.item(row, 2)
        if code_item:
            self._open_wi_url(code_item.text().strip())

    def _filter_wis(self, query: str):
        query = query.strip().lower()
        visible_count = 0
        total_rows = self.wi_table.rowCount()

        for row in range(total_rows):
            acronym_item = self.wi_table.item(row, 1)
            code_item = self.wi_table.item(row, 2)
            name_item = self.wi_table.item(row, 3)

            acronym = acronym_item.text().lower() if acronym_item else ""
            code = code_item.text().lower() if code_item else ""
            name = name_item.text().lower() if name_item else ""

            matches = not query or (query in acronym) or (query in code) or (query in name)
            self.wi_table.setRowHidden(row, not matches)
            if matches:
                visible_count += 1

        if query:
            self.wi_count_badge.setText(f"{visible_count} of {total_rows} WIs")
        else:
            self.wi_count_badge.setText(f"{total_rows} WIs")

    def _show_wi_context_menu(self, pos):
        item = self.wi_table.itemAt(pos)
        if not item:
            return
        row = item.row()
        code_item = self.wi_table.item(row, 2)
        acronym_item = self.wi_table.item(row, 1)
        name_item = self.wi_table.item(row, 3)

        code = code_item.text().strip() if code_item else ""
        acronym = acronym_item.text().strip() if acronym_item else ""
        name = name_item.text().strip() if name_item else ""
        url = f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={code}"

        menu = QMenu(self)

        act_open = menu.addAction("🌐 Open Work Item on 3GPP Portal")
        act_open.triggered.connect(lambda: self._open_wi_url(code))
        menu.addSeparator()

        act_copy_acronym = menu.addAction("📋 Copy Acronym")
        act_copy_acronym.triggered.connect(lambda: QApplication.clipboard().setText(acronym))

        act_copy_code = menu.addAction("📋 Copy WI Code")
        act_copy_code.triggered.connect(lambda: QApplication.clipboard().setText(code))

        act_copy_name = menu.addAction("📋 Copy Description")
        act_copy_name.triggered.connect(lambda: QApplication.clipboard().setText(name))

        act_copy_link = menu.addAction("🔗 Copy Portal Link")
        act_copy_link.triggered.connect(lambda: QApplication.clipboard().setText(url))

        menu.exec_(self.wi_table.viewport().mapToGlobal(pos))

    def _make_key_label(self, text: str) -> QLabel:
        lbl = QLabel(f"<b>{text}</b>")
        lbl.setStyleSheet("color: #718096; font-size: 12px;")
        return lbl

    def _add_row(self, form: QFormLayout, label_text: str, value_text: str):
        val_label = QLabel(value_text)
        val_label.setWordWrap(True)
        val_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow(self._make_key_label(f"{label_text}:"), val_label)


class AdvancedSyncDialog(QDialog):
    """Network Database Sync Dialog with Strict Drop-Down Menus."""

    def __init__(self, db: SpecsDatabase, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Advanced Filtered Sync")
        self.setModal(True)
        self.resize(450, 250)
        self.matching_specs = []

        options = db.get_filter_options()

        layout = QVBoxLayout(self)
        info_label = QLabel(
            "Note: Filters apply to specifications already discovered in your local database. "
            "To discover brand new specifications, run a 'Full Sync' first."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666666; font-style: italic; margin-bottom: 10px;")
        layout.addWidget(info_label)

        form = QFormLayout()

        self.series_combo = QComboBox()
        self.series_combo.addItem("Any")
        self.series_combo.addItems(options["series"])

        self.tech_combo = QComboBox()
        self.tech_combo.addItem("Any")
        self.tech_combo.addItems(options["techs"])

        self.group_combo = QComboBox()
        self.group_combo.addItem("Any")
        self.group_combo.addItems(options["groups"])

        self.type_combo = QComboBox()
        self.type_combo.addItem("Any")
        self.type_combo.addItems(options["types"])

        form.addRow("Series:", self.series_combo)
        form.addRow("Radio Tech:", self.tech_combo)
        form.addRow("Working Group:", self.group_combo)
        form.addRow("Type:", self.type_combo)
        layout.addLayout(form)

        self.count_label = QLabel("Matching specifications: 0")
        self.count_label.setStyleSheet("font-weight: bold; color: #395396; margin-top: 10px;")
        layout.addWidget(self.count_label)

        btn_layout = QHBoxLayout()
        self.sync_btn = QPushButton("🚀 Start Sync")
        self.sync_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(self.sync_btn)
        layout.addLayout(btn_layout)

        self.series_combo.currentTextChanged.connect(self.update_count)
        self.tech_combo.currentTextChanged.connect(self.update_count)
        self.group_combo.currentTextChanged.connect(self.update_count)
        self.type_combo.currentTextChanged.connect(self.update_count)

        self.update_count()

    def update_count(self):
        series = "" if self.series_combo.currentText() == "Any" else self.series_combo.currentText()
        tech = "" if self.tech_combo.currentText() == "Any" else self.tech_combo.currentText()
        group = "" if self.group_combo.currentText() == "Any" else self.group_combo.currentText()
        spec_type = self.type_combo.currentText()

        self.matching_specs = self.db.get_filtered_specs(series, tech, group, spec_type)
        count = len(self.matching_specs)
        self.count_label.setText(f"Matching specifications in local DB: {count}")
        self.sync_btn.setEnabled(count > 0)


class TableFilterDialog(QDialog):
    """Local Table Filter Dialog with Strict Drop-Down Menus."""

    def __init__(self, db: SpecsDatabase, current_filters: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Filter Specifications")
        self.setModal(True)
        self.resize(350, 200)

        options = db.get_filter_options()

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.series_combo = QComboBox()
        self.series_combo.addItem("Any")
        self.series_combo.addItems(options["series"])
        self.series_combo.setCurrentText(current_filters.get("series", "Any") or "Any")

        self.tech_combo = QComboBox()
        self.tech_combo.addItem("Any")
        self.tech_combo.addItems(options["techs"])
        self.tech_combo.setCurrentText(current_filters.get("tech", "Any") or "Any")

        self.group_combo = QComboBox()
        self.group_combo.addItem("Any")
        self.group_combo.addItems(options["groups"])
        self.group_combo.setCurrentText(current_filters.get("group", "Any") or "Any")

        self.type_combo = QComboBox()
        self.type_combo.addItem("Any")
        self.type_combo.addItems(options["types"])
        self.type_combo.setCurrentText(current_filters.get("spec_type", "Any") or "Any")

        form.addRow("Series:", self.series_combo)
        form.addRow("Radio Tech:", self.tech_combo)
        form.addRow("Working Group:", self.group_combo)
        form.addRow("Type:", self.type_combo)
        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        apply_btn = QPushButton("✅ Apply Filters")
        apply_btn.clicked.connect(self.accept)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear_and_accept)

        btn_layout.addStretch()
        btn_layout.addWidget(clear_btn)
        btn_layout.addWidget(apply_btn)
        layout.addLayout(btn_layout)

    def _clear_and_accept(self):
        self.series_combo.setCurrentText("Any")
        self.tech_combo.setCurrentText("Any")
        self.group_combo.setCurrentText("Any")
        self.type_combo.setCurrentText("Any")
        self.accept()

    def get_filters(self) -> dict:
        return {
            "series": (
                "" if self.series_combo.currentText() == "Any" else self.series_combo.currentText()
            ),
            "tech": (
                "" if self.tech_combo.currentText() == "Any" else self.tech_combo.currentText()
            ),
            "group": (
                "" if self.group_combo.currentText() == "Any" else self.group_combo.currentText()
            ),
            "spec_type": self.type_combo.currentText(),
        }


class TargetedSyncDialog(QDialog):
    """Dialog for fetching brand new specifications directly by number or series."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🎯 Quick Fetch Specification")
        self.setModal(True)
        self.resize(350, 150)

        layout = QVBoxLayout(self)

        info_label = QLabel(
            "Enter a specific specification (e.g., <b>23.801-01</b>) or an entire series (e.g., <b>23</b>) "
            "to fetch directly from 3GPP.<br><br><i>You can separate multiple targets with commas.</i>"
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("e.g., 23, 38.331, 23.501")
        layout.addWidget(self.input_field)

        btn_layout = QHBoxLayout()
        self.fetch_btn = QPushButton("🚀 Fetch Now")
        self.fetch_btn.clicked.connect(self.accept)
        self.fetch_btn.setEnabled(False)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(self.fetch_btn)
        layout.addLayout(btn_layout)

        self.input_field.textChanged.connect(lambda t: self.fetch_btn.setEnabled(bool(t.strip())))

    def get_targets(self) -> list:
        raw_text = self.input_field.text()
        return [t.strip() for t in raw_text.split(",") if t.strip()]


class ManualSpecFetcherThread(QThread):
    """
    Background worker that queries 3GPP DynaReport HTML using the shared
    scraper engine with real-time log signaling for UI troubleshooting.
    """

    fetch_finished = pyqtSignal(bool, dict, str)
    log_msg = pyqtSignal(str, int)

    def __init__(self, spec_number: str):
        super().__init__()
        self.spec_number = spec_number.strip()

    def run(self):
        try:
            if self.isInterruptionRequested():
                return

            metadata = fetch_metadata_from_dynareport(self.spec_number, log_cb=self.log_msg.emit)

            if self.isInterruptionRequested():
                return

            if metadata and metadata.get("title"):
                spec_type = metadata.get("type") or "TS"
                msg = f"Fetched metadata for {spec_type} {metadata['number']} successfully."
                self.fetch_finished.emit(True, metadata, msg)
            else:
                err_cause = metadata.get("error") if metadata else ""
                reason = f": {err_cause}" if err_cause else " (Check spec number or network)"
                fail_msg = f"Could not extract metadata for {self.spec_number}{reason}"
                self.fetch_finished.emit(False, metadata or {}, fail_msg)
        except Exception as e:
            self.fetch_finished.emit(False, {}, f"Error fetching details: {e}")


class AddSpecDialog(QDialog):
    """Dialog allowing users to query, preview, manually edit, and register an individual specification."""

    def __init__(self, db: SpecsDatabase, parent=None):
        super().__init__(parent)
        self.db = db
        self.fetch_thread: Optional[ManualSpecFetcherThread] = None
        self.sync_requested = False
        self.saved_spec_number = ""

        self.setWindowTitle("➕ Add / Fetch Specification")
        self.setMinimumWidth(560)
        self.setStyleSheet("""
            QDialog { background-color: #FAFAFA; }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #D0D0D0;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: white;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            QLineEdit, QComboBox { padding: 5px; border: 1px solid #CCC; border-radius: 4px; }
            QLineEdit:focus, QComboBox:focus { border: 1px solid #0078D7; }
        """)

        self._setup_ui()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)

        # 1. Query Section
        query_group = QGroupBox("1. Query 3GPP Specification")
        query_layout = QVBoxLayout(query_group)

        row_layout = QHBoxLayout()
        self.query_input = QLineEdit()
        self.query_input.setPlaceholderText("e.g. 23.501, 38.331, 23.700-01...")
        self.query_input.setToolTip("Enter 3GPP specification number")
        self.query_input.returnPressed.connect(self._start_fetch)

        self.btn_fetch = QPushButton("🔍 Fetch Details")
        self.btn_fetch.setCursor(Qt.PointingHandCursor)
        self.btn_fetch.setStyleSheet("""
            QPushButton {
                background-color: #0078D7;
                color: white;
                font-weight: bold;
                padding: 6px 14px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #005A9E; }
            QPushButton:disabled { background-color: #B0D0F0; }
        """)
        self.btn_fetch.clicked.connect(self._start_fetch)

        row_layout.addWidget(self.query_input, 1)
        row_layout.addWidget(self.btn_fetch)
        query_layout.addLayout(row_layout)

        self.lbl_status = QLabel("Enter a specification number and click 'Fetch Details'.")
        self.lbl_status.setStyleSheet("color: #64748B; font-size: 11px; margin-top: 2px;")
        self.lbl_status.setWordWrap(True)
        query_layout.addWidget(self.lbl_status)
        main_layout.addWidget(query_group)

        # 2. Form Preview
        self.preview_group = QGroupBox("2. Specification Details (Editable)")
        form = QFormLayout(self.preview_group)
        form.setLabelAlignment(Qt.AlignRight)

        self.edit_number = QLineEdit()
        self.edit_number.setPlaceholderText("e.g. 23.501")
        form.addRow("Spec Number *:", self.edit_number)

        self.edit_title = QLineEdit()
        self.edit_title.setPlaceholderText("e.g. System architecture for the 5G System (5GS)")
        form.addRow("Title *:", self.edit_title)

        self.type_combo = QComboBox()
        self.type_combo.addItems(["TS", "TR"])
        form.addRow("Type:", self.type_combo)

        self.edit_group = QLineEdit()
        self.edit_group.setPlaceholderText("e.g. SA2, RAN2, CT1")
        form.addRow("Primary Group:", self.edit_group)

        self.edit_init_rel = QLineEdit()
        self.edit_init_rel.setPlaceholderText("e.g. Rel-15")
        form.addRow("Initial Release:", self.edit_init_rel)

        self.edit_tech = QLineEdit()
        self.edit_tech.setPlaceholderText("e.g. 5G, LTE")
        form.addRow("Radio Technology:", self.edit_tech)

        self.chk_sync_now = QCheckBox("Immediately sync files and releases from 3GPP FTP")
        self.chk_sync_now.setChecked(True)
        form.addRow("", self.chk_sync_now)

        main_layout.addWidget(self.preview_group)

        # 3. Actions
        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("💾 Save Specification")
        self.btn_save.setCursor(Qt.PointingHandCursor)
        self.btn_save.setStyleSheet("""
            QPushButton {
                background-color: #107C41;
                color: white;
                font-weight: bold;
                padding: 7px 18px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #0B5A30; }
        """)
        self.btn_save.clicked.connect(self._save_spec)

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_cancel)
        main_layout.addLayout(btn_layout)

    def _start_fetch(self):
        if self.fetch_thread and self.fetch_thread.isRunning():
            return

        query = self.query_input.text().strip()
        if not query:
            QMessageBox.warning(self, "Input Required", "Please enter a specification number.")
            return

        self.btn_fetch.setEnabled(False)
        self.btn_fetch.setText("⏳ Fetching...")
        self.lbl_status.setText(f"⏳ Connecting to 3GPP DynaReport for '{query}'...")
        self.lbl_status.setStyleSheet("color: #0078D7; font-weight: bold;")

        self.fetch_thread = ManualSpecFetcherThread(query)
        self.fetch_thread.log_msg.connect(self._on_fetch_log)
        self.fetch_thread.fetch_finished.connect(self._on_fetch_finished)
        self.fetch_thread.start()

    def _on_fetch_log(self, msg: str, level: int):
        self.lbl_status.setText(msg)
        if level >= logging.ERROR:
            self.lbl_status.setStyleSheet("color: #DC2626; font-weight: bold;")
        elif level >= logging.WARNING:
            self.lbl_status.setStyleSheet("color: #D97706; font-weight: bold;")
        else:
            self.lbl_status.setStyleSheet("color: #0078D7;")

    def _on_fetch_finished(self, success: bool, data: dict, msg: str):
        self.btn_fetch.setEnabled(True)
        self.btn_fetch.setText("🔍 Fetch Details")

        if success:
            self.lbl_status.setText(f"✅ {msg}")
            self.lbl_status.setStyleSheet("color: #107C41; font-weight: bold;")

            self.edit_number.setText(data.get("number", self.query_input.text().strip()))
            self.edit_title.setText(data.get("title", ""))
            self.type_combo.setCurrentText(data.get("type", "TS"))
            self.edit_group.setText(data.get("primary_group", ""))
            self.edit_init_rel.setText(data.get("initial_release", ""))
            self.edit_tech.setText(data.get("radio_technology", ""))
        else:
            self.lbl_status.setText(f"⚠️ {msg}")
            self.lbl_status.setStyleSheet("color: #DC2626; font-weight: bold;")

            raw_num = self.query_input.text().strip()
            if raw_num:
                self.edit_number.setText(raw_num)
                # Deduce TR for xx.8xx and xx.9xx study items even when the report is missing online
                if re.search(r"\b\d{2}\.[89]\d{2}\b", raw_num):
                    self.type_combo.setCurrentText("TR")
                else:
                    self.type_combo.setCurrentText("TS")

    def _save_spec(self):
        spec_num = self.edit_number.text().strip()
        title = self.edit_title.text().strip()

        if not spec_num:
            QMessageBox.warning(self, "Validation Error", "Specification Number is required.")
            return

        data = {
            "number": spec_num,
            "title": title,
            "type": self.type_combo.currentText(),
            "primary_group": self.edit_group.text().strip(),
            "initial_release": self.edit_init_rel.text().strip(),
            "radio_technology": self.edit_tech.text().strip(),
        }

        if self.db.upsert_manual_spec(data):
            self.saved_spec_number = spec_num
            self.sync_requested = self.chk_sync_now.isChecked()
            QMessageBox.information(
                self, "Success", f"Specification {data['type']} {spec_num} saved to database."
            )
            self.accept()
        else:
            QMessageBox.critical(
                self, "Error", f"Failed to save specification {data['type']} {spec_num} to database."
            )

    def reject(self):
        """Cleanly abort worker thread if dialog is dismissed while fetching."""
        if self.fetch_thread and self.fetch_thread.isRunning():
            self.fetch_thread.requestInterruption()
            self.fetch_thread.quit()
            self.fetch_thread.wait(200)
        super().reject()