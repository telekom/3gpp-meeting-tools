# --- File: src/modules/specifications/ui/spec_info_dialog.py ---
import logging
import webbrowser

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
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
from modules.specifications.ui.spec_annotations_dialog import EditSpecAnnotationsDialog
from modules.specifications.ui.spec_menu import populate_spec_copy_menu
from modules.specifications.utils.utils import normalize_dynareport_spec_number, build_dynareport_url

logger = logging.getLogger(__name__)

class NumericTableWidgetItem(QTableWidgetItem):
    """Table widget item that sorts numerically rather than alphabetically."""

    def __lt__(self, other):
        try:
            return int(self.text()) < int(other.text())
        except (ValueError, TypeError):
            return super().__lt__(other)


class SpecInfoDialog(QDialog):
    """Modernized Specification Details Dialog with dedicated Work Items table and search filter."""

    def __init__(self, details: dict, parent=None, db=None):
        super().__init__(parent)
        self.details = details
        self.db = db or getattr(parent, "db", None)
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

        clean_number = normalize_dynareport_spec_number(spec_num)
        dynareport_url = build_dynareport_url(spec_num)
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
            "user_annotations",
        }
        for key, value in details.items():
            if key not in excluded_keys and value:
                display_key = key.replace("_", " ").title()
                self._add_row(form, display_key, str(value))

        layout.addWidget(details_card)

        # 3. Dedicated User Tags & Notes Card
        self.annotations_card = self._build_annotations_card()
        layout.addWidget(self.annotations_card)

        # 4. Dedicated Work Items Card (Rendered only when WIs exist)
        if self.related_wis:
            wi_card = self._build_work_items_card()
            layout.addWidget(wi_card, stretch=1)

        # 5. Action Buttons Footer
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

        copy_btn = QPushButton("📋 Quick Copy ▾")
        copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_btn.setToolTip("Copy specification metadata in plain text and rich HTML form.")
        copy_menu = QMenu(self)
        copy_menu.setToolTipsVisible(True)
        populate_spec_copy_menu(copy_menu, self, spec_num, details=self.details, db=self.db)
        copy_btn.setMenu(copy_menu)
        btn_layout.addWidget(copy_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _build_annotations_card(self) -> QFrame:
        """Constructs the personal tags and comments card with live edit trigger."""
        card = QFrame()
        card.setObjectName("cardFrame")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(6)

        header_row = QHBoxLayout()
        title_lbl = QLabel("<b>🏷️ Personal Tags & Notes</b>")
        title_lbl.setStyleSheet("font-size: 13px; color: #2D3748;")
        header_row.addWidget(title_lbl)
        header_row.addStretch()

        edit_btn = QPushButton("✏️ Edit Tags & Notes")
        edit_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        edit_btn.setCursor(Qt.PointingHandCursor)
        edit_btn.clicked.connect(self._edit_annotations)
        header_row.addWidget(edit_btn)
        layout.addLayout(header_row)

        self.tags_display_lbl = QLabel()
        self.tags_display_lbl.setWordWrap(True)
        layout.addWidget(self.tags_display_lbl)

        self.notes_display_lbl = QLabel()
        self.notes_display_lbl.setWordWrap(True)
        self.notes_display_lbl.setStyleSheet(
            "color: #4A5568; font-size: 12px; background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 4px; padding: 6px;"
        )
        layout.addWidget(self.notes_display_lbl)

        self._refresh_annotations_card()
        return card

    def _refresh_annotations_card(self):
        spec_num = self.details.get("number", "")
        ann = self.details.get("user_annotations") or {}
        if self.db and spec_num:
            ann = self.db.get_spec_annotations(spec_num)
            self.details["user_annotations"] = ann

        tags = ann.get("tags", [])
        if tags:
            tag_badges = " ".join([
                f"<span style='background-color: #EBF3FC; color: #1E5C99; border: 1px solid #BDD7EE; border-radius: 8px; padding: 2px 6px; font-size: 11px; font-weight: bold;'>🏷️ {t}</span>"
                for t in tags
            ])
            self.tags_display_lbl.setText(f"<b>Tags:</b> {tag_badges}")
        else:
            self.tags_display_lbl.setText("<span style='color: #A0AEC0;'><i>No tags assigned.</i></span>")

        comments = ann.get("comments", "").strip()
        if comments:
            self.notes_display_lbl.setText(comments.replace("\n", "<br>"))
            self.notes_display_lbl.setVisible(True)
        else:
            self.notes_display_lbl.setText("<span style='color: #A0AEC0;'><i>No personal notes recorded yet.</i></span>")

    def _edit_annotations(self):
        spec_num = self.details.get("number", "")
        title = self.details.get("title", "")
        if not spec_num or not self.db:
            return
        dlg = EditSpecAnnotationsDialog(spec_num, title, self.db, self)
        if dlg.exec_() == QDialog.Accepted:
            self._refresh_annotations_card()
            if self.parent() and hasattr(self.parent(), "refresh_table"):
                self.parent().refresh_table()

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
