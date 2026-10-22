# --- File: src/modules/work_items/ui/work_items_dialogs.py ---
import logging
import re
import webbrowser
from datetime import datetime
from typing import Optional

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_scraper import extract_wi_code, fetch_wi_details_from_portal

logger = logging.getLogger(__name__)


class WorkItemInfoDialog(QDialog):
    """Inspector Dialog displaying full Work Item details, linked specs, and remarks history."""

    def __init__(
        self,
        details: dict,
        parent=None,
        is_starred: bool = False,
        is_following: bool = False,
        on_toggle_star=None,
        on_toggle_follow=None,
    ):
        super().__init__(parent)
        self.details = details
        self.is_starred = is_starred
        self.is_following = is_following
        self.on_toggle_star = on_toggle_star
        self.on_toggle_follow = on_toggle_follow

        wi_code = details.get('code', 'Unknown')
        acronym = details.get('acronym', 'Unknown')
        name = details.get('name', 'No Title Available')
        release = details.get('release', '')
        end_date = details.get('end_date', '')

        self.setWindowTitle(f"Work Item Details: {acronym} (#{wi_code})")
        self.setMinimumWidth(620)
        self.setMinimumHeight(480)
        self.setStyleSheet("""
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
            }
            QLabel { font-size: 13px; color: #1E293B; }
            QPushButton#specChipBtn {
                background-color: #EBF3FC;
                border: 1px solid #BFDBFE;
                border-radius: 12px;
                padding: 3px 10px;
                font-size: 11px;
                color: #1E5C99;
                font-weight: bold;
            }
            QPushButton#specChipBtn:hover {
                background-color: #DBEAFE;
                border-color: #1E5C99;
            }
            QScrollArea { border: none; background-color: transparent; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # --- 1. HEADER CARD ---
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(8)

        title_row = QHBoxLayout()

        code_badge = QLabel(f"<b>WI #{wi_code}</b>")
        code_badge.setStyleSheet("""
            background-color: #EBF3FC; color: #1E5C99; border: 1px solid #BFDBFE;
            border-radius: 4px; padding: 2px 6px; font-size: 12px; font-weight: bold;
        """)
        code_badge.setTextInteractionFlags(Qt.TextSelectableByMouse)

        if release:
            rel_badge = QLabel(f"<b>{release}</b>")
            rel_badge.setStyleSheet("""
                background-color: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0;
                border-radius: 4px; padding: 2px 6px; font-size: 12px; font-weight: bold;
            """)
            title_row.addWidget(rel_badge)

        is_finished = bool(end_date and end_date.strip() and end_date < datetime.now().strftime("%Y-%m-%d"))
        status_badge = QLabel("🏁 Finished" if is_finished else "🟢 Active / Ongoing")
        status_badge.setStyleSheet(
            "background-color: #F1F5F9; color: #475569; border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
            if is_finished else
            "background-color: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
        )

        acronym_lbl = QLabel(f"<b>{acronym}</b>")
        acronym_lbl.setStyleSheet("font-size: 16px; color: #1E293B; font-weight: bold;")
        acronym_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)

        title_row.addWidget(code_badge)
        title_row.addWidget(acronym_lbl)
        title_row.addWidget(status_badge)
        title_row.addStretch()

        # Starred Toggle Button
        self.star_btn = QPushButton()
        self.star_btn.setCursor(Qt.PointingHandCursor)
        self.star_btn.setToolTip("Toggle Starred status (Personal Interest / Watchlist)")
        self._update_star_btn_style()
        self.star_btn.clicked.connect(self._handle_toggle_star)
        title_row.addWidget(self.star_btn)

        # Following Toggle Button
        self.follow_btn = QPushButton()
        self.follow_btn.setCursor(Qt.PointingHandCursor)
        self.follow_btn.setToolTip("Toggle Following status (Assigned Delegation Responsibility)")
        self._update_follow_btn_style()
        self.follow_btn.clicked.connect(self._handle_toggle_follow)
        title_row.addWidget(self.follow_btn)

        header_layout.addLayout(title_row)

        desc_label = QLabel(name)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #475569; font-size: 13px; line-height: 1.4;")
        desc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        header_layout.addWidget(desc_label)

        layout.addWidget(header_card)

        # --- 2. DETAILS & TIMELINE CARD ---
        details_card = QFrame()
        details_card.setObjectName("cardFrame")
        form = QFormLayout(details_card)
        form.setContentsMargins(14, 12, 14, 12)
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignRight)

        self._add_row(form, "Working Groups", details.get('wg_names') or '-')
        self._add_row(form, "Start Date", details.get('start_date') or '-')
        self._add_row(form, "Target End Date", details.get('end_date') or '-')
        self._add_row(form, "Latest WID", details.get('latest_wid') or '-')

        linked_specs = details.get('linked_specs', [])
        if linked_specs:
            specs_container = QWidget()
            specs_layout = QHBoxLayout(specs_container)
            specs_layout.setContentsMargins(0, 0, 0, 0)
            specs_layout.setSpacing(6)

            for sp in linked_specs:
                num = sp.get('number', '')
                is_pri = sp.get('is_primary', False)
                clean_num = num.split("-")[0].replace('.', '').strip()
                label = f"⭐ {num}" if is_pri else num

                btn = QPushButton(label)
                btn.setObjectName("specChipBtn")
                btn.setCursor(Qt.PointingHandCursor)
                btn.setToolTip(f"{sp.get('title', '')}\nClick to open 3GPP Portal Report")
                btn.clicked.connect(lambda _, n=clean_num: webbrowser.open(f"https://www.3gpp.org/DynaReport/{n}.htm"))
                specs_layout.addWidget(btn)

            specs_layout.addStretch()
            form.addRow(self._make_key_label("Impacted Specs:"), specs_container)
        else:
            self._add_row(form, "Impacted Specs", "-")

        layout.addWidget(details_card)

        # --- 3. HISTORICAL REMARKS FEED ---
        remarks = details.get('remarks', [])
        if remarks:
            remarks_lbl = QLabel(f"<b>Secretary Remarks ({len(remarks)}):</b>")
            remarks_lbl.setStyleSheet("color: #475569; font-size: 12px; margin-top: 4px;")
            layout.addWidget(remarks_lbl)

            remarks_scroll = QScrollArea()
            remarks_scroll.setWidgetResizable(True)
            remarks_scroll.setMaximumHeight(160)

            remarks_content = QWidget()
            remarks_vbox = QVBoxLayout(remarks_content)
            remarks_vbox.setContentsMargins(0, 0, 0, 0)
            remarks_vbox.setSpacing(6)

            for rm in remarks:
                r_card = QFrame()
                r_card.setStyleSheet("background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px; padding: 6px;")
                r_box = QVBoxLayout(r_card)
                r_box.setContentsMargins(6, 4, 6, 4)
                r_box.setSpacing(3)

                if rm.get('date'):
                    date_lbl = QLabel(f"📅 {rm['date']}")
                    date_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold;")
                    r_box.addWidget(date_lbl)

                text_lbl = QLabel(rm.get('text', ''))
                text_lbl.setWordWrap(True)
                text_lbl.setStyleSheet("color: #1E293B; font-size: 12px;")
                text_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
                r_box.addWidget(text_lbl)

                remarks_vbox.addWidget(r_card)

            remarks_vbox.addStretch()
            remarks_scroll.setWidget(remarks_content)
            layout.addWidget(remarks_scroll)

        # --- 4. ACTION BUTTONS ---
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        portal_btn = QPushButton("🌐 Open 3GPP Portal")
        portal_btn.setObjectName("primaryBtn")
        portal_btn.setCursor(Qt.PointingHandCursor)
        portal_btn.clicked.connect(lambda: webbrowser.open(
            f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}"
        ))
        btn_layout.addWidget(portal_btn)

        copy_btn = QPushButton("📋 Copy Citation")
        copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_btn.clicked.connect(self._copy_citation)
        btn_layout.addWidget(copy_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _update_star_btn_style(self):
        if self.is_starred:
            self.star_btn.setText("⭐ Starred")
            self.star_btn.setStyleSheet("""
                QPushButton {
                    background-color: #FEF3C7; color: #92400E; border: 1px solid #F59E0B;
                    border-radius: 4px; padding: 3px 8px; font-size: 11px; font-weight: bold;
                }
                QPushButton:hover { background-color: #FDE68A; }
            """)
        else:
            self.star_btn.setText("☆ Star")
            self.star_btn.setStyleSheet("""
                QPushButton {
                    background-color: #F8FAFC; color: #64748B; border: 1px solid #CBD5E1;
                    border-radius: 4px; padding: 3px 8px; font-size: 11px;
                }
                QPushButton:hover { background-color: #F1F5F9; color: #1E293B; }
            """)

    def _update_follow_btn_style(self):
        if self.is_following:
            self.follow_btn.setText("🎯 Following")
            self.follow_btn.setStyleSheet("""
                QPushButton {
                    background-color: #DBEAFE; color: #1E40AF; border: 1px solid #3B82F6;
                    border-radius: 4px; padding: 3px 8px; font-size: 11px; font-weight: bold;
                }
                QPushButton:hover { background-color: #BFDBFE; }
            """)
        else:
            self.follow_btn.setText("🎯 Follow")
            self.follow_btn.setStyleSheet("""
                QPushButton {
                    background-color: #F8FAFC; color: #64748B; border: 1px solid #CBD5E1;
                    border-radius: 4px; padding: 3px 8px; font-size: 11px;
                }
                QPushButton:hover { background-color: #F1F5F9; color: #1E293B; }
            """)

    def _handle_toggle_star(self):
        self.is_starred = not self.is_starred
        self._update_star_btn_style()
        if self.on_toggle_star:
            self.on_toggle_star(str(self.details.get('code', '')).strip())

    def _handle_toggle_follow(self):
        self.is_following = not self.is_following
        self._update_follow_btn_style()
        if self.on_toggle_follow:
            self.on_toggle_follow(str(self.details.get('code', '')).strip())

    def _copy_citation(self):
        c = f"WI {self.details.get('code')}: {self.details.get('acronym')} — {self.details.get('name')} ({self.details.get('release')}, {self.details.get('wg_names')})"
        QApplication.clipboard().setText(c)

    def _make_key_label(self, text: str) -> QLabel:
        lbl = QLabel(f"<b>{text}</b>")
        lbl.setStyleSheet("color: #64748B; font-size: 12px;")
        return lbl

    def _add_row(self, form: QFormLayout, label_text: str, value_text: str):
        val_label = QLabel(value_text)
        val_label.setWordWrap(True)
        val_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow(self._make_key_label(f"{label_text}:"), val_label)


class LinkedSpecsDialog(QDialog):
    """Dialog displaying all specifications linked to a given Work Item."""

    def __init__(self, wi_code: str, acronym: str, specs: list, parent=None):
        super().__init__(parent)
        title_str = f"Linked Specifications for WI {acronym} ({wi_code})" if acronym else f"Linked Specifications for WI #{wi_code}"
        self.setWindowTitle(title_str)
        self.setMinimumWidth(580)
        self.setMinimumHeight(380)
        self.setStyleSheet("""
            QFrame#specCard { background-color: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px; }
            QFrame#specCard:hover { border-color: #94A3B8; background-color: #FAFCFF; }
            QLabel { color: #1E293B; }
            QScrollArea { border: 1px solid #CBD5E1; border-radius: 6px; background-color: #FFFFFF; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        header_text = (
            f"Specifications Impacted by Work Item <b>{acronym}</b> ({wi_code}):"
            if acronym else f"Specifications Impacted by Work Item <b>#{wi_code}</b>:"
        )
        header = QLabel(header_text)
        header.setStyleSheet("font-size: 14px; color: #1E293B;")
        header.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(header)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        scroll_content = QWidget()
        cards_layout = QVBoxLayout(scroll_content)
        cards_layout.setContentsMargins(8, 8, 8, 8)
        cards_layout.setSpacing(8)

        if specs:
            for sp in specs:
                num = sp.get('number', 'Unknown')
                title = sp.get('title', 'No Title Available')
                spec_type = sp.get('type', 'TS')
                series_name = sp.get('series_name', num.split('.')[0] if '.' in num else '')
                is_pri = bool(sp.get('is_primary', False))

                card = self._create_spec_card(num, title, spec_type, series_name, is_pri)
                cards_layout.addWidget(card)
        else:
            no_data = QLabel("No linked specifications recorded in the local database for this Work Item.")
            no_data.setStyleSheet("color: #64748B; font-style: italic; padding: 25px;")
            no_data.setAlignment(Qt.AlignCenter)
            cards_layout.addWidget(no_data)

        cards_layout.addStretch()
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

        footer_layout = QHBoxLayout()
        count_lbl = QLabel(f"<b>Total:</b> {len(specs)} specification(s)")
        count_lbl.setStyleSheet("color: #64748B; font-size: 12px;")
        footer_layout.addWidget(count_lbl)
        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        layout.addLayout(footer_layout)

    def _create_spec_card(self, num: str, title: str, spec_type: str, series_name: str, is_primary: bool) -> QFrame:
        card = QFrame()
        card.setObjectName("specCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(6)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        type_badge = QLabel(f"<b>{spec_type}</b>")
        type_badge.setStyleSheet("""
            background-color: #EBF3FC; color: #1E5C99; border: 1px solid #BFDBFE;
            border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;
        """)
        top_row.addWidget(type_badge)

        num_lbl = QLabel(f"<b>{num}</b>")
        num_lbl.setStyleSheet("font-size: 14px; font-weight: bold; color: #1E293B;")
        num_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        top_row.addWidget(num_lbl)

        if is_primary:
            pri_badge = QLabel("⭐ Primary Spec")
            pri_badge.setStyleSheet("""
                background-color: #FEF3C7; color: #92400E; border: 1px solid #FDE68A;
                border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;
            """)
            top_row.addWidget(pri_badge)

        top_row.addStretch()

        clean_num = num.split("-")[0].replace('.', '').strip()
        dyna_url = f"https://www.3gpp.org/DynaReport/{clean_num}.htm" if clean_num else ""
        ftp_url = f"https://www.3gpp.org/ftp/Specs/archive/{series_name}_series/{num}/" if series_name else ""

        if dyna_url:
            dyna_btn = QPushButton("🌐 DynaReport")
            dyna_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            dyna_btn.setCursor(Qt.PointingHandCursor)
            dyna_btn.clicked.connect(lambda: webbrowser.open(dyna_url))
            top_row.addWidget(dyna_btn)

        if ftp_url:
            ftp_btn = QPushButton("📂 FTP Archive")
            ftp_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            ftp_btn.setCursor(Qt.PointingHandCursor)
            ftp_btn.clicked.connect(lambda: webbrowser.open(ftp_url))
            top_row.addWidget(ftp_btn)

        card_layout.addLayout(top_row)

        title_lbl = QLabel(title)
        title_lbl.setWordWrap(True)
        title_lbl.setStyleSheet("font-size: 12px; color: #475569; line-height: 1.4;")
        title_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        card_layout.addWidget(title_lbl)

        return card


class ManualWIFetcherThread(QThread):
    """
    Background worker that fetches and parses Work Item details directly from
    the 3GPP Portal without blocking the GUI thread.
    """
    fetch_finished = pyqtSignal(bool, dict, str)
    log_msg = pyqtSignal(str, int)

    def __init__(self, query: str):
        super().__init__()
        self.query = query.strip()

    def run(self):
        try:
            if self.isInterruptionRequested():
                return

            code = extract_wi_code(self.query)
            if not code:
                self.fetch_finished.emit(False, {}, "Could not determine a valid Work Item ID from the input.")
                return

            self.log_msg.emit(f"Connecting to 3GPP Portal for WI #{code}...", logging.INFO)
            metadata = fetch_wi_details_from_portal(code, log_cb=self.log_msg.emit)

            if self.isInterruptionRequested():
                return

            if metadata and (metadata.get("name") or metadata.get("acronym")):
                acr = metadata.get("acronym") or metadata.get("name")
                remarks_cnt = len(metadata.get("remarks", []))
                msg = f"Fetched metadata for WI #{code} ({acr}) successfully ({remarks_cnt} remarks)."
                self.fetch_finished.emit(True, metadata, msg)
            else:
                err_cause = metadata.get("error") if metadata else ""
                reason = f": {err_cause}" if err_cause else " (Check WI ID or network connection)"
                fail_msg = f"Could not extract details for WI #{code}{reason}"
                self.fetch_finished.emit(False, metadata or {}, fail_msg)
        except Exception as e:
            self.fetch_finished.emit(False, {}, f"Error fetching details: {e}")


class AddWorkItemDialog(QDialog):
    """Dialog allowing users to query, preview, manually edit, and register an individual Work Item."""

    def __init__(self, db: WorkItemsDatabase, initial_query: str = "", parent=None):
        super().__init__(parent)
        self.db = db
        self.fetch_thread: Optional[ManualWIFetcherThread] = None
        self.saved_wi_code = ""
        self.is_starred_requested = False
        self.is_following_requested = False
        self._fetched_remarks = []

        self.setWindowTitle("➕ Add / Fetch Work Item")
        self.setMinimumWidth(580)
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

        if initial_query:
            self.query_input.setText(initial_query)
            self._start_fetch()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)

        # 1. Query Section
        query_group = QGroupBox("1. Query 3GPP Work Item")
        query_layout = QVBoxLayout(query_group)

        row_layout = QHBoxLayout()
        self.query_input = QLineEdit()
        self.query_input.setPlaceholderText("e.g. 1121123, #1121123, or https://portal.3gpp.org/...workitemId=1121123")
        self.query_input.setToolTip("Enter a 3GPP Work Item ID or Portal URL")
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

        self.lbl_status = QLabel("Enter a Work Item ID or URL and click 'Fetch Details'.")
        self.lbl_status.setStyleSheet("color: #64748B; font-size: 11px; margin-top: 2px;")
        self.lbl_status.setWordWrap(True)
        query_layout.addWidget(self.lbl_status)
        main_layout.addWidget(query_group)

        # 2. Form Preview (Editable)
        self.preview_group = QGroupBox("2. Work Item Details (Editable)")
        form = QFormLayout(self.preview_group)
        form.setLabelAlignment(Qt.AlignRight)

        self.edit_code = QLineEdit()
        self.edit_code.setPlaceholderText("e.g. 1121123")
        form.addRow("WI Code *:", self.edit_code)

        self.edit_acronym = QLineEdit()
        self.edit_acronym.setPlaceholderText("e.g. NR_Sensing_bis-Core")
        form.addRow("Acronym *:", self.edit_acronym)

        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("e.g. Core part: Integrated Sensing and communication (ISAC) for NR")
        form.addRow("Title / Name *:", self.edit_name)

        self.edit_release = QLineEdit()
        self.edit_release.setPlaceholderText("e.g. Rel-20, Rel-19, Rel-18...")
        form.addRow("Release:", self.edit_release)

        self.edit_groups = QLineEdit()
        self.edit_groups.setPlaceholderText("e.g. RAN3 (comma-separated for multiple)")
        form.addRow("Responsible WG(s):", self.edit_groups)

        self.edit_status = QLineEdit()
        self.edit_status.setPlaceholderText("e.g. PCG approved, In progress, Finished...")
        form.addRow("Status:", self.edit_status)

        self.edit_start_date = QLineEdit()
        self.edit_start_date.setPlaceholderText("YYYY-MM-DD")
        form.addRow("Start Date:", self.edit_start_date)

        self.edit_end_date = QLineEdit()
        self.edit_end_date.setPlaceholderText("YYYY-MM-DD")
        form.addRow("End Date:", self.edit_end_date)

        self.edit_latest_wid = QLineEdit()
        self.edit_latest_wid.setPlaceholderText("e.g. RP-261901")
        form.addRow("Latest WID:", self.edit_latest_wid)

        tracking_layout = QVBoxLayout()
        tracking_layout.setSpacing(4)
        self.chk_starred = QCheckBox("⭐ Add to Starred (Personal Watchlist)")
        self.chk_following = QCheckBox("🎯 Add to Following (Delegation Responsibilities)")
        tracking_layout.addWidget(self.chk_starred)
        tracking_layout.addWidget(self.chk_following)
        form.addRow("Tracking:", tracking_layout)

        main_layout.addWidget(self.preview_group)

        # 3. Actions
        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("💾 Save Work Item")
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
        self.btn_save.clicked.connect(self._save_wi)

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
        code = extract_wi_code(query)
        if not code:
            QMessageBox.warning(self, "Input Required", "Please enter a valid Work Item ID or 3GPP Portal URL.")
            return

        self.btn_fetch.setEnabled(False)
        self.btn_fetch.setText("⏳ Fetching...")
        self.lbl_status.setText(f"⏳ Connecting to 3GPP Portal for WI #{code}...")
        self.lbl_status.setStyleSheet("color: #0078D7; font-weight: bold;")

        self.fetch_thread = ManualWIFetcherThread(query)
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

            self.edit_code.setText(data.get("code", ""))
            self.edit_acronym.setText(data.get("acronym", ""))
            self.edit_name.setText(data.get("name", ""))
            self.edit_release.setText(data.get("release", ""))
            groups = data.get("responsible_groups", [])
            self.edit_groups.setText(", ".join(groups) if isinstance(groups, list) else str(groups))
            self.edit_status.setText(data.get("status", ""))
            self.edit_start_date.setText(data.get("start_date", ""))
            self.edit_end_date.setText(data.get("end_date", ""))
            self.edit_latest_wid.setText(data.get("latest_wid", ""))
            self._fetched_remarks = data.get("remarks", [])
        else:
            self.lbl_status.setText(f"⚠️ {msg}")
            self.lbl_status.setStyleSheet("color: #DC2626; font-weight: bold;")

    def _save_wi(self):
        code = self.edit_code.text().strip()
        acronym = self.edit_acronym.text().strip()
        name = self.edit_name.text().strip()

        if not code:
            QMessageBox.warning(self, "Validation Error", "Please provide a valid Work Item Code.")
            self.edit_code.setFocus()
            return

        if not acronym and not name:
            QMessageBox.warning(self, "Validation Error", "Please provide at least a Work Item Acronym or Title.")
            self.edit_acronym.setFocus()
            return

        groups_str = self.edit_groups.text().strip()
        groups = [g.strip() for g in re.split(r'[,;/|\s]+', groups_str) if g.strip()]

        wi_data = {
            "code": code,
            "acronym": acronym,
            "name": name,
            "release": self.edit_release.text().strip(),
            "latest_wid": self.edit_latest_wid.text().strip(),
            "start_date": self.edit_start_date.text().strip(),
            "end_date": self.edit_end_date.text().strip(),
            "responsible_groups": groups,
            "remarks": self._fetched_remarks
        }

        if self.db.upsert_single_work_item(wi_data):
            self.saved_wi_code = code
            self.is_starred_requested = self.chk_starred.isChecked()
            self.is_following_requested = self.chk_following.isChecked()
            self.accept()
        else:
            QMessageBox.critical(self, "Database Error", f"Failed to save Work Item #{code} to the database.")

    def closeEvent(self, event):
        if self.fetch_thread and self.fetch_thread.isRunning():
            self.fetch_thread.requestInterruption()
            self.fetch_thread.wait(1000)
        super().closeEvent(event)
