# --- File: src/modules/work_items/ui/work_items_dialogs.py ---
import logging
import webbrowser
from datetime import datetime

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY

logger = logging.getLogger(__name__)

class WorkItemInfoDialog(QDialog):
    """Inspector Dialog displaying full Work Item details, linked specs, and remarks history."""

    def __init__(self, details: dict, parent=None):
        super().__init__(parent)
        self.details = details
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
