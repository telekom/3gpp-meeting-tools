"""Inspector dialog displaying Work Item details, impacted specs, and remarks history."""

from __future__ import annotations

import webbrowser
from datetime import datetime

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.utils.portal_utils import get_portal_work_item_url
from modules.meetings.ui.tdocs_menus import _copy_to_clipboard

from .copy_menus import build_wi_copy_data, populate_wi_copy_menu


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

        wi_code = details.get("code", "Unknown")
        acronym = details.get("acronym", "Unknown")
        name = details.get("name", "No Title Available")
        release = details.get("release", "")
        end_date = details.get("end_date", "")

        self.setWindowTitle(f"Work Item Details: {acronym} (#{wi_code})")
        self.setMinimumWidth(620)
        self.setMinimumHeight(480)
        self.setStyleSheet(f"""
            QFrame#cardFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 8px;
            }}
            QLabel {{ font-size: 13px; color: {C.TEXT_STRONG}; }}
            QPushButton#specChipBtn {{
                background-color: {C.ACCENT_SOFT};
                border: 1px solid {C.INFO_BORDER};
                border-radius: 12px;
                padding: 3px 10px;
                font-size: 11px;
                color: {C.ACCENT_PRIMARY};
                font-weight: bold;
            }}
            QPushButton#specChipBtn:hover {{
                background-color: {C.INFO_SOFT_STRONG};
                border-color: {C.ACCENT_PRIMARY};
            }}
            QScrollArea {{ border: none; background-color: transparent; }}
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
        code_badge.setStyleSheet(f"""
            background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY}; border: 1px solid {C.INFO_BORDER};
            border-radius: 4px; padding: 2px 6px; font-size: 12px; font-weight: bold;
        """)
        code_badge.setTextInteractionFlags(Qt.TextSelectableByMouse)

        if release:
            rel_badge = QLabel(f"<b>{release}</b>")
            rel_badge.setStyleSheet(f"""
                background-color: {C.WG_RAN_BG}; color: {C.WG_RAN_TEXT}; border: 1px solid {C.WG_RAN_BORDER};
                border-radius: 4px; padding: 2px 6px; font-size: 12px; font-weight: bold;
            """)
            title_row.addWidget(rel_badge)

        is_finished = bool(end_date and end_date.strip() and end_date < datetime.now().strftime("%Y-%m-%d"))
        status_badge = QLabel("🏁 Finished" if is_finished else "🟢 Active / Ongoing")
        status_badge.setStyleSheet(
            f"background-color: {C.BG_CANVAS}; color: {C.TEXT_SECONDARY}; border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
            if is_finished
            else f"background-color: {C.WG_RAN_BG}; color: {C.WG_RAN_TEXT}; border: 1px solid {C.WG_RAN_BORDER}; border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
        )

        acronym_lbl = QLabel(f"<b>{acronym}</b>")
        acronym_lbl.setStyleSheet(f"font-size: 16px; color: {C.TEXT_STRONG}; font-weight: bold;")
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
        desc_label.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 13px; line-height: 1.4;")
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

        self._add_row(form, "Working Groups", details.get("wg_names") or "-")
        self._add_row(form, "Start Date", details.get("start_date") or "-")
        self._add_row(form, "Target End Date", details.get("end_date") or "-")
        self._add_row(form, "Latest WID", details.get("latest_wid") or "-")

        linked_specs = details.get("linked_specs", [])
        if linked_specs:
            specs_container = QWidget()
            specs_layout = QHBoxLayout(specs_container)
            specs_layout.setContentsMargins(0, 0, 0, 0)
            specs_layout.setSpacing(6)

            for sp in linked_specs:
                num = sp.get("number", "")
                is_pri = sp.get("is_primary", False)
                clean_num = num.split("-")[0].replace(".", "").strip()
                label = f"⭐ {num}" if is_pri else num

                btn = QPushButton(label)
                btn.setObjectName("specChipBtn")
                btn.setCursor(Qt.PointingHandCursor)
                btn.setToolTip(f"{sp.get('title', '')}\nClick to open 3GPP Portal Report")
                btn.clicked.connect(
                    lambda _, n=clean_num: webbrowser.open(f"https://www.3gpp.org/DynaReport/{n}.htm")
                )
                specs_layout.addWidget(btn)

            specs_layout.addStretch()
            form.addRow(self._make_key_label("Impacted Specs:"), specs_container)
        else:
            self._add_row(form, "Impacted Specs", "-")

        layout.addWidget(details_card)

        # --- 3. HISTORICAL REMARKS FEED ---
        remarks = details.get("remarks", [])
        if remarks:
            remarks_lbl = QLabel(f"<b>Secretary Remarks ({len(remarks)}):</b>")
            remarks_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; margin-top: 4px;")
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
                r_card.setStyleSheet(
                    f"background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; padding: 6px;"
                )
                r_box = QVBoxLayout(r_card)
                r_box.setContentsMargins(6, 4, 6, 4)
                r_box.setSpacing(3)

                if rm.get("date"):
                    date_lbl = QLabel(f"📅 {rm['date']}")
                    date_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; font-weight: bold;")
                    r_box.addWidget(date_lbl)

                text_lbl = QLabel(rm.get("text", ""))
                text_lbl.setWordWrap(True)
                text_lbl.setStyleSheet(f"color: {C.TEXT_STRONG}; font-size: 12px;")
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
        portal_btn.clicked.connect(lambda: webbrowser.open(get_portal_work_item_url(wi_code)))
        btn_layout.addWidget(portal_btn)

        copy_btn = QPushButton("📋 Quick Copy ▾")
        copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_menu = QMenu(copy_btn)
        populate_wi_copy_menu(copy_menu, self.details, self)
        copy_btn.setMenu(copy_menu)
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
            self.star_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.WARNING_SOFT_ALT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_ACCENT};
                    border-radius: 4px; padding: 3px 8px; font-size: 11px; font-weight: bold;
                }}
                QPushButton:hover {{ background-color: {C.WARNING_BORDER}; }}
            """)
        else:
            self.star_btn.setText("☆ Star")
            self.star_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.BG_SURFACE}; color: {C.TEXT_MUTED}; border: 1px solid {C.BORDER_STRONG};
                    border-radius: 4px; padding: 3px 8px; font-size: 11px;
                }}
                QPushButton:hover {{ background-color: {C.BG_CANVAS}; color: {C.TEXT_STRONG}; }}
            """)

    def _update_follow_btn_style(self):
        if self.is_following:
            self.follow_btn.setText("🎯 Following")
            self.follow_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.INFO_SOFT_STRONG}; color: {C.INFO_DARK}; border: 1px solid {C.INFO_ACCENT};
                    border-radius: 4px; padding: 3px 8px; font-size: 11px; font-weight: bold;
                }}
                QPushButton:hover {{ background-color: {C.INFO_BORDER}; }}
            """)
        else:
            self.follow_btn.setText("🎯 Follow")
            self.follow_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.BG_SURFACE}; color: {C.TEXT_MUTED}; border: 1px solid {C.BORDER_STRONG};
                    border-radius: 4px; padding: 3px 8px; font-size: 11px;
                }}
                QPushButton:hover {{ background-color: {C.BG_CANVAS}; color: {C.TEXT_STRONG}; }}
            """)

    def _handle_toggle_star(self):
        self.is_starred = not self.is_starred
        self._update_star_btn_style()
        if self.on_toggle_star:
            self.on_toggle_star(str(self.details.get("code", "")).strip())

    def _handle_toggle_follow(self):
        self.is_following = not self.is_following
        self._update_follow_btn_style()
        if self.on_toggle_follow:
            self.on_toggle_follow(str(self.details.get("code", "")).strip())

    def _copy_citation(self):
        d = build_wi_copy_data(self.details)
        _copy_to_clipboard(d["cit_plain"], d["cit_html"], "WI Citation", self)

    def _make_key_label(self, text: str) -> QLabel:
        lbl = QLabel(f"<b>{text}</b>")
        lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 12px;")
        return lbl

    def _add_row(self, form: QFormLayout, label_text: str, value_text: str):
        val_label = QLabel(value_text)
        val_label.setWordWrap(True)
        val_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow(self._make_key_label(f"{label_text}:"), val_label)
