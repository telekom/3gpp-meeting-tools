# --- File: src/modules/meetings/ui/session_plan/dialogs/delegation_dispatch_dialogs.py ---
"""Availability-driven session appointment and multi-delegate selection dialogs."""

import html
from typing import Dict, List, Optional, Set, Tuple, Union

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationMember,
    SessionSlot,
)


class SessionDispatchDialog(QDialog):
    """
    Availability-driven session appointment dialog.
    Allows appointing a free delegate on a session-per-session basis to resolve collisions.
    """

    def __init__(
        self,
        slot: SessionSlot,
        original_owner: str,
        members: List[DelegationMember],
        availability_map: Dict[str, dict],
        parent=None,
    ):
        super().__init__(parent)
        self.slot = slot
        self.original_owner = original_owner
        self.members = members
        self.availability_map = availability_map  # member_name -> {'status': 'free'|'busy', 'reason': ...}

        self.setWindowTitle(f"⚡ Appoint Available Delegate — {slot.day_name} {slot.quarter}")
        self.resize(540, 420)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Slot Info Card
        card = QFrame()
        card.setStyleSheet(f"background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; padding: 10px;")
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(6, 6, 6, 6)
        card_lay.setSpacing(4)

        time_str = f"{self.slot.start_time_local} – {self.slot.end_time_local}"
        card_lay.addWidget(QLabel(f"<b>📅 {self.slot.day_name} · {self.slot.quarter} ({time_str}) · Stream {self.slot.stream_number}</b>"))
        card_lay.addWidget(QLabel(f"<b>📋 Topic:</b> {html.escape(self.slot.raw_text)}"))
        if self.original_owner:
            card_lay.addWidget(QLabel(f"<b>Original Topic Owner:</b> {html.escape(self.original_owner)} (⚠️ Collision / Unavailable)"))
        layout.addWidget(card)

        layout.addWidget(QLabel("<b>Select Delegate by Real-Time Availability:</b>"))

        # Table of Delegates ranked by availability
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Status", "Delegate", "Current Schedule / Availability"])
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setAlternatingRowColors(True)

        # Sort: Free members first, then Busy
        sorted_members = sorted(
            self.members,
            key=lambda m: 0 if self.availability_map.get(m.name, {}).get("status") == "free" else 1,
        )

        self.table.setRowCount(len(sorted_members))
        default_select_row = -1

        for row, m in enumerate(sorted_members):
            av = self.availability_map.get(m.name, {"status": "free", "reason": "No other sessions in this slot"})
            is_free = (av.get("status") == "free")

            if is_free and default_select_row == -1 and m.name.lower() != self.original_owner.lower():
                default_select_row = row

            status_txt = "🟢 Free" if is_free else "🔴 Busy"
            status_item = QTableWidgetItem(status_txt)
            status_item.setTextAlignment(Qt.AlignCenter)
            status_item.setData(Qt.UserRole, m.name)

            name_item = QTableWidgetItem(m.name)
            name_item.setForeground(QColor(m.color))
            f = name_item.font()
            f.setBold(True)
            name_item.setFont(f)

            reason_txt = "Available (no parallel session)" if is_free else av.get("reason", "Already scheduled")
            reason_item = QTableWidgetItem(reason_txt)

            if not is_free:
                status_item.setForeground(QColor(C.DANGER_ALT))
                reason_item.setForeground(QColor(C.DANGER_DARK))
                reason_item.setBackground(QColor(C.DANGER_SOFT))
            else:
                status_item.setForeground(QColor(C.GREEN_ACCENT))

            self.table.setItem(row, 0, status_item)
            self.table.setItem(row, 1, name_item)
            self.table.setItem(row, 2, reason_item)

        if default_select_row >= 0:
            self.table.selectRow(default_select_row)

        layout.addWidget(self.table)

        # Appointment Reason / Note
        form = QFormLayout()
        self.reason_edit = QLineEdit()
        self.reason_edit.setPlaceholderText("e.g. Collision with Stream 1; Martin covering AI 5.2.1")
        self.reason_edit.setText(f"Collision in parallel stream with {self.original_owner}" if self.original_owner else "Ad-hoc session coverage")
        form.addRow("<b>Coverage Reason:</b>", self.reason_edit)
        layout.addLayout(form)

        # Buttons
        btn_box = QHBoxLayout()
        revert_btn = QPushButton("↺ Revert to Topic Owner")
        revert_btn.setToolTip("Removes any ad-hoc appointment for this slot, reverting it to the default topic owner.")
        revert_btn.clicked.connect(self._revert_to_owner)
        btn_box.addWidget(revert_btn)

        all_team_btn = QPushButton("👥 Appoint All Team")
        all_team_btn.setStyleSheet(f"background: {C.INDIGO_DARK}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
        all_team_btn.setToolTip("Appoints all delegates in the roster to attend this session together.")
        all_team_btn.clicked.connect(self._appoint_all_team)
        btn_box.addWidget(all_team_btn)

        btn_box.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(cancel_btn)

        appoint_btn = QPushButton("⚡ Appoint Selected")
        appoint_btn.setStyleSheet(f"background: {C.EMERALD_PRIMARY}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
        appoint_btn.clicked.connect(self._validate_and_accept)
        btn_box.addWidget(appoint_btn)
        layout.addLayout(btn_box)

    def _revert_to_owner(self):
        self.reverted = True
        self.accept()

    def _appoint_all_team(self):
        self.reverted = False
        self.appointed_delegate = "ALL"
        self.coverage_reason = self.reason_edit.text().strip() or "All team attending"
        self.accept()

    def _validate_and_accept(self):
        selected_rows = sorted(list(set(item.row() for item in self.table.selectedItems())))
        if not selected_rows:
            QMessageBox.warning(self, "No Delegate Selected", "Please select one or more delegates from the table.")
            return

        del_names: List[str] = []
        busy_names: List[str] = []
        for r in selected_rows:
            item = self.table.item(r, 0)
            if item:
                name = item.data(Qt.UserRole)
                del_names.append(name)
                if "Busy" in item.text():
                    busy_names.append(name)

        if busy_names:
            names_str = ", ".join(busy_names)
            ans = QMessageBox.question(
                self,
                "Delegate Double-Booking Warning",
                f"Warning: The following delegate(s) are already busy in another stream at this time:\n{names_str}\n\nDo you still want to appoint them?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if ans != QMessageBox.Yes:
                return

        self.reverted = False
        if len(del_names) == len(self.members) and len(self.members) > 1:
            self.appointed_delegate = "ALL"
        else:
            self.appointed_delegate = ", ".join(del_names)
        self.coverage_reason = self.reason_edit.text().strip()
        self.accept()

    def get_result(self) -> Tuple[Optional[str], str]:
        """Returns: (appointed_delegate_or_None_if_reverted, reason)."""
        if getattr(self, "reverted", False):
            return None, ""
        return getattr(self, "appointed_delegate", None), getattr(self, "coverage_reason", "")



class MultiDelegateSelectDialog(QDialog):
    """
    Dialog to select multiple delegates or 'All Team' for an Agenda Item or session.
    """

    def __init__(
        self,
        topic_title: str,
        members: List[DelegationMember],
        current_selection: Union[str, List[str]],
        parent=None,
    ):
        super().__init__(parent)
        self.topic_title = topic_title
        self.members = members
        self.current_selection = current_selection

        self.setWindowTitle("👥 Select Delegate(s)")
        self.resize(440, 400)
        self._checkboxes: Dict[str, QCheckBox] = {}
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Header info
        header_card = QFrame()
        header_card.setStyleSheet(f"background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; padding: 8px;")
        h_lay = QVBoxLayout(header_card)
        h_lay.setContentsMargins(6, 6, 6, 6)
        h_lay.addWidget(QLabel("<b>Assign Delegate(s) for Topic:</b>"))
        lbl_top = QLabel(self.topic_title)
        lbl_top.setWordWrap(True)
        lbl_top.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 11px;")
        h_lay.addWidget(lbl_top)
        layout.addWidget(header_card)

        # Parse current selection
        initially_checked: Set[str] = set()
        is_all_initial = False
        if isinstance(self.current_selection, str):
            clean_str = self.current_selection.strip()
            if clean_str.upper() in ("ALL", "ALL TEAM", "ALL DELEGATES", "👥 ALL TEAM", "👥 ALL"):
                is_all_initial = True
                initially_checked = {m.name for m in self.members}
            else:
                initially_checked = {p.strip() for p in clean_str.split(",") if p.strip()}
        elif isinstance(self.current_selection, (list, set)):
            initially_checked = {str(x).strip() for x in self.current_selection if str(x).strip()}
            if len(initially_checked) == len(self.members) and len(self.members) > 0:
                is_all_initial = True

        # "All Team" master checkbox
        self.chk_all = QCheckBox("👥 All Team (Everyone in Delegation)")
        self.chk_all.setStyleSheet(f"font-weight: 700; font-size: 12px; color: {C.INFO_DARK}; padding: 4px;")
        self.chk_all.setChecked(is_all_initial)
        self.chk_all.toggled.connect(self._on_all_toggled)
        layout.addWidget(self.chk_all)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color: {C.BORDER_SUBTLE};")
        layout.addWidget(sep)

        # Scrollable individual member checkboxes
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px;")
        member_widget = QWidget()
        member_lay = QVBoxLayout(member_widget)
        member_lay.setContentsMargins(8, 8, 8, 8)
        member_lay.setSpacing(6)

        for m in self.members:
            row_w = QWidget()
            row_lay = QHBoxLayout(row_w)
            row_lay.setContentsMargins(0, 0, 0, 0)
            row_lay.setSpacing(8)

            badge = QLabel("●")
            badge.setStyleSheet(f"color: {m.color}; font-size: 14px;")
            row_lay.addWidget(badge)

            chk = QCheckBox(m.name)
            chk.setStyleSheet("font-size: 12px; font-weight: 500;")
            chk.setChecked(is_all_initial or m.name in initially_checked)
            chk.toggled.connect(self._on_individual_toggled)
            self._checkboxes[m.name] = chk

            row_lay.addWidget(chk)
            row_lay.addStretch()
            member_lay.addWidget(row_w)

        member_lay.addStretch()
        scroll.setWidget(member_widget)
        layout.addWidget(scroll)

        # Quick action buttons
        link_row = QHBoxLayout()
        btn_sel_all = QPushButton("Select All")
        btn_sel_all.setStyleSheet("padding: 2px 8px; font-size: 11px;")
        btn_sel_all.clicked.connect(lambda: self._set_all_individual_checked(True))
        btn_clr_all = QPushButton("Clear All")
        btn_clr_all.setStyleSheet("padding: 2px 8px; font-size: 11px;")
        btn_clr_all.clicked.connect(lambda: self._set_all_individual_checked(False))
        link_row.addWidget(btn_sel_all)
        link_row.addWidget(btn_clr_all)
        link_row.addStretch()
        layout.addLayout(link_row)

        # Dialog buttons
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("Apply Assignment")
        btn_ok.setStyleSheet(f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
        btn_ok.clicked.connect(self.accept)
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_ok)
        layout.addLayout(btn_box)

    def _on_all_toggled(self, checked: bool):
        if getattr(self, "_block_sync", False):
            return
        self._set_all_individual_checked(checked)

    def _set_all_individual_checked(self, checked: bool):
        self._block_sync = True
        for chk in self._checkboxes.values():
            chk.setChecked(checked)
        self.chk_all.setChecked(checked)
        self._block_sync = False

    def _on_individual_toggled(self):
        if getattr(self, "_block_sync", False):
            return
        self._block_sync = True
        all_checked = bool(self._checkboxes) and all(c.isChecked() for c in self._checkboxes.values())
        self.chk_all.setChecked(all_checked)
        self._block_sync = False

    def get_selected_delegates(self) -> List[str]:
        if self.chk_all.isChecked():
            return [m.name for m in self.members]
        return [name for name, chk in self._checkboxes.items() if chk.isChecked()]

    def get_owner_string(self) -> str:
        """Returns 'ALL' if all members are selected, otherwise comma-separated names or empty."""
        selected = self.get_selected_delegates()
        if not selected:
            return ""
        if len(selected) == len(self.members) and len(self.members) > 1:
            return "ALL"
        return ", ".join(selected)

