# --- File: src/modules/meetings/ui/session_plan/dialogs/slot_and_topic_dialogs.py ---
"""Per-slot notes, physical room location notes, and My Tracked Topics selection dialogs."""

from typing import Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from modules.meetings.core.session_plan.session_plan_models import UserSessionAnnotation


class SessionPlanNoteDialog(QDialog):
    """Allows marking priority, assigned colleague, custom slot location, and adding comments to a slot."""

    def __init__(
        self,
        slot_desc: str,
        current_annotation: Optional[UserSessionAnnotation] = None,
        current_slot_location: str = "",
        room_label: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle(f"📝 Session Note — {slot_desc}")
        self.resize(460, 350)
        self._setup_ui(current_annotation, current_slot_location, room_label)

    def _setup_ui(
        self,
        ann: Optional[UserSessionAnnotation],
        current_slot_location: str = "",
        room_label: str = "",
    ):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        form = QFormLayout()

        # Priority Badge
        self.badge_combo = QComboBox()
        self.badge_combo.addItem("Normal / Unmarked", "normal")
        self.badge_combo.addItem("⭐ Must Attend (Priority)", "must_attend")
        self.badge_combo.addItem("👥 Colleague Covering", "colleague")
        self.badge_combo.addItem("📡 Monitor on TOHRU", "tohru")
        self.badge_combo.addItem("✅ Done / Handled", "done")

        if ann:
            idx = self.badge_combo.findData(ann.status_badge)
            if idx >= 0:
                self.badge_combo.setCurrentIndex(idx)
        form.addRow("<b>Priority / Status:</b>", self.badge_combo)

        # Colleague assignment
        self.colleague_input = QLineEdit()
        self.colleague_input.setPlaceholderText("e.g. Martin, Self, or co-delegate")
        if ann:
            self.colleague_input.setText(ann.assigned_colleague)
        form.addRow("Assigned To:", self.colleague_input)

        # Optional per-slot custom location override
        self.slot_location_input = QLineEdit()
        ph = "e.g., Lobby table, 3rd floor foyer (overrides room default for this slot)"
        if room_label:
            ph = f"Overrides '{room_label}' default for this slot only (optional)"
        self.slot_location_input.setPlaceholderText(ph)
        if current_slot_location:
            self.slot_location_input.setText(current_slot_location)
        form.addRow("📍 Slot Location:", self.slot_location_input)
        layout.addLayout(form)

        # Comment
        layout.addWidget(QLabel("<b>Personal Comment / Action Item:</b>"))
        self.comment_edit = QTextEdit()
        self.comment_edit.setPlaceholderText("e.g. Presenting S2-2610444; check with Ericsson during coffee break...")
        if ann:
            self.comment_edit.setPlainText(ann.comment)
        layout.addWidget(self.comment_edit)

        # Buttons
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(cancel_btn)

        save_btn = QPushButton("💾 Save Note")
        save_btn.setStyleSheet(f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
        save_btn.clicked.connect(self.accept)
        btn_box.addWidget(save_btn)
        layout.addLayout(btn_box)

    def get_data(self) -> Tuple[str, str, str]:
        """Returns: (status_badge, assigned_colleague, comment)."""
        return (
            self.badge_combo.currentData(),
            self.colleague_input.text().strip(),
            self.comment_edit.toPlainText().strip(),
        )

    def get_slot_location(self) -> str:
        """Returns optional per-slot custom location override string."""
        return self.slot_location_input.text().strip()



class SessionPlanRoomLocationsDialog(QDialog):
    """Allows editing hand-made floor/wing/hall notes for each physical room in the meeting."""

    def __init__(
        self,
        meeting_number: str,
        room_infos: List[Tuple[str, str]],
        current_locations: Optional[Dict[str, str]] = None,
        focus_room: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self.meeting_number = str(meeting_number)
        self.room_infos: List[Tuple[str, str]] = list(room_infos or [])
        self.current_locations: Dict[str, str] = dict(current_locations or {})
        self.focus_room = (focus_room or "").strip()
        self._room_inputs: Dict[str, QLineEdit] = {}

        self.setWindowTitle(f"📍 Physical Room Locations — SA2 #{self.meeting_number}")
        self.resize(580, 440)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        banner = QFrame()
        banner.setStyleSheet(
            f"background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px; padding: 8px;"
        )
        b_lay = QVBoxLayout(banner)
        b_lay.setContentsMargins(6, 4, 6, 4)
        b_lay.setSpacing(3)
        info_lbl = QLabel(
            "<b>Hand-Made Room Location Notes:</b> Attach floor, wing, or landmark directions "
            "(e.g., <i>\"3rd floor, Ballroom A\"</i> or <i>\"Level 2, near registration\"</i>) to each physical room.<br>"
            f"<span style='color:{C.TEXT_MUTED};'>Saved locally in <code>meeting_locations.json</code> inside the meeting folder. "
            "Room swaps (e.g. <i>\"In Main Room\"</i>) inherit the target room's location note automatically.</span>"
        )
        info_lbl.setWordWrap(True)
        info_lbl.setStyleSheet("font-size: 11.5px;")
        b_lay.addWidget(info_lbl)
        layout.addWidget(banner)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"QScrollArea {{ border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; background: {C.BG_CARD}; }}")
        container = QWidget()
        self._form_layout = QVBoxLayout(container)
        self._form_layout.setContentsMargins(10, 10, 10, 10)
        self._form_layout.setSpacing(8)

        # Ensure focus_room is included even if not in room_infos yet
        known_lower = {r.lower() for r, _ in self.room_infos}
        if self.focus_room and self.focus_room.lower() not in known_lower:
            self.room_infos.append((self.focus_room, "Custom / Ad-hoc Room"))
            known_lower.add(self.focus_room.lower())

        for room_name, subtitle in self.room_infos:
            self._add_room_row(room_name, subtitle, self.current_locations.get(room_name, ""))

        if not self.room_infos:
            empty_lbl = QLabel("No streams/rooms loaded yet. Use '+ Add Custom Room' below to define a room location.")
            empty_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-style: italic;")
            empty_lbl.setWordWrap(True)
            self._form_layout.addWidget(empty_lbl)

        self._form_layout.addStretch(1)
        scroll.setWidget(container)
        layout.addWidget(scroll, 1)

        # Add custom room row
        add_box = QHBoxLayout()
        add_box.setSpacing(6)
        self.custom_room_name_edit = QLineEdit()
        self.custom_room_name_edit.setPlaceholderText("Custom room name (e.g. Foyer, Breakout 3)...")
        self.custom_room_note_edit = QLineEdit()
        self.custom_room_note_edit.setPlaceholderText("Location note (e.g. 4th floor coffee station)...")
        btn_add_room = QPushButton("➕ Add Room")
        btn_add_room.clicked.connect(self._on_add_custom_room)
        add_box.addWidget(self.custom_room_name_edit, 1)
        add_box.addWidget(self.custom_room_note_edit, 2)
        add_box.addWidget(btn_add_room)
        layout.addLayout(add_box)

        # Bottom buttons
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(cancel_btn)

        save_btn = QPushButton("💾 Save Locations")
        save_btn.setStyleSheet(
            f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 14px; border-radius: 6px;"
        )
        save_btn.clicked.connect(self.accept)
        btn_box.addWidget(save_btn)
        layout.addLayout(btn_box)

        # Pre-focus requested room field if any
        if self.focus_room:
            for r_name, ed in self._room_inputs.items():
                if r_name.lower() == self.focus_room.lower():
                    ed.setFocus()
                    ed.selectAll()
                    break

    def _add_room_row(self, room_name: str, subtitle: str, initial_note: str = ""):
        row_frame = QFrame()
        row_frame.setStyleSheet(
            f"background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; padding: 4px;"
        )
        r_lay = QHBoxLayout(row_frame)
        r_lay.setContentsMargins(6, 4, 6, 4)
        r_lay.setSpacing(8)

        lbl_col = QVBoxLayout()
        lbl_col.setSpacing(1)
        lbl_col.setContentsMargins(0, 0, 0, 0)
        name_lbl = QLabel(f"<b>📍 {room_name}</b>")
        name_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; border: none; background: transparent;")
        lbl_col.addWidget(name_lbl)
        if subtitle:
            sub_lbl = QLabel(subtitle)
            sub_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 10.5px; border: none; background: transparent;")
            lbl_col.addWidget(sub_lbl)
        lbl_w = QWidget()
        lbl_w.setLayout(lbl_col)
        lbl_w.setFixedWidth(185)
        r_lay.addWidget(lbl_w)

        note_edit = QLineEdit()
        note_edit.setPlaceholderText("e.g., 3rd floor, East Wing, past registration...")
        note_edit.setText(initial_note)
        r_lay.addWidget(note_edit, 1)

        btn_clr = QPushButton("✕")
        btn_clr.setFixedWidth(26)
        btn_clr.setToolTip("Clear location note")
        btn_clr.clicked.connect(note_edit.clear)
        r_lay.addWidget(btn_clr)

        self._room_inputs[room_name] = note_edit
        insert_idx = max(0, self._form_layout.count() - 1)
        self._form_layout.insertWidget(insert_idx, row_frame)

    def _on_add_custom_room(self):
        r_name = self.custom_room_name_edit.text().strip()
        r_note = self.custom_room_note_edit.text().strip()
        if not r_name:
            return
        for existing_name, ed in self._room_inputs.items():
            if existing_name.lower() == r_name.lower():
                if r_note:
                    ed.setText(r_note)
                ed.setFocus()
                self.custom_room_name_edit.clear()
                self.custom_room_note_edit.clear()
                return
        self._add_room_row(r_name, "Custom / Ad-hoc Room", r_note)
        self.custom_room_name_edit.clear()
        self.custom_room_note_edit.clear()

    def get_locations(self) -> Dict[str, str]:
        """Returns dictionary of `{room_name: non_empty_location_note}`."""
        return {
            r_name: ed.text().strip()
            for r_name, ed in self._room_inputs.items()
            if ed.text().strip()
        }



class SessionPlanTrackedTopicsDialog(QDialog):
    """Allows selecting which Agenda Items and Work Items to follow/track in My Topics."""

    def __init__(
        self,
        agenda_map: dict,
        currently_tracked: Dict[str, str],
        followed_wi_tokens: Optional[Set[str]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.agenda_map = agenda_map or {}
        self.currently_tracked = {k.lower(): v for k, v in currently_tracked.items()}
        self.followed_wi_tokens = {t.lower() for t in (followed_wi_tokens or set())}

        self.setWindowTitle("🎯 Manage My Tracked Topics & Agenda Items")
        self.resize(680, 520)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Header info
        hdr_frame = QFrame()
        hdr_frame.setObjectName("TrackedTopicsHeader")
        hdr_frame.setStyleSheet(
            f"QFrame#TrackedTopicsHeader {{ background: {C.VIOLET_SURFACE}; border: 1px solid {C.VIOLET_BORDER}; border-radius: 8px; }}"
        )
        hdr_layout = QVBoxLayout(hdr_frame)
        hdr_layout.setContentsMargins(12, 10, 12, 10)
        hdr_layout.setSpacing(4)
        t_lbl = QLabel("<b>Tracked Topics & Followed Agenda Items</b>")
        t_lbl.setStyleSheet(f"color: {C.VIOLET_DARK}; font-size: 13px; background: transparent; border: none;")
        hdr_layout.addWidget(t_lbl)
        desc_lbl = QLabel("Tracked topics are highlighted in purple across the Overview Grid and populate your personal <b>'My Topics'</b> schedule and parallel conflict detector.")
        desc_lbl.setStyleSheet(f"color: {C.VIOLET_PRIMARY}; font-size: 11px; background: transparent; border: none;")
        desc_lbl.setWordWrap(True)
        hdr_layout.addWidget(desc_lbl)
        layout.addWidget(hdr_frame)

        # Filter bar & quick action buttons
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(6)
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("🔍 Filter by AI #, acronym, or topic title...")
        self.search_box.setClearButtonEnabled(True)
        self.search_box.textChanged.connect(self._filter_list)
        filter_layout.addWidget(self.search_box)

        btn_auto = QPushButton("⭐ Track Followed WIs")
        btn_auto.setToolTip("Automatically check all Agenda Items matching your tracked Work Items watchlist.")
        btn_auto.setStyleSheet(f"background: {C.VIOLET_SOFT}; color: {C.VIOLET_PRIMARY}; font-weight: 600; padding: 4px 10px; border-radius: 6px; border: 1px solid {C.VIOLET_BORDER};")
        btn_auto.clicked.connect(self._check_followed_wis)
        filter_layout.addWidget(btn_auto)

        btn_all = QPushButton("☑ All Visible")
        btn_all.clicked.connect(self._check_all_visible)
        filter_layout.addWidget(btn_all)

        btn_none = QPushButton("☐ Clear All")
        btn_none.clicked.connect(self._clear_all)
        filter_layout.addWidget(btn_none)

        layout.addLayout(filter_layout)

        # Checklist
        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet(
            f"QListWidget {{ outline: none; background: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; }}"
            f"QListWidget::item {{ padding: 5px 6px; border-bottom: 1px solid {C.BG_CANVAS}; }}"
        )

        # Populate with agenda items
        import re
        def sort_key(s):
            return [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", s)]

        sorted_ais = sorted(self.agenda_map.keys(), key=sort_key)
        for ai_num in sorted_ais:
            item_data = self.agenda_map[ai_num]
            desc = item_data.description
            acronym = getattr(item_data, "acronym", None)

            disp_text = f"AI {ai_num}: {desc}"
            item = QListWidgetItem(disp_text)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)

            # Check if currently tracked
            ai_lower = ai_num.lower()
            is_checked = (ai_lower in self.currently_tracked) or (acronym and acronym.lower() in self.currently_tracked)
            item.setCheckState(Qt.Checked if is_checked else Qt.Unchecked)

            # Store payload
            item.setData(Qt.UserRole, {
                "ai_number": ai_num,
                "acronym": acronym or "",
                "description": desc,
            })
            self.list_widget.addItem(item)

        layout.addWidget(self.list_widget)

        # Counter & Buttons
        bottom_bar = QHBoxLayout()
        self.count_lbl = QLabel()
        self._update_count()
        self.list_widget.itemChanged.connect(lambda: self._update_count())
        bottom_bar.addWidget(self.count_lbl)

        bottom_bar.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        bottom_bar.addWidget(cancel_btn)

        save_btn = QPushButton("💾 Save Tracked Topics")
        save_btn.setStyleSheet(f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
        save_btn.clicked.connect(self.accept)
        bottom_bar.addWidget(save_btn)

        layout.addLayout(bottom_bar)

    def _filter_list(self, query: str):
        q = query.lower().strip()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            txt = item.text().lower()
            item.setHidden(bool(q and q not in txt))

    def _check_followed_wis(self):
        count = 0
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            data = item.data(Qt.UserRole)
            ai_num = data["ai_number"].lower()
            acronym = data["acronym"].lower()
            desc = data["description"].lower()

            matches = (
                ai_num in self.followed_wi_tokens
                or (acronym and acronym in self.followed_wi_tokens)
                or any(t in desc for t in self.followed_wi_tokens)
            )
            if matches:
                item.setCheckState(Qt.Checked)
                count += 1
        self._update_count()
        QMessageBox.information(self, "Followed WIs Checked", f"Matched and checked {count} Agenda Item(s) against your Work Items watchlist.")

    def _check_all_visible(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if not item.isHidden():
                item.setCheckState(Qt.Checked)
        self._update_count()

    def _clear_all(self):
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(Qt.Unchecked)
        self._update_count()

    def _update_count(self):
        checked = sum(1 for i in range(self.list_widget.count()) if self.list_widget.item(i).checkState() == Qt.Checked)
        self.count_lbl.setText(f"<b>{checked}</b> topic(s) selected")

    def get_selected_topics(self) -> Dict[str, str]:
        """Returns dict of topic_token -> topic_label."""
        res = {}
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                data = item.data(Qt.UserRole)
                ai_num = data["ai_number"]
                acronym = data["acronym"]
                label = f"AI {ai_num}" + (f" ({acronym})" if acronym else "")
                res[ai_num.lower()] = label
                if acronym:
                    res[acronym.lower()] = label
        return res

