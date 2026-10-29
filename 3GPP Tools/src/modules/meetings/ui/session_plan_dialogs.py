# --- File: src/modules/meetings/ui/session_plan_dialogs.py ---
"""
Dialogs for SA2 Session Plan Management:
1. SessionPlanTimezoneDialog: Venue vs Austrian Corporate Laptop Clock selector.
2. SessionPlanImportDialog: Manual ingestion via URL, TDoc#, File, or Paste.
3. SessionPlanDiffDialog: Visual version comparison & changelog.
4. SessionPlanNoteDialog: Personal slot notes and colleague delegation.
"""
import logging
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple, Union

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QColorDialog,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from modules.meetings.core.session_plan_diff import SessionPlanDiffEngine, SessionPlanDiffResult
from modules.meetings.core.session_plan_exporter import SessionPlanExporter
from modules.meetings.core.session_plan_models import (
    DELEGATE_PALETTE,
    DelegationMember,
    DelegationPlan,
    SessionAppointment,
    SessionPlanSnapshot,
    SessionSlot,
    TopicOwnership,
    UserSessionAnnotation,
)
from modules.meetings.core.session_plan_timezone import SessionPlanTimezone

logger = logging.getLogger(__name__)

TIMEZONE_PRESETS = [
    (120, "UTC+02:00 — Europe CEST (Prague, Vienna, Paris, Gothenburg - Summer)"),
    (60, "UTC+01:00 — Europe CET (Vienna, Paris, Berlin - Winter)"),
    (540, "UTC+09:00 — Asia JST/KST (Tokyo, Fukuoka, Seoul, Jeju)"),
    (480, "UTC+08:00 — Asia CST (Shanghai, Beijing, Taipei)"),
    (600, "UTC+10:00 — Australia AEST (Brisbane, Gold Coast)"),
    (-420, "UTC-07:00 — US PDT (San Diego, Reno - Summer)"),
    (-480, "UTC-08:00 — US PST (San Diego, Reno - Winter)"),
    (-300, "UTC-05:00 — US CDT / Cancun (Chicago Summer / Florida Winter)"),
    (-240, "UTC-04:00 — US EDT (Orlando, New York - Summer)"),
    (-360, "UTC-06:00 — US CST (Chicago, Dallas - Winter)"),
]


class SessionPlanTimezoneDialog(QDialog):
    """Allows selecting or confirming the Venue Local Timezone."""

    def __init__(
        self,
        meeting_number: str,
        location: str = "",
        initial_offset_min: Optional[int] = None,
        initial_tz_name: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self.meeting_number = meeting_number
        self.location = location
        self.laptop_offset_min = SessionPlanTimezone.get_laptop_tz_offset_minutes()

        self.setWindowTitle(f"🌍 Meeting Timezone — SA2 #{meeting_number}")
        self.resize(540, 360)
        self._setup_ui(initial_offset_min, initial_tz_name)

    def _setup_ui(self, initial_offset_min: Optional[int], initial_tz_name: str):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Header Info Banner
        info_frame = QFrame()
        info_frame.setStyleSheet("background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 10px;")
        info_layout = QVBoxLayout(info_frame)
        info_layout.setSpacing(4)

        loc_lbl = QLabel(f"📍 <b>Meeting Location:</b> {self.location or 'Unknown'}")
        info_layout.addWidget(loc_lbl)

        lap_sign = "+" if self.laptop_offset_min >= 0 else "-"
        lap_h = abs(self.laptop_offset_min) // 60
        lap_m = abs(self.laptop_offset_min) % 60
        lap_str = f"UTC{lap_sign}{lap_h:02d}:{lap_m:02d}"
        laptop_lbl = QLabel(f"💻 <b>Your Corporate Laptop Clock:</b> {lap_str} (Austrian Time)")
        laptop_lbl.setStyleSheet("color: #475569;")
        info_layout.addWidget(laptop_lbl)

        expl = QLabel("3GPP Session Plans are published in <b>Meeting Local Time</b>. Selecting the correct venue timezone ensures live session indicators, reminders, and calendar exports trigger accurately.")
        expl.setWordWrap(True)
        expl.setStyleSheet("font-size: 11px; color: #64748B; margin-top: 4px;")
        info_layout.addWidget(expl)
        layout.addWidget(info_frame)

        # Preset selection
        form_layout = QFormLayout()
        self.preset_combo = QComboBox()
        for off, label in TIMEZONE_PRESETS:
            self.preset_combo.addItem(label, off)
        self.preset_combo.addItem("Custom Offset...", "custom")

        # Select initial or best guess
        target_off = initial_offset_min if initial_offset_min is not None else 120
        idx = -1
        for i in range(self.preset_combo.count()):
            if self.preset_combo.itemData(i) == target_off:
                idx = i
                break
        if idx >= 0:
            self.preset_combo.setCurrentIndex(idx)

        form_layout.addRow("<b>Venue Timezone:</b>", self.preset_combo)

        # Custom spinbox (initially disabled unless 'custom' chosen)
        self.custom_spin = QSpinBox()
        self.custom_spin.setRange(-720, 840)
        self.custom_spin.setSingleStep(30)
        self.custom_spin.setValue(target_off)
        self.custom_spin.setSuffix(" minutes from UTC")
        self.custom_spin.setEnabled(self.preset_combo.currentData() == "custom")
        form_layout.addRow("Custom UTC Offset:", self.custom_spin)

        self.preset_combo.currentIndexChanged.connect(self._on_preset_changed)
        layout.addLayout(form_layout)

        # Difference alert preview
        self.diff_preview_lbl = QLabel()
        self.diff_preview_lbl.setStyleSheet("font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 6px;")
        layout.addWidget(self.diff_preview_lbl)
        self._update_diff_preview()

        layout.addStretch()

        # Buttons
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(cancel_btn)

        save_btn = QPushButton("✔ Confirm Timezone")
        save_btn.setStyleSheet("background: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
        save_btn.clicked.connect(self.accept)
        btn_box.addWidget(save_btn)
        layout.addLayout(btn_box)

    def _on_preset_changed(self):
        is_custom = (self.preset_combo.currentData() == "custom")
        self.custom_spin.setEnabled(is_custom)
        if not is_custom:
            self.custom_spin.setValue(int(self.preset_combo.currentData()))
        self._update_diff_preview()

    def _update_diff_preview(self):
        offset = self.get_selected_offset()
        diff = self.laptop_offset_min - offset
        if diff == 0:
            self.diff_preview_lbl.setText("✨ Venue timezone matches your Austrian laptop clock (Zero offset).")
            self.diff_preview_lbl.setStyleSheet("background: #F0FDF4; color: #166534;")
        else:
            diff_h = abs(diff) // 60
            ahead_behind = "ahead of" if diff < 0 else "behind"
            self.diff_preview_lbl.setText(f"🕒 Venue is {diff_h} hour(s) {ahead_behind} your Austrian laptop clock.")
            self.diff_preview_lbl.setStyleSheet("background: #FFFBEB; color: #B45309;")

    def get_selected_offset(self) -> int:
        data = self.preset_combo.currentData()
        if data == "custom":
            return self.custom_spin.value()
        return int(data)

    def get_selected_tz_name(self) -> str:
        data = self.preset_combo.currentData()
        if data == "custom":
            val = self.custom_spin.value()
            sign = "+" if val >= 0 else "-"
            return f"UTC{sign}{abs(val)//60:02d}:{abs(val)%60:02d} (Custom)"
        return self.preset_combo.currentText().split("—")[0].strip()


class SessionPlanImportDialog(QDialog):
    """Modeless dialog supporting 100% manual ingestion via URL, TDoc#, File, or Paste."""

    import_requested = pyqtSignal(str, str)  # (mode: 'url'|'file'|'paste', payload)

    def __init__(self, meeting_number: str, parent=None):
        super().__init__(parent)
        self.meeting_number = meeting_number
        self.setWindowTitle(f"📥 Import Session Plan — SA2 #{meeting_number}")
        self.resize(580, 420)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        tabs = QTabWidget()

        # Tab 1: URL / TDoc#
        tab_url = QWidget()
        vbox_url = QVBoxLayout(tab_url)
        vbox_url.addWidget(QLabel("<b>Enter 3GPP URL or TDoc Number:</b>"))
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("e.g. S2-2610195 or https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/.../S2-2610195.zip")
        vbox_url.addWidget(self.url_input)

        hint = QLabel("💡 <i>Entering just the TDoc number (e.g. <b>S2-2610195</b>) automatically checks the local on-site meeting server (10.10.10.10), SYNC folder, /INBOX/, and /Docs/.</i>")
        hint.setWordWrap(True)
        hint.setStyleSheet("font-size: 11px; color: #64748B; margin-top: 4px;")
        vbox_url.addWidget(hint)
        vbox_url.addStretch()

        btn_url = QPushButton("⬇️ Download && Import")
        btn_url.setStyleSheet("background: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
        btn_url.clicked.connect(self._on_import_url)
        vbox_url.addWidget(btn_url)
        tabs.addTab(tab_url, "🔗 URL / TDoc#")

        # Tab 2: Browse File
        tab_file = QWidget()
        vbox_file = QVBoxLayout(tab_file)
        vbox_file.addWidget(QLabel("<b>Select Local Session Plan File:</b>"))
        hbox_file = QHBoxLayout()
        self.file_path_input = QLineEdit()
        self.file_path_input.setReadOnly(True)
        hbox_file.addWidget(self.file_path_input)
        browse_btn = QPushButton("📂 Browse...")
        browse_btn.clicked.connect(self._on_browse_file)
        hbox_file.addWidget(browse_btn)
        vbox_file.addLayout(hbox_file)

        file_hint = QLabel("Supports Word Documents (<b>.docx</b>, <b>.doc</b>), <b>.zip</b> archives, or HTML tables.")
        file_hint.setStyleSheet("font-size: 11px; color: #64748B;")
        vbox_file.addWidget(file_hint)
        vbox_file.addStretch()

        btn_file = QPushButton("📥 Load File")
        btn_file.setStyleSheet("background: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
        btn_file.clicked.connect(self._on_import_file)
        vbox_file.addWidget(btn_file)
        tabs.addTab(tab_file, "📂 Browse File")

        # Tab 3: Paste Table
        tab_paste = QWidget()
        vbox_paste = QVBoxLayout(tab_paste)
        vbox_paste.addWidget(QLabel("<b>Paste Table or Markdown:</b>"))
        self.paste_edit = QTextEdit()
        self.paste_edit.setPlaceholderText("Paste markdown or tab-separated text table here (e.g. from ETSI EMEET email or TOHRU)...")
        vbox_paste.addWidget(self.paste_edit)

        btn_paste = QPushButton("📋 Parse Pasted Table")
        btn_paste.setStyleSheet("background: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
        btn_paste.clicked.connect(self._on_import_paste)
        vbox_paste.addWidget(btn_paste)
        tabs.addTab(tab_paste, "📋 Paste Table")

        layout.addWidget(tabs)

    def _on_import_url(self):
        val = self.url_input.text().strip()
        if not val:
            QMessageBox.warning(self, "Input Required", "Please enter a valid URL or TDoc number.")
            return
        self.import_requested.emit("url", val)
        self.accept()

    def _on_browse_file(self):
        f_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Session Plan Document",
            "",
            "Session Plan Files (*.docx *.doc *.zip *.htm *.html);;All Files (*.*)",
        )
        if f_path:
            self.file_path_input.setText(f_path)

    def _on_import_file(self):
        val = self.file_path_input.text().strip()
        if not val or not Path(val).exists():
            QMessageBox.warning(self, "File Not Found", "Please browse and select an existing file.")
            return
        self.import_requested.emit("file", val)
        self.accept()

    def _on_import_paste(self):
        val = self.paste_edit.toPlainText().strip()
        if not val:
            QMessageBox.warning(self, "Text Required", "Please paste session plan table text.")
            return
        self.import_requested.emit("paste", val)
        self.accept()


class SessionPlanDiffDialog(QDialog):
    """Compares two snapshots and displays visual changes and changelog bullets."""

    def __init__(self, diff_result: SessionPlanDiffResult, parent=None):
        super().__init__(parent)
        self.diff = diff_result
        self.setWindowTitle(f"⇄ Compare Session Plans: {diff_result.old_snapshot_label} ➔ {diff_result.new_snapshot_label}")
        self.resize(750, 500)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # KPI Banner
        kpi_frame = QFrame()
        kpi_frame.setStyleSheet("background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 10px;")
        kpi_layout = QHBoxLayout(kpi_frame)

        moved_count = len(self.diff.moved_items)
        vol_count = len(self.diff.volume_changes)
        added_count = len(self.diff.added_slots)
        rem_count = len(self.diff.removed_slots)

        kpi_layout.addWidget(QLabel(f"🔄 <b>{moved_count}</b> Rescheduled"))
        kpi_layout.addWidget(QLabel(f"🟢 <b>{added_count}</b> Added"))
        kpi_layout.addWidget(QLabel(f"🔴 <b>{rem_count}</b> Removed"))
        kpi_layout.addWidget(QLabel(f"🔵 <b>{vol_count}</b> Volume Changes"))
        kpi_layout.addStretch()
        layout.addWidget(kpi_frame)

        # Changelog List
        layout.addWidget(QLabel("<b>Detailed Changelog:</b>"))
        self.change_list = QListWidget()
        if not self.diff.changelog_bullets:
            self.change_list.addItem("✨ No structural changes detected between these two snapshots.")
        else:
            for b in self.diff.changelog_bullets:
                item = QListWidgetItem(b)
                if b.startswith("+"):
                    item.setForeground(Qt.darkGreen)
                elif b.startswith("-"):
                    item.setForeground(Qt.darkRed)
                elif b.startswith("↔"):
                    item.setForeground(Qt.darkYellow)
                self.change_list.addItem(item)
        layout.addWidget(self.change_list)

        # Close button
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_box.addWidget(close_btn)
        layout.addLayout(btn_box)


class SessionPlanNoteDialog(QDialog):
    """Allows marking priority, assigned colleague, and adding comments to a slot."""

    def __init__(
        self,
        slot_desc: str,
        current_annotation: Optional[UserSessionAnnotation] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle(f"📝 Session Note — {slot_desc}")
        self.resize(440, 300)
        self._setup_ui(current_annotation)

    def _setup_ui(self, ann: Optional[UserSessionAnnotation]):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        form = QFormLayout()

        # Priority Badge
        self.badge_combo = QComboBox()
        self.badge_combo.addItem("Normal / Unmarked", "normal")
        self.badge_combo.addItem("⭐ Must Attend (Priority)", "must_attend")
        self.badge_combo.addItem("👥 Colleague Covering", "colleague")
        self.badge_combo.addItem("👀 Monitor on TOHRU", "tohru")
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
        save_btn.setStyleSheet("background: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
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
        hdr_frame.setStyleSheet("background: #F5F3FF; border: 1px solid #DDD6FE; border-radius: 8px; padding: 10px;")
        hdr_layout = QVBoxLayout(hdr_frame)
        hdr_layout.setSpacing(4)
        t_lbl = QLabel("<b>Tracked Topics & Followed Agenda Items</b>")
        t_lbl.setStyleSheet("color: #5B21B6; font-size: 13px;")
        hdr_layout.addWidget(t_lbl)
        desc_lbl = QLabel("Tracked topics are highlighted in purple across the Overview Grid and populate your personal <b>'My Topics'</b> schedule and parallel conflict detector.")
        desc_lbl.setStyleSheet("color: #6D28D9; font-size: 11px;")
        desc_lbl.setWordWrap(True)
        hdr_layout.addWidget(desc_lbl)
        layout.addWidget(hdr_frame)

        # Filter bar & quick action buttons
        filter_layout = QHBoxLayout()
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("🔍 Filter by AI #, acronym, or topic title...")
        self.search_box.setClearButtonEnabled(True)
        self.search_box.textChanged.connect(self._filter_list)
        filter_layout.addWidget(self.search_box)

        btn_auto = QPushButton("⭐ Track Followed WIs")
        btn_auto.setToolTip("Automatically check all Agenda Items matching your tracked Work Items watchlist.")
        btn_auto.setStyleSheet("background: #EDE9FE; color: #6D28D9; font-weight: 600; padding: 4px 10px; border-radius: 6px; border: 1px solid #DDD6FE;")
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
        self.list_widget.setStyleSheet("QListWidget::item { padding: 4px; border-bottom: 1px solid #F1F5F9; }")

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
        save_btn.setStyleSheet("background: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
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


# --- Delegation Management Dialogs ---

class DelegationMemberDialog(QDialog):
    """Dialog to create or edit a delegation member with color swatch picker."""

    def __init__(
        self,
        parent=None,
        member: Optional[DelegationMember] = None,
        existing_names: Optional[List[str]] = None,
    ):
        super().__init__(parent)
        self.member = member
        self.existing_names = [n.lower() for n in (existing_names or []) if not (member and n.lower() == member.name.lower())]
        self.selected_color = member.color if member else DELEGATE_PALETTE[0]

        self.setWindowTitle("👤 Edit Team Member" if member else "➕ Add Team Member")
        self.resize(440, 360)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        form = QFormLayout()
        form.setSpacing(10)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("e.g. Martin, Anna, Thomas")
        if self.member:
            self.name_edit.setText(self.member.name)
        self.name_edit.textChanged.connect(self._update_preview)
        form.addRow("<b>Delegate Name:</b>", self.name_edit)

        self.email_edit = QLineEdit()
        self.email_edit.setPlaceholderText("e.g. name@company.com (optional)")
        if self.member:
            self.email_edit.setText(self.member.email)
        form.addRow("Email Address:", self.email_edit)

        self.role_edit = QLineEdit()
        self.role_edit.setPlaceholderText("e.g. Head of Delegation, RAN Specialist, Architecture")
        if self.member:
            self.role_edit.setText(self.member.notes)
        form.addRow("Role / Notes:", self.role_edit)

        self.self_chk = QCheckBox("This is me (My Identity)")
        self.self_chk.setToolTip("Marking this sets this delegate as 'yourself' for 1-click personal timetable filtering.")
        if self.member and self.member.is_self:
            self.self_chk.setChecked(True)
        self.self_chk.toggled.connect(self._update_preview)
        form.addRow("", self.self_chk)

        layout.addLayout(form)

        # Color Swatches Palette
        layout.addWidget(QLabel("<b>Delegate Color Palette:</b>"))
        swatch_layout = QHBoxLayout()
        swatch_layout.setSpacing(6)
        self.swatch_buttons: List[QPushButton] = []

        for col in DELEGATE_PALETTE:
            btn = QPushButton()
            btn.setFixedSize(26, 26)
            btn.setCursor(Qt.PointingHandCursor)
            is_active = (col.lower() == self.selected_color.lower())
            border_style = "2px solid #0F172A" if is_active else "1px solid #CBD5E1"
            btn.setStyleSheet(f"background-color: {col}; border-radius: 13px; border: {border_style};")
            btn.clicked.connect(lambda _, c=col: self._select_palette_color(c))
            swatch_layout.addWidget(btn)
            self.swatch_buttons.append(btn)

        # Custom color button
        self.btn_custom_col = QPushButton("🎨 Custom...")
        self.btn_custom_col.setFixedHeight(26)
        self.btn_custom_col.clicked.connect(self._pick_custom_color)
        swatch_layout.addWidget(self.btn_custom_col)
        swatch_layout.addStretch()
        layout.addLayout(swatch_layout)

        # Live Badge Preview
        preview_box = QHBoxLayout()
        preview_box.addWidget(QLabel("Live Badge Preview:"))
        self.preview_lbl = QLabel()
        self.preview_lbl.setFixedHeight(28)
        preview_box.addWidget(self.preview_lbl)
        preview_box.addStretch()
        layout.addLayout(preview_box)
        self._update_preview()

        # Dialog Buttons
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(cancel_btn)

        save_btn = QPushButton("💾 Save Member")
        save_btn.setStyleSheet("background: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
        save_btn.clicked.connect(self._validate_and_accept)
        btn_box.addWidget(save_btn)
        layout.addLayout(btn_box)

    def _select_palette_color(self, hex_color: str):
        self.selected_color = hex_color
        import re
        for btn in self.swatch_buttons:
            m = re.search(r"background-color:\s*(#[0-9a-fA-F]+)", btn.styleSheet())
            b_col = m.group(1) if m else "#2563EB"
            is_active = (b_col.lower() == hex_color.lower())
            border_style = "3px solid #0F172A" if is_active else "1px solid #CBD5E1"
            btn.setStyleSheet(f"background-color: {b_col}; border-radius: 13px; border: {border_style};")
        self._update_preview()

    def _pick_custom_color(self):
        col = QColorDialog.getColor(QColor(self.selected_color), self, "Select Delegate Color")
        if col.isValid():
            self._select_palette_color(col.name())

    def _update_preview(self):
        name = self.name_edit.text().strip() or "Delegate"
        is_me = self.self_chk.isChecked()
        suffix = " (Me)" if is_me else ""
        col = self.selected_color
        bg_t = DelegationMember.compute_tint(col, 0.12)
        bdr_t = DelegationMember.compute_tint(col, 0.35)
        self.preview_lbl.setText(f"  ● {name}{suffix}  ")
        self.preview_lbl.setStyleSheet(
            f"background-color: {bg_t}; color: {col}; border: 1px solid {bdr_t}; border-radius: 6px; font-weight: 600; padding: 2px 8px;"
        )

    def _validate_and_accept(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Invalid Name", "Please enter a valid delegate name.")
            return
        if name.lower() in self.existing_names:
            QMessageBox.warning(self, "Duplicate Name", f"A delegate named '{name}' already exists in this meeting roster.")
            return
        self.accept()

    def get_member(self) -> DelegationMember:
        return DelegationMember(
            name=self.name_edit.text().strip(),
            email=self.email_edit.text().strip(),
            color=self.selected_color,
            is_self=self.self_chk.isChecked(),
            notes=self.role_edit.text().strip(),
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
        card.setStyleSheet("background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 10px;")
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

            name_item = QTableWidgetItem(m.name + (" (Me)" if m.is_self else ""))
            name_item.setForeground(QColor(m.color))
            f = name_item.font()
            f.setBold(True)
            name_item.setFont(f)

            reason_txt = "Available (no parallel session)" if is_free else av.get("reason", "Already scheduled")
            reason_item = QTableWidgetItem(reason_txt)

            if not is_free:
                status_item.setForeground(QColor("#DC2626"))
                reason_item.setForeground(QColor("#991B1B"))
                reason_item.setBackground(QColor("#FEF2F2"))
            else:
                status_item.setForeground(QColor("#16A34A"))

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
        all_team_btn.setStyleSheet("background: #4F46E5; color: #FFFFFF; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
        all_team_btn.setToolTip("Appoints all delegates in the roster to attend this session together.")
        all_team_btn.clicked.connect(self._appoint_all_team)
        btn_box.addWidget(all_team_btn)

        btn_box.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(cancel_btn)

        appoint_btn = QPushButton("⚡ Appoint Selected")
        appoint_btn.setStyleSheet("background: #059669; color: #FFFFFF; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
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
        header_card.setStyleSheet("background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 8px;")
        h_lay = QVBoxLayout(header_card)
        h_lay.setContentsMargins(6, 6, 6, 6)
        h_lay.addWidget(QLabel("<b>Assign Delegate(s) for Topic:</b>"))
        lbl_top = QLabel(self.topic_title)
        lbl_top.setWordWrap(True)
        lbl_top.setStyleSheet("color: #334155; font-size: 11px;")
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
        self.chk_all.setStyleSheet("font-weight: 700; font-size: 12px; color: #1E40AF; padding: 4px;")
        self.chk_all.setChecked(is_all_initial)
        self.chk_all.toggled.connect(self._on_all_toggled)
        layout.addWidget(self.chk_all)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #E2E8F0;")
        layout.addWidget(sep)

        # Scrollable individual member checkboxes
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px;")
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

            chk = QCheckBox(m.name + (" (Myself)" if m.is_self else ""))
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
        btn_ok.setStyleSheet("background: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
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


class DelegationImportDialog(QDialog):
    """Dialog to import a delegation plan from pasted email text or JSON file."""

    def __init__(self, current_roster: Optional[List[str]] = None, parent=None):
        super().__init__(parent)
        self.current_roster = current_roster or []
        self.parsed_data: Optional[dict] = None

        self.setWindowTitle("📥 Import Delegation Plan (Email / JSON)")
        self.resize(560, 460)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        header = QLabel(
            "<b>Import Delegation Allocation:</b> Paste the email text containing the delegation plan, "
            "or load an exported <code>.json</code> plan file."
        )
        header.setWordWrap(True)
        layout.addWidget(header)

        # File browse row
        file_box = QHBoxLayout()
        self.btn_file = QPushButton("📁 Browse .json File...")
        self.btn_file.clicked.connect(self._browse_file)
        file_box.addWidget(self.btn_file)
        self.file_lbl = QLabel("No file loaded")
        self.file_lbl.setStyleSheet("color: #64748B; font-style: italic;")
        file_box.addWidget(self.file_lbl)
        file_box.addStretch()
        layout.addLayout(file_box)

        # Text edit for paste
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Paste email body or JSON text here...\n(The tool automatically extracts the embedded plan payload)")
        self.text_edit.textChanged.connect(self._on_text_changed)
        layout.addWidget(self.text_edit)

        # Status / Summary Banner
        self.status_banner = QFrame()
        self.status_banner.setStyleSheet("background: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 6px; padding: 8px;")
        sb_lay = QVBoxLayout(self.status_banner)
        sb_lay.setContentsMargins(6, 4, 6, 4)
        self.status_lbl = QLabel("⏳ Waiting for plan input...")
        self.status_lbl.setStyleSheet("color: #475569; font-weight: 500;")
        sb_lay.addWidget(self.status_lbl)
        layout.addWidget(self.status_banner)

        # Personal Identity Selection
        self.identity_box = QFrame()
        id_lay = QHBoxLayout(self.identity_box)
        id_lay.setContentsMargins(0, 0, 0, 0)
        id_lay.addWidget(QLabel("<b>👤 Which delegate are you?</b>"))
        self.identity_combo = QComboBox()
        self.identity_combo.setToolTip("Select your own name from the imported roster to personalize your local view.")
        id_lay.addWidget(self.identity_combo, 1)
        layout.addWidget(self.identity_box)
        self.identity_box.setEnabled(False)

        # Merge Options
        merge_box = QHBoxLayout()
        self.rb_replace = QRadioButton("Replace current team plan")
        self.rb_replace.setChecked(True)
        self.rb_merge = QRadioButton("Merge into current team plan")
        merge_box.addWidget(self.rb_replace)
        merge_box.addWidget(self.rb_merge)
        merge_box.addStretch()
        layout.addLayout(merge_box)

        # Buttons
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(cancel_btn)

        self.import_btn = QPushButton("📥 Import Plan")
        self.import_btn.setStyleSheet("background: #2563EB; color: #FFFFFF; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
        self.import_btn.setEnabled(False)
        self.import_btn.clicked.connect(self.accept)
        btn_box.addWidget(self.import_btn)
        layout.addLayout(btn_box)

    def _browse_file(self):
        fpath, _ = QFileDialog.getOpenFileName(self, "Open Delegation Plan JSON", "", "JSON Files (*.json);;All Files (*.*)")
        if fpath:
            try:
                content = Path(fpath).read_text(encoding="utf-8")
                self.file_lbl.setText(Path(fpath).name)
                self.text_edit.setPlainText(content)
            except Exception as e:
                QMessageBox.critical(self, "File Error", f"Failed to read file: {e}")

    def _on_text_changed(self):
        raw = self.text_edit.toPlainText().strip()
        if not raw:
            self.status_lbl.setText("⏳ Waiting for plan input...")
            self.status_banner.setStyleSheet("background: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 6px; padding: 8px;")
            self.import_btn.setEnabled(False)
            self.identity_box.setEnabled(False)
            self.parsed_data = None
            return

        parsed = SessionPlanExporter.parse_delegation_plan_json(raw)
        if parsed:
            self.parsed_data = parsed
            members = parsed.get("members", [])
            topics = parsed.get("topic_owners", {})
            appts = parsed.get("appointments", {})
            mtg = parsed.get("meeting_number", "")

            self.status_lbl.setText(
                f"✅ Valid Plan: SA2 #{mtg} · <b>{len(members)}</b> Delegates · <b>{len(topics)}</b> Topic Owners · <b>{len(appts)}</b> Ad-Hoc Coverages"
            )
            self.status_banner.setStyleSheet("background: #ECFDF5; border: 1px solid #A7F3D0; border-radius: 6px; padding: 8px; color: #065F46;")
            self.import_btn.setEnabled(True)
            self.identity_box.setEnabled(True)

            self.identity_combo.clear()
            self.identity_combo.addItem("— (Don't claim identity) —", "")
            for m in members:
                m_name = m.get("name", "") if isinstance(m, dict) else str(m)
                if m_name:
                    self.identity_combo.addItem(m_name, m_name)
        else:
            self.parsed_data = None
            self.status_lbl.setText("❌ No valid delegation plan detected in text. Ensure it contains 3GPP-DELEGATION-PLAN payload.")
            self.status_banner.setStyleSheet("background: #FEF2F2; border: 1px solid #FECACA; border-radius: 6px; padding: 8px; color: #991B1B;")
            self.import_btn.setEnabled(False)
            self.identity_box.setEnabled(False)

    def get_data(self) -> Tuple[dict, bool, str]:
        """Returns: (parsed_data_dict, is_merge_bool, claimed_self_delegate_name)."""
        claimed_self = self.identity_combo.currentData() or ""
        return self.parsed_data, self.rb_merge.isChecked(), claimed_self

