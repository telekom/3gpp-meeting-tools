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
from typing import Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from modules.meetings.core.session_plan_diff import SessionPlanDiffEngine, SessionPlanDiffResult
from modules.meetings.core.session_plan_models import (
    SessionPlanSnapshot,
    SessionSlot,
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
