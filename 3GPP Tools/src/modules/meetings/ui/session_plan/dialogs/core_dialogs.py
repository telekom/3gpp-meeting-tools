# --- File: src/modules/meetings/ui/session_plan/dialogs/core_dialogs.py ---
"""Timezone selector, manual Session Plan importer, and snapshot diff dialogs."""

from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
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
    QSpinBox,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from modules.meetings.core.session_plan.session_plan_diff import SessionPlanDiffResult
from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone


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
        info_frame.setStyleSheet(f"background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px; padding: 10px;")
        info_layout = QVBoxLayout(info_frame)
        info_layout.setSpacing(4)

        loc_lbl = QLabel(f"📍 <b>Meeting Location:</b> {self.location or 'Unknown'}")
        info_layout.addWidget(loc_lbl)

        lap_sign = "+" if self.laptop_offset_min >= 0 else "-"
        lap_h = abs(self.laptop_offset_min) // 60
        lap_m = abs(self.laptop_offset_min) % 60
        lap_str = f"UTC{lap_sign}{lap_h:02d}:{lap_m:02d}"
        laptop_lbl = QLabel(f"💻 <b>Your Corporate Laptop Clock:</b> {lap_str} (Austrian Time)")
        laptop_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY};")
        info_layout.addWidget(laptop_lbl)

        expl = QLabel("3GPP Session Plans are published in <b>Meeting Local Time</b>. Selecting the correct venue timezone ensures live session indicators, reminders, and calendar exports trigger accurately.")
        expl.setWordWrap(True)
        expl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; margin-top: 4px;")
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
        save_btn.setStyleSheet(f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
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
            self.diff_preview_lbl.setStyleSheet(f"background: {C.SUCCESS_SOFT}; color: {C.SUCCESS_DARK};")
        else:
            diff_h = abs(diff) // 60
            ahead_behind = "ahead of" if diff < 0 else "behind"
            self.diff_preview_lbl.setText(f"🕒 Venue is {diff_h} hour(s) {ahead_behind} your Austrian laptop clock.")
            self.diff_preview_lbl.setStyleSheet(f"background: {C.WARNING_SOFT}; color: {C.WARNING};")

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
        hint.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; margin-top: 4px;")
        vbox_url.addWidget(hint)
        vbox_url.addStretch()

        btn_url = QPushButton("⬇️ Download && Import")
        btn_url.setStyleSheet(f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
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
        file_hint.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED};")
        vbox_file.addWidget(file_hint)
        vbox_file.addStretch()

        btn_file = QPushButton("📥 Load File")
        btn_file.setStyleSheet(f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
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
        btn_paste.setStyleSheet(f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 12px; border-radius: 6px;")
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
        kpi_frame.setStyleSheet(f"background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px; padding: 10px;")
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

