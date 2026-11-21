# --- File: src/modules/meetings/ui/session_plan/dialogs/delegation_roster_dialogs.py ---
"""Team delegate member editor (color swatch picker) and Delegation Plan importer dialogs."""

from pathlib import Path
from typing import List, Optional, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QColorDialog,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTextEdit,
    QVBoxLayout,
)

from core.ui.theme import C
from modules.meetings.core.session_plan.session_plan_exporter import SessionPlanExporter
from modules.meetings.core.session_plan.session_plan_models import (
    DELEGATE_PALETTE,
    DelegationMember,
)


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
            border_style = f"2px solid {C.TEXT_PRIMARY}" if is_active else f"1px solid {C.BORDER_STRONG}"
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
        save_btn.setStyleSheet(f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
        save_btn.clicked.connect(self._validate_and_accept)
        btn_box.addWidget(save_btn)
        layout.addLayout(btn_box)

    def _select_palette_color(self, hex_color: str):
        self.selected_color = hex_color
        import re
        for btn in self.swatch_buttons:
            m = re.search(r"background-color:\s*(#[0-9a-fA-F]+)", btn.styleSheet())
            b_col = m.group(1) if m else C.ACCENT_ALT
            is_active = (b_col.lower() == hex_color.lower())
            border_style = f"3px solid {C.TEXT_PRIMARY}" if is_active else f"1px solid {C.BORDER_STRONG}"
            btn.setStyleSheet(f"background-color: {b_col}; border-radius: 13px; border: {border_style};")
        self._update_preview()

    def _pick_custom_color(self):
        col = QColorDialog.getColor(QColor(self.selected_color), self, "Select Delegate Color")
        if col.isValid():
            self._select_palette_color(col.name())

    def _update_preview(self):
        name = self.name_edit.text().strip() or "Delegate"
        col = self.selected_color
        bg_t = DelegationMember.compute_tint(col, 0.12)
        bdr_t = DelegationMember.compute_tint(col, 0.35)
        self.preview_lbl.setText(f"  ● {name}  ")
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
            is_self=self.member.is_self if self.member else False,
            notes=self.role_edit.text().strip(),
        )



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
        self.file_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-style: italic;")
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
        self.status_banner.setStyleSheet(f"background: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; padding: 8px;")
        sb_lay = QVBoxLayout(self.status_banner)
        sb_lay.setContentsMargins(6, 4, 6, 4)
        self.status_lbl = QLabel("⏳ Waiting for plan input...")
        self.status_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-weight: 500;")
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
        self.import_btn.setStyleSheet(f"background: {C.ACCENT_ALT}; color: {C.BG_CARD}; font-weight: 600; padding: 6px 14px; border-radius: 6px;")
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
            self.status_banner.setStyleSheet(f"background: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; padding: 8px;")
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
            self.status_banner.setStyleSheet(f"background: {C.WG_RAN_BG}; border: 1px solid {C.WG_RAN_BORDER}; border-radius: 6px; padding: 8px; color: {C.WG_RAN_TEXT};")
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
            self.status_banner.setStyleSheet(f"background: {C.DANGER_SOFT}; border: 1px solid {C.DANGER_BORDER}; border-radius: 6px; padding: 8px; color: {C.DANGER_DARK};")
            self.import_btn.setEnabled(False)
            self.identity_box.setEnabled(False)

    def get_data(self) -> Tuple[dict, bool, str]:
        """Returns: (parsed_data_dict, is_merge_bool, claimed_self_delegate_name)."""
        claimed_self = self.identity_combo.currentData() or ""
        return self.parsed_data, self.rb_merge.isChecked(), claimed_self

