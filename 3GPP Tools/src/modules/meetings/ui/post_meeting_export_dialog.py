# --- File: src/modules/meetings/ui/post_meeting_export_dialog.py ---
import datetime
from pathlib import Path
from typing import Dict, List, Set, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QGroupBox,
    QRadioButton, QCheckBox, QPushButton, QLineEdit, QFileDialog,
    QListWidget, QListWidgetItem, QFrame, QScrollArea, QWidget,
    QButtonGroup, QMessageBox
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.ui.tdocs_models import TDocsTableModel


class PostMeetingExportDialog(QDialog):
    """
    Configuration dialog to customize and trigger post-meeting Markdown dossiers.
    Supports scoping by Followed/Starred Work Items, filtered view, or whole meeting,
    with per-release and per-WI directory organization.
    """

    def __init__(
        self,
        parent,
        meeting_dir: Path,
        all_rows: list,
        visible_rows: list,
        agenda_map: dict,
        mtg_info: dict,
        followed_tokens: set,
        starred_tokens: set,
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.all_rows = all_rows or []
        self.visible_rows = visible_rows or []
        self.agenda_map = agenda_map or {}
        self.mtg_info = mtg_info or {}
        self.followed_tokens = set(followed_tokens or [])
        self.starred_tokens = set(starred_tokens or [])
        self.tracked_tokens = self.followed_tokens | self.starred_tokens

        self.setWindowTitle("📑 Post-Meeting Report & Markdown Dossier Export")
        self.resize(580, 720)
        self.setStyleSheet("""
            QDialog { background-color: #F8FAFC; }
            QGroupBox {
                font-weight: bold;
                color: #1E293B;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 14px;
                background-color: #FFFFFF;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 6px;
                background-color: #FFFFFF;
            }
            QLabel { color: #334155; }
            QRadioButton, QCheckBox {
                color: #1E293B;
                padding: 2px;
            }
        """)

        # Extract available meeting Work Items for dynamic counts and custom selection
        self.meeting_wis = self._extract_meeting_work_items()
        self.followed_wis = [w for w in self.meeting_wis if (TDocsTableModel._extract_base_tokens(w) & self.followed_tokens)]
        self.starred_wis = [w for w in self.meeting_wis if (TDocsTableModel._extract_base_tokens(w) & self.starred_tokens)]
        self.tracked_wis = [w for w in self.meeting_wis if (TDocsTableModel._extract_base_tokens(w) & self.tracked_tokens)]

        self._setup_ui()

    def _extract_meeting_work_items(self) -> List[str]:
        """Collects all unique canonical Work Items present across the meeting's contributions."""
        wis = set()
        for r in self.all_rows:
            raw_wis = TDocsTableModel._split_related_wis(r.get("Related WIs", ""))
            for rw in raw_wis:
                clean = rw.strip()
                if clean and clean.lower() not in {"(no wi)", "none", "nan", "-"}:
                    wis.add(clean)

            ai_num = str(r.get("Agenda Item", "")).strip()
            if ai_num and self.agenda_map:
                item = self.agenda_map.get(ai_num)
                if item and getattr(item, "acronym", None):
                    wis.add(item.acronym.strip())

            desc = str(r.get("Agenda item description", "")).strip()
            if desc:
                import re
                m = re.search(r'\(([A-Za-z0-9_\-\/]+)\)', desc)
                if m:
                    wis.add(m.group(1).strip())

        return sorted(list(wis), key=lambda x: x.lower())

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(14, 14, 14, 14)

        # --- Header Card ---
        header_frame = QFrame()
        header_frame.setStyleSheet("""
            QFrame {
                background-color: #EFF6FF;
                border: 1px solid #BFDBFE;
                border-radius: 6px;
                padding: 8px;
            }
        """)
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(6, 4, 6, 4)
        header_layout.setSpacing(2)

        wg = str(self.mtg_info.get("wg_name", "3GPP")).strip()
        mtg = str(self.mtg_info.get("meeting_number", "")).strip()
        title_text = f"<b>{wg} #{mtg} Post-Meeting Report</b>" if (wg or mtg) else "<b>Post-Meeting Report</b>"
        lbl_title = QLabel(title_text)
        lbl_title.setStyleSheet("font-size: 14px; color: #1E40AF; font-weight: bold; border: none;")
        lbl_sub = QLabel("Generates a structured, per-release and per-WI Markdown dossier with an executive summary cover page.")
        lbl_sub.setStyleSheet("font-size: 11px; color: #3B82F6; border: none;")
        lbl_sub.setWordWrap(True)

        header_layout.addWidget(lbl_title)
        header_layout.addWidget(lbl_sub)
        main_layout.addWidget(header_frame)

        # --- Group 1: Scope Selection ---
        scope_group = QGroupBox("1. Select Export Scope")
        scope_layout = QVBoxLayout(scope_group)
        scope_layout.setSpacing(6)

        self.scope_btn_group = QButtonGroup(self)

        n_followed = len(self.followed_wis)
        n_starred = len(self.starred_wis)
        n_tracked = len(self.tracked_wis)
        n_visible = len(self.visible_rows)
        n_total = len(self.all_rows)
        n_all_wis = len(self.meeting_wis)

        self.rb_followed = QRadioButton(f"🎯 Followed Work Items only ({n_followed} WIs in meeting)")
        self.rb_whole = QRadioButton(f"🌐 Whole Meeting ({n_all_wis} WIs, {n_total} TDocs total)")
        self.rb_starred = QRadioButton(f"⭐ Starred Work Items only ({n_starred} WIs in meeting)")
        self.rb_tracked = QRadioButton(f"🎯 Followed & ⭐ Starred Work Items ({n_tracked} WIs)")
        self.rb_visible = QRadioButton(f"🔍 Current Filtered Table View ({n_visible} TDocs visible)")
        self.rb_custom = QRadioButton("📋 Custom Work Item Selection...")

        self.scope_btn_group.addButton(self.rb_followed)
        self.scope_btn_group.addButton(self.rb_whole)
        self.scope_btn_group.addButton(self.rb_starred)
        self.scope_btn_group.addButton(self.rb_tracked)
        self.scope_btn_group.addButton(self.rb_visible)
        self.scope_btn_group.addButton(self.rb_custom)

        # Default selection: Followed if available, else Whole Meeting
        if n_followed > 0:
            self.rb_followed.setChecked(True)
        else:
            self.rb_whole.setChecked(True)

        scope_layout.addWidget(self.rb_followed)
        scope_layout.addWidget(self.rb_whole)
        scope_layout.addWidget(self.rb_starred)
        scope_layout.addWidget(self.rb_tracked)
        scope_layout.addWidget(self.rb_visible)
        scope_layout.addWidget(self.rb_custom)

        # Custom WI Selection Widget (Collapsible)
        self.custom_wi_widget = QWidget()
        custom_layout = QVBoxLayout(self.custom_wi_widget)
        custom_layout.setContentsMargins(18, 4, 4, 4)
        custom_layout.setSpacing(4)

        self.wi_search = QLineEdit()
        self.wi_search.setPlaceholderText("Filter Work Items list...")
        self.wi_search.textChanged.connect(self._filter_custom_wi_list)
        custom_layout.addWidget(self.wi_search)

        self.wi_list = QListWidget()
        self.wi_list.setMaximumHeight(110)
        for wi in self.meeting_wis:
            item = QListWidgetItem(wi)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.wi_list.addItem(item)
        custom_layout.addWidget(self.wi_list)

        self.custom_wi_widget.setVisible(False)
        self.rb_custom.toggled.connect(self.custom_wi_widget.setVisible)
        scope_layout.addWidget(self.custom_wi_widget)

        main_layout.addWidget(scope_group)

        # --- Group 2: Document Types & Statuses ---
        docs_group = QGroupBox("2. Target Document Types & Statuses")
        docs_layout = QVBoxLayout(docs_group)
        docs_layout.setSpacing(4)

        self.chk_ls_in = QCheckBox("📥 Incoming Liaison Statements (LS in) — All received")
        self.chk_ls_in.setChecked(True)
        self.chk_ls_in.setToolTip("Include all received/incoming Liaison Statements.")

        self.chk_pcr = QCheckBox("📝 Agreed / Approved Pseudo-CRs (pCR)")
        self.chk_pcr.setChecked(True)
        self.chk_pcr.setToolTip("Include study-item pseudo-CRs agreed or approved by the meeting.")

        self.chk_cr = QCheckBox("📄 Agreed / Approved Change Requests (CR)")
        self.chk_cr.setChecked(True)
        self.chk_cr.setToolTip("Include normative Change Requests agreed or approved by the meeting.")

        self.chk_ls_out = QCheckBox("📤 Agreed / Approved Outgoing Liaison Statements (LS out)")
        self.chk_ls_out.setChecked(True)
        self.chk_ls_out.setToolTip("Include outgoing Liaison Statements approved by the meeting.")

        self.chk_general_ls = QCheckBox("📦 Include General / Overarching LS in (without specific WI)")
        self.chk_general_ls.setChecked(True)
        self.chk_general_ls.setToolTip("Places general incoming LS (e.g. AI 4.1) under a dedicated 'General/LS_IN/' folder.")

        docs_layout.addWidget(self.chk_ls_in)
        docs_layout.addWidget(self.chk_pcr)
        docs_layout.addWidget(self.chk_cr)
        docs_layout.addWidget(self.chk_ls_out)
        docs_layout.addWidget(self.chk_general_ls)

        main_layout.addWidget(docs_group)

        # --- Group 3: Conversion & Output Options ---
        opts_group = QGroupBox("3. Organization & Conversion Options")
        opts_layout = QVBoxLayout(opts_group)
        opts_layout.setSpacing(4)

        self.chk_local_indices = QCheckBox("Generate local index.md cover inside each Work Item directory")
        self.chk_local_indices.setChecked(True)

        self.chk_download = QCheckBox("Download missing TDoc ZIP packages automatically from FTP")
        self.chk_download.setChecked(True)

        self.chk_force = QCheckBox("Force reconvert (overwrite existing local markdown caches)")
        self.chk_force.setChecked(False)

        self.chk_auto_open = QCheckBox("Open export directory in Windows Explorer upon completion")
        self.chk_auto_open.setChecked(True)

        opts_layout.addWidget(self.chk_local_indices)
        opts_layout.addWidget(self.chk_download)
        opts_layout.addWidget(self.chk_force)
        opts_layout.addWidget(self.chk_auto_open)

        # Destination Folder row
        dest_layout = QHBoxLayout()
        dest_layout.addWidget(QLabel("Destination:"))
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        folder_name = f"PostMeeting_{wg}_{mtg}_{timestamp}" if (wg or mtg) else f"PostMeeting_{timestamp}"
        default_dir = self.meeting_dir / "Export" / folder_name
        self.dest_edit = QLineEdit(str(default_dir))
        self.dest_edit.setReadOnly(False)
        dest_layout.addWidget(self.dest_edit)

        btn_browse = QPushButton("📁 Browse...")
        btn_browse.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_browse.clicked.connect(self._browse_destination)
        dest_layout.addWidget(btn_browse)

        opts_layout.addSpacing(4)
        opts_layout.addLayout(dest_layout)

        main_layout.addWidget(opts_group)

        # --- Buttons ---
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_cancel = QPushButton("Cancel")
        btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)

        btn_start = QPushButton("🚀 Start Export")
        btn_start.setStyleSheet("""
            QPushButton {
                background-color: #2563EB;
                color: #FFFFFF;
                font-weight: bold;
                border-radius: 4px;
                padding: 6px 16px;
                border: 1px solid #1D4ED8;
            }
            QPushButton:hover {
                background-color: #1D4ED8;
            }
            QPushButton:pressed {
                background-color: #1E40AF;
            }
        """)
        btn_start.setDefault(True)
        btn_start.clicked.connect(self._on_start_clicked)
        btn_layout.addWidget(btn_start)

        main_layout.addLayout(btn_layout)

    def _filter_custom_wi_list(self, text: str):
        search = text.strip().lower()
        for i in range(self.wi_list.count()):
            item = self.wi_list.item(i)
            item.setHidden(bool(search and search not in item.text().lower()))

    def _browse_destination(self):
        d = QFileDialog.getExistingDirectory(self, "Select Export Directory", str(self.meeting_dir / "Export"))
        if d:
            self.dest_edit.setText(d)

    def _on_start_clicked(self):
        # Validate that at least one document type is selected
        if not any([
            self.chk_ls_in.isChecked(),
            self.chk_pcr.isChecked(),
            self.chk_cr.isChecked(),
            self.chk_ls_out.isChecked(),
        ]):
            QMessageBox.warning(self, "No Types Selected", "Please select at least one document category to export.")
            return

        if self.rb_custom.isChecked():
            checked_wis = self.get_custom_wis()
            if not checked_wis:
                QMessageBox.warning(self, "No Work Items Selected", "Please check at least one Work Item from the custom list.")
                return

        self.accept()

    def get_scope_mode(self) -> str:
        if self.rb_followed.isChecked():
            return "followed"
        if self.rb_whole.isChecked():
            return "whole"
        if self.rb_starred.isChecked():
            return "starred"
        if self.rb_tracked.isChecked():
            return "tracked"
        if self.rb_visible.isChecked():
            return "visible"
        if self.rb_custom.isChecked():
            return "custom"
        return "followed"

    def get_custom_wis(self) -> Set[str]:
        selected = set()
        for i in range(self.wi_list.count()):
            item = self.wi_list.item(i)
            if item.checkState() == Qt.Checked:
                selected.add(item.text().strip().lower())
        return selected

    def get_export_config(self) -> dict:
        dest_text = self.dest_edit.text().strip()
        output_dir = Path(dest_text) if dest_text else None

        return {
            "scope_mode": self.get_scope_mode(),
            "custom_wis": self.get_custom_wis(),
            "include_ls_in": self.chk_ls_in.isChecked(),
            "include_pcr": self.chk_pcr.isChecked(),
            "include_cr": self.chk_cr.isChecked(),
            "include_ls_out": self.chk_ls_out.isChecked(),
            "include_general_ls": self.chk_general_ls.isChecked(),
            "generate_local_indices": self.chk_local_indices.isChecked(),
            "download_missing": self.chk_download.isChecked(),
            "force_reconvert": self.chk_force.isChecked(),
            "auto_open": self.chk_auto_open.isChecked(),
            "output_dir": output_dir,
        }
