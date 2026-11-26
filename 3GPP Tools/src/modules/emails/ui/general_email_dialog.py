# --- File: src/modules/emails/ui/general_email_dialog.py ---

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QColorDialog,
    QDateEdit,
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import scale_px
from core.ui.theme import C
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    TABLE_STYLE_CLEAN,
)
from modules.emails.core.email_config import load_wg_email_config, save_wg_email_config
from modules.emails.ui.config_dialog import OutlookFolderPickerDialog

# Backward compatibility re-exports for dialogs extracted to dedicated module
from modules.emails.ui.tdoc_emails_dialog import StandaloneEmailReaderWindow, TDocEmailsDialog

_PRIMARY_BTN_STYLE = (
    f"QPushButton {{ background-color: {C.OFFICE_BLUE}; color: {C.TEXT_ON_ACCENT}; "
    f"font-weight: bold; padding: 6px 16px; border-radius: 5px; border: none; }} "
    f"QPushButton:hover {{ background-color: {C.INFO}; }}"
)

__all__ = [
    "GeneralEmailFoldersDialog",
    "GeneralEmailSyncDialog",
    "StandaloneEmailReaderWindow",
    "TDocEmailsDialog",
    "load_wg_email_config",
    "save_wg_email_config",
]


class GeneralEmailFoldersDialog(QDialog):
    """Modernized dialog for configuring WG Outlook folders, subfolder inclusion, tags, and badge colors."""

    def __init__(self, wg: str, parent=None):
        super().__init__(parent)
        self.wg = str(wg or "SA2").strip().upper()
        self.setWindowTitle(f"⚙️ Outlook Folders Configuration: {self.wg}")
        self.resize(scale_px(860), scale_px(490))
        self.setMinimumSize(scale_px(640), scale_px(380))
        self.setStyleSheet(f"QDialog {{ background-color: {C.BG_CANVAS}; }}")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        # Header banner
        hdr_frame = QFrame(self)
        hdr_frame.setStyleSheet(CARD_FRAME_STYLE)
        hdr_layout = QVBoxLayout(hdr_frame)
        hdr_layout.setContentsMargins(12, 10, 12, 10)
        hdr_layout.setSpacing(3)

        title_lbl = QLabel(f"📂 Configured Outlook Folders for {self.wg}")
        title_lbl.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        hdr_layout.addWidget(title_lbl)

        desc = QLabel(
            "Configure Outlook folders to scan for meeting emails and threads. "
            "<b>Subfolders are included by default</b> (more specific folder paths take tag priority and are never scanned twice)."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")
        hdr_layout.addWidget(desc)
        layout.addWidget(hdr_frame)

        # 4-Column Table: Outlook Folder Path | Subfolders | Tag | Tag Color
        self.table = QTableWidget(0, 4, self)
        self.table.setHorizontalHeaderLabels([
            "Outlook Folder Path",
            "Subfolders",
            "Tag",
            "Tag Color",
        ])
        self.table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setTextElideMode(Qt.ElideMiddle)
        self.table.verticalHeader().setDefaultSectionSize(scale_px(38))
        self.table.verticalHeader().setMinimumSectionSize(scale_px(34))

        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.Fixed)
        self.table.setColumnWidth(1, scale_px(118))
        hdr.setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.setColumnWidth(2, scale_px(130))
        hdr.setSectionResizeMode(3, QHeaderView.Fixed)
        self.table.setColumnWidth(3, scale_px(118))

        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)
        layout.addWidget(self.table, 1)

        # Toolbar row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_add = QPushButton("➕ Add Folder via Outlook...")
        self.btn_add.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_add.setToolTip("Browse your Microsoft Outlook mailbox tree and add a folder.")
        self.btn_add.clicked.connect(self._add_folder)

        self.btn_browse = QPushButton("📂 Change Selected Folder...")
        self.btn_browse.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_browse.setToolTip("Replace the folder path on the selected row via the Outlook folder picker.")
        self.btn_browse.clicked.connect(self._change_selected_folder)

        self.btn_del = QPushButton("🗑️ Remove Selected")
        self.btn_del.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_del.setToolTip("Remove the selected folder row(s) from this Working Group.")
        self.btn_del.clicked.connect(self._remove_selected)

        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_browse)
        btn_row.addWidget(self.btn_del)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # Footer row
        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(8)
        hint_lbl = QLabel("💡 Tip: Hover a folder path to inspect the full path, or double-click a path cell to browse.")
        hint_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        bottom_row.addWidget(hint_lbl)
        bottom_row.addStretch()

        btn_cancel = QPushButton("Cancel")
        btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_cancel.clicked.connect(self.reject)

        btn_save = QPushButton("💾 Save && Close")
        btn_save.setStyleSheet(_PRIMARY_BTN_STYLE)
        btn_save.clicked.connect(self._save_and_close)

        bottom_row.addWidget(btn_cancel)
        bottom_row.addWidget(btn_save)
        layout.addLayout(bottom_row)

        self._populate_table()

    def _populate_table(self):
        folders = load_wg_email_config(self.wg)
        self.table.setRowCount(0)
        for item in folders:
            if isinstance(item, dict):
                fpath = str(item.get("folder_path", "") or "").strip()
                tag = str(item.get("tag", "WG") or "WG").strip()
                color = str(item.get("color", C.OFFICE_BLUE) or C.OFFICE_BLUE).strip()
                inc_sub = bool(item.get("include_subfolders", True))
            else:
                fpath = str(item or "").strip()
                tag = "WG"
                color = C.OFFICE_BLUE
                inc_sub = True
            self._append_folder_row(fpath, tag=tag, color=color, include_subfolders=inc_sub)

    def _append_folder_row(self, folder_path: str, *, tag: str = "WG", color: str = "", include_subfolders: bool = True):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setRowHeight(row, scale_px(38))

        clean_path = str(folder_path or "").strip()
        leaf_name = clean_path.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1] if clean_path else ""
        path_item = QTableWidgetItem(clean_path)
        if clean_path:
            path_item.setToolTip(
                f"Folder: {leaf_name}\nFull Path: {clean_path}\n\n"
                "(Double-click to browse Outlook folders, or edit path text directly)"
            )
        self.table.setItem(row, 0, path_item)

        self._set_subfolders_widget(row, include_subfolders)

        tag_item = QTableWidgetItem(tag or "WG")
        tag_item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        tag_item.setToolTip("Display badge tag shown on emails synced from this folder (e.g. WG, Discussion, Inbox).")
        self.table.setItem(row, 2, tag_item)

        self._set_color_widget(row, color or C.OFFICE_BLUE)

    def _set_subfolders_widget(self, row: int, checked: bool):
        container = QWidget(self.table)
        container.setStyleSheet(".QWidget { background: transparent; border: none; }")
        lay = QHBoxLayout(container)
        lay.setContentsMargins(6, 2, 6, 2)
        lay.setAlignment(Qt.AlignCenter)
        chk = QCheckBox("Include", container)
        chk.setChecked(bool(checked))
        chk.setToolTip(
            "When checked, scans this folder AND all of its nested subfolders.\n"
            "Uncheck if you only want to scan messages directly inside this top-level folder."
        )
        chk.setStyleSheet(f"QCheckBox {{ color: {C.TEXT_PRIMARY}; font-size: 12px; background: transparent; }}")
        lay.addWidget(chk)
        container.setProperty("_checkbox", chk)
        self.table.setCellWidget(row, 1, container)

    def _get_row_include_subfolders(self, row: int) -> bool:
        w = self.table.cellWidget(row, 1)
        if w is None:
            return True
        chk = w.property("_checkbox")
        if isinstance(chk, QCheckBox):
            return bool(chk.isChecked())
        found = w.findChild(QCheckBox)
        return bool(found.isChecked()) if found is not None else True

    def _set_color_widget(self, row: int, hex_color: str):
        container = QWidget(self.table)
        container.setStyleSheet(".QWidget { background: transparent; border: none; }")
        lay = QHBoxLayout(container)
        lay.setContentsMargins(6, 4, 6, 4)

        btn_color = QPushButton(str(hex_color or C.OFFICE_BLUE).upper(), container)
        btn_color.setCursor(Qt.PointingHandCursor)
        btn_color.setMinimumHeight(scale_px(26))
        self._apply_color_button_style(btn_color, btn_color.text())
        btn_color.clicked.connect(lambda _=False, b=btn_color: self._pick_color(b))
        lay.addWidget(btn_color)
        container.setProperty("_color_btn", btn_color)
        self.table.setCellWidget(row, 3, container)

    @staticmethod
    def _apply_color_button_style(btn: QPushButton, hex_code: str):
        qcol = QColor(hex_code) if QColor(hex_code).isValid() else QColor(C.OFFICE_BLUE)
        lum = qcol.red() * 0.299 + qcol.green() * 0.587 + qcol.blue() * 0.114
        fg = C.TEXT_ON_ACCENT if lum < 155 else C.TEXT_PRIMARY
        btn.setStyleSheet(
            f"QPushButton {{ background-color: {qcol.name().upper()}; color: {fg}; "
            f"font-weight: bold; font-size: 11px; border-radius: 4px; padding: 3px 8px; "
            f"border: 1px solid {C.BORDER_STRONG}; }}"
        )

    def _get_row_color(self, row: int) -> str:
        w = self.table.cellWidget(row, 3)
        if w is None:
            return C.OFFICE_BLUE
        if isinstance(w, QPushButton):
            return w.text().strip() or C.OFFICE_BLUE
        btn = w.property("_color_btn") or w.findChild(QPushButton)
        return (btn.text().strip() if btn is not None else "") or C.OFFICE_BLUE

    def _pick_color(self, btn: QPushButton):
        curr = QColor(btn.text()) if QColor(btn.text()).isValid() else QColor(C.OFFICE_BLUE)
        col = QColorDialog.getColor(curr, self, "Pick Tag Color")
        if col.isValid():
            hex_code = col.name().upper()
            btn.setText(hex_code)
            self._apply_color_button_style(btn, hex_code)

    def _add_folder(self):
        dialog = OutlookFolderPickerDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            path = dialog.get_selected_path()
            if path:
                leaf = path.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1]
                default_tag = "Inbox" if leaf.lower() == "inbox" else "WG"
                self._append_folder_row(path, tag=default_tag, color=C.OFFICE_BLUE, include_subfolders=True)
                self.table.selectRow(self.table.rowCount() - 1)

    def _change_selected_folder(self):
        row = self.table.currentRow()
        if row < 0:
            self._add_folder()
            return
        self._browse_for_row(row)

    def _on_cell_double_clicked(self, row: int, col: int):
        if col == 0 and 0 <= row < self.table.rowCount():
            self._browse_for_row(row)

    def _browse_for_row(self, row: int):
        dialog = OutlookFolderPickerDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            path = dialog.get_selected_path()
            if path:
                leaf = path.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1]
                item = self.table.item(row, 0)
                if item is None:
                    item = QTableWidgetItem(path)
                    self.table.setItem(row, 0, item)
                else:
                    item.setText(path)
                item.setToolTip(f"Folder: {leaf}\nFull Path: {path}")

    def _remove_selected(self):
        sel_Model = self.table.selectionModel()
        rows = sorted({idx.row() for idx in sel_Model.selectedRows()}, reverse=True) if sel_Model else []
        if not rows and self.table.currentRow() >= 0:
            rows = [self.table.currentRow()]
        for idx in rows:
            self.table.removeRow(idx)

    def _save_and_close(self):
        folders = []
        for r in range(self.table.rowCount()):
            p_item = self.table.item(r, 0)
            t_item = self.table.item(r, 2)
            p = p_item.text().strip() if p_item else ""
            t = t_item.text().strip() if t_item else "WG"
            c = self._get_row_color(r)
            inc_sub = self._get_row_include_subfolders(r)
            if p:
                folders.append({
                    "folder_path": p,
                    "tag": t or "WG",
                    "color": c,
                    "include_subfolders": inc_sub,
                })
        save_wg_email_config(self.wg, folders)
        self.accept()


class GeneralEmailSyncDialog(QDialog):
    """Dialog for configuring date range and ±5 day default buffer before running a WG Outlook folder sync."""

    DEFAULT_BUFFER_DAYS = 5

    def __init__(self, wg: str, start_date: str, end_date: str, parent=None):
        super().__init__(parent)
        self.wg = str(wg or "SA2").strip().upper()
        self._init_start = QDate.fromString(start_date, Qt.ISODate) if start_date else QDate()
        self._init_end = QDate.fromString(end_date, Qt.ISODate) if end_date else QDate()
        if not self._init_start.isValid():
            self._init_start = QDate.currentDate().addDays(-14)
        if not self._init_end.isValid():
            self._init_end = QDate.currentDate().addDays(7)

        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowTitle(f"🔄 Sync Emails: {self.wg}")
        self.setMinimumWidth(scale_px(550))
        self.setStyleSheet(f"QDialog {{ background-color: {C.BG_CANVAS}; }}")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        hdr = QLabel(f"📡 Sync Outlook Emails for {self.wg}", self)
        hdr.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        layout.addWidget(hdr)

        # Compact Folder Summary + Quick Configure Card
        folder_card = QFrame(self)
        folder_card.setStyleSheet(CARD_FRAME_STYLE)
        folder_lay = QHBoxLayout(folder_card)
        folder_lay.setContentsMargins(12, 8, 12, 8)
        folder_lay.setSpacing(8)

        self.lbl_folders = QLabel("", folder_card)
        self.lbl_folders.setWordWrap(True)
        self.lbl_folders.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")
        folder_lay.addWidget(self.lbl_folders, 1)

        self.btn_configure_folders = QPushButton("⚙️ Folders...", folder_card)
        self.btn_configure_folders.setMinimumHeight(scale_px(28))
        self.btn_configure_folders.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_configure_folders.setToolTip(f"Configure Outlook folders and subfolder scope for {self.wg}.")
        self.btn_configure_folders.clicked.connect(self._on_configure_folders)
        folder_lay.addWidget(self.btn_configure_folders)
        layout.addWidget(folder_card)

        # Date Range + Buffer Card
        card = QFrame(self)
        card.setStyleSheet(CARD_FRAME_STYLE)
        grid = QGridLayout(card)
        grid.setContentsMargins(12, 10, 12, 10)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)

        lbl_dates = QLabel("Meeting Dates:", card)
        lbl_dates.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-weight: bold;")
        grid.addWidget(lbl_dates, 0, 0)

        dates_row = QHBoxLayout()
        dates_row.setSpacing(6)
        ctrl_h = scale_px(30)

        self.dt_start = QDateEdit(card)
        self.dt_start.setDisplayFormat("yyyy-MM-dd")
        self.dt_start.setCalendarPopup(True)
        self.dt_start.setMinimumWidth(scale_px(142))
        self.dt_start.setMinimumHeight(ctrl_h)
        self.dt_start.setDate(self._init_start)

        self.dt_end = QDateEdit(card)
        self.dt_end.setDisplayFormat("yyyy-MM-dd")
        self.dt_end.setCalendarPopup(True)
        self.dt_end.setMinimumWidth(scale_px(142))
        self.dt_end.setMinimumHeight(ctrl_h)
        self.dt_end.setDate(self._init_end)

        dates_row.addWidget(self.dt_start)
        dates_row.addWidget(QLabel("to", card))
        dates_row.addWidget(self.dt_end)

        self.btn_reset_dates = QPushButton("↺ Reset", card)
        self.btn_reset_dates.setMinimumHeight(scale_px(26))
        self.btn_reset_dates.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_reset_dates.setToolTip("Reset dates to the meeting bounds and ±5 day default buffer.")
        self.btn_reset_dates.clicked.connect(self._on_reset_defaults)
        dates_row.addWidget(self.btn_reset_dates)
        dates_row.addStretch()
        grid.addLayout(dates_row, 0, 1)

        lbl_buf = QLabel("Buffer Window (±):", card)
        lbl_buf.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-weight: bold;")
        grid.addWidget(lbl_buf, 1, 0)

        buf_row = QHBoxLayout()
        buf_row.setSpacing(6)
        self.spin_buf = QSpinBox(card)
        self.spin_buf.setRange(0, 60)
        self.spin_buf.setValue(self.DEFAULT_BUFFER_DAYS)
        self.spin_buf.setSuffix(" days")
        self.spin_buf.setMinimumWidth(scale_px(110))
        self.spin_buf.setMinimumHeight(ctrl_h)
        buf_row.addWidget(self.spin_buf)

        self._preset_buttons = []
        for days_val in (0, 3, 5, 10, 14):
            btn_p = QPushButton(f"±{days_val}d", card)
            btn_p.setCheckable(True)
            btn_p.setMinimumHeight(scale_px(26))
            btn_p.setStyleSheet(
                f"QPushButton {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_STRONG}; "
                f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 2px 8px; font-size: 11px; }} "
                f"QPushButton:hover {{ background-color: {C.BG_HOVER}; }} "
                f"QPushButton:checked {{ background-color: {C.INFO_SOFT}; color: {C.INFO_DARK}; "
                f"border-color: {C.INFO_BORDER}; font-weight: bold; }} "
                f"QPushButton:checked:hover {{ background-color: {C.INFO_SOFT_STRONG}; }}"
            )
            btn_p.clicked.connect(lambda _=False, d=days_val: self.spin_buf.setValue(d))
            buf_row.addWidget(btn_p)
            self._preset_buttons.append((days_val, btn_p))
        buf_row.addStretch()
        grid.addLayout(buf_row, 1, 1)

        self.lbl_effective_window = QLabel("", card)
        self.lbl_effective_window.setWordWrap(True)
        grid.addWidget(self.lbl_effective_window, 2, 0, 1, 2)
        layout.addWidget(card)

        # Footer Row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        tip_lbl = QLabel("Includes Sent Items & deduplicates messages across folders.", self)
        tip_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px;")
        btn_row.addWidget(tip_lbl)
        btn_row.addStretch()

        btn_cancel = QPushButton("Cancel", self)
        btn_cancel.setMinimumHeight(ctrl_h)
        btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_cancel.clicked.connect(self.reject)

        self.btn_sync = QPushButton("🚀 Start Sync", self)
        self.btn_sync.setMinimumHeight(ctrl_h)
        self.btn_sync.setStyleSheet(_PRIMARY_BTN_STYLE)
        self.btn_sync.setDefault(True)
        self.btn_sync.clicked.connect(self.accept)

        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(self.btn_sync)
        layout.addLayout(btn_row)

        self.dt_start.dateChanged.connect(self._update_preview_and_presets)
        self.dt_end.dateChanged.connect(self._update_preview_and_presets)
        self.spin_buf.valueChanged.connect(self._update_preview_and_presets)

        self._has_folders = False
        self._refresh_folders_summary()
        self._update_preview_and_presets()
        self.adjustSize()

    def _refresh_folders_summary(self) -> None:
        folders = load_wg_email_config(self.wg)
        self._has_folders = len(folders) > 0
        if not self._has_folders:
            self.lbl_folders.setText(
                f"⚠️ No Outlook folders are configured for <b>{self.wg}</b>. Click <b>Folders...</b> to add one."
            )
            self.lbl_folders.setStyleSheet(
                f"background-color: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; "
                f"border: 1px solid {C.WARNING_BORDER}; border-radius: 4px; padding: 4px 8px; font-size: 12px;"
            )
        else:
            sub_cnt = sum(1 for f in folders if isinstance(f, dict) and bool(f.get("include_subfolders", True)))
            tags_str = ", ".join(str(f.get("tag", "WG")) for f in folders if isinstance(f, dict)) or "WG"
            self.lbl_folders.setText(
                f"📂 Configured Folders: <b>{len(folders)}</b> ({tags_str}) · "
                f"Subfolders enabled on <b>{sub_cnt}/{len(folders)}</b>"
            )
            self.lbl_folders.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")

    def _on_configure_folders(self) -> None:
        dlg = GeneralEmailFoldersDialog(self.wg, self)
        if dlg.exec_() == QDialog.Accepted:
            self._refresh_folders_summary()
            self._update_preview_and_presets()

    def _on_reset_defaults(self) -> None:
        self.dt_start.setDate(self._init_start)
        self.dt_end.setDate(self._init_end)
        self.spin_buf.setValue(self.DEFAULT_BUFFER_DAYS)

    def _update_preview_and_presets(self) -> None:
        s_date = self.dt_start.date()
        e_date = self.dt_end.date()
        buf = int(self.spin_buf.value())

        for days_val, btn_p in self._preset_buttons:
            btn_p.blockSignals(True)
            btn_p.setChecked(days_val == buf)
            btn_p.blockSignals(False)

        invalid_range = e_date < s_date
        if invalid_range:
            self.lbl_effective_window.setText("⚠️ End date is before start date — adjust the Meeting Dates range.")
            self.lbl_effective_window.setStyleSheet(
                f"background-color: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; "
                f"border: 1px solid {C.WARNING_BORDER}; border-radius: 4px; padding: 5px 8px; font-size: 12px; font-weight: bold;"
            )
        else:
            eff_start = s_date.addDays(-buf)
            eff_end = e_date.addDays(buf)
            span_days = eff_start.daysTo(eff_end) + 1
            self.lbl_effective_window.setText(
                f"📅 Effective scan window: <b>{eff_start.toString(Qt.ISODate)} → {eff_end.toString(Qt.ISODate)}</b> "
                f"({span_days} days total, ±{buf}d buffer)"
            )
            self.lbl_effective_window.setStyleSheet(
                f"background-color: {C.INFO_SOFT}; color: {C.INFO_DARK}; "
                f"border: 1px solid {C.INFO_BORDER}; border-radius: 4px; padding: 5px 8px; font-size: 12px;"
            )
        self.btn_sync.setEnabled(self._has_folders and not invalid_range)

    def get_params(self):
        return {
            "start_date": self.dt_start.date().toString(Qt.ISODate),
            "end_date": self.dt_end.date().toString(Qt.ISODate),
            "buffer": self.spin_buf.value(),
        }
