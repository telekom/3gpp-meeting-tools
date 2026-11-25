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
    """Dialog for configuring date range and buffer window before running a WG Outlook folder sync."""

    def __init__(self, wg: str, start_date: str, end_date: str, parent=None):
        super().__init__(parent)
        self.wg = str(wg or "SA2").strip().upper()
        self.setWindowTitle(f"🔄 Sync Emails: {self.wg}")
        self.resize(scale_px(480), scale_px(270))
        self.setStyleSheet(f"QDialog {{ background-color: {C.BG_CANVAS}; }}")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        hdr = QLabel(f"📡 Sync Outlook Emails for {self.wg}")
        hdr.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        layout.addWidget(hdr)

        folders = load_wg_email_config(self.wg)
        sub_cnt = sum(1 for f in folders if isinstance(f, dict) and bool(f.get("include_subfolders", True)))
        tags_str = ", ".join(str(f.get("tag", "WG")) for f in folders if isinstance(f, dict)) or "None"
        lbl_folders = QLabel(
            f"Configured Folders: <b>{len(folders)}</b> ({tags_str}) · "
            f"Subfolders enabled on <b>{sub_cnt}/{len(folders)}</b>"
        )
        lbl_folders.setWordWrap(True)
        lbl_folders.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")
        layout.addWidget(lbl_folders)

        card = QFrame(self)
        card.setStyleSheet(CARD_FRAME_STYLE)
        form = QVBoxLayout(card)
        form.setContentsMargins(12, 10, 12, 10)
        form.setSpacing(8)

        dt_layout = QHBoxLayout()
        dt_layout.addWidget(QLabel("Date Range:"))
        self.dt_start = QDateEdit()
        self.dt_start.setCalendarPopup(True)
        self.dt_end = QDateEdit()
        self.dt_end.setCalendarPopup(True)

        if start_date:
            self.dt_start.setDate(QDate.fromString(start_date, Qt.ISODate))
        else:
            self.dt_start.setDate(QDate.currentDate().addDays(-14))

        if end_date:
            self.dt_end.setDate(QDate.fromString(end_date, Qt.ISODate))
        else:
            self.dt_end.setDate(QDate.currentDate().addDays(7))

        dt_layout.addWidget(self.dt_start)
        dt_layout.addWidget(QLabel("to"))
        dt_layout.addWidget(self.dt_end)
        dt_layout.addStretch()
        form.addLayout(dt_layout)

        buf_layout = QHBoxLayout()
        buf_layout.addWidget(QLabel("Buffer Window (Days ±):"))
        self.spin_buf = QSpinBox()
        self.spin_buf.setRange(0, 30)
        self.spin_buf.setValue(3)
        buf_layout.addWidget(self.spin_buf)
        buf_layout.addStretch()
        form.addLayout(buf_layout)
        layout.addWidget(card)

        layout.addStretch()

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_cancel.clicked.connect(self.reject)

        btn_sync = QPushButton("🚀 Start Sync")
        btn_sync.setStyleSheet(_PRIMARY_BTN_STYLE)
        btn_sync.clicked.connect(self.accept)

        btn_row.addStretch()
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_sync)
        layout.addLayout(btn_row)

    def get_params(self):
        return {
            "start_date": self.dt_start.date().toString(Qt.ISODate),
            "end_date": self.dt_end.date().toString(Qt.ISODate),
            "buffer": self.spin_buf.value(),
        }
