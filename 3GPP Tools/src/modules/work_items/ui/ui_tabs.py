# --- File: src/modules/work_items/ui/ui_tabs.py ---
import webbrowser
from datetime import datetime
from pathlib import Path
import logging

from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QTimer, QEvent, QRect, pyqtSignal
from PyQt5.QtGui import QColor, QPalette
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTableView, QHeaderView,
    QPushButton, QProgressBar, QMessageBox, QLineEdit, QMenu, QStyle,
    QApplication, QStyledItemDelegate, QComboBox, QDialog, QFrame,
    QScrollArea, QFormLayout
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_scraper import WorkItemsScraperThread, TargetedWIScraperThread
from modules.work_items.core.wi_settings import WorkItemsSettings


class WidDelegate(QStyledItemDelegate):
    """Renders the Latest WID as a clickable hyperlink using the common palette."""
    link_clicked = pyqtSignal(str, str)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text:
            return

        painter.save()
        font = option.font
        font.setUnderline(True)
        painter.setFont(font)

        if option.state & QStyle.State_Selected:
            painter.setPen(option.palette.color(QPalette.HighlightedText))
        else:
            painter.setPen(QColor("#1E5C99"))

        painter.drawText(option.rect, Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            text = index.data(Qt.DisplayRole)
            if text:
                self.link_clicked.emit(text, "open_doc")
                return True
        return super().editorEvent(event, model, option, index)


class SpecsDelegate(QStyledItemDelegate):
    """Renders the linked specifications count as a standardized pill button."""
    specs_clicked = pyqtSignal(int)

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)
        text = index.data(Qt.DisplayRole)
        if not text or text == "-":
            painter.save()
            painter.setPen(QColor("#94A3B8"))
            painter.drawText(option.rect, Qt.AlignCenter, "-")
            painter.restore()
            return

        badge_rect = QRect(option.rect.left() + 4, option.rect.top() + 4, option.rect.width() - 8, option.rect.height() - 8)

        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        painter.setBrush(QColor("#EBF3FC"))
        painter.setPen(QColor("#BFDBFE"))
        painter.drawRoundedRect(badge_rect, 10, 10)

        painter.setPen(QColor("#1E5C99"))
        font = option.font
        font.setBold(True)
        font.setPointSize(font.pointSize() - 1 if font.pointSize() > 8 else 8)
        painter.setFont(font)
        painter.drawText(badge_rect, Qt.AlignCenter, text)
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            text = index.data(Qt.DisplayRole)
            if text and text != "-":
                self.specs_clicked.emit(index.row())
                return True
        return super().editorEvent(event, model, option, index)


class RemarksDelegate(QStyledItemDelegate):
    """Draws the latest remark text with a clickable history badge."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.button_width = 45
        self.button_margin = 4

    def get_button_rect(self, option):
        rect = option.rect
        return QRect(
            rect.right() - self.button_width - self.button_margin,
            rect.top() + self.button_margin,
            self.button_width,
            rect.height() - (2 * self.button_margin)
        )

    def paint(self, painter, option, index):
        QApplication.style().drawControl(QStyle.CE_ItemViewItem, option, painter)

        remarks_list = index.data(Qt.UserRole + 1)
        if not remarks_list:
            return

        latest_remark = remarks_list[0]
        count = len(remarks_list)

        text_rect = option.rect.adjusted(6, 0, -(self.button_width + 10), 0)
        if option.state & QStyle.State_Selected:
            painter.setPen(option.palette.color(QPalette.HighlightedText))
        else:
            painter.setPen(option.palette.color(QPalette.Text))

        elided_text = option.fontMetrics.elidedText(latest_remark, Qt.ElideRight, text_rect.width())
        painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, elided_text)

        btn_rect = self.get_button_rect(option)
        painter.save()
        painter.setRenderHint(painter.Antialiasing)
        painter.setBrush(QColor("#EBF3FC"))
        painter.setPen(QColor("#BFDBFE"))
        painter.drawRoundedRect(btn_rect, 4, 4)
        painter.setPen(QColor("#1E5C99"))
        painter.drawText(btn_rect, Qt.AlignCenter, f"💬 {count}")
        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            btn_rect = self.get_button_rect(option)
            if btn_rect.contains(event.pos()):
                remarks_list = index.data(Qt.UserRole + 1)
                if remarks_list:
                    menu = QMenu()
                    for remark in remarks_list:
                        menu.addAction(remark)
                    menu.exec_(event.globalPos())
                return True
        return super().editorEvent(event, model, option, index)


class WorkItemInfoDialog(QDialog):
    """Inspector Dialog displaying full Work Item details, linked specs, and remarks history."""

    def __init__(self, details: dict, parent=None):
        super().__init__(parent)
        self.details = details
        wi_code = details.get('code', 'Unknown')
        acronym = details.get('acronym', 'Unknown')
        name = details.get('name', 'No Title Available')
        release = details.get('release', '')
        end_date = details.get('end_date', '')

        self.setWindowTitle(f"Work Item Details: {acronym} (#{wi_code})")
        self.setMinimumWidth(620)
        self.setMinimumHeight(480)
        self.setStyleSheet("""
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
            }
            QLabel { font-size: 13px; color: #1E293B; }
            QPushButton#specChipBtn {
                background-color: #EBF3FC;
                border: 1px solid #BFDBFE;
                border-radius: 12px;
                padding: 3px 10px;
                font-size: 11px;
                color: #1E5C99;
                font-weight: bold;
            }
            QPushButton#specChipBtn:hover {
                background-color: #DBEAFE;
                border-color: #1E5C99;
            }
            QScrollArea { border: none; background-color: transparent; }
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
        code_badge.setStyleSheet("""
            background-color: #EBF3FC; color: #1E5C99; border: 1px solid #BFDBFE;
            border-radius: 4px; padding: 2px 6px; font-size: 12px; font-weight: bold;
        """)
        code_badge.setTextInteractionFlags(Qt.TextSelectableByMouse)

        if release:
            rel_badge = QLabel(f"<b>{release}</b>")
            rel_badge.setStyleSheet("""
                background-color: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0;
                border-radius: 4px; padding: 2px 6px; font-size: 12px; font-weight: bold;
            """)
            title_row.addWidget(rel_badge)

        is_finished = bool(end_date and end_date.strip() and end_date < datetime.now().strftime("%Y-%m-%d"))
        status_badge = QLabel("🏁 Finished" if is_finished else "🟢 Active / Ongoing")
        status_badge.setStyleSheet(
            "background-color: #F1F5F9; color: #475569; border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
            if is_finished else
            "background-color: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
        )

        acronym_lbl = QLabel(f"<b>{acronym}</b>")
        acronym_lbl.setStyleSheet("font-size: 16px; color: #1E293B; font-weight: bold;")
        acronym_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)

        title_row.addWidget(code_badge)
        title_row.addWidget(acronym_lbl)
        title_row.addWidget(status_badge)
        title_row.addStretch()
        header_layout.addLayout(title_row)

        desc_label = QLabel(name)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #475569; font-size: 13px; line-height: 1.4;")
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

        self._add_row(form, "Working Groups", details.get('wg_names') or '-')
        self._add_row(form, "Start Date", details.get('start_date') or '-')
        self._add_row(form, "Target End Date", details.get('end_date') or '-')
        self._add_row(form, "Latest WID", details.get('latest_wid') or '-')

        linked_specs = details.get('linked_specs', [])
        if linked_specs:
            specs_container = QWidget()
            specs_layout = QHBoxLayout(specs_container)
            specs_layout.setContentsMargins(0, 0, 0, 0)
            specs_layout.setSpacing(6)

            for sp in linked_specs:
                num = sp.get('number', '')
                is_pri = sp.get('is_primary', False)
                clean_num = num.split("-")[0].replace('.', '').strip()
                label = f"⭐ {num}" if is_pri else num

                btn = QPushButton(label)
                btn.setObjectName("specChipBtn")
                btn.setCursor(Qt.PointingHandCursor)
                btn.setToolTip(f"{sp.get('title', '')}\nClick to open 3GPP Portal Report")
                btn.clicked.connect(lambda _, n=clean_num: webbrowser.open(f"https://www.3gpp.org/DynaReport/{n}.htm"))
                specs_layout.addWidget(btn)

            specs_layout.addStretch()
            form.addRow(self._make_key_label("Impacted Specs:"), specs_container)
        else:
            self._add_row(form, "Impacted Specs", "-")

        layout.addWidget(details_card)

        # --- 3. HISTORICAL REMARKS FEED ---
        remarks = details.get('remarks', [])
        if remarks:
            remarks_lbl = QLabel(f"<b>Secretary Remarks ({len(remarks)}):</b>")
            remarks_lbl.setStyleSheet("color: #475569; font-size: 12px; margin-top: 4px;")
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
                r_card.setStyleSheet("background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px; padding: 6px;")
                r_box = QVBoxLayout(r_card)
                r_box.setContentsMargins(6, 4, 6, 4)
                r_box.setSpacing(3)

                if rm.get('date'):
                    date_lbl = QLabel(f"📅 {rm['date']}")
                    date_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold;")
                    r_box.addWidget(date_lbl)

                text_lbl = QLabel(rm.get('text', ''))
                text_lbl.setWordWrap(True)
                text_lbl.setStyleSheet("color: #1E293B; font-size: 12px;")
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
        portal_btn.clicked.connect(lambda: webbrowser.open(
            f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}"
        ))
        btn_layout.addWidget(portal_btn)

        copy_btn = QPushButton("📋 Copy Citation")
        copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_btn.clicked.connect(self._copy_citation)
        btn_layout.addWidget(copy_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _copy_citation(self):
        c = f"WI {self.details.get('code')}: {self.details.get('acronym')} — {self.details.get('name')} ({self.details.get('release')}, {self.details.get('wg_names')})"
        QApplication.clipboard().setText(c)

    def _make_key_label(self, text: str) -> QLabel:
        lbl = QLabel(f"<b>{text}</b>")
        lbl.setStyleSheet("color: #64748B; font-size: 12px;")
        return lbl

    def _add_row(self, form: QFormLayout, label_text: str, value_text: str):
        val_label = QLabel(value_text)
        val_label.setWordWrap(True)
        val_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow(self._make_key_label(f"{label_text}:"), val_label)


class LinkedSpecsDialog(QDialog):
    """Dialog displaying all specifications linked to a given Work Item."""

    def __init__(self, wi_code: str, acronym: str, specs: list, parent=None):
        super().__init__(parent)
        title_str = f"Linked Specifications for WI {acronym} ({wi_code})" if acronym else f"Linked Specifications for WI #{wi_code}"
        self.setWindowTitle(title_str)
        self.setMinimumWidth(580)
        self.setMinimumHeight(380)
        self.setStyleSheet("""
            QFrame#specCard { background-color: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px; }
            QFrame#specCard:hover { border-color: #94A3B8; background-color: #FAFCFF; }
            QLabel { color: #1E293B; }
            QScrollArea { border: 1px solid #CBD5E1; border-radius: 6px; background-color: #FFFFFF; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        header_text = (
            f"Specifications Impacted by Work Item <b>{acronym}</b> ({wi_code}):"
            if acronym else f"Specifications Impacted by Work Item <b>#{wi_code}</b>:"
        )
        header = QLabel(header_text)
        header.setStyleSheet("font-size: 14px; color: #1E293B;")
        header.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(header)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        scroll_content = QWidget()
        cards_layout = QVBoxLayout(scroll_content)
        cards_layout.setContentsMargins(8, 8, 8, 8)
        cards_layout.setSpacing(8)

        if specs:
            for sp in specs:
                num = sp.get('number', 'Unknown')
                title = sp.get('title', 'No Title Available')
                spec_type = sp.get('type', 'TS')
                series_name = sp.get('series_name', num.split('.')[0] if '.' in num else '')
                is_pri = bool(sp.get('is_primary', False))

                card = self._create_spec_card(num, title, spec_type, series_name, is_pri)
                cards_layout.addWidget(card)
        else:
            no_data = QLabel("No linked specifications recorded in the local database for this Work Item.")
            no_data.setStyleSheet("color: #64748B; font-style: italic; padding: 25px;")
            no_data.setAlignment(Qt.AlignCenter)
            cards_layout.addWidget(no_data)

        cards_layout.addStretch()
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

        footer_layout = QHBoxLayout()
        count_lbl = QLabel(f"<b>Total:</b> {len(specs)} specification(s)")
        count_lbl.setStyleSheet("color: #64748B; font-size: 12px;")
        footer_layout.addWidget(count_lbl)
        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        layout.addLayout(footer_layout)

    def _create_spec_card(self, num: str, title: str, spec_type: str, series_name: str, is_primary: bool) -> QFrame:
        card = QFrame()
        card.setObjectName("specCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(6)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        type_badge = QLabel(f"<b>{spec_type}</b>")
        type_badge.setStyleSheet("""
            background-color: #EBF3FC; color: #1E5C99; border: 1px solid #BFDBFE;
            border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;
        """)
        top_row.addWidget(type_badge)

        num_lbl = QLabel(f"<b>{num}</b>")
        num_lbl.setStyleSheet("font-size: 14px; font-weight: bold; color: #1E293B;")
        num_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        top_row.addWidget(num_lbl)

        if is_primary:
            pri_badge = QLabel("⭐ Primary Spec")
            pri_badge.setStyleSheet("""
                background-color: #FEF3C7; color: #92400E; border: 1px solid #FDE68A;
                border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;
            """)
            top_row.addWidget(pri_badge)

        top_row.addStretch()

        clean_num = num.split("-")[0].replace('.', '').strip()
        dyna_url = f"https://www.3gpp.org/DynaReport/{clean_num}.htm" if clean_num else ""
        ftp_url = f"https://www.3gpp.org/ftp/Specs/archive/{series_name}_series/{num}/" if series_name else ""

        if dyna_url:
            dyna_btn = QPushButton("🌐 DynaReport")
            dyna_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            dyna_btn.setCursor(Qt.PointingHandCursor)
            dyna_btn.clicked.connect(lambda: webbrowser.open(dyna_url))
            top_row.addWidget(dyna_btn)

        if ftp_url:
            ftp_btn = QPushButton("📂 FTP Archive")
            ftp_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            ftp_btn.setCursor(Qt.PointingHandCursor)
            ftp_btn.clicked.connect(lambda: webbrowser.open(ftp_url))
            top_row.addWidget(ftp_btn)

        card_layout.addLayout(top_row)

        title_lbl = QLabel(title)
        title_lbl.setWordWrap(True)
        title_lbl.setStyleSheet("font-size: 12px; color: #475569; line-height: 1.4;")
        title_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        card_layout.addWidget(title_lbl)

        return card


class WorkItemsTableModel(QAbstractTableModel):
    """Fast tabular data model supporting chunked row insertions."""

    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = ["Code", "Acronym", "Name", "WG", "Latest WID", "Release", "Specs", "Start Date", "End Date", "Remarks"]

    def data(self, index, role):
        if not index.isValid():
            return None
        row = self._data[index.row()]
        col_name = self._headers[index.column()]

        if role in (Qt.DisplayRole, Qt.UserRole):
            key_map = {
                "Code": "code", "Acronym": "acronym", "Name": "name",
                "WG": "wg_names", "Latest WID": "latest_wid", "Release": "release",
                "Specs": "spec_count", "Start Date": "start_date", "End Date": "end_date",
                "Remarks": "remarks"
            }
            if col_name == "Remarks" and role == Qt.DisplayRole:
                return ""

            if col_name == "Specs" and role == Qt.DisplayRole:
                cnt = row.get("spec_count", 0)
                return f"📑 {cnt}" if cnt > 0 else "-"

            val = row.get(key_map.get(col_name, ""), "")
            return str(val).strip() if val is not None else ""

        elif role == Qt.UserRole + 1 and col_name == "Remarks":
            raw_remarks = row.get("remarks")
            if raw_remarks:
                bundled_remarks = raw_remarks.split("|||")
                parsed_remarks = []
                for item in bundled_remarks:
                    parts = item.split(":::", 1)
                    parsed_remarks.append((parts[0], parts[1]) if len(parts) == 2 else ("", item))
                parsed_remarks.sort(key=lambda x: x[0], reverse=True)
                return [item[1] for item in parsed_remarks]
            return []

        elif role == Qt.ToolTipRole:
            if col_name == "Name":
                return f"<b>{row.get('acronym', '')}</b><br>{row.get('name', '')}"
            elif col_name == "Acronym":
                return row.get('name', '')
            elif col_name == "Specs":
                cnt = row.get("spec_count", 0)
                return f"Click to view all {cnt} linked specifications" if cnt > 0 else "No linked specifications"
            elif col_name == "Latest WID":
                wid = row.get("latest_wid", "")
                return f"Click to download/open latest Work Item Description ({wid})" if wid else ""
            elif col_name == "Remarks":
                raw = row.get("remarks", "")
                if raw:
                    entries = [r.split(":::")[-1] for r in raw.split("|||")]
                    return "<br>• " + "<br>• ".join(entries[:5])
            return None

        elif role == Qt.TextAlignmentRole:
            if col_name in ["Name", "Remarks"]:
                return Qt.AlignLeft | Qt.AlignVCenter
            return Qt.AlignCenter
        return None

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._headers[section]
        return None

    def get_row_data(self, row_idx: int) -> dict:
        if 0 <= row_idx < len(self._data):
            return self._data[row_idx]
        return {}

    def update_data(self, new_data):
        self.beginResetModel()
        self._data = new_data
        self.endResetModel()

    def append_data(self, new_rows):
        if not new_rows:
            return
        start_row = len(self._data)
        end_row = start_row + len(new_rows) - 1
        self.beginInsertRows(QModelIndex(), start_row, end_row)
        self._data.extend(new_rows)
        self.endInsertRows()


class WorkItemsTab(QWidget):
    global_action_requested = pyqtSignal(str, str)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db_path = db_path
        self.db = WorkItemsDatabase(db_path)
        self.settings = WorkItemsSettings()

        self._chunk_size = 60
        self._current_offset = 0
        self._total_count = 0
        self._loaded_data = []
        self._is_loading_chunk = False

        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(350)
        self.search_timer.timeout.connect(lambda: self.refresh_table(reset_pagination=True))

        self.save_filters_timer = QTimer()
        self.save_filters_timer.setSingleShot(True)
        self.save_filters_timer.setInterval(1000)
        self.save_filters_timer.timeout.connect(self._save_filters)

        self._setup_ui()
        self._populate_filters()
        self._load_filters()
        self.refresh_table(reset_pagination=True)

    def _save_filters(self):
        filters = {
            "search": self.search_input.text().strip(),
            "releases": self.release_combo.getCheckedItems(),
            "wgs": self.wg_combo.getCheckedItems(),
            "status": self.status_combo.currentData() or "active"
        }
        self.settings.save_filters(filters)

    def _load_filters(self):
        filters = self.settings.get_filters()
        if not filters:
            idx = self.status_combo.findData("active")
            if idx >= 0:
                self.status_combo.setCurrentIndex(idx)
            return

        if "search" in filters:
            self.search_input.setText(filters["search"])

        if "releases" in filters:
            self._apply_checked_items(self.release_combo, filters["releases"])

        if "wgs" in filters:
            self._apply_checked_items(self.wg_combo, filters["wgs"])

        if "status" in filters:
            idx = self.status_combo.findData(filters["status"])
            if idx >= 0:
                self.status_combo.setCurrentIndex(idx)
        else:
            idx = self.status_combo.findData("active")
            if idx >= 0:
                self.status_combo.setCurrentIndex(idx)

    def _apply_checked_items(self, combo, items_to_check):
        model = combo.model()
        items_to_check_stripped = [str(x).strip() for x in items_to_check]

        for row in range(model.rowCount()):
            if hasattr(model, 'item'):
                item = model.item(row)
                if item.text().strip() in items_to_check_stripped:
                    item.setCheckState(Qt.Checked)
                else:
                    item.setCheckState(Qt.Unchecked)
            else:
                index = model.index(row, 0)
                text = str(model.data(index, Qt.DisplayRole)).strip()
                state = Qt.Checked if text in items_to_check_stripped else Qt.Unchecked
                model.setData(index, state, Qt.CheckStateRole)

        if hasattr(combo, 'updateText'):
            combo.updateText()
        elif hasattr(combo, 'repaint'):
            combo.repaint()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # --- HEADER & CONTROLS ---
        header_layout = QHBoxLayout()
        header_lbl = QLabel("<b>📋 3GPP Work Items (WIs)</b>")
        header_lbl.setStyleSheet("font-size: 16px; color: #1E293B;")

        self.sync_btn = QPushButton("🔄 Sync 3GPP WIs")
        self.sync_btn.setObjectName("primaryBtn")
        self.sync_btn.setToolTip("Click to download and synchronize 3GPP Work Items in parallel from the server.")
        self.sync_btn.clicked.connect(self._start_sync)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(200)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color: #64748B; font-style: italic;")

        header_layout.addWidget(header_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.status_lbl)
        header_layout.addWidget(self.progress_bar)
        header_layout.addWidget(self.sync_btn)
        main_layout.addLayout(header_layout)

        # --- INLINE SEARCH & FILTER BAR ---
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("<b>🔍 Local Search:</b>"))

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search Code, Acronym, or Name...")
        self.search_input.setToolTip("Filter the table instantly by typing keywords.")
        self.search_input.textChanged.connect(lambda: self.search_timer.start())
        search_layout.addWidget(self.search_input)

        self.release_combo = CheckableComboBox("Release")
        self.release_combo.setToolTip("Filter by 3GPP Release")
        self.release_combo.setMinimumWidth(150)
        self.release_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.release_combo)

        self.wg_combo = CheckableComboBox("WG")
        self.wg_combo.setToolTip("Filter by Working Group")
        self.wg_combo.setMinimumWidth(130)
        self.wg_combo.selectionChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.wg_combo)

        self.status_combo = QComboBox()
        self.status_combo.setToolTip("Filter by Work Item completion status (End Date)")
        self.status_combo.setMinimumWidth(150)
        self.status_combo.addItem("🟢 Active / Ongoing", "active")
        self.status_combo.addItem("All Statuses", "all")
        self.status_combo.addItem("🏁 Finished", "finished")
        self.status_combo.currentIndexChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.status_combo)

        self.clear_filters_btn = QPushButton("❌ Clear")
        self.clear_filters_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_filters_btn.setToolTip("Reset all search fields and active filters.")
        self.clear_filters_btn.setCursor(Qt.PointingHandCursor)
        self.clear_filters_btn.clicked.connect(self._clear_all_filters)
        search_layout.addWidget(self.clear_filters_btn)

        main_layout.addLayout(search_layout)

        # --- RESULTS COUNTER ---
        self.count_label = QLabel("Showing 0 Work Items")
        self.count_label.setStyleSheet("font-weight: bold; color: #64748B; margin-top: 4px;")
        count_layout = QHBoxLayout()
        count_layout.addStretch()
        count_layout.addWidget(self.count_label)
        main_layout.addLayout(count_layout)

        # --- TABLE VIEW ---
        self.table = QTableView()
        self.table_model = WorkItemsTableModel()
        self.table.setModel(self.table_model)

        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setStyleSheet("""
            QTableView {
                border: 1px solid #CBD5E1;
                gridline-color: #F1F5F9;
                background-color: #FFFFFF;
            }
            QTableView::item:selected {
                background-color: #EBF3FC;
                color: #1E293B;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                padding: 4px;
                font-weight: bold;
                border: 1px solid #E2E8F0;
                color: #1E293B;
            }
        """)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.doubleClicked.connect(self._on_row_double_clicked)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Stretch)

        # Column Delegates
        self.wid_delegate = WidDelegate(self.table)
        self.wid_delegate.link_clicked.connect(self._log_and_forward_action)
        self.table.setItemDelegateForColumn(4, self.wid_delegate)

        self.specs_delegate = SpecsDelegate(self.table)
        self.specs_delegate.specs_clicked.connect(self._on_specs_clicked)
        self.table.setItemDelegateForColumn(6, self.specs_delegate)

        header.setSectionResizeMode(9, QHeaderView.Stretch)
        self.table.setItemDelegateForColumn(9, RemarksDelegate(self.table))

        self.table.verticalScrollBar().valueChanged.connect(self._on_scroll)
        main_layout.addWidget(self.table)

    def _clear_all_filters(self):
        self.search_input.clear()
        self._apply_checked_items(self.release_combo, [])
        self._apply_checked_items(self.wg_combo, [])
        self.status_combo.setCurrentIndex(0)
        self.refresh_table(reset_pagination=True)

    def _on_scroll(self, value):
        max_val = self.table.verticalScrollBar().maximum()
        if value >= max_val - 5 and not self._is_loading_chunk:
            if len(self._loaded_data) < self._total_count:
                self._load_next_chunk()

    def _log_and_forward_action(self, text, action):
        self.global_action_requested.emit(text, action)

    def _populate_filters(self):
        options = self.db.get_filter_options()

        self.release_combo.blockSignals(True)
        self.release_combo.updateItems(options.get('releases', []))
        self.release_combo.blockSignals(False)

        self.wg_combo.blockSignals(True)
        self.wg_combo.updateItems(options.get('groups', []))
        self.wg_combo.blockSignals(False)

    def refresh_table(self, reset_pagination: bool = True):
        self.save_filters_timer.start()

        if reset_pagination:
            self._current_offset = 0
            self._loaded_data = []

        search_term = self.search_input.text().strip()
        selected_releases = self.release_combo.getCheckedItems()
        selected_wgs = self.wg_combo.getCheckedItems()
        selected_status = self.status_combo.currentData() or "active"

        self._total_count = self.db.count_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status
        )

        chunk = self.db.search_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status,
            limit=self._chunk_size,
            offset=self._current_offset
        )

        self._loaded_data = chunk
        self.table_model.update_data(self._loaded_data)
        self._update_count_label()

    def _load_next_chunk(self):
        self._is_loading_chunk = True
        self._current_offset += self._chunk_size

        search_term = self.search_input.text().strip()
        selected_releases = self.release_combo.getCheckedItems()
        selected_wgs = self.wg_combo.getCheckedItems()
        selected_status = self.status_combo.currentData() or "active"

        chunk = self.db.search_work_items(
            search_term=search_term if search_term else None,
            releases=selected_releases,
            wg_names=selected_wgs,
            status=selected_status,
            limit=self._chunk_size,
            offset=self._current_offset
        )

        if chunk:
            self.table_model.append_data(chunk)
            self._loaded_data.extend(chunk)
            self._update_count_label()

        self._is_loading_chunk = False

    def _update_count_label(self):
        loaded = len(self._loaded_data)
        if self._total_count > loaded:
            self.count_label.setText(f"Showing {loaded} of {self._total_count} Work Items (scroll down to load more)")
        else:
            self.count_label.setText(f"Showing {self._total_count} Work Items")

    def _on_row_double_clicked(self, index):
        if not index.isValid():
            return
        row_data = self.table_model.get_row_data(index.row())
        wi_code = row_data.get('code')
        if wi_code:
            self._show_wi_info(wi_code)

    def _on_specs_clicked(self, row_idx: int):
        row_data = self.table_model.get_row_data(row_idx)
        wi_code = row_data.get('code')
        acronym = row_data.get('acronym', '')
        if wi_code:
            specs = self.db.get_linked_specs_for_wi(wi_code)
            dialog = LinkedSpecsDialog(wi_code, acronym, specs, self)
            dialog.exec_()

    def _show_wi_info(self, wi_code: str):
        details = self.db.get_work_item_details(wi_code)
        if details:
            dialog = WorkItemInfoDialog(details, self)
            dialog.exec_()

    def _start_sync(self):
        self.sync_btn.setEnabled(False)
        self.sync_btn.setText("⏳ Syncing...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.scraper_thread = WorkItemsScraperThread(self.db_path, self)
        self.scraper_thread.progress.connect(self._update_progress)
        self.scraper_thread.finished_sync.connect(self._on_sync_finished)
        self.scraper_thread.start()

    def _update_progress(self, current: int, total: int, msg: str):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_lbl.setText(msg)

    def _on_sync_finished(self, success: bool, msg: str):
        self.sync_btn.setEnabled(True)
        self.sync_btn.setText("🔄 Sync 3GPP WIs")
        self.progress_bar.setVisible(False)
        self.status_lbl.setText("")

        self._populate_filters()
        self.refresh_table(reset_pagination=True)

        if success:
            QMessageBox.information(self, "Sync Complete", msg)
        else:
            QMessageBox.warning(self, "Sync Failed", msg)

    def _show_context_menu(self, position):
        selected_indexes = self.table.selectionModel().selectedRows()
        if not selected_indexes:
            return

        menu = QMenu(self)
        len_indexes = len(selected_indexes)
        wi_code_list = [
            self.table_model.data(self.table_model.index(e.row(), 0), Qt.DisplayRole)
            for e in selected_indexes
        ]
        wi_code_list = [code for code in wi_code_list if code]

        if len_indexes == 1:
            row_idx = selected_indexes[0].row()
            row_data = self.table_model.get_row_data(row_idx)
            wi_code = row_data.get('code')
            acronym = row_data.get('acronym', '')
            name = row_data.get('name', '')
            wg = row_data.get('wg_names', '')
            rel = row_data.get('release', '')

            if not wi_code:
                return

            info_action = menu.addAction("ℹ️ View WI Details")
            wi_page_action = menu.addAction("🌐 Open WI Page (Portal)")
            local_specs_action = menu.addAction("📑 View Linked Specifications in DB")
            specs_action = menu.addAction("📂 Specifications Resulting from this WI (Web)")
            crs_action = menu.addAction("📄 CRs Related to this WI (Web)")
            menu.addSeparator()
            citation_action = menu.addAction("📋 Copy WI Citation")
            menu.addSeparator()
            update_action = menu.addAction("🔄 Update WI")
            delete_action = menu.addAction("🗑️ Delete this Work Item")
        else:
            info_action = None
            wi_page_action = None
            local_specs_action = None
            specs_action = None
            crs_action = None
            citation_action = None
            wi_code = None
            update_action = menu.addAction(f"🔄 Update WIs ({len_indexes} WIs)")
            delete_action = menu.addAction(f"🗑️ Delete selected Work Items ({len_indexes} WIs)")

        action = menu.exec_(self.table.viewport().mapToGlobal(position))

        if len_indexes == 1:
            if action == info_action:
                self._show_wi_info(wi_code)
            elif action == wi_page_action:
                webbrowser.open(f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}")
            elif action == local_specs_action:
                specs = self.db.get_linked_specs_for_wi(wi_code)
                dialog = LinkedSpecsDialog(wi_code, acronym, specs, self)
                dialog.exec_()
            elif action == specs_action:
                webbrowser.open(f"https://portal.3gpp.org/Specifications.aspx?q=1&WiUid={wi_code}")
            elif action == crs_action:
                webbrowser.open(f"https://portal.3gpp.org/ChangeRequests.aspx?q=1&workitem={wi_code}")
            elif action == citation_action:
                citation = f"WI {wi_code}: {acronym} — {name} ({rel}, {wg})"
                QApplication.clipboard().setText(citation)
            elif action == delete_action:
                confirm = QMessageBox.question(
                    self, "Confirm Deletion",
                    f"Are you sure you want to delete Work Item '{wi_code}' from the database?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if confirm == QMessageBox.Yes:
                    self.db.delete_work_item(wi_code)
                    self.refresh_table(reset_pagination=True)
            elif action == update_action:
                self._start_targeted_sync([wi_code])
        else:
            if action == delete_action:
                confirm = QMessageBox.question(
                    self, "Confirm Batch Deletion",
                    f"Are you sure you want to delete {len(wi_code_list)} selected Work Items from the database?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if confirm == QMessageBox.Yes:
                    self.db.delete_work_items(wi_code_list)
                    self.refresh_table(reset_pagination=True)
            elif action == update_action:
                self._start_targeted_sync(wi_code_list)

    def _start_targeted_sync(self, wi_codes: list):
        self.sync_btn.setEnabled(False)
        self.sync_btn.setText("⏳ Updating...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.targeted_thread = TargetedWIScraperThread(self.db_path, wi_codes, self)
        self.targeted_thread.progress.connect(self._update_progress)
        self.targeted_thread.finished_sync.connect(self._on_sync_finished)
        self.targeted_thread.start()