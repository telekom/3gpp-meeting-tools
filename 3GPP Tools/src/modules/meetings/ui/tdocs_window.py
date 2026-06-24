# --- File: modules/meetings/ui/tdocs_window.py ---
import os
import re
import webbrowser
import zipfile
import requests
from pathlib import Path

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableView,
                             QHeaderView, QLabel, QLineEdit, QComboBox, QFrame,
                             QPushButton, QMessageBox, QStyledItemDelegate, QStyleOptionViewItem, QStyle)
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QTextDocument, QColor, QPainter, QPen
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel, QEvent, pyqtSignal, QRectF, QSize, \
    QThread, QTimer

from core.network.session import NetworkSession


# ==========================================
# --- BACKGROUND WORKER: TDOC DOWNLOADER ---
# ==========================================
class TDocActionThread(QThread):
    finished_action = pyqtSignal(str, bool, str)

    def __init__(self, tdoc: str, docs_url: str, meeting_dir: Path):
        super().__init__()
        self.tdoc = tdoc
        self.docs_url = docs_url
        self.tdoc_dir = meeting_dir / tdoc
        self.zip_path = self.tdoc_dir / f"{tdoc}.zip"

    def run(self):
        try:
            # 1. Check if viewable files are already extracted
            extracted_files = []
            if self.tdoc_dir.exists():
                valid_exts = ('.doc', '.docx', '.pdf', '.ppt', '.pptx')
                extracted_files = [f for f in self.tdoc_dir.iterdir()
                                   if f.is_file() and f.suffix.lower() in valid_exts and not f.name.startswith('~$')]

            # 2. Download and Extract if necessary
            if not extracted_files:
                if not self.zip_path.exists():
                    self.tdoc_dir.mkdir(parents=True, exist_ok=True)
                    dl_url = self.docs_url.rstrip('/') + f"/{self.tdoc}.zip"

                    session = NetworkSession.get_instance()
                    NetworkSession.apply_humanness(session)
                    response = session.get(dl_url, stream=True, timeout=30)
                    response.raise_for_status()

                    with open(self.zip_path, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=16384):
                            if chunk: f.write(chunk)

                # Extract all valid documents flatly (ignoring internal ZIP folder structures)
                with zipfile.ZipFile(self.zip_path, 'r') as z:
                    for info in z.infolist():
                        if '__MACOSX' in info.filename or info.filename.startswith('._'):
                            continue
                        if info.filename.lower().endswith(('.doc', '.docx', '.pdf', '.ppt', '.pptx')):
                            out_path = self.tdoc_dir / Path(info.filename).name
                            with open(out_path, 'wb') as f:
                                f.write(z.read(info.filename))
                            extracted_files.append(out_path)

            if not extracted_files:
                self.finished_action.emit(self.tdoc, False,
                                          "No viewable documents (.doc, .pdf, .ppt) found inside the ZIP.")
                return

            # 3. Open extracted files
            for doc in extracted_files:
                if hasattr(os, 'startfile'):
                    os.startfile(str(doc))
                else:
                    webbrowser.open(f"file:///{doc}")

            self.finished_action.emit(self.tdoc, True, "Opened successfully.")

        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                self.finished_action.emit(self.tdoc, False, "TDoc ZIP not found on the server (404 Error).")
            else:
                self.finished_action.emit(self.tdoc, False, f"Network error: {e}")
        except Exception as e:
            self.finished_action.emit(self.tdoc, False, str(e))


# ==========================================
# --- UPGRADED: BULLETPROOF MULTI-SELECT ---
# ==========================================
class CheckableComboBox(QComboBox):
    selectionChanged = pyqtSignal(list)

    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.title = title
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self._updating = False

        self.lineEdit().installEventFilter(self)
        self.setModel(QStandardItemModel(self))
        self.view().viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        if obj == self.lineEdit() and event.type() == QEvent.MouseButtonPress:
            self.showPopup()
            return True

        if obj == self.view().viewport() and event.type() == QEvent.MouseButtonRelease:
            index = self.view().indexAt(event.pos())
            if index.isValid():
                item = self.model().itemFromIndex(index)
                if item:
                    state = item.checkState()
                    new_state = Qt.Unchecked if (state == Qt.Checked or state == 2) else Qt.Checked

                    self._updating = True
                    if index.row() == 0:
                        item.setCheckState(new_state)
                        for i in range(1, self.model().rowCount()):
                            self.model().item(i).setCheckState(new_state)
                    else:
                        item.setCheckState(new_state)
                        all_checked = True
                        for i in range(1, self.model().rowCount()):
                            if self.model().item(i).checkState() not in (Qt.Checked, 2):
                                all_checked = False
                                break
                        self.model().item(0).setCheckState(Qt.Checked if all_checked else Qt.Unchecked)

                    self._updating = False
                    self.updateText()
                    self.selectionChanged.emit(self.getCheckedItems())
            return True
        return super().eventFilter(obj, event)

    def addItems(self, items):
        item_all = QStandardItem("(Select All)")
        item_all.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
        item_all.setCheckState(Qt.Checked)
        item_all.setData("ALL", Qt.UserRole)
        self.model().appendRow(item_all)

        for text in items:
            display_text = str(text) if text else "(Empty)"
            item = QStandardItem(display_text)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            item.setData(str(text), Qt.UserRole)
            self.model().appendRow(item)
        self.updateText()

    def getCheckedItems(self):
        checked = []
        for i in range(1, self.model().rowCount()):
            item = self.model().item(i)
            state = item.checkState()
            if state == Qt.Checked or state == 2:
                checked.append(item.data(Qt.UserRole))
        return checked

    def updateText(self):
        if self._updating: return
        checked = self.getCheckedItems()
        total = self.model().rowCount() - 1

        if total == 0:
            self.lineEdit().setText(f"{self.title}: None")
        elif len(checked) == total:
            self.lineEdit().setText(f"{self.title}: All")
        elif len(checked) == 1:
            display = checked[0] if checked[0] else "(Empty)"
            self.lineEdit().setText(f"{self.title}: {display}")
        else:
            self.lineEdit().setText(f"{self.title}: {len(checked)} selected")


# ==========================================
# --- DELEGATES ---
# ==========================================
class HtmlDelegate(QStyledItemDelegate):
    linkClicked = pyqtSignal(str)

    def paint(self, painter, option, index):
        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        painter.save()
        doc = QTextDocument()
        doc.setDocumentMargin(4)
        doc.setDefaultFont(options.font)
        doc.setHtml(options.text)

        options.text = ""
        options.widget.style().drawControl(QStyle.CE_ItemViewItem, options, painter)

        painter.translate(options.rect.left(), options.rect.top())
        clip = QRectF(0, 0, options.rect.width(), options.rect.height())
        doc.drawContents(painter, clip)
        painter.restore()

    def sizeHint(self, option, index):
        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        doc = QTextDocument()
        doc.setDocumentMargin(4)
        doc.setDefaultFont(options.font)
        doc.setHtml(options.text)
        return QSize(int(doc.idealWidth()), int(doc.size().height()))

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease:
            options = QStyleOptionViewItem(option)
            self.initStyleOption(options, index)
            doc = QTextDocument()
            doc.setDocumentMargin(4)
            doc.setDefaultFont(options.font)
            doc.setHtml(options.text)

            pos = event.pos() - option.rect.topLeft()
            anchor = doc.documentLayout().anchorAt(pos)
            if anchor:
                self.linkClicked.emit(anchor)
                return True
        return super().editorEvent(event, model, option, index)


class TDocActionDelegate(QStyledItemDelegate):
    actionClicked = pyqtSignal(str)

    def paint(self, painter, option, index):
        options = QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        options.text = ""
        options.widget.style().drawControl(QStyle.CE_ItemViewItem, options, painter)

        tdoc = index.data(Qt.UserRole)
        state = index.data(Qt.UserRole + 1)
        if not tdoc: return

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)

        # Add internal padding to make it float like a pill
        rect = option.rect.adjusted(4, 4, -4, -4)

        if state == 'EXISTS':
            bg_color, border_color, text_color = QColor("#E6F4E6"), QColor("#A3DDA3"), QColor("#0C6B0C")
            text = "✓ Open"
        elif state == 'LOADING':
            bg_color, border_color, text_color = QColor("#FFF4CE"), QColor("#F3C74C"), QColor("#B85C00")
            text = "⏳ Fetching"
        else:
            bg_color, border_color, text_color = QColor("#E1F0FF"), QColor("#99C9FF"), QColor("#005A9E")
            text = "⬇ Download"

        painter.setBrush(bg_color)
        painter.setPen(QPen(border_color, 1))
        painter.drawRoundedRect(rect, 10, 10)

        painter.setPen(text_color)
        font = painter.font()
        font.setBold(True)
        font.setPointSize(8)
        painter.setFont(font)
        painter.drawText(rect, Qt.AlignCenter, text)

        painter.restore()

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease:
            rect = option.rect.adjusted(4, 4, -4, -4)
            if rect.contains(event.pos()):
                tdoc = index.data(Qt.UserRole)
                if tdoc:
                    self.actionClicked.emit(tdoc)
                    return True
        return super().editorEvent(event, model, option, index)


# ==========================================
# --- DATA MODELS ---
# ==========================================
class TDocsTableModel(QAbstractTableModel):
    def __init__(self, meeting_dir: Path, data=None):
        super().__init__()
        self.meeting_dir = meeting_dir
        self._data = data or []
        # INDEX 0 is now the Action Button Column
        self._headers = [
            "", "TDoc", "Title", "Source", "Type", "For",
            "Abstract", "Secretary Remarks", "Agenda Item", "TDoc Status", "Related TDocs"
        ]
        self.valid_tdocs = {str(r.get("TDoc", "")) for r in self._data if r.get("TDoc")}
        self.loading_tdocs = set()

    def set_loading(self, tdoc: str, is_loading: bool):
        if is_loading:
            self.loading_tdocs.add(tdoc)
        else:
            self.loading_tdocs.discard(tdoc)

        for r in range(self.rowCount()):
            if self._data[r].get("TDoc") == tdoc:
                idx = self.index(r, 0)
                self.dataChanged.emit(idx, idx)
                break

    def _linkify(self, prefix: str, text: str, html: bool) -> str:
        if not text: return ""
        if not html: return f"{prefix}: {text}"

        def repl(match):
            tdoc = match.group(0)
            if tdoc in self.valid_tdocs:
                return f'<a href="{tdoc}" style="color: #005A9E; font-weight: bold; text-decoration: underline;">{tdoc}</a>'
            else:
                return f'<span style="color: #999999;">{tdoc}</span>'

        linked_text = re.sub(r'[a-zA-Z0-9]+-\d+', repl, text)
        return f"<span style='color: #444;'><b>{prefix}:</b></span> {linked_text}"

    def _format_related_tdocs(self, row_data: dict, html=False) -> str:
        parts = []
        r_rev_of = row_data.get("Is revision of")
        r_rev_to = row_data.get("Revised to")
        r_orig = row_data.get("Original LS")
        r_reply = row_data.get("Reply in")

        if r_rev_of: parts.append(self._linkify("⬅️ Rev of", r_rev_of, html))
        if r_rev_to: parts.append(self._linkify("➡️ Rev to", r_rev_to, html))
        if r_orig: parts.append(self._linkify("✉️ Orig LS", r_orig, html))
        if r_reply: parts.append(self._linkify("↩️ Reply", r_reply, html))

        separator = "<br>" if html else "\n"
        return separator.join(parts)

    def data(self, index, role):
        if not index.isValid(): return None
        row = self._data[index.row()]
        col_name = self._headers[index.column()]

        # Handle the new Action Column
        if col_name == "":
            if role == Qt.UserRole:
                return row.get("TDoc", "")
            if role == Qt.UserRole + 1:
                tdoc = row.get("TDoc", "")
                if tdoc in self.loading_tdocs: return "LOADING"
                zip_path = self.meeting_dir / tdoc / f"{tdoc}.zip"
                return "EXISTS" if zip_path.exists() else "MISSING"
            return None

        # Display Role (What the user visually sees)
        if role == Qt.DisplayRole:
            if col_name == "Related TDocs":
                return self._format_related_tdocs(row, html=True)

            val = row.get(col_name, "")
            val_str = str(val).strip() if val is not None else ""

            # ---> FIX: Flatten and truncate the Abstract so rows stay short!
            if col_name == "Abstract" and len(val_str) > 90:
                flat_str = val_str.replace('\n', ' ').replace('\r', '')
                return flat_str[:87] + "..."

            return val_str

        # User Role (Raw data used by the background Search Engine)
        elif role == Qt.UserRole:
            if col_name == "Related TDocs":
                return self._format_related_tdocs(row, html=False)
            val = row.get(col_name, "")
            return str(val).strip() if val is not None else ""

        # ---> NEW: ToolTip Role (Shows the full text on hover!)
        elif role == Qt.ToolTipRole:
            val = row.get(col_name, "")
            val_str = str(val).strip() if val is not None else ""

            if col_name == "Abstract" and val_str:
                # Wrap in a fixed-width HTML div so PyQt automatically word-wraps the tooltip beautifully
                return f"<div style='width: 400px; white-space: pre-wrap;'>{val_str}</div>"
            elif col_name in ["Title", "Source", "Secretary Remarks"] and len(val_str) > 30:
                return val_str
            return None

        # Alignment
        elif role == Qt.TextAlignmentRole:
            if col_name in ["TDoc", "Type", "For", "Agenda Item", "TDoc Status", "Related TDocs"]:
                return Qt.AlignCenter
            return Qt.AlignLeft | Qt.AlignVCenter

        return None

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole: return self._headers[section]
        return None


class TDocsFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.global_filter = ""
        self.type_filters = set()
        self.status_filters = set()
        self.ai_filters = set()

    def setGlobalFilter(self, text):
        self.global_filter = str(text).lower().strip()
        self.invalidateFilter()

    def setTypeFilters(self, types):
        self.type_filters = set(types)
        self.invalidateFilter()

    def setStatusFilters(self, statuses):
        self.status_filters = set(statuses)
        self.invalidateFilter()

    def setAIFilters(self, ais):
        self.ai_filters = set(ais)
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()

        # Columns shifted +1 due to the new Action button at index 0
        type_data = model.data(model.index(source_row, 4, source_parent), Qt.UserRole)
        if type_data not in self.type_filters: return False

        ai_data = model.data(model.index(source_row, 8, source_parent), Qt.UserRole)
        if ai_data not in self.ai_filters: return False

        status_data = model.data(model.index(source_row, 9, source_parent), Qt.UserRole)
        if status_data not in self.status_filters: return False

        if self.global_filter:
            match_found = False
            # Search TDoc, Title, Source, Abstract, and Related
            for col in [1, 2, 3, 6, 10]:
                data = model.data(model.index(source_row, col, source_parent), Qt.UserRole)
                if data and self.global_filter in str(data).lower():
                    match_found = True
                    break
            if not match_found: return False

        return True


# ==========================================
# --- TDOCS WINDOW ---
# ==========================================
class TDocsWindow(QWidget):
    def __init__(self, mtg_info: dict, tdocs_data: list, filepath: str):
        super().__init__()
        self.mtg_info = mtg_info
        self.filepath = filepath
        self.meeting_dir = Path(filepath).parent.parent
        self.active_threads = {}

        title = f"TDocs: {mtg_info.get('wg_name', '')} {mtg_info.get('meeting_number', '')}"
        self.setWindowTitle(title)
        self.resize(1400, 750)
        self.setStyleSheet("QWidget { background-color: #FAFAFA; }")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # --- HEADER & COUNT ---
        header_layout = QHBoxLayout()
        title_lbl = QLabel(f"<b>{title}</b>")
        title_lbl.setStyleSheet("font-size: 18px; color: #333;")

        self.folder_btn = QPushButton("📂 Meeting Folder")
        self.folder_btn.setCursor(Qt.PointingHandCursor)
        self.folder_btn.setStyleSheet("""
            QPushButton {
                font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold;
                border-radius: 6px; padding: 5px 12px;
                color: #005A9E; background-color: #E1F0FF; border: 1px solid #99C9FF;
            }
            QPushButton:hover { background-color: #CCE4FF; border: 1px solid #005A9E; }
        """)
        self.folder_btn.clicked.connect(self._open_meeting_folder)

        self.excel_btn = QPushButton("📗 Open in Excel")
        self.excel_btn.setCursor(Qt.PointingHandCursor)
        self.excel_btn.setStyleSheet("""
            QPushButton {
                font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; font-weight: bold;
                border-radius: 6px; padding: 5px 12px;
                color: #0C6B0C; background-color: #E6F4E6; border: 1px solid #A3DDA3;
            }
            QPushButton:hover { background-color: #D1EED1; border: 1px solid #0C6B0C; }
        """)
        self.excel_btn.clicked.connect(self._open_excel)

        self.count_lbl = QLabel(f"Showing {len(tdocs_data)} of {len(tdocs_data)} TDocs")
        self.count_lbl.setStyleSheet("font-size: 13px; color: #666;")

        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        header_layout.addWidget(self.folder_btn)
        header_layout.addWidget(self.excel_btn)
        header_layout.addSpacing(15)
        header_layout.addWidget(self.count_lbl)
        main_layout.addLayout(header_layout)

        # --- MODERN FILTER BAR ---
        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; }
            QLabel { font-weight: bold; color: #555; border: none; }
            QLineEdit, QComboBox { padding: 6px; border: 1px solid #CCC; border-radius: 4px; background: #FFF; }
            QLineEdit:focus { border: 1px solid #0078D7; }
        """)
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(15, 10, 15, 10)

        filter_layout.addWidget(QLabel("🔍 Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search TDoc number, title, source, or abstract...")
        self.search_input.setMinimumWidth(250)
        self.search_input.textChanged.connect(self._on_search_changed)
        filter_layout.addWidget(self.search_input)

        def sanitize(val): return str(val).strip() if val is not None else ""

        self.type_combo = CheckableComboBox("Type")
        self.type_combo.setMinimumWidth(150)
        unique_types = sorted(list(set(sanitize(r.get("Type", "")) for r in tdocs_data)))
        self.type_combo.addItems(unique_types)
        self.type_combo.selectionChanged.connect(self._on_type_changed)
        filter_layout.addWidget(self.type_combo)

        self.ai_combo = CheckableComboBox("AI")
        self.ai_combo.setMinimumWidth(150)
        unique_ais = sorted(list(set(sanitize(r.get("Agenda Item", "")) for r in tdocs_data)))
        self.ai_combo.addItems(unique_ais)
        self.ai_combo.selectionChanged.connect(self._on_ai_changed)
        filter_layout.addWidget(self.ai_combo)

        self.status_combo = CheckableComboBox("Status")
        self.status_combo.setMinimumWidth(150)
        unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in tdocs_data)))
        self.status_combo.addItems(unique_statuses)
        self.status_combo.selectionChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

        main_layout.addWidget(filter_frame)

        # --- TABLE SETUP ---
        self.table = QTableView()
        self.model = TDocsTableModel(self.meeting_dir, tdocs_data)

        self.proxy = TDocsFilterProxyModel()
        self.proxy.setSourceModel(self.model)
        self.proxy.layoutChanged.connect(self._update_count_label)

        self.proxy.setTypeFilters(unique_types)
        self.proxy.setAIFilters(unique_ais)
        self.proxy.setStatusFilters(unique_statuses)

        self.table.setModel(self.proxy)

        self.table.setSelectionMode(QTableView.NoSelection)
        self.table.setFocusPolicy(Qt.NoFocus)

        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setStyleSheet("""
            QTableView { gridline-color: #E0E0E0; border: 1px solid #E0E0E0; background-color: #FFFFFF; }
            QHeaderView::section { background-color: #F5F5F5; padding: 4px; font-weight: bold; border: 1px solid #E0E0E0; }
        """)

        self.table.setWordWrap(True)
        self.table.verticalHeader().setDefaultSectionSize(20)
        self.table.resizeRowsToContents()

        # Action Delegate (Index 0)
        self.action_delegate = TDocActionDelegate(self.table)
        self.action_delegate.actionClicked.connect(self._handle_tdoc_action)
        self.table.setItemDelegateForColumn(0, self.action_delegate)

        # HTML Link Delegate (Index 10)
        self.html_delegate = HtmlDelegate(self.table)
        self.html_delegate.linkClicked.connect(self._scroll_to_tdoc)
        self.table.setItemDelegateForColumn(10, self.html_delegate)
        self.table.viewport().setMouseTracking(True)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.resizeSection(0, 85)  # Action Button
        header.resizeSection(1, 110)  # TDoc
        header.resizeSection(2, 200)  # Title (Narrowed)
        header.resizeSection(3, 100)  # Source (Constrained)

        # Abstract stretches to consume all remaining space dynamically
        header.setSectionResizeMode(6, QHeaderView.Stretch)
        header.resizeSection(10, 160)  # Related TDocs

        main_layout.addWidget(self.table)

    # --- ACTIONS & TRIGGERS ---
    def _handle_tdoc_action(self, tdoc: str):
        if tdoc in self.model.loading_tdocs:
            return

        docs_url = self.mtg_info.get("docs_folder_url")
        if not docs_url:
            QMessageBox.warning(self, "Missing URL", "This meeting does not have a Docs/ URL mapped in the database.")
            return

        self.model.set_loading(tdoc, True)
        QTimer.singleShot(0, self.table.resizeRowsToContents) # <--- NEW: Stop row from shrinking!

        thread = TDocActionThread(tdoc, docs_url, self.meeting_dir)
        thread.finished_action.connect(self._on_tdoc_action_finished)
        self.active_threads[tdoc] = thread
        thread.start()

    def _on_tdoc_action_finished(self, tdoc: str, success: bool, msg: str):
        if tdoc in self.active_threads:
            del self.active_threads[tdoc]

        self.model.set_loading(tdoc, False)
        QTimer.singleShot(0, self.table.resizeRowsToContents)  # <--- NEW: Stop row from shrinking!

        if not success:
            QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)

    def _scroll_to_tdoc(self, target_tdoc: str):
        for row in range(self.proxy.rowCount()):
            idx = self.proxy.index(row, 1)  # Column 1 is TDoc string
            if self.proxy.data(idx, Qt.UserRole) == target_tdoc:
                self.table.scrollTo(idx, QTableView.PositionAtCenter)
                return
        QMessageBox.information(self, "Hidden", f"TDoc '{target_tdoc}' is currently hidden by your filters.")

    def _open_meeting_folder(self):
        if self.meeting_dir.exists():
            if hasattr(os, 'startfile'):
                os.startfile(str(self.meeting_dir))
            else:
                webbrowser.open(f"file:///{self.meeting_dir}")
        else:
            QMessageBox.warning(self, "Not Found", "The root meeting folder has not been created yet.")

    def _open_excel(self):
        try:
            if hasattr(os, 'startfile'):
                os.startfile(self.filepath)
            else:
                webbrowser.open(f"file:///{self.filepath}")
        except Exception as e:
            QMessageBox.warning(self, "Open Error", f"Could not open the Excel file:\n{e}")

    def _on_search_changed(self, text):
        self.proxy.setGlobalFilter(text)

    def _on_type_changed(self, types):
        self.proxy.setTypeFilters(types)

    def _on_ai_changed(self, ais):
        self.proxy.setAIFilters(ais)

    def _on_status_changed(self, statuses):
        self.proxy.setStatusFilters(statuses)

    def _update_count_label(self):
        visible = self.proxy.rowCount()
        total = self.model.rowCount()
        self.count_lbl.setText(f"Showing {visible} of {total} TDocs")

        # <--- NEW: Force resize AFTER the filter completely finishes updating the UI!
        if hasattr(self, 'table'):
            QTimer.singleShot(0, self.table.resizeRowsToContents)