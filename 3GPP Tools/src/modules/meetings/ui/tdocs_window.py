# --- File: modules/meetings/ui/tdocs_window.py ---
import os
import re
import webbrowser
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableView,
                             QHeaderView, QLabel, QLineEdit, QComboBox, QFrame,
                             QPushButton, QMessageBox, QStyledItemDelegate, QStyleOptionViewItem, QStyle)
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QTextDocument
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel, QEvent, pyqtSignal, QRectF, QSize


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
        self._updating = False  # Prevent recursive loops when bulk-checking

        self.lineEdit().installEventFilter(self)
        self.setModel(QStandardItemModel(self))
        self.view().viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        # 1. Force the dropdown to open if the user clicks the text box
        if obj == self.lineEdit() and event.type() == QEvent.MouseButtonPress:
            self.showPopup()
            return True

        # 2. Handle the checkboxes inside the dropdown
        if obj == self.view().viewport() and event.type() == QEvent.MouseButtonRelease:
            index = self.view().indexAt(event.pos())
            if index.isValid():
                item = self.model().itemFromIndex(index)
                if item:
                    state = item.checkState()
                    new_state = Qt.Unchecked if (state == Qt.Checked or state == 2) else Qt.Checked

                    self._updating = True

                    # If user clicked the "(Select All)" item at index 0
                    if index.row() == 0:
                        item.setCheckState(new_state)
                        # Cascade state to all other items
                        for i in range(1, self.model().rowCount()):
                            self.model().item(i).setCheckState(new_state)
                    else:
                        # User clicked a standard item
                        item.setCheckState(new_state)
                        # Re-evaluate the "(Select All)" checkbox state
                        all_checked = True
                        for i in range(1, self.model().rowCount()):
                            if self.model().item(i).checkState() not in (Qt.Checked, 2):
                                all_checked = False
                                break
                        self.model().item(0).setCheckState(Qt.Checked if all_checked else Qt.Unchecked)

                    self._updating = False

                    self.updateText()
                    self.selectionChanged.emit(self.getCheckedItems())
            return True  # Consume event to prevent popup from closing

        return super().eventFilter(obj, event)

    def addItems(self, items):
        # ADDED: Master "(Select All)" item at the very top
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
        # Skip index 0 because it's the "(Select All)" master toggle
        for i in range(1, self.model().rowCount()):
            item = self.model().item(i)
            state = item.checkState()
            if state == Qt.Checked or state == 2:
                checked.append(item.data(Qt.UserRole))
        return checked

    def updateText(self):
        if self._updating: return
        checked = self.getCheckedItems()
        total = self.model().rowCount() - 1  # Exclude "(Select All)" from math

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
# --- HTML CLICKABLE CELL DELEGATE  ---
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

        # Draw the standard background (maintains alternating row colors cleanly)
        options.text = ""
        options.widget.style().drawControl(QStyle.CE_ItemViewItem, options, painter)

        # Draw the HTML text over it
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


# ==========================================
# --- DATA MODELS ---
# ==========================================
class TDocsTableModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        self._data = data or []
        self._headers = [
            "TDoc", "Title", "Source", "Type", "For",
            "Abstract", "Secretary Remarks", "Agenda Item", "TDoc Status", "Related TDocs"
        ]
        self.valid_tdocs = {str(r.get("TDoc", "")) for r in self._data if r.get("TDoc")}

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

        if role == Qt.DisplayRole:
            if col_name == "Related TDocs":
                return self._format_related_tdocs(row, html=True)
            val = row.get(col_name, "")
            return str(val).strip() if val is not None else ""

        elif role == Qt.UserRole:
            if col_name == "Related TDocs":
                return self._format_related_tdocs(row, html=False)
            val = row.get(col_name, "")
            return str(val).strip() if val is not None else ""

        elif role == Qt.TextAlignmentRole:
            if col_name in ["TDoc", "Type", "For", "Agenda Item", "TDoc Status"]:
                return Qt.AlignCenter
            # ---> FIX: Vertically center all other text columns!
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

        type_data = model.data(model.index(source_row, 3, source_parent), Qt.UserRole)
        if type_data not in self.type_filters: return False

        ai_data = model.data(model.index(source_row, 7, source_parent), Qt.UserRole)
        if ai_data not in self.ai_filters: return False

        status_data = model.data(model.index(source_row, 8, source_parent), Qt.UserRole)
        if status_data not in self.status_filters: return False

        if self.global_filter:
            match_found = False
            for col in [0, 1, 2, 5, 9]:
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
        self.filepath = filepath
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

        self.table = QTableView()
        self.model = TDocsTableModel(tdocs_data)

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

        # ---> FIX 1: Remove bulky row padding and tightly wrap the text
        self.table.setWordWrap(True)
        self.table.verticalHeader().setDefaultSectionSize(20)  # Small, clean default height
        self.table.resizeRowsToContents()  # Calculates exact heights based on content

        self.html_delegate = HtmlDelegate(self.table)
        self.html_delegate.linkClicked.connect(self._scroll_to_tdoc)
        self.table.setItemDelegateForColumn(9, self.html_delegate)
        self.table.viewport().setMouseTracking(True)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)

        # ---> FIX 2: Shift the flexible "Stretch" space from Title to Abstract
        header.resizeSection(0, 110)  # TDoc
        header.resizeSection(1, 250)  # Title (Now narrower and fixed width)
        header.setSectionResizeMode(5, QHeaderView.Stretch)  # Abstract (Now takes all remaining space)
        header.resizeSection(9, 160)  # Related TDocs

        main_layout.addWidget(self.table)

    # --- ACTIONS & TRIGGERS ---
    def _scroll_to_tdoc(self, target_tdoc: str):
        """Finds the TDoc and beautifully scrolls it to the exact center of the screen."""
        for row in range(self.proxy.rowCount()):
            idx = self.proxy.index(row, 0)
            if self.proxy.data(idx, Qt.UserRole) == target_tdoc:
                # FIXED 2: No more selectRow(). We just visually scroll to it.
                self.table.scrollTo(idx, QTableView.PositionAtCenter)
                return

        QMessageBox.information(self, "Hidden",
                                f"TDoc '{target_tdoc}' exists, but is currently hidden by your filters.")

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

        # ---> FIX: Force PyQt to recalculate text-wrap heights whenever rows are filtered!
        if hasattr(self, 'table'):
            self.table.resizeRowsToContents()