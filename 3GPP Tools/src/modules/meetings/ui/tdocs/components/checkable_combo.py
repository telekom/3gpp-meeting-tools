# --- File: src/modules/meetings/ui/tdocs/components/checkable_combo.py ---
"""Multi-select CheckableComboBox with summary sizing, activeFilter state, and availability dimming."""

import logging
import re

from PyQt5.QtCore import QEvent, QSize, Qt, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QPalette, QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import (
    QComboBox,
    QListView,
    QStyle,
    QStyleOptionComboBox,
    QStylePainter,
)

from core.ui.font_scale import get_font_scale_notifier, scale_px
from core.ui.theme import C

from .combo_delegate import CheckableComboItemDelegate


class CheckableComboBox(QComboBox):
    """Multi-selection combo box displaying summary labels and theme-aware checkable popup rows."""

    selectionChanged = pyqtSignal(list)

    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.title = title

        self.setEditable(False)
        self.setMaxVisibleItems(25)

        view = QListView(self)
        view.setObjectName("checkableComboPopupView")
        view.setUniformItemSizes(True)
        view.setStyleSheet(f"""
            QListView#checkableComboPopupView {{
                background-color: {C.BG_CARD};
                color: {C.TEXT_STRONG};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 3px;
                outline: none;
            }}
            QListView#checkableComboPopupView::item {{
                min-height: 20px;
                padding: 0px 4px;
                border-radius: 3px;
            }}
            QListView#checkableComboPopupView::item:hover {{
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_PRIMARY};
            }}
            QListView#checkableComboPopupView::item:selected {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
            }}
        """)
        self._item_delegate = CheckableComboItemDelegate(view)
        view.setItemDelegate(self._item_delegate)
        self.setView(view)
        self.setModel(QStandardItemModel(self))

        self._updating = False
        self._current_summary_text = f"{self.title}: All"

        self.view().viewport().installEventFilter(self)
        get_font_scale_notifier().scale_changed.connect(self._on_font_scale_changed)

    def _on_font_scale_changed(self, _scale: float = 0.0) -> None:
        """Recompute custom summary width/height when global text scaling changes."""
        self.updateGeometry()
        self.update()

    def _compute_summary_size_hint(self, minimum: bool = False) -> QSize:
        """Compute sizeHint from rendered summary labels rather than wide popup items."""
        fm = self.fontMetrics()
        current_text = getattr(self, "_current_summary_text", "") or f"{self.title}: All"
        if minimum:
            candidates = (f"{self.title}: All", current_text)
        else:
            candidates = (
                current_text,
                f"{self.title}: All",
                f"{self.title}: 99 sel",
                f"{self.title}: None",
            )
        max_text_w = max(fm.horizontalAdvance(text) for text in candidates)
        opt = QStyleOptionComboBox()
        self.initStyleOption(opt)
        styled = self.style().sizeFromContents(
            QStyle.CT_ComboBox,
            opt,
            QSize(max_text_w + scale_px(18), max(fm.height(), scale_px(16))),
            self,
        )
        pad_w = scale_px(32 if minimum else 42)
        width = max(styled.width(), max_text_w + pad_w)
        max_w = self.maximumWidth()
        if 0 < max_w < 16777215:
            width = min(width, max_w)
        height = max(styled.height(), fm.height() + scale_px(10))
        return QSize(width, height)

    def sizeHint(self) -> QSize:
        return self._compute_summary_size_hint(minimum=False)

    def minimumSizeHint(self) -> QSize:
        return self._compute_summary_size_hint(minimum=True)

    def showPopup(self):
        if self.view() is not None:
            fm = self.view().fontMetrics()
            max_item_w = 0
            for row in range(self.model().rowCount()):
                item = self.model().item(row)
                if item and item.text():
                    max_item_w = max(max_item_w, fm.horizontalAdvance(item.text()))
            self.view().setMinimumWidth(
                min(scale_px(480), max(self.width(), max_item_w + scale_px(68)))
            )
        super().showPopup()
        if self.view():
            self.view().scrollToTop()

    def _refresh_text_and_tooltip(self):
        """Calculates custom text and tooltip based on checked items and updates activeFilter state."""
        checked = self.getCheckedItems()
        total = max(0, self.model().rowCount() - 1)
        is_filtered = (0 < len(checked) < total) or (total > 0 and len(checked) == 0)

        if total == 0:
            text = f"{self.title}: None"
        elif len(checked) == total:
            text = f"{self.title}: All"
        elif len(checked) == 1:
            raw_val = str(checked[0]) if checked[0] else "(Empty)"
            if len(raw_val) > 14:
                display_val = raw_val[:12] + "…"
            else:
                display_val = raw_val
            text = f"{self.title}: {display_val}"
            self.setToolTip(f"{self.title}: {raw_val}")
        else:
            text = f"{self.title}: {len(checked)} sel"
            if len(checked) <= 6:
                self.setToolTip(f"{self.title} ({len(checked)} selected):\n" + "\n".join(f"• {c}" for c in checked))
            else:
                self.setToolTip(f"{self.title}: {len(checked)} selected")

        if getattr(self, "_active_filter_state", None) != is_filtered:
            self._active_filter_state = is_filtered
            self.setProperty("activeFilter", "true" if is_filtered else "false")
            self.style().unpolish(self)
            self.style().polish(self)

        prev_text = getattr(self, "_current_summary_text", None)
        self._current_summary_text = text
        if prev_text != text:
            self.updateGeometry()
        return text

    def paintEvent(self, event):
        """Overrides the default rendering to display custom summary text natively."""
        painter = QStylePainter(self)
        painter.setPen(self.palette().color(QPalette.Text))

        opt = QStyleOptionComboBox()
        self.initStyleOption(opt)

        opt.currentText = self._refresh_text_and_tooltip()

        self.style().drawComplexControl(QStyle.CC_ComboBox, opt, painter, self)
        self.style().drawControl(QStyle.CE_ComboBoxLabel, opt, painter, self)

    def eventFilter(self, obj, event):
        try:
            if obj == self.view().viewport():
                if event.type() == QEvent.MouseButtonRelease:
                    return True

                elif event.type() == QEvent.MouseButtonPress:
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
                                    child = self.model().item(i)
                                    if child:
                                        child.setCheckState(new_state)
                            else:
                                item.setCheckState(new_state)
                                all_checked = True
                                for i in range(1, self.model().rowCount()):
                                    child = self.model().item(i)
                                    if child and child.checkState() not in (Qt.Checked, 2):
                                        all_checked = False
                                        break

                                root_item = self.model().item(0)
                                if root_item:
                                    root_item.setCheckState(Qt.Checked if all_checked else Qt.Unchecked)

                            self._updating = False
                            self._refresh_text_and_tooltip()
                            self.update()
                            self.selectionChanged.emit(self.getCheckedItems())
                    return True

        except Exception as e:
            logging.error(f"[CheckableComboBox] Event filter error: {e}", exc_info=True)

        return super().eventFilter(obj, event)

    def addItems(self, items, display_labels=None, item_tooltips=None):
        display_labels = display_labels or {}
        item_tooltips = item_tooltips or {}
        item_all = QStandardItem("(Select All)")
        item_all.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
        item_all.setCheckState(Qt.Checked)
        item_all.setData("ALL", Qt.UserRole)
        self.model().appendRow(item_all)

        for text in items:
            raw_value = str(text)
            display_text = display_labels.get(raw_value, raw_value if raw_value else "(Empty)")
            display_text = re.sub(r'[\r\n]+', ' ', str(display_text)).strip() or "(Empty)"
            item = QStandardItem(display_text)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            item.setData(str(text), Qt.UserRole)
            if raw_value in item_tooltips:
                item.setToolTip(item_tooltips[raw_value])
            self.model().appendRow(item)
        super().setCurrentIndex(0)
        self._refresh_text_and_tooltip()
        self.update()

    def updateItems(self, items, display_labels=None, item_tooltips=None):
        previously_checked = set(self.getCheckedItems())
        previously_all = {
            str(self.model().item(i).data(Qt.UserRole))
            for i in range(1, self.model().rowCount())
            if self.model().item(i) is not None
        }
        display_labels = display_labels or {}
        item_tooltips = item_tooltips or {}

        was_all_checked = True
        if self.model().rowCount() > 0:
            root_item = self.model().item(0)
            if root_item:
                was_all_checked = root_item.checkState() == Qt.Checked

        self._updating = True
        try:
            if self.model().rowCount() > 0:
                self.model().removeRows(0, self.model().rowCount())

            item_all = QStandardItem("(Select All)")
            item_all.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            item_all.setData("ALL", Qt.UserRole)
            self.model().appendRow(item_all)

            all_checked_now = True
            for text in items:
                raw_value = str(text)
                display_text = display_labels.get(raw_value, raw_value if raw_value else "(Empty)")
                display_text = re.sub(r'[\r\n]+', ' ', str(display_text)).strip() or "(Empty)"
                item = QStandardItem(display_text)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)

                if was_all_checked or raw_value in previously_checked or (previously_all and raw_value not in previously_all):
                    item.setCheckState(Qt.Checked)
                else:
                    item.setCheckState(Qt.Unchecked)
                    all_checked_now = False

                item.setData(raw_value, Qt.UserRole)
                if raw_value in item_tooltips:
                    item.setToolTip(item_tooltips[raw_value])
                self.model().appendRow(item)

            root_item = self.model().item(0)
            if root_item:
                root_item.setCheckState(Qt.Checked if all_checked_now else Qt.Unchecked)
        finally:
            self._updating = False
        super().setCurrentIndex(0)
        self._refresh_text_and_tooltip()
        self.update()

    def currentData(self, role=Qt.UserRole):
        """Backward compatibility helper returning selected values as comma-separated string."""
        checked = self.getCheckedItems()
        if not checked:
            return ""
        return ",".join(str(c) for c in checked)

    def findData(self, data, role=Qt.UserRole):
        """Backward compatibility helper finding item index by user role value."""
        data_str = str(data).strip()
        for i in range(1, self.model().rowCount()):
            item = self.model().item(i)
            if item and str(item.data(role)).strip() == data_str:
                return i
        return -1

    def setCurrentIndex(self, index: int):
        """Backward compatibility helper setting single item or clearing all."""
        if index <= 0:
            self.setCheckedItems([])
        else:
            item = self.model().item(index)
            if item:
                val = item.data(Qt.UserRole)
                if val:
                    self.setCheckedItems([str(val)])

    def setAvailability(self, available_values):
        """Grey zero-intersection choices while keeping every option selectable."""
        available = {str(value) for value in (available_values or [])}
        muted_brush = QBrush(QColor(C.TEXT_FAINT))
        normal_brush = QBrush(self.palette().color(QPalette.Text))

        for i in range(1, self.model().rowCount()):
            item = self.model().item(i)
            if not item:
                continue
            value = str(item.data(Qt.UserRole))
            is_available = value in available
            item.setForeground(normal_brush if is_available else muted_brush)
            item.setData(is_available, Qt.UserRole + 1)

        self._refresh_text_and_tooltip()
        self.update()

    def getCheckedItems(self):
        checked = []
        for i in range(1, self.model().rowCount()):
            item = self.model().item(i)
            if item:
                state = item.checkState()
                if state == Qt.Checked or state == 2:
                    checked.append(item.data(Qt.UserRole))
        return checked

    def isAllChecked(self) -> bool:
        """Return True if all data items are currently checked."""
        total = max(0, self.model().rowCount() - 1)
        if total == 0:
            return True
        return len(self.getCheckedItems()) == total

    def isFiltered(self) -> bool:
        """Return True if the combobox is actively filtering (i.e. not all items are checked)."""
        total = max(0, self.model().rowCount() - 1)
        if total == 0:
            return False
        return len(self.getCheckedItems()) < total

    def setCheckedItems(self, items: list, allow_empty: bool = False):
        """Set the check state of items matching the given list of raw values/strings."""
        raw_list = list(items or [])
        items_set = {str(x).strip() for x in raw_list}
        model = self.model()
        row_count = model.rowCount()
        if row_count <= 1:
            return

        self._updating = True
        model.blockSignals(True)
        try:
            if not raw_list and not allow_empty:
                for row in range(row_count):
                    item = model.item(row)
                    if item:
                        item.setCheckState(Qt.Checked)
            else:
                all_checked = True
                for row in range(1, row_count):
                    item = model.item(row)
                    if item:
                        raw_val = item.data(Qt.UserRole)
                        val = str(item.text() if raw_val is None else raw_val).strip()
                        is_checked = val in items_set
                        item.setCheckState(Qt.Checked if is_checked else Qt.Unchecked)
                        if not is_checked:
                            all_checked = False

                root = model.item(0)
                if root:
                    root.setCheckState(Qt.Checked if all_checked else Qt.Unchecked)
        finally:
            model.blockSignals(False)
            self._updating = False

        self._refresh_text_and_tooltip()
        self.update()

    def updateText(self):
        # Deprecated: Legacy support for external calls. Rendering is now automatic.
        self._refresh_text_and_tooltip()
        self.update()
