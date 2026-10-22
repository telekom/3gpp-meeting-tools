# --- File: src/modules/meetings/ui/tdocs_components.py ---
import logging
from PyQt5.QtCore import pyqtSignal, QEvent, Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QPalette, QColor, QBrush
from PyQt5.QtWidgets import QComboBox, QListView, QStylePainter, QStyleOptionComboBox, QStyle


class CheckableComboBox(QComboBox):
    selectionChanged = pyqtSignal(list)

    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.title = title

        # Native, non-editable combobox behavior
        self.setEditable(False)

        self.setView(QListView(self))
        self.setModel(QStandardItemModel(self))

        self._updating = False

        # Install filter ONLY to catch clicks on the dropdown list (viewport)
        self.view().viewport().installEventFilter(self)

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

        self._current_summary_text = text
        return text

    def paintEvent(self, event):
        """Overrides the default rendering to display custom summary text natively."""
        painter = QStylePainter(self)
        painter.setPen(self.palette().color(QPalette.Text))

        opt = QStyleOptionComboBox()
        self.initStyleOption(opt)

        opt.currentText = self._refresh_text_and_tooltip()

        # Draw the combobox natively with the custom text
        self.style().drawComplexControl(QStyle.CC_ComboBox, opt, painter, self)
        self.style().drawControl(QStyle.CE_ComboBoxLabel, opt, painter, self)

    def eventFilter(self, obj, event):
        try:
            if obj == self.view().viewport():
                # ---> THE FIX 1: Consume the release event so the menu doesn't close natively,
                # and it prevents the initial "click release" from toggling the first item.
                if event.type() == QEvent.MouseButtonRelease:
                    return True

                # ---> THE FIX 2: Handle toggling checkboxes explicitly on PRESS instead!
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
                                    if child: child.setCheckState(new_state)
                            else:
                                item.setCheckState(new_state)
                                all_checked = True
                                for i in range(1, self.model().rowCount()):
                                    child = self.model().item(i)
                                    if child and child.checkState() not in (Qt.Checked, 2):
                                        all_checked = False
                                        break

                                root_item = self.model().item(0)
                                if root_item: root_item.setCheckState(Qt.Checked if all_checked else Qt.Unchecked)

                            self._updating = False
                            self._refresh_text_and_tooltip()
                            self.update()  # Trigger paintEvent to update the button text
                            self.selectionChanged.emit(self.getCheckedItems())
                    return True

        except Exception as e:
            logging.error(f"[CheckableComboBox] Event filter error: {e}", exc_info=True)

        return super().eventFilter(obj, event)

    def addItems(self, items, display_labels=None):
        display_labels = display_labels or {}
        item_all = QStandardItem("(Select All)")
        item_all.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
        item_all.setCheckState(Qt.Checked)
        item_all.setData("ALL", Qt.UserRole)
        self.model().appendRow(item_all)

        for text in items:
            raw_value = str(text)
            display_text = display_labels.get(raw_value, raw_value if raw_value else "(Empty)")
            item = QStandardItem(display_text)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            item.setData(str(text), Qt.UserRole)
            self.model().appendRow(item)
        self._refresh_text_and_tooltip()
        self.update()

    def updateItems(self, items, display_labels=None):
        previously_checked = set(self.getCheckedItems())
        display_labels = display_labels or {}

        was_all_checked = True
        if self.model().rowCount() > 0:
            root_item = self.model().item(0)
            if root_item:
                was_all_checked = (root_item.checkState() == Qt.Checked)

        self.model().blockSignals(True)

        # Safely remove rows to keep viewport intact
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
            item = QStandardItem(display_text)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)

            if was_all_checked or raw_value in previously_checked:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
                all_checked_now = False

            item.setData(raw_value, Qt.UserRole)
            self.model().appendRow(item)

        root_item = self.model().item(0)
        if root_item:
            root_item.setCheckState(Qt.Checked if all_checked_now else Qt.Unchecked)

        self.model().blockSignals(False)
        self._refresh_text_and_tooltip()
        self.update()  # Trigger paintEvent to update the button text

    def setAvailability(self, available_values):
        """Grey zero-intersection choices while keeping every option selectable."""
        available = {str(value) for value in (available_values or [])}
        muted_brush = QBrush(QColor("#94A3B8"))
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

    def setCheckedItems(self, items: list):
        """Set the check state of items matching the given list of raw values/strings."""
        items_set = {str(x).strip() for x in (items or []) if str(x).strip()}
        model = self.model()
        row_count = model.rowCount()
        if row_count <= 1:
            return

        self._updating = True
        model.blockSignals(True)
        try:
            if not items_set:
                for row in range(row_count):
                    item = model.item(row)
                    if item:
                        item.setCheckState(Qt.Checked)
            else:
                all_checked = True
                for row in range(1, row_count):
                    item = model.item(row)
                    if item:
                        val = str(item.data(Qt.UserRole) or item.text()).strip()
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