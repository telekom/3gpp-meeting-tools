# --- File: src/modules/meetings/ui/tdocs_components.py ---
import logging
from PyQt5.QtCore import pyqtSignal, QEvent, Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QPalette
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

    def paintEvent(self, event):
        """Overrides the default rendering to display custom summary text natively."""
        painter = QStylePainter(self)
        painter.setPen(self.palette().color(QPalette.Text))

        opt = QStyleOptionComboBox()
        self.initStyleOption(opt)

        # Calculate custom text based on checked items
        checked = self.getCheckedItems()
        total = max(0, self.model().rowCount() - 1)

        if total == 0:
            text = f"{self.title}: None"
        elif len(checked) == total:
            text = f"{self.title}: All"
        elif len(checked) == 1:
            text = f"{self.title}: {checked[0] if checked[0] else '(Empty)'}"
        else:
            text = f"{self.title}: {len(checked)} selected"

        opt.currentText = text

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
                            self.update()  # Trigger paintEvent to update the button text
                            self.selectionChanged.emit(self.getCheckedItems())
                    return True

        except Exception as e:
            logging.error(f"[CheckableComboBox] Event filter error: {e}", exc_info=True)

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
        self.update()

    def updateItems(self, items):
        previously_checked = set(self.getCheckedItems())

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
            display_text = str(text) if text else "(Empty)"
            item = QStandardItem(display_text)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)

            if was_all_checked or text in previously_checked:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
                all_checked_now = False

            item.setData(str(text), Qt.UserRole)
            self.model().appendRow(item)

        root_item = self.model().item(0)
        if root_item:
            root_item.setCheckState(Qt.Checked if all_checked_now else Qt.Unchecked)

        self.model().blockSignals(False)
        self.update()  # Trigger paintEvent to update the button text

    def getCheckedItems(self):
        checked = []
        for i in range(1, self.model().rowCount()):
            item = self.model().item(i)
            if item:
                state = item.checkState()
                if state == Qt.Checked or state == 2:
                    checked.append(item.data(Qt.UserRole))
        return checked

    def updateText(self):
        # Deprecated: Legacy support for external calls. Rendering is now automatic.
        self.update()