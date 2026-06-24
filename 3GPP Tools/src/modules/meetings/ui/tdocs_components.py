from PyQt5.QtCore import pyqtSignal, QEvent, Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtWidgets import QComboBox


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

    def updateItems(self, items):
        previously_checked = set(self.getCheckedItems())
        was_all_checked = (self.model().item(0).checkState() == Qt.Checked) if self.model().rowCount() > 0 else True

        self.model().blockSignals(True)
        self.model().clear()

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

        self.model().item(0).setCheckState(Qt.Checked if all_checked_now else Qt.Unchecked)
        self.model().blockSignals(False)
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
            self.lineEdit().setText(f"{self.title}: {checked[0] if checked[0] else '(Empty)'}")
        else:
            self.lineEdit().setText(f"{self.title}: {len(checked)} selected")
