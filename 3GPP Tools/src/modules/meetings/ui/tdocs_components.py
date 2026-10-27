# --- File: src/modules/meetings/ui/tdocs_components.py ---
import logging
from pathlib import Path
from PyQt5.QtCore import pyqtSignal, QEvent, Qt, QPoint
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QPalette, QColor, QBrush, QCursor
from PyQt5.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListView,
    QMenu,
    QMessageBox,
    QPushButton,
    QStyle,
    QStyleOptionComboBox,
    QStylePainter,
    QToolTip,
    QWidget,
)


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


def populate_slot_word_menu(menu: QMenu, slot: str, parent_widget: QWidget = None):
    """
    Populates a QMenu with options to stage the currently active Word document,
    any open Word document from the running Word session, or a local file into ``slot`` ('A' or 'B').
    """
    from modules.meetings.core.compare_manager import ComparisonManager
    from modules.word_tools.core.document_opener import list_open_word_documents

    mgr = ComparisonManager.get_instance()
    slot_label = f"Slot {slot}" if slot in ("A", "B") else "Cart"
    role_hint = "Original / Baseline" if slot == "A" else ("Revised / Target" if slot == "B" else "Next Slot")

    act_active = menu.addAction(f"🪟 Use Active Word Document ({slot_label}: {role_hint})")
    act_active.setToolTip(
        "Stage the document currently active in Microsoft Word (automatically snapshots unsaved changes)."
    )

    def _on_stage_active():
        doc = mgr.stage_active_word_document(slot=slot if slot in ("A", "B") else None)
        if not doc:
            QMessageBox.warning(
                parent_widget,
                "No Active Word Document",
                "No active Microsoft Word document could be detected.\n\n"
                "Please ensure Microsoft Word is running and a document is open.",
            )
        else:
            QToolTip.showText(
                QCursor.pos(),
                f"🪟 Staged '{doc.get('name')}' into {slot_label}",
                parent_widget,
            )

    act_active.triggered.connect(_on_stage_active)

    open_docs = list_open_word_documents()
    if open_docs:
        open_sub = menu.addMenu(f"🖥️ Select Open Word Document ({len(open_docs)})")
        open_sub.setToolTipsVisible(True)
        for d in open_docs:
            name = d.get("name", "Untitled")
            full_name = d.get("full_name") or name
            is_act = d.get("is_active", False)
            is_mod = d.get("modified", False)
            prefix = "🟢 [Active] " if is_act else "📄 "
            suffix = " (unsaved edits)" if is_mod else ""
            act_doc = open_sub.addAction(f"{prefix}{name}{suffix}")
            act_doc.setToolTip(full_name)
            act_doc.triggered.connect(
                lambda _, ident=full_name, nm=name: [
                    mgr.stage_open_word_document(ident, slot=slot if slot in ("A", "B") else None),
                    QToolTip.showText(QCursor.pos(), f"🖥️ Staged '{nm}' into {slot_label}", parent_widget),
                ]
            )

    act_browse = menu.addAction(f"📂 Browse Local Word Document for {slot_label}...")
    act_browse.setToolTip("Select a local .docx or .doc file from disk.")

    def _on_browse_local():
        file_path, _ = QFileDialog.getOpenFileName(
            parent_widget,
            f"Select Word Document for {slot_label}",
            "",
            "Word Documents (*.docx *.doc *.docm);;All Files (*.*)",
        )
        if file_path:
            p = Path(file_path)
            mgr.add_to_cart(p.name, str(p), slot=slot if slot in ("A", "B") else None)

    act_browse.triggered.connect(_on_browse_local)

    current_slot_data = mgr.slot_a if slot == "A" else (mgr.slot_b if slot == "B" else None)
    if current_slot_data:
        menu.addSeparator()
        act_clear_slot = menu.addAction(f"🗑️ Clear {slot_label} ({current_slot_data.get('name', '')})")
        act_clear_slot.triggered.connect(lambda: mgr.clear_slot(slot))


class ComparisonCartWidget(QFrame):
    """
    Reusable, compact Comparison Cart bar for both MeetingsTab and TDocsWindow.
    Synchronizes across windows via ComparisonManager, fixes button padding/disabled
    quirks, and lets users stage the currently active/open Word document directly into
    Slot A or Slot B with 1 click.
    """

    def __init__(self, parent=None, show_launcher_when_empty: bool = True):
        super().__init__(parent)
        self.setObjectName("cartFrame")
        self.show_launcher_when_empty = show_launcher_when_empty
        self.cmp_thread = None

        self.setStyleSheet("""
            QFrame#cartFrame {
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 0px;
            }
            QFrame#cartFrame QLabel {
                color: #1E293B;
                border: none;
                background: transparent;
                font-size: 11px;
                padding: 0px 2px;
            }
            QPushButton#cartSlotEmptyBtn {
                background-color: #FFFFFF;
                color: #64748B;
                border: 1px dashed #94A3B8;
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 11px;
                font-weight: 500;
                min-height: 20px;
                max-height: 22px;
            }
            QPushButton#cartSlotEmptyBtn:hover {
                background-color: #EFF6FF;
                color: #1E5C99;
                border: 1px solid #3B82F6;
            }
            QPushButton#cartSlotFilledBtn {
                background-color: #EBF3FC;
                color: #1E40AF;
                border: 1px solid #93C5FD;
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 11px;
                font-weight: bold;
                min-height: 20px;
                max-height: 22px;
            }
            QPushButton#cartSlotFilledBtn:hover {
                background-color: #DBEAFE;
                border-color: #2563EB;
            }
            QPushButton#cartSwapBtn {
                background-color: #F1F5F9;
                color: #475569;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 1px 6px;
                font-size: 11px;
                font-weight: bold;
                min-height: 20px;
                max-height: 22px;
            }
            QPushButton#cartSwapBtn:hover {
                background-color: #E2E8F0;
                color: #0F172A;
                border-color: #94A3B8;
            }
            QPushButton#cartCompareBtn {
                background-color: #1E5C99;
                color: #FFFFFF;
                border: 1px solid #15426E;
                border-radius: 4px;
                padding: 2px 10px;
                font-size: 11px;
                font-weight: bold;
                min-height: 20px;
                max-height: 22px;
            }
            QPushButton#cartCompareBtn:hover {
                background-color: #15426E;
            }
            QPushButton#cartCompareBtn:disabled {
                background-color: #E2E8F0;
                color: #94A3B8;
                border: 1px solid #CBD5E1;
            }
            QPushButton#cartSecondaryBtn {
                background-color: #FFFFFF;
                color: #334155;
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 11px;
                font-weight: 600;
                min-height: 20px;
                max-height: 22px;
            }
            QPushButton#cartSecondaryBtn:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
                color: #0F172A;
            }
        """)

        from modules.meetings.core.compare_manager import ComparisonManager

        self.mgr = ComparisonManager.get_instance()
        self.launcher_btn = None
        self._setup_ui()
        self.mgr.cart_updated.connect(self.update_cart_ui)
        self.update_cart_ui(self.mgr.slot_a, self.mgr.slot_b)

    def create_launcher_button(self, parent: QWidget = None) -> QPushButton:
        """
        Creates a compact companion button that remains visible when the cart is empty,
        allowing the user to stage the currently active/open Word document into Slot A
        to start a comparison session at any time.
        """
        btn = QPushButton("⚖️ Stage Active Word ▾", parent or self.parentWidget())
        btn.setCursor(Qt.PointingHandCursor)
        btn.setFixedHeight(24)
        btn.setToolTip(
            "Stage your currently active Microsoft Word document (or any open Word document) into the Comparison Cart."
        )
        btn.setStyleSheet("""
            QPushButton {
                background-color: #F8FAFC;
                color: #475569;
                border: 1px solid #CBD5E1;
                border-radius: 5px;
                padding: 2px 10px;
                font-size: 11px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #EFF6FF;
                color: #1E5C99;
                border-color: #93C5FD;
            }
        """)

        def _show_launcher_menu():
            menu = QMenu(btn)
            menu.setToolTipsVisible(True)
            populate_slot_word_menu(menu, "A", self)
            menu.exec_(btn.mapToGlobal(QPoint(0, -menu.sizeHint().height())))

        btn.clicked.connect(_show_launcher_menu)
        self.launcher_btn = btn
        has_items = bool(self.mgr.slot_a or self.mgr.slot_b)
        btn.setVisible(not has_items)
        return btn

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 3, 8, 3)
        layout.setSpacing(5)

        self.lbl_cart_title = QLabel("<b>⚖️ Cart:</b>")
        self.lbl_cart_title.setToolTip(
            "Word Comparison Cart: Stage any two TDocs or open Word documents into Slot A (Original) and Slot B (Revised)."
        )
        layout.addWidget(self.lbl_cart_title)

        self.btn_slot_a = QPushButton("[Slot A: Empty ▾]")
        self.btn_slot_a.setObjectName("cartSlotEmptyBtn")
        self.btn_slot_a.setCursor(Qt.PointingHandCursor)
        self.btn_slot_a.clicked.connect(lambda: self._show_slot_menu("A", self.btn_slot_a))
        layout.addWidget(self.btn_slot_a)

        self.btn_swap = QPushButton("⇄")
        self.btn_swap.setObjectName("cartSwapBtn")
        self.btn_swap.setCursor(Qt.PointingHandCursor)
        self.btn_swap.setToolTip("Swap Slot A (Original) and Slot B (Revised) to invert the redline direction.")
        self.btn_swap.clicked.connect(self.mgr.swap_slots)
        layout.addWidget(self.btn_swap)

        self.btn_slot_b = QPushButton("[Slot B: Empty ▾]")
        self.btn_slot_b.setObjectName("cartSlotEmptyBtn")
        self.btn_slot_b.setCursor(Qt.PointingHandCursor)
        self.btn_slot_b.clicked.connect(lambda: self._show_slot_menu("B", self.btn_slot_b))
        layout.addWidget(self.btn_slot_b)

        self.btn_active_word = QPushButton("🪟 Active Word")
        self.btn_active_word.setObjectName("cartSecondaryBtn")
        self.btn_active_word.setCursor(Qt.PointingHandCursor)
        self.btn_active_word.setToolTip(
            "Stage the currently active Microsoft Word document into the next available cart slot (or right-click to choose slot)."
        )
        self.btn_active_word.clicked.connect(self._on_quick_stage_active_word)
        layout.addWidget(self.btn_active_word)

        self.btn_compare = QPushButton("⚖️ Compare in Word")
        self.btn_compare.setObjectName("cartCompareBtn")
        self.btn_compare.setCursor(Qt.PointingHandCursor)
        self.btn_compare.setEnabled(False)
        self.btn_compare.setToolTip(
            "Generate a native Microsoft Word Track Changes redline comparing Slot A (Original) vs Slot B (Revised)."
        )
        self.btn_compare.clicked.connect(self.run_comparison)
        layout.addWidget(self.btn_compare)

        self.btn_clear_cart = QPushButton("✖ Clear")
        self.btn_clear_cart.setObjectName("cartSecondaryBtn")
        self.btn_clear_cart.setCursor(Qt.PointingHandCursor)
        self.btn_clear_cart.setToolTip("Clear all items from the comparison cart.")
        self.btn_clear_cart.clicked.connect(self.mgr.clear_cart)
        layout.addWidget(self.btn_clear_cart)

        # Backwards-compatibility aliases
        self.lbl_slot_a = self.btn_slot_a
        self.lbl_slot_b = self.btn_slot_b

    def _show_slot_menu(self, slot: str, anchor_btn: QPushButton):
        menu = QMenu(self)
        menu.setToolTipsVisible(True)
        populate_slot_word_menu(menu, slot, self)
        menu.exec_(anchor_btn.mapToGlobal(QPoint(0, -menu.sizeHint().height())))

    def _on_quick_stage_active_word(self):
        target_slot = "A" if not self.mgr.slot_a else ("B" if not self.mgr.slot_b else None)
        doc = self.mgr.stage_active_word_document(slot=target_slot)
        if not doc:
            QMessageBox.warning(
                self,
                "No Active Word Document",
                "No active Microsoft Word document could be detected.\n\n"
                "Please ensure Microsoft Word is running and a document is open on your screen.",
            )
        else:
            QToolTip.showText(
                QCursor.pos(),
                f"🪟 Staged '{doc.get('name')}' into Comparison Cart",
                self,
            )

    @staticmethod
    def _elide_label(text: str, max_len: int = 24) -> str:
        text = str(text or "").strip()
        if len(text) <= max_len:
            return text
        return text[: max_len - 1] + "…"

    def update_cart_ui(self, slot_a: dict, slot_b: dict):
        has_items = bool(slot_a or slot_b)
        self.setVisible(has_items)
        if self.launcher_btn is not None:
            self.launcher_btn.setVisible(not has_items)

        if slot_a:
            raw_a = slot_a.get("name", "Doc A")
            self.btn_slot_a.setText(f"A: {self._elide_label(raw_a)} ▾")
            self.btn_slot_a.setObjectName("cartSlotFilledBtn")
            self.btn_slot_a.setToolTip(
                f"Slot A (Original / Baseline): {raw_a}\nPath: {slot_a.get('path', '')}\n\nClick to change or clear Slot A."
            )
        else:
            self.btn_slot_a.setText("[Slot A: Empty ▾]")
            self.btn_slot_a.setObjectName("cartSlotEmptyBtn")
            self.btn_slot_a.setToolTip(
                "Slot A (Original / Baseline) is empty.\nClick to stage the Active Word Document, an Open Word Document, or a Local File."
            )

        if slot_b:
            raw_b = slot_b.get("name", "Doc B")
            self.btn_slot_b.setText(f"B: {self._elide_label(raw_b)} ▾")
            self.btn_slot_b.setObjectName("cartSlotFilledBtn")
            self.btn_slot_b.setToolTip(
                f"Slot B (Revised / Target): {raw_b}\nPath: {slot_b.get('path', '')}\n\nClick to change or clear Slot B."
            )
        else:
            self.btn_slot_b.setText("[Slot B: Empty ▾]")
            self.btn_slot_b.setObjectName("cartSlotEmptyBtn")
            self.btn_slot_b.setToolTip(
                "Slot B (Revised / Target) is empty.\nClick to stage the Active Word Document, an Open Word Document, or a Local File."
            )

        # Re-polish buttons after objectName change so Qt applies the updated selector
        for btn in (self.btn_slot_a, self.btn_slot_b):
            btn.style().unpolish(btn)
            btn.style().polish(btn)

        both_ready = bool(slot_a and slot_b)
        self.btn_compare.setEnabled(both_ready)
        if both_ready:
            self.btn_compare.setToolTip(
                f"Compare '{slot_a.get('name')}' (Original) vs '{slot_b.get('name')}' (Revised) in Microsoft Word."
            )
        else:
            self.btn_compare.setToolTip(
                "Populate both Slot A and Slot B to enable Microsoft Word redline comparison."
            )

    def run_comparison(self):
        from modules.word_tools.core.word_comparator import WordComparatorThread

        path_a, path_b = self.mgr.resolve_comparison_paths()
        if not path_a or not path_b:
            return

        self.btn_compare.setText("⏳ Comparing...")
        self.btn_compare.setEnabled(False)

        self.cmp_thread = WordComparatorThread(path_a, path_b)
        self.cmp_thread.finished.connect(self._on_comparison_finished)
        self.cmp_thread.start()

    def _on_comparison_finished(self):
        self.btn_compare.setText("⚖️ Compare in Word")
        self.btn_compare.setEnabled(bool(self.mgr.slot_a and self.mgr.slot_b))
