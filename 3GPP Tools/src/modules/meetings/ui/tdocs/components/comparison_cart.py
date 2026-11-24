# --- File: src/modules/meetings/ui/tdocs/components/comparison_cart.py ---
"""ComparisonCartWidget, Word slot menu helpers, and rich tooltip formatter."""

from pathlib import Path

from PyQt5.QtCore import QPoint, Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QToolTip,
    QWidget,
)

from core.ui.font_scale import scale_px
from core.ui.theme import C


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

        self.setStyleSheet(f"""
            QFrame#cartFrame {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                padding: 0px;
            }}
            QFrame#cartFrame QLabel {{
                color: {C.TEXT_STRONG};
                border: none;
                background: transparent;
                font-size: 11px;
                padding: 0px 2px;
            }}
            QPushButton#cartSlotEmptyBtn {{
                background-color: {C.BG_CARD};
                color: {C.TEXT_MUTED};
                border: 1px dashed {C.TEXT_FAINT};
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 11px;
                font-weight: 500;
                min-height: 20px;
                max-height: 22px;
            }}
            QPushButton#cartSlotEmptyBtn:hover {{
                background-color: {C.INFO_SOFT};
                color: {C.ACCENT_PRIMARY};
                border: 1px solid {C.INFO_ACCENT};
            }}
            QPushButton#cartSlotFilledBtn {{
                background-color: {C.ACCENT_SOFT};
                color: {C.INFO_DARK};
                border: 1px solid {C.INFO_BORDER_MEDIUM};
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 11px;
                font-weight: bold;
                min-height: 20px;
                max-height: 22px;
            }}
            QPushButton#cartSlotFilledBtn:hover {{
                background-color: {C.INFO_SOFT_STRONG};
                border-color: {C.ACCENT_ALT};
            }}
            QPushButton#cartSwapBtn {{
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_SECONDARY};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 1px 6px;
                font-size: 11px;
                font-weight: bold;
                min-height: 20px;
                max-height: 22px;
            }}
            QPushButton#cartSwapBtn:hover {{
                background-color: {C.BORDER_SUBTLE};
                color: {C.TEXT_PRIMARY};
                border-color: {C.TEXT_FAINT};
            }}
            QPushButton#cartCompareBtn {{
                background-color: {C.ACCENT_PRIMARY};
                color: {C.BG_CARD};
                border: 1px solid {C.ACCENT_HOVER};
                border-radius: 4px;
                padding: 2px 10px;
                font-size: 11px;
                font-weight: bold;
                min-height: 20px;
                max-height: 22px;
            }}
            QPushButton#cartCompareBtn:hover {{
                background-color: {C.ACCENT_HOVER};
            }}
            QPushButton#cartCompareBtn:disabled {{
                background-color: {C.BORDER_SUBTLE};
                color: {C.TEXT_FAINT};
                border: 1px solid {C.BORDER_STRONG};
            }}
            QPushButton#cartSecondaryBtn {{
                background-color: {C.BG_CARD};
                color: {C.TEXT_BODY};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 2px 8px;
                font-size: 11px;
                font-weight: 600;
                min-height: 20px;
                max-height: 22px;
            }}
            QPushButton#cartSecondaryBtn:hover {{
                background-color: {C.BG_CANVAS};
                border-color: {C.TEXT_FAINT};
                color: {C.TEXT_PRIMARY};
            }}
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
        btn.setMinimumHeight(scale_px(24))
        btn.setToolTip(
            "Stage your currently active Microsoft Word document (or any open Word document) into the Comparison Cart."
        )
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_SECONDARY};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 5px;
                padding: 2px 10px;
                font-size: 11px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {C.INFO_SOFT};
                color: {C.ACCENT_PRIMARY};
                border-color: {C.INFO_BORDER_MEDIUM};
            }}
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


def make_rich_tooltip(title: str, description: str, tip: str = "", shortcut: str = "") -> str:
    """Creates a beautifully styled, structured multi-line HTML card for widget tooltips."""
    import html

    shortcut_html = (
        f' <span style="color: {C.TEXT_MUTED}; font-size: 11px; font-weight: normal;">({html.escape(shortcut)})</span>'
        if shortcut
        else ""
    )
    tip_html = (
        f"""
  <div style="margin-top: 6px; padding-top: 4px; border-top: 1px solid {C.BORDER_SUBTLE}; font-size: 11px; color: {C.TEXT_SECONDARY};">
    💡 <b>Tip:</b> {html.escape(tip)}
  </div>"""
        if tip
        else ""
    )

    return f"""<div style="font-family: -apple-system, Segoe UI, sans-serif; font-size: 12px; line-height: 1.4; color: {C.TEXT_STRONG}; max-width: 330px;">
  <div style="font-weight: bold; font-size: 13px; color: {C.TEXT_PRIMARY}; margin-bottom: 3px;">
    {title}{shortcut_html}
  </div>
  <div style="color: {C.TEXT_BODY};">{html.escape(description)}</div>{tip_html}
</div>"""
