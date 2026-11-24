"""Docking, split layout, and interactive prompt/evidence handlers for AssistantTab."""

from __future__ import annotations

import logging

from PyQt5.QtCore import Qt, QTimer, QUrl
from PyQt5.QtWidgets import QApplication, QTreeWidgetItem

logger = logging.getLogger(__name__)


class AssistantDockingLayoutMixin:
    """Handles dock-mode reparenting, splitter toggling, anchors, and evidence selection."""

    @property
    def is_docked(self) -> bool:
        return self._is_docked

    def set_cockpit_visible_in_dock(self, visible: bool) -> None:
        """Show or collapse the lower Sources/Activity drawer while in compact dock mode."""
        visible = bool(visible)
        if hasattr(self, "dock_sources_toggle_btn"):
            self.dock_sources_toggle_btn.blockSignals(True)
            self.dock_sources_toggle_btn.setChecked(visible)
            self.dock_sources_toggle_btn.blockSignals(False)
        self.cockpit_tabs.setVisible(visible)
        if self._is_docked:
            self.main_splitter.setOrientation(Qt.Vertical)
            self.main_splitter.setSizes([380, 240] if visible else [600, 0])

    def set_docked_mode(self, docked: bool) -> None:
        """Toggle compact layout when Assistant body is hosted inside RightSidePanel."""
        docked = bool(docked)
        changed = docked != self._is_docked
        self._is_docked = docked

        self.ribbon_card.setVisible(not docked)
        self.dock_compact_strip.setVisible(docked)
        self.docked_placeholder_card.setVisible(docked)

        if docked:
            self._body_layout.setContentsMargins(S_XS(), S_XS(), S_XS(), S_XS())
            self.main_splitter.setOrientation(Qt.Vertical)
            self.set_cockpit_visible_in_dock(False)
        else:
            self._body_layout.setContentsMargins(0, 0, 0, 0)
            self.cockpit_tabs.setVisible(True)
            self._toggle_layout(self.layout_toggle_btn.isChecked())

        if changed:
            self.docked_changed.emit(docked)

    def restore_body_into_tab(self) -> None:
        """Re-attach ``body_container`` into this tab and exit compact dock mode."""
        if self._host_layout.indexOf(self.body_container) < 0:
            self._host_layout.insertWidget(0, self.body_container, 1)
        self.body_container.setVisible(True)
        self.set_docked_mode(False)

    def _set_status_text(self, text: str) -> None:
        self.status_label.setText(text)
        if hasattr(self, "dock_status_label"):
            self.dock_status_label.setText(text)

    def _update_source_counts(self) -> None:
        count = len(self._evidence)
        label = f"📚 Sources ({count})"
        self.cockpit_tabs.setTabText(0, label)
        if hasattr(self, "dock_sources_toggle_btn"):
            self.dock_sources_toggle_btn.setText(label)

    def _toggle_layout(self, checked: bool) -> None:
        """Toggle between Side-by-Side (Horizontal) and Bottom Split (Vertical)."""
        if checked:
            self.main_splitter.setOrientation(Qt.Horizontal)
            self.main_splitter.setSizes([680, 420])
            self.layout_toggle_btn.setText("⬒ Bottom Split")
            self.layout_toggle_btn.setToolTip("Switch to stacked top/bottom layout")
        else:
            self.main_splitter.setOrientation(Qt.Vertical)
            self.main_splitter.setSizes([450, 280])
            self.layout_toggle_btn.setText("◫ Side Drawer")
            self.layout_toggle_btn.setToolTip("Switch to side-by-side drawer layout")

    def _on_prompt_changed(self) -> None:
        text = self.question_input.toPlainText()
        count = len(text)
        self.char_count_label.setText(f"{count} char" + ("s" if count != 1 else ""))
        self.clear_input_btn.setEnabled(count > 0)

    def _on_starter_chip_selected(self, prompt: str) -> None:
        self.question_input.setPlainText(prompt)
        self._on_send()

    def _on_anchor_clicked(self, url: QUrl) -> None:
        raw_url = url.toString() if hasattr(url, "toString") else str(url)
        if raw_url.startswith("evidence:"):
            eid = raw_url.split(":", 1)[1].strip()
            self._select_evidence_by_id(eid)
        elif raw_url.startswith("copy_msg:"):
            try:
                idx = int(raw_url.split(":", 1)[1].strip())
                if 0 <= idx < len(self._conversation):
                    text = self._conversation[idx].get("content", "")
                    QApplication.clipboard().setText(text)
                    self.status_label.setText("Answer copied to clipboard!")
                    QTimer.singleShot(
                        2000,
                        lambda: self.status_label.setText(
                            "Ready" if not (self._worker and self._worker.isRunning()) else "Researching…"
                        ),
                    )
            except Exception as exc:
                logger.error("Failed to copy message: %s", exc)

    def _select_evidence_by_id(self, eid: str) -> None:
        """Highlights the evidence row in sources tree, displays it in inspector, and switches to Sources tab."""
        if self._is_docked and not self.cockpit_tabs.isVisible():
            self.set_cockpit_visible_in_dock(True)
        self.cockpit_tabs.setCurrentIndex(0)  # Tab 0 is Sources
        for i in range(self.sources_tree.topLevelItemCount()):
            item = self.sources_tree.topLevelItem(i)
            if item.data(0, Qt.UserRole) == eid:
                self.sources_tree.setCurrentItem(item)
                self.sources_tree.scrollToItem(item)
                ev = self._evidence.get(eid)
                self.evidence_inspector.set_evidence(ev)
                break

    def _on_source_tree_item_clicked(self, item: QTreeWidgetItem, column: int = 0) -> None:
        evidence_id = str(item.data(0, Qt.UserRole) or "")
        evidence = self._evidence.get(evidence_id)
        self.evidence_inspector.set_evidence(evidence)


def S_XS() -> int:
    """Return spacing token S.XS without importing theme into top-level cycles."""
    from core.ui.theme import S

    return S.XS
