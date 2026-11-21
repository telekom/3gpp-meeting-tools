"""Scrollable container of specification cards and legacy tree-compat wrappers."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QFrame, QScrollArea, QVBoxLayout, QWidget

from core.ui.theme import C
from .spec_card import (
    NASReleasePillWidget,
    NASSpecCardWidget,
    _extract_major_release,
)


class _NASCardCompatChild:
    def __init__(self, pill: NASReleasePillWidget):
        self.pill = pill

    def text(self, col: int = 0) -> str:
        rel = _extract_major_release(self.pill.version)
        return f"{rel} · v{self.pill.version}"

    def checkState(self, col: int = 0) -> Qt.CheckState:
        return Qt.Checked if self.pill.isChecked() else Qt.Unchecked


class _NASCardCompatSpec:
    def __init__(self, card: NASSpecCardWidget):
        self.card = card

    def text(self, col: int = 0) -> str:
        wg = self.card.meta.get("wg", "3GPP")
        acr = self.card.meta.get("acronym", "")
        return f"[{wg}] TS {self.card.spec_number} — {acr}".strip()

    def childCount(self) -> int:
        return len(self.card.pills)

    def child(self, idx: int) -> _NASCardCompatChild:
        return _NASCardCompatChild(self.card.pills[idx])


class NASVersionCardsWidget(QScrollArea):
    """
    Modern Spec-Search-aligned release pills selector for Protocols (`NASTab`).
    Replaces the deep vertical tree with compact specification cards and wrapping release pills.
    """

    selection_changed = pyqtSignal()
    reload_version_requested = pyqtSignal(str, str)
    reload_spec_requested = pyqtSignal(str, int)
    delete_version_requested = pyqtSignal(str, str)
    delete_spec_requested = pyqtSignal(str, int)
    structure_changed = pyqtSignal()
    open_spec_requested = pyqtSignal(str, str, str, str)
    show_in_specs_requested = pyqtSignal(str, str)
    search_in_spec_search_requested = pyqtSignal(str, str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._updating_checks: bool = False
        self._cards: Dict[str, NASSpecCardWidget] = {}

        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setStyleSheet(f"""
            QScrollArea {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background-color: {C.BG_SURFACE};
            }}
        """)

        self.container = QWidget()
        self.container.setObjectName("nasCardsContainer")
        self.container.setStyleSheet(
            f"QWidget#nasCardsContainer {{ background-color: {C.BG_SURFACE}; }}"
        )
        self.cards_layout = QVBoxLayout(self.container)
        self.cards_layout.setContentsMargins(4, 4, 4, 4)
        self.cards_layout.setSpacing(6)
        self.cards_layout.addStretch()
        self.setWidget(self.container)

    def populate(
        self,
        versions: List[Dict[str, Any]],
        saved_checked: Optional[List[Dict[str, Any]]] = None,
        saved_collapsed: Optional[Set[str]] = None,
    ):
        self._updating_checks = True
        self.clear()
        if not versions:
            self._updating_checks = False
            self.selection_changed.emit()
            return

        saved_tuples: Optional[Set[Tuple[str, str]]] = None
        if saved_checked is not None:
            saved_tuples = {(item.get("spec_number"), item.get("version")) for item in saved_checked if item}

        specs_map: Dict[str, List[Dict[str, Any]]] = {}
        for v in versions:
            specs_map.setdefault(v["spec_number"], []).append(v)

        # Remove trailing stretch before adding cards
        count = self.cards_layout.count()
        if count > 0 and self.cards_layout.itemAt(count - 1).spacerItem():
            self.cards_layout.removeItem(self.cards_layout.itemAt(count - 1))

        for spec_num in sorted(specs_map.keys()):
            card = NASSpecCardWidget(spec_num, specs_map[spec_num], self.container)
            card.selection_changed.connect(self._on_card_selection_changed)
            card.collapsed_changed.connect(self.structure_changed.emit)
            card.reload_version_requested.connect(self.reload_version_requested.emit)
            card.reload_spec_requested.connect(self.reload_spec_requested.emit)
            card.delete_version_requested.connect(self.delete_version_requested.emit)
            card.delete_spec_requested.connect(self.delete_spec_requested.emit)
            card.open_spec_requested.connect(self.open_spec_requested.emit)
            card.show_in_specs_requested.connect(self.show_in_specs_requested.emit)
            card.search_in_spec_search_requested.connect(self.search_in_spec_search_requested.emit)

            card._updating_checks = True
            for pill in card.pills:
                if saved_tuples is not None:
                    is_chk = (str(spec_num), str(pill.version)) in saved_tuples
                else:
                    is_chk = True
                pill.setChecked(is_chk)
            card._updating_checks = False
            card._update_header_state()
            if saved_collapsed and spec_num in saved_collapsed:
                card.set_collapsed(True, emit_signal=False)

            self._cards[spec_num] = card
            self.cards_layout.addWidget(card)

        self.cards_layout.addStretch()
        self._updating_checks = False
        self.selection_changed.emit()

    def clear(self):
        while self.cards_layout.count():
            item = self.cards_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._cards.clear()
        self.cards_layout.addStretch()

    def _on_card_selection_changed(self):
        if self._updating_checks:
            return
        self.selection_changed.emit()
        self.structure_changed.emit()

    def filter_tree(self, query: str):
        q = str(query or "").strip().lower()
        if not q:
            for card in self._cards.values():
                card.setVisible(True)
            return
        for spec_num, card in self._cards.items():
            hay = " ".join([
                spec_num,
                str(card.meta.get("wg", "")),
                str(card.meta.get("acronym", "")),
                str(card.meta.get("title", "")),
                " ".join(p.text() for p in card.pills),
            ]).lower()
            card.setVisible(all(tok in hay for tok in q.split()))

    def select_all(self):
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                p.setChecked(True)
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def deselect_all(self):
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                p.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def select_none(self):
        self.deselect_all()

    def select_latest_only(self):
        """Selects only the single newest version for each indexed protocol specification."""
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            for idx, p in enumerate(card.pills):
                p.setChecked(idx == 0)
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def select_latest_per_release(self):
        """Selects the newest version of each major Release per specification for evolution diffs."""
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            seen = set()
            for p in card.pills:
                rel = _extract_major_release(p.version)
                if rel not in seen:
                    seen.add(rel)
                    p.setChecked(True)
                else:
                    p.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def select_releases_min(self, min_rel_num: int):
        """Selects only releases with major release >= min_rel_num (e.g. Rel-18+)."""
        self._updating_checks = True
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                rel = _extract_major_release(p.version)
                m = re.search(r"\d+", rel)
                p.setChecked(bool(m and int(m.group()) >= min_rel_num))
            card._updating_checks = False
            card._update_header_state()
        self._updating_checks = False
        self.selection_changed.emit()
        self.structure_changed.emit()

    def get_selected_version_ids(self) -> List[int]:
        ids: List[int] = []
        for card in self._cards.values():
            for p in card.pills:
                if p.isChecked() and p.version_id is not None:
                    ids.append(p.version_id)
        return ids

    def get_checked_versions_info(self) -> List[Dict[str, Any]]:
        checked: List[Dict[str, Any]] = []
        for spec_num in sorted(self._cards.keys()):
            card = self._cards[spec_num]
            for p in card.pills:
                if p.isChecked():
                    checked.append({"spec_number": spec_num, "version": p.version})
        return checked

    def get_all_versions_for_spec(self, spec_number: str) -> List[Dict[str, Any]]:
        card = self._cards.get(spec_number)
        if not card:
            return []
        return [{"spec_number": card.spec_number, "version": p.version} for p in card.pills]

    def get_collapsed_spec_numbers(self) -> List[str]:
        return [spec_num for spec_num, card in self._cards.items() if card.is_collapsed()]

    def topLevelItemCount(self) -> int:
        return 1 + len(self._cards)

    def topLevelItem(self, idx: int):
        keys = sorted(self._cards.keys())
        card_idx = idx - 1
        if 0 <= card_idx < len(keys):
            return _NASCardCompatSpec(self._cards[keys[card_idx]])
        return None
