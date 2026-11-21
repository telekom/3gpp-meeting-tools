"""Concept B Release Pills Flow container and legacy tree compatibility adapters."""

import re
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QFrame, QScrollArea, QVBoxLayout, QWidget

from core.ui.theme import C
from modules.spec_search.core.spec_search_db import extract_major_release

from .spec_card import ReleasePillWidget, SpecCardWidget


class _TreeCompatChildItem:
    def __init__(self, pill: ReleasePillWidget):
        self.pill = pill

    def text(self, col: int = 0) -> str:
        date_lbl = f" ({self.pill.release_date})" if self.pill.release_date else ""
        return f"v{self.pill.version}{date_lbl}"

    def toolTip(self, col: int = 0) -> str:
        return self.pill.toolTip()

    def data(self, col: int, role: int = Qt.UserRole) -> Any:
        return {
            "type": "version",
            "id": self.pill.version_id,
            "spec_number": self.pill.spec_number,
            "version": self.pill.version,
            "release_date": self.pill.release_date,
        }

    def checkState(self, col: int = 0) -> Qt.CheckState:
        return Qt.Checked if self.pill.isChecked() else Qt.Unchecked

    def setCheckState(self, col: int, state: Qt.CheckState):
        self.pill.setChecked(state == Qt.Checked)


class _TreeCompatSpecItem:
    def __init__(self, card: SpecCardWidget):
        self.card = card

    def text(self, col: int = 0) -> str:
        t = self.card.raw_title
        if t:
            clean = re.sub(r'^(NR;\s*|LTE;\s*|Universal Mobile Telecommunications System \(UMTS\);\s*)', '', t, flags=re.IGNORECASE).strip()
            parts = [p.strip() for p in clean.split(';') if p.strip()]
            display_title = parts[0] if parts else clean
            if len(display_title) > 40:
                display_title = display_title[:39].rstrip() + "…"
            return f"TS {self.card.spec_number}: {display_title}  ({len(self.card.pills)})"
        return f"TS {self.card.spec_number}  ({len(self.card.pills)})"

    def toolTip(self, col: int = 0) -> str:
        if self.card.raw_title:
            return f"TS {self.card.spec_number}: {self.card.raw_title}\n{len(self.card.pills)} indexed release(s)"
        return f"TS {self.card.spec_number}\n{len(self.card.pills)} indexed release(s)"

    def data(self, col: int, role: int = Qt.UserRole) -> Any:
        return {
            "type": "spec",
            "spec_number": self.card.spec_number,
            "title": self.card.raw_title,
        }

    def isHidden(self) -> bool:
        return self.card.isHidden()

    def setHidden(self, val: bool):
        self.card.setHidden(val)

    def setExpanded(self, val: bool):
        pass

    def childCount(self) -> int:
        return len(self.card.pills)

    def child(self, idx: int) -> _TreeCompatChildItem:
        return _TreeCompatChildItem(self.card.pills[idx])

    def checkState(self, col: int = 0) -> Qt.CheckState:
        return self.card.chk_all.checkState()

    def setCheckState(self, col: int, state: Qt.CheckState):
        if state == Qt.Checked:
            self.card.select_all()
        else:
            self.card.deselect_all()


class _TreeCompatRootItem:
    def __init__(self, widget: "SpecSearchVersionCardsWidget"):
        self.widget = widget

    def text(self, col: int = 0) -> str:
        return "All Indexed Specifications"

    def data(self, col: int, role: int = Qt.UserRole) -> Any:
        return {"type": "all"}

    def isHidden(self) -> bool:
        return False

    def setHidden(self, val: bool):
        pass

    def checkState(self, col: int = 0) -> Qt.CheckState:
        total = sum(len(c.pills) for c in self.widget._cards.values())
        checked = len(self.widget.get_selected_version_ids())
        if checked == total and total > 0:
            return Qt.Checked
        elif checked == 0:
            return Qt.Unchecked
        return Qt.PartiallyChecked

    def setCheckState(self, col: int, state: Qt.CheckState):
        if state == Qt.Checked:
            self.widget.select_all()
        else:
            self.widget.deselect_all()


class SpecSearchVersionCardsWidget(QScrollArea):
    """
    Modern Concept B: Release Pills Flow Widget.
    Replaces the deep vertical tree with clean, compact specification cards featuring
    interactive release pills, 1-click status badges, and quick-filtering.
    Includes full backward-compatibility adapters for legacy tree tests and callers.
    """
    selection_changed = pyqtSignal()
    delete_version_requested = pyqtSignal(str, str)
    delete_spec_requested = pyqtSignal(str, int)
    prune_spec_requested = pyqtSignal(str)
    smart_import_requested = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
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
        self.container.setStyleSheet(f"background-color: {C.BG_SURFACE};")
        self.cards_layout = QVBoxLayout(self.container)
        self.cards_layout.setContentsMargins(4, 4, 4, 4)
        self.cards_layout.setSpacing(6)
        self.cards_layout.addStretch()
        self.setWidget(self.container)

        self._cards: Dict[str, SpecCardWidget] = {}
        self._all_item_compat = _TreeCompatRootItem(self)

    def populate(
        self,
        versions: List[Dict[str, Any]],
        saved_checked: Optional[List[Dict[str, Any]]] = None,
        spec_titles: Optional[Dict[str, str]] = None,
    ):
        self.clear()
        if not versions:
            self.selection_changed.emit()
            return

        saved_tuples = (
            {(str(i.get("spec_number")), str(i.get("version"))) for i in saved_checked if i}
            if saved_checked is not None
            else None
        )

        specs_map: Dict[str, List[Dict[str, Any]]] = {}
        for v in versions:
            specs_map.setdefault(v["spec_number"], []).append(v)

        # Remove stretch before adding cards
        count = self.cards_layout.count()
        if count > 0:
            last = self.cards_layout.itemAt(count - 1)
            if last.spacerItem():
                self.cards_layout.removeItem(last)

        for spec_num in sorted(specs_map.keys()):
            title = ""
            if spec_titles:
                clean_num = re.sub(r'^(TS|TR)\s*', '', spec_num.strip(), flags=re.IGNORECASE)
                title = spec_titles.get(clean_num, "")
                if not title:
                    base_num = clean_num.split('-')[0]
                    title = spec_titles.get(base_num, "")

            card = SpecCardWidget(spec_num, title, specs_map[spec_num], self.container)
            card.selection_changed.connect(self.selection_changed.emit)
            card.delete_version_requested.connect(self.delete_version_requested.emit)
            card.delete_spec_requested.connect(self.delete_spec_requested.emit)
            card.prune_spec_requested.connect(self.prune_spec_requested.emit)
            card.smart_import_requested.connect(self.smart_import_requested.emit)

            # Restore checked states
            for pill in card.pills:
                if saved_tuples is not None:
                    is_chk = (str(spec_num), str(pill.version)) in saved_tuples
                else:
                    is_chk = True
                pill.setChecked(is_chk)
            card._update_header_state()

            self._cards[spec_num] = card
            self.cards_layout.addWidget(card)

        self.cards_layout.addStretch()
        self.selection_changed.emit()

    def clear(self):
        while self.cards_layout.count():
            item = self.cards_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._cards.clear()
        self.cards_layout.addStretch()

    def filter_tree(self, query: str):
        query = query.strip().lower()
        if not query:
            for card in self._cards.values():
                card.setVisible(True)
            return

        for spec_num, card in self._cards.items():
            num_match = query in spec_num.lower()
            title_match = query in card.raw_title.lower()
            pill_match = any(query in p.version.lower() or query in p.text().lower() for p in card.pills)
            card.setVisible(num_match or title_match or pill_match)

    def select_all(self):
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                p.setChecked(True)
            card._updating_checks = False
            card._update_header_state()
        self.selection_changed.emit()

    def deselect_all(self):
        for card in self._cards.values():
            card._updating_checks = True
            for p in card.pills:
                p.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self.selection_changed.emit()

    def select_latest_per_release(self):
        """Checks the newest version of each major Release for each specification."""
        for card in self._cards.values():
            card._updating_checks = True
            seen_releases = set()
            for pill in card.pills:
                rel = extract_major_release(pill.version)
                if rel not in seen_releases:
                    seen_releases.add(rel)
                    pill.setChecked(True)
                else:
                    pill.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self.selection_changed.emit()

    def select_releases_min(self, min_rel_num: int):
        """Checks only releases with major release number >= min_rel_num (e.g. 18 for Rel-18+)."""
        for card in self._cards.values():
            card._updating_checks = True
            for pill in card.pills:
                rel = extract_major_release(pill.version)
                match = re.search(r'\d+', rel)
                if match and int(match.group()) >= min_rel_num:
                    pill.setChecked(True)
                else:
                    pill.setChecked(False)
            card._updating_checks = False
            card._update_header_state()
        self.selection_changed.emit()

    def get_selected_version_ids(self) -> List[int]:
        ids = []
        for card in self._cards.values():
            for p in card.pills:
                if p.isChecked() and p.version_id:
                    ids.append(p.version_id)
        return ids

    def get_checked_versions_info(self) -> List[Dict[str, Any]]:
        info = []
        for card in self._cards.values():
            for p in card.pills:
                if p.isChecked():
                    info.append({
                        "spec_number": card.spec_number,
                        "version": p.version,
                    })
        return info

    def get_all_versions_for_spec(self, spec_number: str) -> List[Dict[str, Any]]:
        card = self._cards.get(spec_number)
        if not card:
            return []
        return [{"spec_number": card.spec_number, "version": p.version} for p in card.pills]

    def focus_spec(self, spec_number: str, version: Optional[str] = None) -> bool:
        """Scrolls to and highlights a specific specification card and version."""
        card = self._cards.get(spec_number)
        if not card:
            return False
        card.setVisible(True)
        if version:
            card.set_pill_state(version, True)
        self.ensureWidgetVisible(card)
        return True

    # --- Tree Compatibility Adapters ---
    def topLevelItemCount(self) -> int:
        return 1 + len(self._cards)

    def topLevelItem(self, idx: int):
        if idx == 0:
            return self._all_item_compat
        spec_keys = sorted(self._cards.keys())
        card_idx = idx - 1
        if 0 <= card_idx < len(spec_keys):
            return _TreeCompatSpecItem(self._cards[spec_keys[card_idx]])
        return None

    def setCurrentItem(self, item):
        pass

    def scrollToItem(self, item):
        pass


# Backward compatibility alias
SpecSearchVersionTreeWidget = SpecSearchVersionCardsWidget
