"""Faceted search bar and categorized specification list for the Specification Import Studio."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.components.styles import SEARCH_INPUT_STYLE
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from core.ui.theme import C
from modules.nas.core.spec_catalog import (
    get_spec_metadata,
    importable_spec_order,
    spec_generation,
    spec_matches_filters,
)
from .chips import make_badge, make_chip

# (facet key, label, [(chip text, filter value or None for "All")])
FACETS: List[Tuple[str, str, List[Tuple[str, Optional[str]]]]] = [
    ("generation", "Tech", [("All", None), ("5G", "5G"), ("LTE", "LTE"), ("3G", "3G/2G")]),
    ("domain", "Domain", [("All", None), ("RAN", "RAN"), ("Core", "Core")]),
    ("plane", "Plane", [("All", None), ("CP", "CP"), ("UP", "UP")]),
    (
        "layer",
        "Layer",
        [
            ("All", None), ("L2", "L2"), ("L3 / RRC", "L3 / RRC"), ("NAS", "NAS"), ("UE Policy", "UE Policy"),
            ("App Protocol", "App Protocol"), ("Transport", "Transport"), ("Tunnel / Framing", "Tunnel / Framing"),
        ],
    ),
]


class SpecFacetBar(QWidget):
    """Search box + exclusive filter chips (Tech / Domain / Plane / Layer)."""

    changed = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._groups: Dict[str, QButtonGroup] = {}
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        self.search_edit = QLineEdit(self)
        self.search_edit.setPlaceholderText("🔎  Search specifications (e.g. 38.323, PDCP, URSP, RAN2, tunnel…)")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setMinimumHeight(scale_px(30))
        self.search_edit.setStyleSheet(SEARCH_INPUT_STYLE)
        self.search_edit.textChanged.connect(lambda _t: self.changed.emit())
        lay.addWidget(self.search_edit)

        host = QWidget(self)
        flow = FlowLayout(host, margin=0, h_spacing=14, v_spacing=4)
        for key, label, options in FACETS:
            group_widget = QWidget(host)
            row = QHBoxLayout(group_widget)
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(4)
            lbl = QLabel(label + ":", group_widget)
            lbl.setStyleSheet(f"font-size: 10.5px; font-weight: 600; color: {C.TEXT_MUTED};")
            row.addWidget(lbl)
            grp = QButtonGroup(self)
            grp.setExclusive(True)
            for idx, (text, value) in enumerate(options):
                chip = make_chip(text, parent=group_widget)
                chip.setProperty("filter_value", value)
                if idx == 0:
                    chip.setChecked(True)
                grp.addButton(chip)
                row.addWidget(chip)
            grp.buttonClicked.connect(lambda _b: self.changed.emit())
            self._groups[key] = grp
            flow.addWidget(group_widget)
        lay.addWidget(host)

    def criteria(self) -> Dict[str, Optional[str]]:
        out: Dict[str, Optional[str]] = {"text": self.search_edit.text()}
        for key, grp in self._groups.items():
            btn = grp.checkedButton()
            out[key] = btn.property("filter_value") if btn else None
        return out

    def reset(self) -> None:
        self.search_edit.clear()
        for grp in self._groups.values():
            first = grp.buttons()[0] if grp.buttons() else None
            if first:
                first.setChecked(True)
        self.changed.emit()


class SpecRowWidget(QFrame):
    """Two-line spec card: `TS 38.323 · NR PDCP` + title + facet / status badges."""

    def __init__(self, spec_number: str, ingested: int, queued: int, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        meta = get_spec_metadata(spec_number)
        importable = bool(meta.get("importable", True))
        self.setStyleSheet("QFrame { background: transparent; border: none; }")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 5, 8, 5)
        lay.setSpacing(2)

        fg = C.TEXT_STRONG if importable else C.TEXT_FAINT
        head = QLabel(f"<b>TS {spec_number}</b> · {meta['acronym']}", self)
        head.setStyleSheet(f"font-size: 12px; color: {fg};")
        lay.addWidget(head)
        sub = QLabel(str(meta["title"]), self)
        sub.setStyleSheet(f"font-size: 10.5px; color: {C.TEXT_MUTED if importable else C.TEXT_FAINT};")
        lay.addWidget(sub)

        badges = QHBoxLayout()
        badges.setSpacing(4)
        badges.addWidget(make_badge(str(meta["wg"]), "wg"))
        badges.addWidget(make_badge(spec_generation(spec_number), "gen"))
        if meta.get("plane"):
            badges.addWidget(make_badge(str(meta["plane"]), "plane"))
        if meta.get("layer"):
            badges.addWidget(make_badge(str(meta["layer"]), "layer"))
        if not importable:
            badges.addWidget(make_badge("reference only", "muted"))
        if ingested:
            badges.addWidget(make_badge(f"✓ {ingested} ingested", "ok"))
        if queued:
            badges.addWidget(make_badge(f"● {queued} queued", "queued"))
        badges.addStretch()
        lay.addLayout(badges)


class SpecListPanel(QListWidget):
    """Categorized, filterable list of catalogued specifications."""

    spec_selected = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setSpacing(2)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setStyleSheet(
            f"QListWidget {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; outline: 0; }} "
            f"QListWidget::item {{ border: 1px solid transparent; border-radius: 5px; }} "
            f"QListWidget::item:hover {{ background-color: {C.BG_HOVER}; }} "
            f"QListWidget::item:selected {{ background-color: {C.ACCENT_SOFT}; border: 1px solid {C.ACCENT_PRIMARY}; }}"
        )
        self.currentItemChanged.connect(self._on_current_changed)
        self._ingested: Dict[str, int] = {}
        self._queued: Dict[str, int] = {}
        self._criteria: Dict[str, Optional[str]] = {}

    def set_status(self, ingested: Dict[str, int], queued: Dict[str, int]) -> None:
        self._ingested, self._queued = ingested, queued

    def current_spec(self) -> str:
        item = self.currentItem()
        return str(item.data(Qt.UserRole)) if item else ""

    def refresh_status(self) -> None:
        """Re-render badges (ingested / queued counts) without changing selection or emitting signals."""
        for row in range(self.count()):
            item = self.item(row)
            spec = str(item.data(Qt.UserRole))
            widget = SpecRowWidget(spec, self._ingested.get(spec, 0), self._queued.get(spec, 0))
            item.setSizeHint(widget.sizeHint())
            self.setItemWidget(item, widget)

    def rebuild(self, criteria: Dict[str, Optional[str]], keep: str = "") -> int:
        """Rebuild rows for `criteria`; returns visible row count (selection kept/defaulted)."""
        self._criteria = criteria
        wanted = keep or self.current_spec()
        self.blockSignals(True)
        self.clear()
        first_row, wanted_row = -1, -1
        for spec in importable_spec_order():
            if not spec_matches_filters(
                spec,
                text=criteria.get("text") or "",
                generation=criteria.get("generation"),
                domain=criteria.get("domain"),
                plane=criteria.get("plane"),
                layer=criteria.get("layer"),
            ):
                continue
            meta = get_spec_metadata(spec)
            importable = bool(meta.get("importable", True))
            item = QListWidgetItem()
            item.setData(Qt.UserRole, spec)
            item.setToolTip(str(meta.get("note") or meta["description"]))
            widget = SpecRowWidget(spec, self._ingested.get(spec, 0), self._queued.get(spec, 0))
            item.setSizeHint(widget.sizeHint())
            self.addItem(item)
            self.setItemWidget(item, widget)
            if not importable:
                item.setFlags(item.flags() & ~Qt.ItemIsSelectable)
                continue
            row = self.count() - 1
            if first_row < 0:
                first_row = row
            if spec == wanted:
                wanted_row = row
        self.blockSignals(False)
        target = wanted_row if wanted_row >= 0 else first_row
        if target >= 0:
            self.setCurrentRow(target)
            self.spec_selected.emit(self.current_spec())
        else:
            self.spec_selected.emit("")
        return self.count()

    def _on_current_changed(self, current: Optional[QListWidgetItem], _prev: Optional[QListWidgetItem]) -> None:
        self.spec_selected.emit(str(current.data(Qt.UserRole)) if current else "")
