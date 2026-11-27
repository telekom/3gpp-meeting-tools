"""Right-hand release browser (header card, release chips, checkable version table)."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.components.styles import TABLE_STYLE_CLEAN
from core.ui.flow_layout import FlowLayout
from core.ui.theme import C
from modules.nas.core.spec_catalog import get_spec_metadata, spec_generation
from .chips import make_badge, make_chip

COLUMNS = ["", "Release", "Version", "Uploaded", "File", "Local Cache", "DB Status"]
RowInfo = Dict[str, Any]


def release_of(version: str) -> str:
    major = str(version).lstrip("vV").split(".")[0]
    return f"Rel-{major}" if major.isdigit() else "—"


class ReleasePanel(QWidget):
    """Shows versions of the selected spec; checked rows form the multi-spec ingest queue."""

    queue_changed = pyqtSignal(str, list)  # spec_number, [row_info dicts checked for that spec]
    row_double_clicked = pyqtSignal(dict)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._spec = ""
        self._rows: List[RowInfo] = []
        self._visible: List[RowInfo] = []
        self._checked: Set[Tuple[str, str]] = set()
        self._filling = False
        self._build_ui()

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)

        self.header = QFrame(self)
        self.header.setObjectName("releaseHeaderCard")
        self.header.setStyleSheet(
            f"#releaseHeaderCard {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; "
            f"border-left: 4px solid {C.ACCENT_PRIMARY}; border-radius: 6px; }}"
        )
        hlay = QVBoxLayout(self.header)
        hlay.setContentsMargins(12, 8, 12, 8)
        hlay.setSpacing(3)
        self.title_lbl = QLabel("Select a specification", self.header)
        self.title_lbl.setWordWrap(True)
        self.title_lbl.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        hlay.addWidget(self.title_lbl)
        self.badge_host = QWidget(self.header)
        self.badge_flow = FlowLayout(self.badge_host, margin=0, h_spacing=4, v_spacing=3)
        hlay.addWidget(self.badge_host)
        self.desc_lbl = QLabel("", self.header)
        self.desc_lbl.setWordWrap(True)
        self.desc_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_SECONDARY};")
        hlay.addWidget(self.desc_lbl)
        lay.addWidget(self.header)

        bar = QHBoxLayout()
        bar.setSpacing(6)
        lbl = QLabel("Releases:", self)
        lbl.setStyleSheet(f"font-size: 10.5px; font-weight: 600; color: {C.TEXT_MUTED};")
        bar.addWidget(lbl)
        self.chip_host = QWidget(self)
        self.chip_flow = FlowLayout(self.chip_host, margin=0, h_spacing=4, v_spacing=4)
        self._chip_group = QButtonGroup(self)
        self._chip_group.setExclusive(True)
        self._chip_group.buttonClicked.connect(lambda _b: self._apply_release_filter())
        bar.addWidget(self.chip_host, 1)
        self.queue_visible_btn = make_chip("☑ Queue visible", "Check every release currently shown", checkable=False, parent=self)
        self.queue_visible_btn.clicked.connect(self._check_visible)
        self.clear_btn = make_chip("☐ Clear", "Uncheck all releases of this specification", checkable=False, parent=self)
        self.clear_btn.clicked.connect(self._clear_spec)
        bar.addWidget(self.queue_visible_btn)
        bar.addWidget(self.clear_btn)
        lay.addLayout(bar)

        self.table = QTableWidget(0, len(COLUMNS), self)
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setStyleSheet(TABLE_STYLE_CLEAN)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.itemChanged.connect(self._on_item_changed)
        self.table.itemDoubleClicked.connect(self._on_double_clicked)
        lay.addWidget(self.table, 1)

    # ------------------------------------------------------------------ public API
    def current_spec(self) -> str:
        return self._spec

    def set_checked(self, checked: Set[Tuple[str, str]]) -> None:
        self._checked = set(checked)
        self._refill_table()

    def selected_rows(self) -> List[RowInfo]:
        idxs = sorted({i.row() for i in self.table.selectedItems()})
        return [self._visible[i] for i in idxs if 0 <= i < len(self._visible)]

    def show_spec(self, spec_number: str, rows: List[RowInfo], checked: Set[Tuple[str, str]]) -> None:
        self._spec = spec_number
        self._rows = rows
        self._checked = set(checked)
        meta = get_spec_metadata(spec_number) if spec_number else {}
        self.title_lbl.setText(f"TS {spec_number}: {meta['full_title']}" if spec_number else "No specification selected")
        while self.badge_flow.count():
            item = self.badge_flow.takeAt(0)
            if item and item.widget():
                item.widget().deleteLater()
        if spec_number:
            for text, kind in (
                (str(meta["wg"]), "wg"), (spec_generation(spec_number), "gen"), (str(meta["domain"]), "domain"),
                (str(meta.get("plane", "")), "plane"), (str(meta.get("layer", "")), "layer"),
            ):
                if text:
                    self.badge_flow.addWidget(make_badge(text, kind))
        self.desc_lbl.setText(str(meta.get("description", "")) if spec_number else "")
        self._rebuild_release_chips()
        self._apply_release_filter()

    # ------------------------------------------------------------------ internals
    def _rebuild_release_chips(self) -> None:
        for btn in list(self._chip_group.buttons()):
            self._chip_group.removeButton(btn)
        while self.chip_flow.count():
            item = self.chip_flow.takeAt(0)
            if item and item.widget():
                item.widget().deleteLater()
        majors: List[str] = []
        for r in self._rows:
            rel = release_of(r["version"])
            if rel not in majors:
                majors.append(rel)
        options = [("All", "all"), ("Latest / Rel", "latest"), ("Cached", "cached")] + [(m, m) for m in majors[:6]]
        for idx, (text, value) in enumerate(options):
            chip = make_chip(text, parent=self.chip_host)
            chip.setProperty("rel_filter", value)
            chip.setChecked(idx == 0)
            self._chip_group.addButton(chip)
            self.chip_flow.addWidget(chip)

    def _apply_release_filter(self) -> None:
        btn = self._chip_group.checkedButton()
        mode = str(btn.property("rel_filter")) if btn else "all"
        if mode == "latest":
            seen: Set[str] = set()
            self._visible = []
            for r in self._rows:
                rel = release_of(r["version"])
                if rel not in seen:
                    seen.add(rel)
                    self._visible.append(r)
        elif mode == "cached":
            self._visible = [r for r in self._rows if r["cached"]]
        elif mode.startswith("Rel-"):
            self._visible = [r for r in self._rows if release_of(r["version"]) == mode]
        else:
            self._visible = list(self._rows)
        self._refill_table()

    def _refill_table(self) -> None:
        self._filling = True
        self.table.setRowCount(len(self._visible))
        for row_idx, r in enumerate(self._visible):
            key = (r["spec_number"], r["version"])
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            chk.setCheckState(Qt.Checked if key in self._checked else Qt.Unchecked)
            self.table.setItem(row_idx, 0, chk)
            self.table.setItem(row_idx, 1, QTableWidgetItem(release_of(r["version"])))
            ver = QTableWidgetItem(f"v{r['version']}")
            ver.setFont(self._bold_font(ver))
            self.table.setItem(row_idx, 2, ver)
            self.table.setItem(row_idx, 3, QTableWidgetItem(r["uploaded"]))
            self.table.setItem(row_idx, 4, QTableWidgetItem(r["filename"]))
            cache = QTableWidgetItem(f"🟢 Cached ({r['cache_ext']})" if r["cached"] else "🌐 Remote (FTP)")
            cache.setForeground(self._color(C.SUCCESS_DARK if r["cached"] else C.TEXT_MUTED))
            self.table.setItem(row_idx, 5, cache)
            db = QTableWidgetItem("✅ Ingested" if r["in_db"] else "⚪ Ready")
            db.setForeground(self._color(C.ACCENT_PRIMARY if r["in_db"] else C.TEXT_MUTED))
            self.table.setItem(row_idx, 6, db)
        self._filling = False

    @staticmethod
    def _bold_font(item: QTableWidgetItem):
        f = item.font()
        f.setBold(True)
        return f

    @staticmethod
    def _color(hex_token: str) -> QBrush:
        return QBrush(QColor(hex_token))

    def _emit_queue(self) -> None:
        checked_here = [r for r in self._rows if (r["spec_number"], r["version"]) in self._checked]
        self.queue_changed.emit(self._spec, checked_here)

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if self._filling or item.column() != 0 or not (0 <= item.row() < len(self._visible)):
            return
        r = self._visible[item.row()]
        key = (r["spec_number"], r["version"])
        if item.checkState() == Qt.Checked:
            self._checked.add(key)
        else:
            self._checked.discard(key)
        self._emit_queue()

    def _check_visible(self) -> None:
        for r in self._visible:
            self._checked.add((r["spec_number"], r["version"]))
        self._refill_table()
        self._emit_queue()

    def _clear_spec(self) -> None:
        for r in self._rows:
            self._checked.discard((r["spec_number"], r["version"]))
        self._refill_table()
        self._emit_queue()

    def _on_double_clicked(self, item: QTableWidgetItem) -> None:
        if item.column() == 0:
            return
        if 0 <= item.row() < len(self._visible):
            self.row_double_clicked.emit(self._visible[item.row()])
