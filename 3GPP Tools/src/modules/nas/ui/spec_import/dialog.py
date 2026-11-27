"""Modern two-pane Specification Import Studio (faceted spec browser + multi-spec ingest queue)."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from core.ui.components.styles import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from core.ui.theme import C
from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
from modules.nas.core.nas_threads import SpecCacheIndex
from modules.nas.core.spec_catalog import SPEC_BUNDLES, get_spec_metadata
from modules.specifications.core.database import SpecsDatabase
from .chips import make_chip, primary_button_qss
from .release_panel import ReleasePanel, RowInfo
from .spec_list import SpecFacetBar, SpecListPanel


class NASVersionSelectDialog(QDialog):
    """Select protocol specification releases (across several specs) to download and ingest."""

    def __init__(
        self,
        specs_db: SpecsDatabase,
        nas_db: NASDatabase,
        cache_dir: Path,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.specs_db = specs_db
        self.nas_db = nas_db
        self.cache_dir = cache_dir
        self.selected_files_info: List[Dict[str, Any]] = []
        self._queue: Dict[Tuple[str, str], RowInfo] = {}
        self._row_cache: Dict[str, List[RowInfo]] = {}
        self._imported: Set[Tuple[str, str]] = set()
        self._cache_index = SpecCacheIndex()

        self.setWindowTitle("Import 3GPP Specifications")
        self.setMinimumSize(scale_px(860), scale_px(560))
        self.resize(scale_px(1100), scale_px(720))
        self._refresh_imported()
        self._setup_ui()
        self.facets.reset()
        self._update_footer()

    # ------------------------------------------------------------------ UI
    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 12)
        root.setSpacing(8)

        title = QLabel("📥  Import 3GPP Specifications", self)
        title.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        sub = QLabel(
            "Pick specifications on the left, tick the releases you need on the right (across several specs), "
            "then fetch & ingest them in one go. Releases not cached locally are downloaded from the 3GPP FTP server.",
            self,
        )
        sub.setWordWrap(True)
        sub.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED};")
        root.addWidget(title)
        root.addWidget(sub)

        self.facets = SpecFacetBar(self)
        self.facets.changed.connect(self._on_filters_changed)
        root.addWidget(self.facets)

        bundle_host = QWidget(self)
        bundle_row = QHBoxLayout(bundle_host)
        bundle_row.setContentsMargins(0, 0, 0, 0)
        bundle_row.setSpacing(6)
        blabel = QLabel("Quick bundles (latest release of each):", bundle_host)
        blabel.setStyleSheet(f"font-size: 10.5px; font-weight: 600; color: {C.TEXT_MUTED};")
        bundle_row.addWidget(blabel)
        flow_host = QWidget(bundle_host)
        flow = FlowLayout(flow_host, margin=0, h_spacing=5, v_spacing=4)
        for label, tip, specs in SPEC_BUNDLES:
            chip = make_chip(label, f"{tip}\nQueues: " + ", ".join(f"TS {s}" for s in specs), checkable=False, parent=flow_host)
            chip.clicked.connect(lambda _c=False, sp=tuple(specs): self._queue_bundle(list(sp)))
            flow.addWidget(chip)
        bundle_row.addWidget(flow_host, 1)
        root.addWidget(bundle_host)

        self.splitter = QSplitter(Qt.Horizontal, self)
        self.splitter.setChildrenCollapsible(False)
        self.spec_list = SpecListPanel(self.splitter)
        self.spec_list.setMinimumWidth(scale_px(330))
        self.spec_list.spec_selected.connect(self._on_spec_selected)
        self.release_panel = ReleasePanel(self.splitter)
        self.release_panel.queue_changed.connect(self._on_queue_changed)
        self.release_panel.row_double_clicked.connect(self._on_row_double_clicked)
        self.splitter.addWidget(self.spec_list)
        self.splitter.addWidget(self.release_panel)
        self.splitter.setSizes([scale_px(430), scale_px(650)])
        root.addWidget(self.splitter, 1)

        footer = QFrame(self)
        footer.setObjectName("importFooter")
        footer.setStyleSheet(
            f"#importFooter {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        frow = QHBoxLayout(footer)
        frow.setContentsMargins(10, 6, 10, 6)
        frow.setSpacing(8)
        self.summary_lbl = QLabel("", footer)
        self.summary_lbl.setStyleSheet(f"font-size: 11.5px; font-weight: 600; color: {C.TEXT_BODY};")
        frow.addWidget(self.summary_lbl, 1)
        self.clear_queue_btn = QPushButton("🗑 Clear queue", footer)
        self.clear_queue_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_queue_btn.setCursor(Qt.PointingHandCursor)
        self.clear_queue_btn.clicked.connect(self._clear_queue)
        frow.addWidget(self.clear_queue_btn)
        self.cancel_btn = QPushButton("Cancel", footer)
        self.cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cancel_btn.clicked.connect(self.reject)
        frow.addWidget(self.cancel_btn)
        self.ok_btn = QPushButton("📥 Fetch && Ingest", footer)
        self.ok_btn.setStyleSheet(primary_button_qss())
        self.ok_btn.setCursor(Qt.PointingHandCursor)
        self.ok_btn.setDefault(True)
        self.ok_btn.clicked.connect(self._on_accept)
        frow.addWidget(self.ok_btn)
        root.addWidget(footer)

    # ------------------------------------------------------------------ data
    def _refresh_imported(self) -> None:
        try:
            self._imported = {(v["spec_number"], v["version"]) for v in self.nas_db.get_imported_versions()}
        except Exception:
            self._imported = set()

    def _load_rows(self, spec_number: str) -> List[RowInfo]:
        if spec_number in self._row_cache:
            return self._row_cache[spec_number]
        files = self.specs_db.search_files(spec_number=spec_number) if self.specs_db else []
        files = sorted(files, key=lambda row: parse_version_tuple(row[5]), reverse=True)
        rows: List[RowInfo] = []
        for row_data in files:
            _, s_num, _, _, filename, version, url, upload_date, *_rest = row_data
            cached = self._cache_index.find(filename, s_num)
            rows.append({
                "spec_number": s_num,
                "version": version,
                "filename": filename,
                "file_url": url,
                "uploaded": str(upload_date or "")[:10],
                "cached": bool(cached),
                "cache_ext": cached.suffix[1:].upper() if cached else "",
                "in_db": (s_num, version) in self._imported,
            })
        self._row_cache[spec_number] = rows
        return rows

    def _queue_counts(self) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        for spec, _ver in self._queue:
            counts[spec] = counts.get(spec, 0) + 1
        return counts

    def _ingested_counts(self) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        for spec, _ver in self._imported:
            counts[spec] = counts.get(spec, 0) + 1
        return counts

    # ------------------------------------------------------------------ slots
    def _on_filters_changed(self) -> None:
        self.spec_list.set_status(self._ingested_counts(), self._queue_counts())
        self.spec_list.rebuild(self.facets.criteria())

    def _on_spec_selected(self, spec_number: str) -> None:
        if spec_number and spec_number == self.release_panel.current_spec() and self.release_panel._rows:
            return
        rows = self._load_rows(spec_number) if spec_number else []
        self.release_panel.show_spec(spec_number, rows, set(self._queue))

    def _on_queue_changed(self, spec_number: str, rows: List[RowInfo]) -> None:
        for key in [k for k in self._queue if k[0] == spec_number]:
            del self._queue[key]
        for r in rows:
            self._queue[(r["spec_number"], r["version"])] = r
        self._after_queue_change()

    def _queue_bundle(self, specs: List[str]) -> None:
        added = 0
        for spec in specs:
            rows = self._load_rows(spec)
            if rows:
                self._queue[(spec, rows[0]["version"])] = rows[0]
                added += 1
        if not added:
            QMessageBox.information(self, "Bundle Unavailable", "No releases for this bundle were found in the specification archive.")
        self.release_panel.set_checked(set(self._queue))
        self._after_queue_change()

    def _clear_queue(self) -> None:
        self._queue.clear()
        self.release_panel.set_checked(set())
        self._after_queue_change()

    def _after_queue_change(self) -> None:
        self.spec_list.set_status(self._ingested_counts(), self._queue_counts())
        self.spec_list.refresh_status()
        self._update_footer()

    def _update_footer(self) -> None:
        n = len(self._queue)
        specs = self._queue_counts()
        if n:
            self.summary_lbl.setText(f"📦 {n} release{'s' if n != 1 else ''} queued from {len(specs)} specification{'s' if len(specs) != 1 else ''}")
            self.summary_lbl.setToolTip("\n".join(f"TS {s} v{v}" for s, v in self._queue))
        else:
            self.summary_lbl.setText("Nothing queued — tick releases (or use a quick bundle) to build the ingest queue")
            self.summary_lbl.setToolTip("")
        self.ok_btn.setText(f"📥 Fetch && Ingest ({n})" if n else "📥 Fetch && Ingest")
        self.clear_queue_btn.setEnabled(bool(n))

    def _on_row_double_clicked(self, row: RowInfo) -> None:
        key = (row["spec_number"], row["version"])
        if not self._queue:
            self._queue[key] = row
            self._on_accept()
            return
        if key in self._queue:
            del self._queue[key]
        else:
            self._queue[key] = row
        self.release_panel.set_checked(set(self._queue))
        self._after_queue_change()

    def _on_accept(self) -> None:
        if not self._queue:
            for r in self.release_panel.selected_rows():
                self._queue[(r["spec_number"], r["version"])] = r
        if not self._queue:
            QMessageBox.warning(self, "Selection Required", "Please queue at least one specification release to import.")
            return
        self.selected_files_info = [
            {
                "spec_number": r["spec_number"],
                "version": r["version"],
                "filename": r["filename"],
                "file_url": r["file_url"],
            }
            for r in self._queue.values()
        ]
        self.accept()

    def preselect_spec(self, spec_number: str, major_version: int = 0) -> bool:
        """Select `spec_number` in the left list and queue its latest (or `major_version`) release."""
        clean = str(spec_number or "").replace("TS", "").strip()
        if not clean:
            return False
        self.facets.reset()
        target_row = -1
        for idx in range(self.spec_list.count()):
            item = self.spec_list.item(idx)
            if item and str(item.data(Qt.UserRole)) == clean:
                target_row = idx
                self.spec_list.setCurrentRow(idx)
                self.spec_list.scrollToItem(item)
                break
        if target_row < 0:
            return False

        rows = self._load_rows(clean)
        picked: Optional[RowInfo] = None
        if major_version > 0:
            for r in rows:
                vt = parse_version_tuple(r["version"])
                if vt:
                    maj = vt[0][1] if isinstance(vt[0], tuple) else vt[0]
                    if maj == major_version:
                        picked = r
                        break
        if picked is None and rows:
            picked = rows[0]

        if picked is not None:
            self._queue[(clean, picked["version"])] = picked
            self.release_panel.set_checked(set(self._queue))
            self._after_queue_change()
            for r_idx, vis in enumerate(self.release_panel._visible):
                if vis["version"] == picked["version"]:
                    self.release_panel.table.selectRow(r_idx)
                    item = self.release_panel.table.item(r_idx, 0)
                    if item is not None:
                        self.release_panel.table.scrollToItem(item)
                    break
        return True

    # Backwards-compatible helper used by older callers/tests
    def current_spec_meta(self) -> Dict[str, Any]:
        return get_spec_metadata(self.release_panel.current_spec() or "38.331")  # type: ignore[return-value]
