# --- File: src/modules/spec_search/ui/quick_search_panel.py ---
"""
Compact docked Quick Spec Search panel for the contextual RightSidePanel (Phase 3.3).

Queries indexed 3GPP specification clauses non-blockingly via SpecSearchDatabase
and SpecSearchQueryWorker with 250ms debounce, stale-request discarding,
default Latest-per-Spec scoping, collapsible clause preview, and deep links.
"""

from __future__ import annotations

import html
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QRectF, QSize, Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QAbstractTextDocumentLayout, QColor, QPen, QTextDocument
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStyledItemDelegate,
    QStyle,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.icons import lucide_icon
from core.ui.theme import C, R, S, T
from core.ui.toast import notify_success
from modules.spec_search.core.spec_search_db import SpecSearchDatabase, parse_version_tuple
from modules.spec_search.core.spec_search_threads import SpecSearchQueryWorker

logger = logging.getLogger(__name__)

_ROLE_HTML = Qt.UserRole + 10
_ROLE_META = Qt.UserRole + 11


class _QuickSearchResultDelegate(QStyledItemDelegate):
    """Lightweight HTML delegate rendering citation header and <mark> snippet."""

    def paint(self, painter, option, index) -> None:
        painter.save()
        rect = option.rect
        selected = bool(option.state & QStyle.State_Selected)
        hovered = bool(option.state & QStyle.State_MouseOver)

        if selected:
            bg = C.ACCENT_SOFT
            border = C.ACCENT_PRIMARY
        elif hovered:
            bg = C.BG_SURFACE
            border = C.BORDER_SUBTLE
        else:
            bg = C.BG_CARD
            border = C.BORDER_SUBTLE

        painter.setPen(Qt.NoPen)
        painter.setBrush(option.palette.base())
        card_rect = rect.adjusted(4, 3, -4, -3)

        painter.setBrush(QColor(bg))
        painter.setPen(QPen(QColor(border), 1))
        painter.drawRoundedRect(card_rect, R.MD, R.MD)

        html_text = index.data(_ROLE_HTML) or ""
        doc = QTextDocument()
        doc.setDocumentMargin(0)
        doc.setDefaultFont(option.font)
        doc.setTextWidth(max(120, card_rect.width() - 16))
        doc.setHtml(html_text)

        painter.translate(card_rect.left() + 8, card_rect.top() + 6)
        ctx = QAbstractTextDocumentLayout.PaintContext()
        clip = QRectF(0, 0, card_rect.width() - 16, card_rect.height() - 12)
        painter.setClipRect(clip)
        doc.documentLayout().draw(painter, ctx)
        painter.restore()

    def sizeHint(self, option, index) -> QSize:
        html_text = index.data(_ROLE_HTML) or ""
        doc = QTextDocument()
        doc.setDocumentMargin(0)
        doc.setDefaultFont(option.font)
        width = max(240, option.rect.width() - 24) if option.rect.width() > 60 else 320
        doc.setTextWidth(width)
        doc.setHtml(html_text)
        height = int(doc.size().height()) + 18
        return QSize(width, max(54, min(110, height)))


def _format_snippet_html(snippet: str) -> str:
    """Escape snippet HTML safely while styling FTS5 <mark>...</mark> hits."""
    if not snippet:
        return ""
    tokens = re.split(r"(<mark>.*?</mark>)", str(snippet), flags=re.IGNORECASE | re.DOTALL)
    parts: List[str] = []
    for tok in tokens:
        low = tok.lower()
        if low.startswith("<mark>") and low.endswith("</mark>"):
            inner = html.escape(tok[6:-7])
            parts.append(
                f'<span style="background-color:{C.WARNING_SOFT};color:{C.TEXT_STRONG};'
                f'font-weight:600;border-radius:2px;">{inner}</span>'
            )
        else:
            parts.append(html.escape(tok))
    return "".join(parts)


class QuickSpecSearchPanel(QWidget):
    """Compact sidebar quick search over indexed 3GPP clauses."""

    open_in_spec_search_requested = pyqtSignal(str, str, str)  # (spec, version, clause)
    open_in_specs_requested = pyqtSignal(str, str)             # (spec, version)

    MAX_VISIBLE_RESULTS = 80
    DEBOUNCE_MS = 250

    def __init__(
        self,
        spec_search_db_path: Path,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("quickSpecSearchPanel")
        self.db_path = Path(spec_search_db_path)
        self.db = SpecSearchDatabase(self.db_path)

        self._versions: List[Dict[str, Any]] = []
        self._req_id: int = 0
        self._worker: Optional[SpecSearchQueryWorker] = None
        self._retired_workers: List[SpecSearchQueryWorker] = []
        self._selected_meta: Optional[Dict[str, Any]] = None
        self._selected_content: str = ""

        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.setInterval(self.DEBOUNCE_MS)
        self._debounce_timer.timeout.connect(self._execute_search)

        self._setup_ui()
        self._apply_styles()
        self.refresh_versions()

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(S.SM, S.SM, S.SM, S.SM)
        root.setSpacing(S.XS)

        # --- Search & Filter Header Card ---
        filter_card = QFrame(self)
        filter_card.setObjectName("quickSearchFilterCard")
        fc_layout = QVBoxLayout(filter_card)
        fc_layout.setContentsMargins(S.SM, S.SM, S.SM, S.SM)
        fc_layout.setSpacing(S.XS)

        self.search_input = QLineEdit(filter_card)
        self.search_input.setObjectName("quickSearchInput")
        self.search_input.setPlaceholderText("Search clause text (e.g. registration, DNN, PEI)…")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._schedule_search)
        self.search_input.returnPressed.connect(self._execute_search)
        fc_layout.addWidget(self.search_input)

        row2 = QHBoxLayout()
        row2.setContentsMargins(0, 0, 0, 0)
        row2.setSpacing(S.XS)

        self.spec_filter_input = QLineEdit(filter_card)
        self.spec_filter_input.setObjectName("quickSearchSubInput")
        self.spec_filter_input.setPlaceholderText("Spec: e.g. 23.501")
        self.spec_filter_input.setClearButtonEnabled(True)
        self.spec_filter_input.textChanged.connect(self._schedule_search)
        row2.addWidget(self.spec_filter_input, 1)

        self.clause_filter_input = QLineEdit(filter_card)
        self.clause_filter_input.setObjectName("quickSearchSubInput")
        self.clause_filter_input.setPlaceholderText("Clause: e.g. 5.5")
        self.clause_filter_input.setClearButtonEnabled(True)
        self.clause_filter_input.textChanged.connect(self._schedule_search)
        row2.addWidget(self.clause_filter_input, 1)

        self.latest_only_btn = QPushButton("Latest per Spec", filter_card)
        self.latest_only_btn.setObjectName("quickSearchTogglePill")
        self.latest_only_btn.setCheckable(True)
        self.latest_only_btn.setChecked(True)
        self.latest_only_btn.setCursor(Qt.PointingHandCursor)
        self.latest_only_btn.setToolTip(
            "When checked, limits results to the newest indexed version per specification."
        )
        self.latest_only_btn.toggled.connect(self._schedule_search)
        row2.addWidget(self.latest_only_btn)

        fc_layout.addLayout(row2)
        root.addWidget(filter_card)

        # --- Status Strip ---
        self.status_label = QLabel("Type keywords to search indexed clauses.", self)
        self.status_label.setObjectName("quickSearchStatusLabel")
        self.status_label.setWordWrap(True)
        root.addWidget(self.status_label)

        # --- Results + Collapsible Preview Splitter ---
        self.results_splitter = QSplitter(Qt.Vertical, self)
        self.results_splitter.setChildrenCollapsible(False)

        self.results_list = QListWidget(self.results_splitter)
        self.results_list.setObjectName("quickSearchResultsList")
        self.results_list.setUniformItemSizes(False)
        self.results_list.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        self.results_list.setItemDelegate(_QuickSearchResultDelegate(self.results_list))
        self.results_list.itemClicked.connect(self._on_result_selected)
        self.results_list.currentItemChanged.connect(self._on_current_item_changed)
        self.results_splitter.addWidget(self.results_list)

        self.preview_frame = QFrame(self.results_splitter)
        self.preview_frame.setObjectName("quickSearchPreviewFrame")
        prev_layout = QVBoxLayout(self.preview_frame)
        prev_layout.setContentsMargins(S.SM, S.SM, S.SM, S.SM)
        prev_layout.setSpacing(S.XS)

        prev_hdr = QHBoxLayout()
        prev_hdr.setContentsMargins(0, 0, 0, 0)
        prev_hdr.setSpacing(S.XS)

        self.preview_title_label = QLabel("Clause Preview", self.preview_frame)
        self.preview_title_label.setObjectName("quickSearchPreviewTitle")
        self.preview_title_label.setWordWrap(True)
        prev_hdr.addWidget(self.preview_title_label, 1)

        self.close_preview_btn = QPushButton("", self.preview_frame)
        self.close_preview_btn.setObjectName("quickSearchIconBtn")
        self.close_preview_btn.setIcon(
            lucide_icon("x", color=C.TEXT_MUTED, active_color=C.TEXT_PRIMARY, size=13)
        )
        self.close_preview_btn.setToolTip("Collapse clause preview")
        self.close_preview_btn.setCursor(Qt.PointingHandCursor)
        self.close_preview_btn.clicked.connect(lambda: self.preview_frame.setVisible(False))
        prev_hdr.addWidget(self.close_preview_btn)
        prev_layout.addLayout(prev_hdr)

        self.preview_browser = QTextBrowser(self.preview_frame)
        self.preview_browser.setObjectName("quickSearchPreviewBrowser")
        self.preview_browser.setOpenExternalLinks(False)
        prev_layout.addWidget(self.preview_browser, 1)

        action_row = QHBoxLayout()
        action_row.setContentsMargins(0, 0, 0, 0)
        action_row.setSpacing(S.XS)

        self.open_search_btn = QPushButton("Spec Search", self.preview_frame)
        self.open_search_btn.setObjectName("quickSearchActionBtn")
        self.open_search_btn.setIcon(
            lucide_icon("search", color=C.ACCENT_PRIMARY, size=13)
        )
        self.open_search_btn.setCursor(Qt.PointingHandCursor)
        self.open_search_btn.setToolTip("Open this spec and clause in full Spec Search tab")
        self.open_search_btn.clicked.connect(self._on_open_in_spec_search)
        action_row.addWidget(self.open_search_btn)

        self.open_specs_btn = QPushButton("Specs", self.preview_frame)
        self.open_specs_btn.setObjectName("quickSearchActionBtn")
        self.open_specs_btn.setIcon(
            lucide_icon("book-open", color=C.TEXT_SECONDARY, size=13)
        )
        self.open_specs_btn.setCursor(Qt.PointingHandCursor)
        self.open_specs_btn.setToolTip("Open this specification in the Specifications tab")
        self.open_specs_btn.clicked.connect(self._on_open_in_specs)
        action_row.addWidget(self.open_specs_btn)

        action_row.addStretch(1)

        self.copy_citation_btn = QPushButton("Copy", self.preview_frame)
        self.copy_citation_btn.setObjectName("quickSearchActionBtn")
        self.copy_citation_btn.setIcon(
            lucide_icon("copy", color=C.TEXT_SECONDARY, size=13)
        )
        self.copy_citation_btn.setCursor(Qt.PointingHandCursor)
        self.copy_citation_btn.setToolTip("Copy clause citation and excerpt to clipboard")
        self.copy_citation_btn.clicked.connect(self._on_copy_citation)
        action_row.addWidget(self.copy_citation_btn)

        prev_layout.addLayout(action_row)
        self.preview_frame.setVisible(False)
        self.results_splitter.addWidget(self.preview_frame)
        self.results_splitter.setStretchFactor(0, 1)
        self.results_splitter.setStretchFactor(1, 0)

        root.addWidget(self.results_splitter, 1)

    def _apply_styles(self) -> None:
        self.setStyleSheet(f"""
            QWidget#quickSpecSearchPanel {{
                background-color: {C.BG_CANVAS};
            }}
            QFrame#quickSearchFilterCard, QFrame#quickSearchPreviewFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.LG}px;
            }}
            QLineEdit#quickSearchInput, QLineEdit#quickSearchSubInput {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_PRIMARY};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.MD}px;
                padding: 5px 8px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_BODY}px;
            }}
            QLineEdit#quickSearchInput:focus, QLineEdit#quickSearchSubInput:focus {{
                border-color: {C.ACCENT_PRIMARY};
                background-color: {C.BG_CARD};
            }}
            QPushButton#quickSearchTogglePill {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_SECONDARY};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.MD}px;
                padding: 4px 8px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QPushButton#quickSearchTogglePill:checked {{
                background-color: {C.ACCENT_SOFT};
                color: {C.ACCENT_PRIMARY};
                border-color: {C.ACCENT_PRIMARY};
            }}
            QLabel#quickSearchStatusLabel {{
                color: {C.TEXT_MUTED};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                padding: 0 2px;
            }}
            QListWidget#quickSearchResultsList {{
                background-color: transparent;
                border: none;
                outline: none;
            }}
            QLabel#quickSearchPreviewTitle {{
                color: {C.TEXT_STRONG};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_TABLE}px;
                font-weight: 600;
            }}
            QTextBrowser#quickSearchPreviewBrowser {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_PRIMARY};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.MD}px;
                padding: 6px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_TABLE}px;
            }}
            QPushButton#quickSearchActionBtn {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_PRIMARY};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.MD}px;
                padding: 4px 8px;
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                font-weight: 600;
            }}
            QPushButton#quickSearchActionBtn:hover {{
                background-color: {C.ACCENT_SOFT};
                border-color: {C.ACCENT_PRIMARY};
                color: {C.ACCENT_PRIMARY};
            }}
            QPushButton#quickSearchIconBtn {{
                background-color: transparent;
                border: none;
                border-radius: {R.SM}px;
                padding: 2px;
            }}
            QPushButton#quickSearchIconBtn:hover {{
                background-color: {C.BG_SURFACE};
            }}
        """)

    # --- Public Helpers ---

    def current_query(self) -> str:
        return self.search_input.text().strip()

    def focus_input(self) -> None:
        self.search_input.setFocus()
        self.search_input.selectAll()

    def refresh_versions(self) -> None:
        """Reload indexed specification versions from SpecSearchDatabase."""
        try:
            self._versions = self.db.get_imported_versions()
        except Exception as exc:
            logger.debug("QuickSpecSearchPanel failed to read imported versions: %s", exc)
            self._versions = []

    def resolve_version_ids(self) -> List[int]:
        """Compute active version IDs honoring Spec filter and Latest-per-Spec."""
        if not self._versions:
            self.refresh_versions()

        spec_q = self.spec_filter_input.text().strip().lower()
        for prefix in ("ts ", "tr ", "ts", "tr"):
            if spec_q.startswith(prefix) and len(spec_q) > len(prefix):
                rest = spec_q[len(prefix):].strip()
                if rest and rest[0].isdigit():
                    spec_q = rest
                    break

        filtered = [
            row
            for row in self._versions
            if not spec_q or spec_q in str(row.get("spec_number", "")).lower()
        ]

        if self.latest_only_btn.isChecked():
            best_by_spec: Dict[str, Dict[str, Any]] = {}
            for row in filtered:
                spec_num = str(row.get("spec_number", ""))
                existing = best_by_spec.get(spec_num)
                if existing is None or parse_version_tuple(row.get("version", "")) > parse_version_tuple(
                    existing.get("version", "")
                ):
                    best_by_spec[spec_num] = row
            filtered = list(best_by_spec.values())

        return [int(r["id"]) for r in filtered if r.get("id") is not None]

    # --- Asynchronous Search Execution ---

    def _schedule_search(self) -> None:
        self._debounce_timer.start()

    def _execute_search(self) -> None:
        self._debounce_timer.stop()
        query = self.search_input.text().strip()
        if not query:
            self._req_id += 1
            self.results_list.clear()
            self.preview_frame.setVisible(False)
            self._selected_meta = None
            self._selected_content = ""
            if not self._versions:
                self.refresh_versions()
            if not self._versions:
                self.status_label.setText(
                    "No specifications indexed in Spec Search yet."
                )
            else:
                self.status_label.setText(
                    f"Ready across {len(self._versions)} indexed specification version(s)."
                )
            return

        version_ids = self.resolve_version_ids()
        if not version_ids:
            self._req_id += 1
            self.results_list.clear()
            self.preview_frame.setVisible(False)
            if not self._versions:
                self.status_label.setText("No specifications indexed in Spec Search yet.")
            else:
                self.status_label.setText("No indexed versions match the Spec filter.")
            return

        clause_filter = self.clause_filter_input.text().strip() or None
        self._req_id += 1
        req_id = self._req_id
        self.status_label.setText("Searching…")

        prev_worker = self._worker
        if prev_worker is not None and prev_worker.isRunning():
            prev_worker.requestInterruption()
            self._retired_workers.append(prev_worker)

        worker = SpecSearchQueryWorker(
            self.db,
            query,
            version_ids,
            clause_filter,
            req_id,
        )
        self._worker = worker
        worker.results_ready.connect(self._on_results_ready)
        worker.finished.connect(lambda w=worker: self._on_worker_finished(w))
        worker.start()

    def _on_worker_finished(self, worker: SpecSearchQueryWorker) -> None:
        if worker in self._retired_workers:
            self._retired_workers.remove(worker)
        if self._worker is worker:
            self._worker = None
        worker.deleteLater()

    def _on_results_ready(self, df: Any, request_id: int) -> None:
        if request_id != self._req_id:
            return

        self.results_list.clear()
        self.preview_frame.setVisible(False)
        self._selected_meta = None
        self._selected_content = ""

        if df is None or getattr(df, "empty", True):
            self.status_label.setText("0 matching clauses found.")
            return

        total = len(df)
        display_df = df.head(self.MAX_VISIBLE_RESULTS)
        for _, row in display_df.iterrows():
            spec_num = str(row.get("spec_number", "") or "")
            version = str(row.get("version", "") or "")
            clause_num = str(row.get("clause_number", "") or "")
            clause_title = str(row.get("clause_title", "") or "")
            snippet_raw = str(row.get("snippet_text", "") or "")
            clause_pk = int(row.get("clause_pk", 0) or 0)

            ver_lbl = f"v{version.lstrip('vV')}" if version else ""
            header_parts = [f"TS {html.escape(spec_num)}"]
            if ver_lbl:
                header_parts.append(html.escape(ver_lbl))
            if clause_num:
                header_parts.append(f"Clause {html.escape(clause_num)}")
            header = " · ".join(header_parts)
            if clause_title:
                header += f" — {html.escape(clause_title)}"

            snippet_html = _format_snippet_html(snippet_raw)
            card_html = (
                f'<div style="font-size:{T.SIZE_TABLE}px;color:{C.ACCENT_PRIMARY};'
                f'font-weight:600;margin-bottom:2px;">{header}</div>'
                f'<div style="font-size:{T.SIZE_CAPTION}px;color:{C.TEXT_SECONDARY};">'
                f'{snippet_html}</div>'
            )

            item = QListWidgetItem(self.results_list)
            item.setData(_ROLE_HTML, card_html)
            item.setData(
                _ROLE_META,
                {
                    "clause_pk": clause_pk,
                    "spec_number": spec_num,
                    "version": version,
                    "clause_number": clause_num,
                    "clause_title": clause_title,
                    "snippet_text": snippet_raw,
                },
            )

        if total > self.MAX_VISIBLE_RESULTS:
            self.status_label.setText(
                f"Showing top {self.MAX_VISIBLE_RESULTS} of {total} matching clauses."
            )
        else:
            self.status_label.setText(f"{total} matching clause{'s' if total != 1 else ''}.")

    # --- Preview & Actions ---

    def _on_current_item_changed(self, current: Optional[QListWidgetItem], _prev) -> None:
        if current is not None:
            self._on_result_selected(current)

    def _on_result_selected(self, item: QListWidgetItem) -> None:
        meta = item.data(_ROLE_META)
        if not isinstance(meta, dict):
            return
        self._selected_meta = meta

        spec_num = meta.get("spec_number", "")
        version = str(meta.get("version", "")).lstrip("vV")
        clause_num = meta.get("clause_number", "")
        clause_title = meta.get("clause_title", "")
        clause_pk = int(meta.get("clause_pk") or 0)

        header = f"TS {spec_num} v{version} · Clause {clause_num}"
        if clause_title:
            header += f" — {clause_title}"
        self.preview_title_label.setText(header)

        rec = self.db.get_clause_content(clause_pk) if clause_pk > 0 else None
        full_text = (rec or {}).get("content") or meta.get("snippet_text") or ""
        self._selected_content = str(full_text)

        query = self.current_query()
        escaped = html.escape(self._selected_content)
        if query:
            try:
                pattern = re.compile(re.escape(html.escape(query)), re.IGNORECASE)
                escaped = pattern.sub(
                    lambda m: (
                        f'<span style="background-color:{C.WARNING_SOFT};'
                        f'font-weight:600;">{m.group(0)}</span>'
                    ),
                    escaped,
                )
            except re.error:
                pass

        escaped = escaped.replace("\n", "<br>")
        self.preview_browser.setHtml(
            f'<div style="line-height:1.45;color:{C.TEXT_PRIMARY};">{escaped}</div>'
        )
        self.preview_frame.setVisible(True)
        self.results_splitter.setSizes([240, 220])

    def _on_open_in_spec_search(self) -> None:
        if not self._selected_meta:
            return
        self.open_in_spec_search_requested.emit(
            self._selected_meta.get("spec_number", ""),
            self._selected_meta.get("version", ""),
            self._selected_meta.get("clause_number", ""),
        )

    def _on_open_in_specs(self) -> None:
        if not self._selected_meta:
            return
        self.open_in_specs_requested.emit(
            self._selected_meta.get("spec_number", ""),
            self._selected_meta.get("version", ""),
        )

    def _on_copy_citation(self) -> None:
        if not self._selected_meta:
            return
        spec_num = self._selected_meta.get("spec_number", "")
        version = str(self._selected_meta.get("version", "")).lstrip("vV")
        clause_num = self._selected_meta.get("clause_number", "")
        title = self._selected_meta.get("clause_title", "")
        citation = f"TS {spec_num} v{version} Clause {clause_num}"
        if title:
            citation += f" ({title})"
        payload = f"{citation}\n\n{self._selected_content}".strip()
        QApplication.clipboard().setText(payload)
        notify_success(self, "Citation Copied", f"Copied {citation} to clipboard.")

    def shutdown(self, wait_ms: int = 500) -> None:
        """Stop debounce timer and cooperatively shut down query worker(s)."""
        self._debounce_timer.stop()
        workers = [w for w in ([self._worker, *self._retired_workers]) if w is not None]
        for w in workers:
            if w.isRunning():
                w.requestInterruption()
                w.wait(wait_ms)
