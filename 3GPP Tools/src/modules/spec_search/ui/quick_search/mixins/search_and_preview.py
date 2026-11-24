"""Async FTS5 query execution, result rendering, preview highlighting, and deep links."""

from __future__ import annotations

import html
import logging
import re
from typing import Any, Dict, List, Optional

from PyQt5.QtWidgets import QApplication, QListWidgetItem

from core.ui.theme import C, T
from core.ui.toast import notify_success
from modules.spec_search.core.spec_search_db import parse_version_tuple
from modules.spec_search.core.spec_search_threads import SpecSearchQueryWorker
from ..delegate import _ROLE_HTML, _ROLE_META, _format_snippet_html

logger = logging.getLogger(__name__)


class QuickSearchLogicMixin:
    """Handles version scoping, debounced query workers, result preview, and shutdown."""

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
                rest = spec_q[len(prefix) :].strip()
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
                self.status_label.setText("No specifications indexed in Spec Search yet.")
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
                f"{snippet_html}</div>"
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
