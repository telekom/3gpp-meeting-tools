"""Navigation, worker execution, streaming, and link-dispatch mixin for EmailBriefingsDialog."""

from __future__ import annotations

import html
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QApplication,
    QFileDialog,
    QMenu,
    QTreeWidgetItem,
)

from core.ai.llm_runner import ModelSpec, resolve_effective_model_spec
from core.ui.theme import C
from core.ui.toast import notify_success
from modules.emails.core.ai_summary import (
    AiBriefingWorker,
    EmailSummaryRecord,
    EmailSummaryTaskWorker,
    SummaryScope,
    build_ai_handoff_packet,
    build_leaf_handoff_packet,
)
from modules.emails.core.markdown_export import format_compact_timestamp
from modules.emails.ui.ai_summary import M365CopilotHandoffDialog, populate_summary_model_menu

from ..buckets import build_meeting_ai_buckets


class BriefingsActionsMixin:
    """Handles AI/Selection tree loading, Map-Reduce workers, exports, and drill-down links."""

    def refresh_navigation(self, select_kind: str = "", select_key: str = "") -> None:
        self._ai_buckets = build_meeting_ai_buckets(
            tracked_db=self.tracked_db,
            general_email_db=self.general_email_db,
            meeting_id=self.meeting_id,
            tdoc_ai_lookup=self.tdoc_ai_lookup,
            tdoc_meta_map=self.tdoc_meta_map,
        )
        ai_recs = {r.scope_key: r for r in self.summary_store.list_summaries("ai")}
        sel_recs = self.summary_store.list_summaries("selection")

        self.nav_tree.blockSignals(True)
        self.nav_tree.clear()

        ai_root = QTreeWidgetItem(self.nav_tree, ["📌 Agenda Items"])
        ai_root.setFlags(ai_root.flags() & ~Qt.ItemIsSelectable)
        ai_root.setExpanded(True)

        all_ais = sorted(
            set(self._ai_buckets.keys()) | set(ai_recs.keys()),
            key=lambda a: (0 if a in self.followed_ais else 1, a),
        )
        target_item: Optional[QTreeWidgetItem] = None

        for ai_code in all_ais:
            buckets = self._ai_buckets.get(ai_code, {})
            all_msgs = [m for b in buckets.values() for m in b.get("messages", [])]
            fresh = self.summary_store.compute_freshness("ai", ai_code, all_msgs)
            if not fresh.has_summary:
                pill = "⚪"
            elif fresh.new_count > 0:
                pill = f"🟡+{fresh.new_count}"
            else:
                pill = "✅"
            star = "👀 " if ai_code in self.followed_ais else ""
            raw_lbl = self.ai_label_map.get(ai_code, ai_code)
            short_lbl = (raw_lbl[:38] + "…") if len(raw_lbl) > 40 else raw_lbl
            item = QTreeWidgetItem(ai_root, [f"{pill} {star}{short_lbl} ({len(buckets)})"])
            item.setData(0, Qt.UserRole, ("ai", ai_code))
            item.setToolTip(0, f"{raw_lbl}\nTopics: {len(buckets)} | Emails: {fresh.total_messages}")
            if select_kind == "ai" and select_key == ai_code:
                target_item = item

        sel_root = QTreeWidgetItem(self.nav_tree, ["📑 Saved Selections"])
        sel_root.setFlags(sel_root.flags() & ~Qt.ItemIsSelectable)
        sel_root.setExpanded(True)
        for rec in sel_recs:
            item = QTreeWidgetItem(sel_root, [f"📑 {rec.title} ({rec.message_count})"])
            item.setData(0, Qt.UserRole, ("selection", rec.scope_key))
            item.setToolTip(0, f"{rec.title}\nUpdated: {rec.updated_at}\nFocus: {rec.focus_prompt or '(none)'}")
            if select_kind == "selection" and select_key == rec.scope_key:
                target_item = item

        if target_item is None and ai_root.childCount() > 0:
            target_item = ai_root.child(0)
        elif target_item is None and sel_root.childCount() > 0:
            target_item = sel_root.child(0)

        self.nav_tree.blockSignals(False)
        if target_item is not None:
            self.nav_tree.setCurrentItem(target_item)
            self._load_selected_scope()
        else:
            self._render_empty_scope()

    def _filter_nav_tree(self, text: str) -> None:
        q = (text or "").strip().lower()
        root = self.nav_tree.invisibleRootItem()
        for g_idx in range(root.childCount()):
            grp = root.child(g_idx)
            for c_idx in range(grp.childCount()):
                ch = grp.child(c_idx)
                match = (not q) or (q in ch.text(0).lower()) or (q in ch.toolTip(0).lower())
                ch.setHidden(not match)

    def _on_nav_selection_changed(self, current: Optional[QTreeWidgetItem], _prev: Optional[QTreeWidgetItem]) -> None:
        if current is None or not current.data(0, Qt.UserRole):
            return
        self._load_selected_scope()

    def _active_scope_tuple(self) -> Optional[tuple[str, str]]:
        cur = self.nav_tree.currentItem()
        if cur is None:
            return None
        data = cur.data(0, Qt.UserRole)
        if isinstance(data, tuple) and len(data) == 2:
            return str(data[0]), str(data[1])
        return None

    def _render_empty_scope(self) -> None:
        self.scope_title_lbl.setText("No AI discussions or saved selections found yet")
        self.freshness_lbl.setText("Ready")
        self.briefing_browser.setMarkdown(
            "*Sync meeting emails or select rows in the Unified Email Inbox to generate AI briefings.*"
        )
        self.btn_update.setEnabled(False)
        self.btn_rebuild.setEnabled(False)
        self.btn_m365_copilot.setEnabled(False)
        self.btn_copy.setEnabled(False)
        self.btn_export_md.setEnabled(False)

    def _load_selected_scope(self) -> None:
        tup = self._active_scope_tuple()
        if tup is None:
            self._render_empty_scope()
            return
        kind, key = tup
        rec = self.summary_store.get_summary(kind, key)
        self._active_record = rec

        if kind == "ai":
            buckets = self._ai_buckets.get(key, {})
            all_msgs = [m for b in buckets.values() for m in b.get("messages", [])]
            fresh = self.summary_store.compute_freshness("ai", key, all_msgs)
            ai_label = self.ai_label_map.get(key, f"AI {key}")
            self.scope_title_lbl.setText(f"🧭 {ai_label}")
            self.btn_update.setEnabled(bool(buckets) and not self._is_running())
            self.btn_rebuild.setEnabled(bool(buckets) and not self._is_running())
            self.btn_m365_copilot.setEnabled(bool(buckets) and not self._is_running())
            if rec and rec.focus_prompt:
                self.focus_prompt_edit.setText(rec.focus_prompt)
            if not self.btn_focus_toggle.isChecked():
                self.focus_row_widget.setVisible(False)
            self._render_ai_leaves_drawer(buckets)
        else:
            all_msgs = self._pending_selection_messages if (key == self._pending_selection_key) else []
            fresh = self.summary_store.compute_freshness("selection", key, all_msgs) if all_msgs else None
            title = rec.title if rec else (self._pending_selection_title or key)
            self.scope_title_lbl.setText(f"📑 {title}")
            can_run = bool(all_msgs) and not self._is_running()
            self.btn_update.setEnabled(can_run)
            self.btn_rebuild.setEnabled(can_run)
            self.btn_m365_copilot.setEnabled(can_run)
            self.focus_row_widget.setVisible(True)
            self.btn_focus_toggle.setChecked(True)
            if rec and rec.focus_prompt:
                self.focus_prompt_edit.setText(rec.focus_prompt)
            self._render_ai_leaves_drawer({})

        if rec is not None:
            cutoff_disp = format_compact_timestamp(rec.cutoff_utc) if rec.cutoff_utc else "unknown"
            prov = f"{rec.provider} · {rec.model or 'default'}" + (f" · {rec.reasoning_effort}" if rec.reasoning_effort else "")
            if fresh and fresh.new_count > 0:
                self.freshness_lbl.setText(f"🟡 +{fresh.new_count} new since {cutoff_disp} · {prov}")
                self.freshness_lbl.setStyleSheet(
                    f"background-color: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; "
                    f"border: 1px solid {C.WARNING_BORDER}; border-radius: 4px; padding: 2px 8px; font-size: 11px; font-weight: bold;"
                )
                self.btn_update.setText(f"⚡ Update (+{fresh.new_count})")
            else:
                self.freshness_lbl.setText(f"✅ Up to date ({rec.message_count} msgs) · {prov}")
                self.freshness_lbl.setStyleSheet(
                    f"background-color: {C.SUCCESS_SOFT}; color: {C.SUCCESS_DARK}; "
                    f"border: 1px solid {C.SUCCESS_BORDER_MEDIUM}; border-radius: 4px; padding: 2px 8px; font-size: 11px; font-weight: bold;"
                )
                self.btn_update.setText("⚡ Refresh Briefing")
            self.briefing_browser.setMarkdown(rec.summary_md)
            self.btn_copy.setEnabled(True)
            self.btn_export_md.setEnabled(True)
        else:
            spec = self._active_synthesis_spec()
            self.freshness_lbl.setText(f"⚪ Not generated · {spec.short_badge()}")
            self.freshness_lbl.setStyleSheet(
                f"background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; "
                f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; padding: 2px 8px; font-size: 11px; font-weight: bold;"
            )
            self.btn_update.setText("✨ Generate Briefing")
            self.briefing_browser.setMarkdown(
                "*No briefing saved for this scope yet. Click **✨ Generate Briefing** above to synthesize across all underlying discussions.*"
            )
            self.btn_copy.setEnabled(False)
            self.btn_export_md.setEnabled(False)

    def _render_ai_leaves_drawer(self, buckets: Dict[str, Dict[str, Any]]) -> None:
        if not buckets:
            self.leaves_frame.setVisible(False)
            return
        self.leaves_frame.setVisible(True)
        arrow = "▾" if self.leaves_browser.isVisible() else "▸"
        self.btn_toggle_leaves.setText(f"{arrow} Underlying Thread / TDoc Summaries ({len(buckets)})")
        lines: List[str] = []
        for b in buckets.values():
            stype = str(b.get("scope_type") or "thread")
            skey = str(b.get("scope_key") or "")
            msgs = b.get("messages") or []
            fresh = self.summary_store.compute_freshness(stype, skey, msgs)
            badge = "⚪ none" if not fresh.has_summary else (f"🟡 +{fresh.new_count}" if fresh.new_count > 0 else "✅ current")
            href = f"tdoc:{html.escape(skey)}" if stype == "tdoc" else f"thread:{html.escape(skey)}"
            lines.append(
                f"- [{html.escape(skey)}]({href}) — {html.escape(str(b.get('title') or skey))} "
                f"*({len(msgs)} msgs · {badge})*"
            )
        self.leaves_browser.setMarkdown("\n".join(lines))

    def _on_toggle_leaves_drawer(self) -> None:
        vis = not self.leaves_browser.isVisible()
        self.leaves_browser.setVisible(vis)
        tup = self._active_scope_tuple()
        cnt = len(self._ai_buckets.get(tup[1], {})) if (tup and tup[0] == "ai") else 0
        arrow = "▾" if vis else "▸"
        self.btn_toggle_leaves.setText(f"{arrow} Underlying Thread / TDoc Summaries ({cnt})")

    def _active_synthesis_spec(self) -> ModelSpec:
        return resolve_effective_model_spec(role="synthesis", override=self._synth_model_override)

    def _populate_model_menu(self) -> None:
        populate_summary_model_menu(
            self.model_menu,
            role="synthesis",
            active_spec=self._active_synthesis_spec(),
            on_selected=self._on_model_selected,
            on_defaults_saved=self._load_selected_scope,
            parent=self,
        )

    def _on_model_selected(self, spec: ModelSpec) -> None:
        self._synth_model_override = spec
        self._load_selected_scope()

    def _is_running(self) -> bool:
        return bool(self._worker is not None and self._worker.isRunning())

    def _set_running_ui(self, running: bool) -> None:
        self.btn_stop.setVisible(running)
        self.btn_update.setEnabled(not running)
        self.btn_rebuild.setEnabled(not running)
        self.btn_model.setEnabled(not running)
        self.btn_m365_copilot.setEnabled(not running)
        self.progress_lbl.setVisible(running)

    def _open_m365_copilot_handoff(self, *, auto_launch: bool = True):
        tup = self._active_scope_tuple()
        if tup is None:
            return None
        kind, key = tup
        focus_txt = self.focus_prompt_edit.text().strip()
        if kind == "ai":
            buckets = self._ai_buckets.get(key, {})
            if not buckets:
                notify_info(self, "No Emails", f"No email discussions found under AI {key}.")
                return None
            ai_label = self.ai_label_map.get(key, f"AI {key}")
            packet = build_ai_handoff_packet(
                self.summary_store,
                key,
                ai_label,
                buckets,
                meeting_dir=self.meeting_dir,
                wg=self.wg,
                meeting_id=self.meeting_id,
                tdoc_meta_map=self.tdoc_meta_map,
                focus_prompt=focus_txt,
            )
        else:
            msgs = self._pending_selection_messages if key == self._pending_selection_key else []
            if not msgs:
                notify_info(self, "Selection Messages Unavailable", "Re-select those rows from the Unified Inbox to rebuild the selection prompt.")
                return None
            title = self._pending_selection_title or (self._active_record.title if self._active_record else key)
            packet = build_leaf_handoff_packet(
                self.summary_store,
                SummaryScope(scope_type="selection", scope_key=key, title=title, focus_prompt=focus_txt),
                msgs,
                meeting_dir=self.meeting_dir,
                wg=self.wg,
                meeting_id=self.meeting_id,
                mode="full",
            )

        if self._m365_dialog is not None:
            try:
                self._m365_dialog.close()
            except Exception:
                pass
            self._m365_dialog = None

        dlg = M365CopilotHandoffDialog(
            self.summary_store,
            packet,
            auto_launch=auto_launch,
            meeting_dir=self.meeting_dir,
            owner=self,
        )
        dlg.summary_saved.connect(self._on_worker_completed)
        dlg.destroyed.connect(lambda _=None: setattr(self, "_m365_dialog", None))
        self._m365_dialog = dlg
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
        return dlg

    def _run_active_briefing(self, *, mode: str = "auto") -> None:
        if self._is_running():
            return
        tup = self._active_scope_tuple()
        if tup is None:
            return
        kind, key = tup
        focus_txt = self.focus_prompt_edit.text().strip()
        self._stream_chunks = []
        self.briefing_browser.clear()
        self.progress_lbl.setText("⏳ Starting synthesis...")
        self._set_running_ui(True)

        if kind == "ai":
            buckets = self._ai_buckets.get(key, {})
            ai_label = self.ai_label_map.get(key, f"AI {key}")
            self._worker = AiBriefingWorker(
                self.summary_store.db_path,
                key,
                ai_label,
                buckets,
                meeting_dir=self.meeting_dir,
                wg=self.wg,
                meeting_id=self.meeting_id,
                mode=mode,
                leaf_model_spec=resolve_effective_model_spec("leaf"),
                synthesis_model_spec=self._active_synthesis_spec(),
                tdoc_meta_map=self.tdoc_meta_map,
                focus_prompt=focus_txt,
                parent=self,
            )
        else:
            msgs = self._pending_selection_messages if key == self._pending_selection_key else []
            title = self._pending_selection_title or (self._active_record.title if self._active_record else key)
            self._worker = EmailSummaryTaskWorker(
                self.summary_store.db_path,
                SummaryScope(scope_type="selection", scope_key=key, title=title, focus_prompt=focus_txt),
                msgs,
                meeting_dir=self.meeting_dir,
                wg=self.wg,
                meeting_id=self.meeting_id,
                mode=mode,
                model_spec=self._active_synthesis_spec(),
                parent=self,
            )

        self._worker.progress_changed.connect(self._on_worker_progress)
        self._worker.token_received.connect(self._on_worker_token)
        self._worker.summary_completed.connect(self._on_worker_completed)
        self._worker.summary_failed.connect(self._on_worker_failed)
        self._worker.finished.connect(lambda: self._set_running_ui(False))
        self._worker.start()

    def _on_worker_progress(self, msg: str) -> None:
        self.progress_lbl.setText(f"⏳ {msg}")

    def _on_worker_token(self, chunk: str) -> None:
        if chunk:
            self._stream_chunks.append(chunk)
            self.briefing_browser.insertPlainText(chunk)
            self.briefing_browser.ensureCursorVisible()

    def _on_worker_completed(self, record: EmailSummaryRecord) -> None:
        self._worker = None
        self._set_running_ui(False)
        self.refresh_navigation(select_kind=record.scope_type, select_key=record.scope_key)
        self.briefing_saved.emit(record)

    def _on_worker_failed(self, err: str) -> None:
        self._worker = None
        self._set_running_ui(False)
        self.progress_lbl.setVisible(True)
        self.progress_lbl.setText(f"⚠️ {err}")
        if self._active_record is not None:
            self.briefing_browser.setMarkdown(self._active_record.summary_md)

    def _cancel_active_worker(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.request_cancel()
            self.progress_lbl.setText("Cancelling...")

    def _on_nav_context_menu(self, pos) -> None:
        item = self.nav_tree.itemAt(pos)
        if item is None or not item.data(0, Qt.UserRole):
            return
        kind, key = item.data(0, Qt.UserRole)
        if kind != "selection":
            return
        menu = QMenu(self)
        act_del = menu.addAction("🗑️ Delete Saved Selection")
        if menu.exec_(QCursor.pos()) == act_del:
            self.summary_store.delete_summary("selection", key)
            self.refresh_navigation()

    def _on_copy_briefing(self) -> None:
        if self._active_record and self._active_record.summary_md:
            QApplication.clipboard().setText(self._active_record.summary_md)
            notify_success(self, "Briefing Copied", "Copied briefing Markdown to clipboard.")

    def _on_export_briefing_md(self) -> None:
        if not self._active_record or not self._active_record.summary_md:
            return
        safe_key = "".join(c if c.isalnum() or c in "-_." else "_" for c in self._active_record.scope_key)
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export AI Briefing",
            f"briefing_{self._active_record.scope_type}_{safe_key}.md",
            "Markdown (*.md)",
        )
        if path:
            Path(path).write_text(self._active_record.summary_md, encoding="utf-8")
            notify_success(self, "Briefing Exported", f"Saved to {Path(path).name}")

    def _on_anchor_clicked(self, url: QUrl) -> None:
        scheme = url.scheme().lower()
        target = url.path().strip() or url.toString().split(":", 1)[1].strip()
        if scheme == "tdoc" and target:
            if self._jump_tdoc_callback:
                self._jump_tdoc_callback(target.upper())
            elif self._open_tdoc_callback:
                self._open_tdoc_callback(target.upper())
        elif scheme == "thread" and target and self._open_thread_callback:
            self._open_thread_callback(target)
        elif scheme in ("http", "https"):
            from PyQt5.QtGui import QDesktopServices
            QDesktopServices.openUrl(url)
