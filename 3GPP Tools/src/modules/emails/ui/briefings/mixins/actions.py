"""Navigation, worker execution, streaming, and link-dispatch mixin for EmailBriefingsDialog."""

from __future__ import annotations

import html
from pathlib import Path
import re
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QColor, QCursor
from PyQt5.QtWidgets import QApplication, QFileDialog, QMenu, QTreeWidgetItem

from core.ai.llm_runner import ModelSpec, resolve_effective_model_spec
from core.ui.font_scale import scale_px
from core.ui.theme import C
from core.ui.toast import notify_info, notify_success
from modules.emails.core.ai_summary import (
    AiBriefingWorker,
    EmailSummaryRecord,
    EmailSummaryTaskWorker,
    SummaryScope,
    build_ai_handoff_packet,
    build_leaf_handoff_packet,
    normalize_summary_for_display,
)
from modules.emails.core.markdown_export import format_compact_timestamp
from modules.emails.ui.ai_summary import M365CopilotHandoffDialog, populate_summary_model_menu

from ..buckets import build_meeting_ai_buckets


def _natural_ai_sort_key(code: str) -> tuple:
    return tuple((0, int(tok)) if tok.isdigit() else (1, tok.lower()) for tok in re.split(r"([0-9]+)", str(code or "")))


def _compact_ai_nav_label(ai_code: str, raw_label: str) -> str:
    txt = str(raw_label or ai_code or "").strip()
    if " — " in txt:
        code_part, rest = txt.split(" — ", 1)
        parts = [p.strip() for p in re.split(r"\s+[—–-]\s+", rest) if p.strip()]
        tail = parts[-1] if parts else rest.strip()
        if len(tail) > 26:
            tail = tail[:24].rstrip() + "…"
        return f"{code_part} — {tail}"
    return (txt[:38] + "…") if len(txt) > 40 else txt


class BriefingsActionsMixin:
    """Handles AI/Selection tree loading, Map-Reduce workers, exports, and drill-down links."""

    def _populate_ai_nav_item(self, parent_node: QTreeWidgetItem, ai_code: str, *, is_followed: bool, select_kind: str, select_key: str) -> Optional[QTreeWidgetItem]:
        buckets = self._ai_buckets.get(ai_code, {})
        all_msgs = [m for b in buckets.values() for m in b.get("messages", [])]
        fresh = self.summary_store.compute_freshness("ai", ai_code, all_msgs)
        if not fresh.has_summary:
            dot, color, status_desc = "○", C.TEXT_MUTED, "Not generated yet"
        elif fresh.new_count > 0:
            dot, color, status_desc = f"◐ +{fresh.new_count}", C.WARNING_DARK, f"{fresh.new_count} new email(s) since last briefing"
        else:
            dot, color, status_desc = "●", C.SUCCESS_DARK, "Up to date"
        raw_lbl = self.ai_label_map.get(ai_code, ai_code)
        short_lbl = _compact_ai_nav_label(ai_code, raw_lbl)
        item = QTreeWidgetItem(parent_node, [f"{dot}  {short_lbl} ({len(buckets)})"])
        item.setData(0, Qt.UserRole, ("ai", ai_code))
        item.setForeground(0, QColor(color))
        if is_followed:
            f = item.font(0)
            f.setBold(True)
            item.setFont(0, f)
        item.setToolTip(
            0,
            f"{raw_lbl}\n{len(buckets)} discussion(s) · {fresh.total_messages} email(s)\n"
            f"Status: {status_desc}\nClick to view or synthesize briefing.",
        )
        return item if (select_kind == "ai" and select_key == ai_code) else None

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

        all_ais = set(self._ai_buckets.keys()) | set(ai_recs.keys())
        followed_list = sorted([a for a in all_ais if a in self.followed_ais], key=_natural_ai_sort_key)
        other_list = sorted([a for a in all_ais if a not in self.followed_ais], key=_natural_ai_sort_key)

        target_item: Optional[QTreeWidgetItem] = None
        first_ai_item: Optional[QTreeWidgetItem] = None

        if followed_list:
            f_root = QTreeWidgetItem(self.nav_tree, [f"★ Followed AIs ({len(followed_list)})"])
            f_root.setFlags(f_root.flags() & ~Qt.ItemIsSelectable)
            f_root.setExpanded(True)
            for ai_code in followed_list:
                matched = self._populate_ai_nav_item(f_root, ai_code, is_followed=True, select_kind=select_kind, select_key=select_key)
                if matched is not None:
                    target_item = matched
                if first_ai_item is None:
                    first_ai_item = f_root.child(f_root.childCount() - 1)

        if other_list or not followed_list:
            o_title = f"Other Agenda Items ({len(other_list)})" if followed_list else f"Agenda Items ({len(other_list)})"
            o_root = QTreeWidgetItem(self.nav_tree, [o_title])
            o_root.setFlags(o_root.flags() & ~Qt.ItemIsSelectable)
            o_root.setExpanded(True)
            for ai_code in other_list:
                matched = self._populate_ai_nav_item(o_root, ai_code, is_followed=False, select_kind=select_kind, select_key=select_key)
                if matched is not None:
                    target_item = matched
                if first_ai_item is None:
                    first_ai_item = o_root.child(o_root.childCount() - 1)

        sel_root = QTreeWidgetItem(self.nav_tree, [f"Saved Selections ({len(sel_recs)})"])
        sel_root.setFlags(sel_root.flags() & ~Qt.ItemIsSelectable)
        sel_root.setExpanded(True)
        self._sel_root = sel_root
        for rec in sel_recs:
            item = QTreeWidgetItem(sel_root, [f"• {rec.title} ({rec.message_count})"])
            item.setData(0, Qt.UserRole, ("selection", rec.scope_key))
            item.setToolTip(0, f"{rec.title}\nUpdated: {rec.updated_at}\nFocus: {rec.focus_prompt or '(none)'}")
            if select_kind == "selection" and select_key == rec.scope_key:
                target_item = item

        if target_item is None:
            target_item = first_ai_item or (sel_root.child(0) if sel_root.childCount() > 0 else None)

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
        self.error_banner.clear_error()
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
            prov_norm = str(rec.provider or "").strip().lower().replace("_", "-")
            if prov_norm == "m365-copilot":
                prov = "💠 M365 Copilot"
            else:
                prov = f"{rec.provider} · {rec.model or 'default'}" + (
                    f" · {rec.reasoning_effort}" if rec.reasoning_effort else ""
                )
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
            self.briefing_browser.setMarkdown(normalize_summary_for_display(rec.summary_md, rec.provider))
            self.btn_copy.setEnabled(True)
            self.btn_export_md.setEnabled(True)
        else:
            spec = self._active_synthesis_spec()
            disc_cnt = len(self._ai_buckets.get(key, {})) if kind == "ai" else 0
            msg_cnt = fresh.total_messages if fresh else 0
            self.freshness_lbl.setText(f"○ Not generated · {spec.short_badge()}")
            self.freshness_lbl.setStyleSheet(
                f"background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; "
                f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; padding: 2px 8px; font-size: 11px; font-weight: bold;"
            )
            self.btn_update.setText("✨ Generate Briefing")
            self.briefing_browser.setMarkdown(
                f"### No Briefing Generated Yet\n\n"
                f"**{disc_cnt} deduplicated discussion(s)** ({msg_cnt} emails) are indexed under this scope.\n\n"
                f"- **⚡ Generate Briefing** — runs local Ollama / CLI synthesis in-app (updates stale threads first).\n"
                f"- **💠 M365 Copilot ▾** — prepares a clean, uncut `.md` pack for Copilot in browser and lets you paste the result back.\n"
                f"- Expand **Underlying Discussions ({disc_cnt})** below to read individual TDoc or Tracked Thread emails directly."
            )
            self.btn_copy.setEnabled(False)
            self.btn_export_md.setEnabled(False)

    def _on_toggle_sidebar(self, visible: bool) -> None:
        vis = bool(visible)
        if hasattr(self, "_left_panel"):
            self._left_panel.setVisible(vis)
        if hasattr(self, "btn_toggle_sidebar"):
            self.btn_toggle_sidebar.setText("◀ AIs" if vis else "▶ AIs")
        if hasattr(self, "_settings"):
            self._settings.setValue("sidebar_visible", vis)

    def _render_ai_leaves_drawer(self, buckets: Dict[str, Dict[str, Any]]) -> None:
        if not buckets:
            self.leaves_frame.setVisible(False)
            return
        self.leaves_frame.setVisible(True)
        arrow = "▾" if self.leaves_browser.isVisible() else "▸"
        self.btn_toggle_leaves.setText(f"{arrow} Underlying Discussions ({len(buckets)})")
        lines: List[str] = []
        for b in buckets.values():
            stype = str(b.get("scope_type") or "thread")
            skey = str(b.get("scope_key") or "")
            msgs = b.get("messages") or []
            alias_tdocs = [str(t).strip().upper() for t in (b.get("alias_tdocs") or []) if str(t).strip()]
            fresh = self.summary_store.compute_freshness(stype, skey, msgs)
            badge = "○ none" if not fresh.has_summary else (f"◐ +{fresh.new_count} new" if fresh.new_count > 0 else "● current")
            if stype == "tdoc":
                primary_link = f"[{html.escape(skey)}](tdoc-emails:{html.escape(skey)})"
                table_part = f" · [Table ↗](tdoc-table:{html.escape(skey)})" if self._jump_tdoc_callback else ""
            else:
                primary_link = f"[{html.escape(skey)}](thread:{html.escape(skey)})"
                if alias_tdocs:
                    shown_tds = ", ".join(f"[{html.escape(td)}](tdoc-emails:{html.escape(td)})" for td in alias_tdocs[:4])
                    extra_td = f" (+{len(alias_tdocs) - 4})" if len(alias_tdocs) > 4 else ""
                    table_part = f" · TDocs: {shown_tds}{extra_td}"
                else:
                    table_part = ""
            lines.append(
                f"- {primary_link} — {html.escape(str(b.get('title') or skey))} "
                f"*({len(msgs)} msgs · {badge})*{table_part}"
            )
        self.leaves_browser.setMarkdown("\n".join(lines))

    def _on_toggle_leaves_drawer(self) -> None:
        vis = not self.leaves_browser.isVisible()
        self.leaves_browser.setVisible(vis)
        self.leaves_frame.setMaximumHeight(16777215 if vis else scale_px(34))
        self.right_vsplit.setSizes([scale_px(420), scale_px(180)] if vis else [scale_px(560), scale_px(34)])
        tup = self._active_scope_tuple()
        cnt = len(self._ai_buckets.get(tup[1], {})) if (tup and tup[0] == "ai") else 0
        arrow = "▾" if vis else "▸"
        self.btn_toggle_leaves.setText(f"{arrow} Underlying Discussions ({cnt})")

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
        packet_builder = None
        if kind == "ai":
            buckets = self._ai_buckets.get(key, {})
            if not buckets:
                notify_info(self, "No Emails", f"No email discussions found under AI {key}.")
                return None
            ai_label = self.ai_label_map.get(key, f"AI {key}")
            def _make_ai_pkt(pm: str):
                return build_ai_handoff_packet(
                    self.summary_store,
                    key,
                    ai_label,
                    buckets,
                    meeting_dir=self.meeting_dir,
                    wg=self.wg,
                    meeting_id=self.meeting_id,
                    tdoc_meta_map=self.tdoc_meta_map,
                    focus_prompt=focus_txt,
                    pack_mode=pm,
                )
            packet_builder = _make_ai_pkt
            packet = packet_builder("full_uncut")
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
            packet_builder=packet_builder,
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
        self.error_banner.clear_error()
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
        self.error_banner.clear_error()
        self.refresh_navigation(select_kind=record.scope_type, select_key=record.scope_key)
        self.briefing_saved.emit(record)

    def _on_worker_failed(self, err: str) -> None:
        self._worker = None
        self._set_running_ui(False)
        self.progress_lbl.setVisible(False)
        self.error_banner.show_error(str(err or "Briefing generation failed."))
        if self._active_record is not None:
            self.briefing_browser.setMarkdown(
                normalize_summary_for_display(self._active_record.summary_md, self._active_record.provider)
            )
        else:
            self.briefing_browser.setMarkdown(
                "*Generation did not complete. Click **🔁 Retry** above or try **💠 Try M365 Copilot**.*"
            )

    def _on_dismiss_error(self) -> None:
        self.error_banner.clear_error()
        self._load_selected_scope()

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
            clean_md = normalize_summary_for_display(self._active_record.summary_md, self._active_record.provider)
            QApplication.clipboard().setText(clean_md)
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
            clean_md = normalize_summary_for_display(self._active_record.summary_md, self._active_record.provider)
            Path(path).write_text(clean_md, encoding="utf-8")
            notify_success(self, "Briefing Exported", f"Saved to {Path(path).name}")

    def _open_tdoc_emails_for_tdoc(self, tdoc_id: str) -> None:
        clean_td = str(tdoc_id or "").strip().upper()
        if not clean_td:
            return
        cb = getattr(self, "_open_tdoc_emails_callback", None)
        if cb is not None:
            cb(clean_td)
            return
        gdb_p = getattr(self.general_email_db, "db_path", None) if self.general_email_db is not None else None
        if gdb_p:
            existing = self._child_email_dialogs.get(clean_td)
            if existing is not None:
                try:
                    existing.setWindowState(existing.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
                    existing.raise_()
                    existing.activateWindow()
                    return
                except RuntimeError:
                    self._child_email_dialogs.pop(clean_td, None)
            from modules.emails.ui.tdoc_emails_dialog import TDocEmailsDialog
            tmeta = (self.tdoc_meta_map or {}).get(clean_td) or {}
            dlg = TDocEmailsDialog(
                clean_td, [clean_td], gdb_p, wg=self.wg, parent=None,
                tdoc_info=tmeta or None, open_tdoc_callback=self._open_tdoc_callback,
            )
            self._child_email_dialogs[clean_td] = dlg
            dlg.finished.connect(lambda _=None, k=clean_td: self._child_email_dialogs.pop(k, None))
            dlg.show(); dlg.raise_(); dlg.activateWindow()
            return
        if self._jump_tdoc_callback:
            self._jump_tdoc_callback(clean_td)
        elif self._open_tdoc_callback:
            self._open_tdoc_callback(clean_td)

    def _on_anchor_clicked(self, url: QUrl) -> None:
        scheme = url.scheme().lower()
        target = url.path().strip() or url.toString().split(":", 1)[1].strip()
        if scheme in ("tdoc-emails", "tdoc") and target:
            self._open_tdoc_emails_for_tdoc(target.upper())
        elif scheme == "tdoc-table" and target:
            if self._jump_tdoc_callback:
                self._jump_tdoc_callback(target.upper())
            elif self._open_tdoc_callback:
                self._open_tdoc_callback(target.upper())
        elif scheme == "thread" and target:
            if self._open_thread_callback:
                self._open_thread_callback(target)
            elif self.tracked_db is not None:
                from modules.emails.ui.tracked_thread_dialog import TrackedThreadDialog
                dlg = TrackedThreadDialog(self.tracked_db, target, parent=None, open_tdoc_callback=self._open_tdoc_callback)
                self._child_email_dialogs[f"THR::{target}"] = dlg
                dlg.finished.connect(lambda _=None, k=f"THR::{target}": self._child_email_dialogs.pop(k, None))
                dlg.show(); dlg.raise_(); dlg.activateWindow()
        elif scheme in ("http", "https"):
            from PyQt5.QtGui import QDesktopServices
            QDesktopServices.openUrl(url)
