"""Selection handling, inline assignment bar population, and Markdown reader rendering mixin."""

from __future__ import annotations

import html
from typing import List
import urllib.parse

from PyQt5.QtCore import QTimer

from core.ui.theme import C
from modules.emails.core.emeeting_signals import (
    build_message_jump_strip_html,
    extract_base_tdoc_id,
)
from modules.emails.core.thread_markdown_exporter import (
    export_thread_to_markdown,
    find_unread_heading_blocks,
    format_compact_timestamp,
    format_emeeting_body_with_callouts,
    format_sender_display,
    inject_html_tables_into_body,
    partition_email_attachments,
    scroll_browser_to_block,
)
from modules.emails.core.tracked_threads_service import normalize_subject


class InboxDetailRenderingMixin:
    """Renders single/multi selection summaries, assignment bars, and thread/message Markdown previews."""

    def _on_selection_changed(self):
        self._flush_pending_note_save(reload_view=False)
        rows = sorted({idx.row() for idx in self.table.selectionModel().selectedRows()}) if self.table.selectionModel() else []
        selected = [self._visible_items[r] for r in rows if 0 <= r < len(self._visible_items)]
        self._selected_items = selected
        if not selected:
            self._current_item = None
            self._render_empty_detail()
            return
        self._current_item = selected[0]
        if len(selected) == 1:
            self._render_detail(selected[0])
        else:
            self._render_multi_detail(selected)

    def _render_empty_detail(self):
        self.detail_subject_lbl.setText("No matching emails in this filter")
        self.detail_meta_lbl.setText("")
        self.detail_links_lbl.setText("")
        if hasattr(self, "detail_jump_strip_lbl"):
            self.detail_jump_strip_lbl.setText("")
            self.detail_jump_strip_lbl.setVisible(False)
        self.assign_frame.setVisible(False)
        self.note_frame.setVisible(False)
        self.preview_browser.setMarkdown("*Select another filter tab or adjust search criteria.*")
        for btn_name in (
            "btn_mark_next",
            "btn_toggle_read",
            "btn_toggle_flag",
            "btn_toggle_objected",
            "btn_reply",
            "btn_copy_start_template",
            "btn_open_outlook",
            "btn_open_context",
            "btn_jump_tdoc_table",
            "btn_llm_export",
            "btn_sync_selected",
            "btn_toggle_note",
            "btn_merge_selected",
        ):
            btn = getattr(self, btn_name, None)
            if btn is not None:
                btn.setEnabled(False)
        self.btn_merge_selected.setVisible(False)
        if hasattr(self, "btn_jump_tdoc_table"):
            self.btn_jump_tdoc_table.setVisible(False)
        if hasattr(self, "detail_status_combo"):
            self.detail_status_combo.setEnabled(False)
        self._update_selected_sync_controls(0)

    def _on_toggle_assign_bar(self, checked: bool):
        self._show_assign_bar = bool(checked)
        self.assign_frame.setVisible(self._show_assign_bar)
        self.btn_toggle_assign_bar.setText("🏷️ AI / WI ▴" if self._show_assign_bar else "🏷️ AI / WI ▾")
        try:
            self._settings.setValue("show_assign_bar", self._show_assign_bar)
            self._settings.sync()
        except Exception:
            pass

    def _refresh_assign_sub_suggestions(self, selected_ais: List[str]):
        allowed: List[str] = []
        seen = set()
        for ai in selected_ais:
            for s in self._ai_to_sub_items.get(ai, []):
                if s not in seen:
                    seen.add(s)
                    allowed.append(s)
        self.assign_subs_pills.set_suggestions(allowed)
        if not selected_ais:
            self.assign_subs_pills.set_empty_dropdown_label("Assign AI")
        elif not allowed:
            self.assign_subs_pills.set_empty_dropdown_label("No sub-topics")

    def _populate_assign_bar(self, it: dict):
        self._updating_assign_bar = True
        try:
            self.assign_frame.setVisible(self._show_assign_bar)
            ais = list(it.get("linked_ais") or [])
            subs = list(it.get("linked_sub_items") or [])
            wis = list(it.get("linked_wis") or [])
            self._multi_base_common = {"ais": [], "subs": [], "wis": []}
            self._refresh_assign_sub_suggestions(ais)
            self.assign_ais_pills.setText(", ".join(ais))
            self.assign_subs_pills.setText(", ".join(subs))
            self.assign_wis_pills.setText(", ".join(wis))
            self.btn_pin_inline.setVisible(not it.get("in_tracked_db") and self.tracked_db is not None)
            self.btn_pin_inline.setEnabled(not it.get("in_tracked_db") and self.tracked_db is not None)
            self.btn_auto_assign.setEnabled(self.tracked_db is not None)
        finally:
            self._updating_assign_bar = False

    def _populate_assign_bar_multi(self, selected: List[dict]):
        self._updating_assign_bar = True
        try:
            self.assign_frame.setVisible(self._show_assign_bar)
            common_ais = [
                a for a in (selected[0].get("linked_ais") or [])
                if all(a in (it.get("linked_ais") or []) for it in selected[1:])
            ]
            common_subs = [
                s for s in (selected[0].get("linked_sub_items") or [])
                if all(s in (it.get("linked_sub_items") or []) for it in selected[1:])
            ]
            common_wis = [
                w for w in (selected[0].get("linked_wis") or [])
                if all(w in (it.get("linked_wis") or []) for it in selected[1:])
            ]
            union_ais = list(dict.fromkeys(
                a for it in selected for a in (it.get("linked_ais") or []) if a
            ))
            self._multi_base_common = {
                "ais": list(common_ais),
                "subs": list(common_subs),
                "wis": list(common_wis),
            }
            self._refresh_assign_sub_suggestions(union_ais)
            self.assign_ais_pills.setText(", ".join(common_ais))
            self.assign_subs_pills.setText(", ".join(common_subs))
            self.assign_wis_pills.setText(", ".join(common_wis))
            has_unpinned = any(not it.get("in_tracked_db") for it in selected)
            self.btn_pin_inline.setVisible(has_unpinned and self.tracked_db is not None)
            self.btn_pin_inline.setEnabled(has_unpinned and self.tracked_db is not None)
            self.btn_auto_assign.setEnabled(self.tracked_db is not None)
        finally:
            self._updating_assign_bar = False

    def _render_multi_detail(self, selected: List[dict]):
        if self._view_mode == "TDOCS":
            noun = "TDocs"
        elif self._view_mode == "THREADS":
            noun = "Threads"
        else:
            noun = "Emails"
        unread_cnt = sum(1 for it in selected if not it["is_read"])
        flagged_cnt = sum(1 for it in selected if it["is_flagged"])
        unpinned_cnt = sum(1 for it in selected if not it.get("in_tracked_db"))
        pinned_cnt = len(selected) - unpinned_cnt

        self.detail_subject_lbl.setText(f"📑 {len(selected)} Selected {noun}")
        self.detail_meta_lbl.setText(
            f"{unread_cnt} unread  |  {flagged_cnt} flagged  |  "
            f"{pinned_cnt} tracked, {unpinned_cnt} unpinned"
        )
        if hasattr(self, "detail_jump_strip_lbl"):
            self.detail_jump_strip_lbl.setText("")
            self.detail_jump_strip_lbl.setVisible(False)

        cleaned_tdocs = list(dict.fromkeys(
            self._clean_tdoc_id(td)
            for it in selected
            for td in (it.get("linked_tdocs") or [])
            if self._clean_tdoc_id(td)
        ))
        links_html = [f"Batch selection ({len(selected)} items) — adding AIs/Sub-Topics/WIs applies to all selected items"]
        if cleaned_tdocs:
            if self._open_tdoc_callback:
                td_anchors = ", ".join(
                    f'<a href="tdoc:{html.escape(td)}" '
                    f'style="color: {C.INFO}; font-weight: 600; text-decoration: underline;">'
                    f'{html.escape(td)}</a>'
                    for td in cleaned_tdocs[:10]
                )
            else:
                td_anchors = ", ".join(html.escape(td) for td in cleaned_tdocs[:10])
            if len(cleaned_tdocs) > 10:
                td_anchors += f" (+{len(cleaned_tdocs) - 10} more)"
            links_html.append(f"TDocs: {td_anchors}")
        self.detail_links_lbl.setText(" &nbsp;|&nbsp; ".join(links_html))
        self.detail_links_lbl.setVisible(True)

        self._populate_assign_bar_multi(selected)
        self._populate_status_and_note_bars(selected)

        self.btn_mark_next.setEnabled(False)
        for bname in ("btn_reply", "btn_copy_start_template", "btn_open_outlook", "btn_open_context", "btn_jump_tdoc_table", "btn_llm_export"):
            b = getattr(self, bname, None)
            if b is not None:
                b.setEnabled(False)
                b.setVisible(False)
        self.btn_toggle_read.setEnabled(True)
        self.btn_toggle_read.setText(
            f"✓ Read ({unread_cnt})" if unread_cnt > 0 else f"🔵 Unread ({len(selected)})"
        )
        self.btn_toggle_flag.setEnabled(True)
        unflagged_cnt = len(selected) - flagged_cnt
        self.btn_toggle_flag.setText(
            f"🚩 ({unflagged_cnt})" if unflagged_cnt > 0 else f"🏳️ ({len(selected)})"
        )
        self._update_selected_sync_controls(len(selected))
        merge_targets = self._distinct_mergeable_targets(selected)
        can_merge = bool(self.tracked_db is not None and len(merge_targets) >= 2)
        self.btn_merge_selected.setVisible(can_merge)
        self.btn_merge_selected.setEnabled(can_merge)
        if can_merge:
            self.btn_merge_selected.setText(f"🔗 Merge ({len(merge_targets)})")

        md_lines = [f"### Batch Triage ({len(selected)} selected {noun.lower()})", ""]
        md_lines.append("Use the **AIs**, **Sub-Topics**, **WIs**, **Auto-Assign**, **Pin**, **Merge**, **Mark Read**, **Flag**, or **Objected** controls above to update all selected items together.\n")
        for idx, it in enumerate(selected, 1):
            subj = str(it.get("subject") or "(No Subject)")
            comp = str(it.get("company") or "Unknown")
            ais = ", ".join(it.get("linked_ais") or []) or "Unassigned"
            wis = ", ".join(it.get("linked_wis") or []) or "None"
            status_tag = "🧵 Tracked" if it.get("in_tracked_db") else "📄 Unpinned"
            md_lines.append(f"{idx}. **{subj}** — *{comp}* · {status_tag} · AI: `{ais}` · WI: `{wis}`")
        self.preview_browser.setMarkdown("\n".join(md_lines))

    def _render_detail(self, it: dict):
        is_tdoc_bucket = it.get("_group_kind") == "TDOC" and bool(it.get("_primary_tdoc"))
        primary_tdoc = str(it.get("_primary_tdoc") or "").strip().upper()
        if not primary_tdoc and it.get("linked_tdocs"):
            primary_tdoc = extract_base_tdoc_id(it["linked_tdocs"][0])
        sig = it.get("_signals") or {}

        subject_text = str(it.get("subject") or "(No Subject)")
        if is_tdoc_bucket:
            type_badge = f" [{it['_tdoc_type']}]" if it.get("_tdoc_type") else ""
            self.detail_subject_lbl.setText(f"📄 {primary_tdoc}{type_badge} — {subject_text}")
        else:
            self.detail_subject_lbl.setText(subject_text)

        if it.get("_is_group"):
            raw_comps = [str(c).strip() for c in (it.get("_companies") or []) if str(c).strip()]
            comp_disp = ", ".join(raw_comps[:5]) + (f" (+{len(raw_comps) - 5})" if len(raw_comps) > 5 else "")
            if not comp_disp:
                comp_disp = it.get("company") or "Unknown"
            meta_parts = []
            if is_tdoc_bucket and it.get("_tdoc_source"):
                meta_parts.append(f"Author: {it['_tdoc_source']}")
            if is_tdoc_bucket and it.get("_tdoc_status"):
                meta_parts.append(f"Status: {it['_tdoc_status']}")
            meta_parts.append(f"{it.get('_total_count', 1)} msgs ({it.get('_unread_count', 0)} unread)")
            meta_parts.append(format_compact_timestamp(it.get("date_str")))
            meta_parts.append(f"Contributors: {comp_disp}")
            self.detail_meta_lbl.setText(" · ".join(meta_parts))
            self.detail_meta_lbl.setToolTip(", ".join(raw_comps) if raw_comps else comp_disp)
        else:
            sender_disp = format_sender_display(
                it.get("sender_name") or it.get("sender_email") or "Unknown",
                it.get("company") or "",
            )
            badges = []
            if not it["is_read"]:
                badges.append("🔵 NEW")
            if it["is_flagged"]:
                badges.append("🚩")
            badge_str = f" · {' '.join(badges)}" if badges else " · ✓"
            self.detail_meta_lbl.setText(
                f"{sender_disp} · {format_compact_timestamp(it.get('date_str'))}{badge_str}"
            )
            self.detail_meta_lbl.setToolTip("")

        links_html = []
        if it.get("in_tracked_db") and it.get("thread_title") and not is_tdoc_bucket:
            tt = str(it["thread_title"]).strip()
            if tt.lower() == subject_text.strip().lower():
                links_html.append("🧵 Tracked")
            else:
                links_html.append(f"🧵 {html.escape(tt)}")
        elif not is_tdoc_bucket:
            links_html.append("📄 Unpinned")

        cleaned_tdocs = list(dict.fromkeys(
            self._clean_tdoc_id(td) for td in (it.get("linked_tdocs") or []) if self._clean_tdoc_id(td)
        ))
        if cleaned_tdocs:
            if self._open_tdoc_callback:
                td_anchors = ", ".join(
                    f'<a href="tdoc:{html.escape(td)}" '
                    f'style="color: {C.INFO}; font-weight: 600; text-decoration: underline;">'
                    f'{html.escape(td)}</a>'
                    for td in cleaned_tdocs[:6]
                )
            else:
                td_anchors = ", ".join(html.escape(td) for td in cleaned_tdocs[:6])
            if len(cleaned_tdocs) > 6:
                td_anchors += f" (+{len(cleaned_tdocs) - 6})"
            links_html.append(f"TDocs: {td_anchors}")

        rev_tokens = sig.get("rev_tokens") or []
        if rev_tokens:
            if self._open_tdoc_callback:
                rev_anchors = ", ".join(
                    f'<a href="tdoc:{html.escape(rt.upper())}" '
                    f'style="color: {C.SUCCESS_DARK}; font-weight: bold; text-decoration: underline;">'
                    f'🆕 {html.escape(rt)}</a>'
                    for rt in rev_tokens[:5]
                )
            else:
                rev_anchors = ", ".join(f"🆕 {html.escape(rt)}" for rt in rev_tokens[:5])
            links_html.append(f"Revs: {rev_anchors}")

        if sig.get("is_objected"):
            if sig.get("unparked_by_new_rev"):
                links_html.append(
                    f'<span style="color:{C.WARNING_DARK}; font-weight:bold;">'
                    f'🛑→🆕 Unparked ({html.escape(str(sig.get("max_rev_tag") or ""))} after objection)</span>'
                )
            else:
                rev_at = f" at {sig['objected_at_rev']}" if sig.get("objected_at_rev") else ""
                links_html.append(
                    f'<span style="color:{C.DANGER}; font-weight:bold;">🛑 Objected / Parked{html.escape(rev_at)}</span>'
                )

        self.detail_links_lbl.setText(" · ".join(links_html))
        self.detail_links_lbl.setVisible(bool(links_html))

        if hasattr(self, "detail_jump_strip_lbl"):
            if it.get("_is_group"):
                jump_html = build_message_jump_strip_html(
                    it.get("_group_messages") or [],
                    primary_tdoc=primary_tdoc,
                    newest_first=True,
                )
                self.detail_jump_strip_lbl.setText(jump_html)
                self.detail_jump_strip_lbl.setVisible(bool(jump_html))
            else:
                self.detail_jump_strip_lbl.setText("")
                self.detail_jump_strip_lbl.setVisible(False)

        self._populate_assign_bar(it)
        self._populate_status_and_note_bars([it])

        can_outlook = not str(it.get("id") or "").startswith("MSG-") and not it.get("_is_group")
        has_context = bool(
            (it.get("thread_id") and self._open_thread_callback)
            or (cleaned_tdocs and self._open_tdoc_callback)
        )
        self.btn_mark_next.setEnabled(True)
        self.btn_toggle_read.setEnabled(True)
        self.btn_toggle_read.setText("✓ Read" if not it["is_read"] else "🔵 Unread")
        self.btn_toggle_flag.setEnabled(True)
        self.btn_toggle_flag.setText("🏳️" if it["is_flagged"] else "🚩")
        self.btn_reply.setEnabled(can_outlook)
        self.btn_reply.setVisible(can_outlook)
        self.btn_open_outlook.setEnabled(can_outlook)
        self.btn_open_outlook.setVisible(can_outlook)
        if hasattr(self, "btn_copy_start_template"):
            self.btn_copy_start_template.setEnabled(True)
            self.btn_copy_start_template.setVisible(True)
        if hasattr(self, "btn_llm_export"):
            self.btn_llm_export.setEnabled(True)
            self.btn_llm_export.setVisible(True)

        self.btn_open_context.setEnabled(has_context)
        self.btn_open_context.setVisible(has_context)
        if is_tdoc_bucket and cleaned_tdocs and self._open_tdoc_callback:
            latest_rev_tok = next(
                (rt for rt in reversed(rev_tokens) if rt.upper().startswith(primary_tdoc)),
                primary_tdoc,
            )
            self.btn_open_context.setText(f"📄 Open {latest_rev_tok}")
        elif it.get("thread_id") and self._open_thread_callback:
            self.btn_open_context.setText("📂 Thread")
        elif cleaned_tdocs and self._open_tdoc_callback:
            self.btn_open_context.setText(f"📄 {cleaned_tdocs[0]}")
        else:
            self.btn_open_context.setText("📂 Context")

        if hasattr(self, "btn_jump_tdoc_table"):
            can_jump = bool(primary_tdoc and getattr(self, "_jump_tdoc_callback", None))
            self.btn_jump_tdoc_table.setVisible(can_jump)
            self.btn_jump_tdoc_table.setEnabled(can_jump)

        self.btn_merge_selected.setVisible(False)
        self.btn_merge_selected.setEnabled(False)
        self._update_selected_sync_controls(1)

        if it.get("_is_group"):
            synth = {
                "thread_id": it.get("thread_id") or it["id"],
                "title": it.get("thread_title") or it.get("subject") or "Thread",
                "normalized_subject": normalize_subject(it.get("subject") or "").lower(),
                "linked_tdocs": it.get("linked_tdocs") or [],
            }
            group_msgs = [
                {
                    "id": m["id"],
                    "sender_name": m.get("sender_name", ""),
                    "sender_email": m.get("sender_email", ""),
                    "company": m.get("company", ""),
                    "date_sent": m.get("date_str", ""),
                    "subject": m.get("subject", ""),
                    "body_text": m.get("raw_body_text") or m.get("body_text", ""),
                    "clean_body": m.get("clean_body") or m.get("body_text", ""),
                    "is_read": 1 if m["is_read"] else 0,
                    "flag_status": 2 if m["is_flagged"] else 0,
                    "attachments": m.get("attachments") or [],
                    "_signals": m.get("_signals"),
                }
                for m in (it.get("_group_messages") or [])
            ]
            md = export_thread_to_markdown(
                synth,
                group_msgs,
                deduplicate_quotes=True,
                interactive_links=True,
                newest_first=True,
                include_metadata=False,
                enable_tdoc_links=bool(self._open_tdoc_callback),
                include_discussion_heading=False,
            )
            self.preview_browser.setMarkdown(md)
            blocks = find_unread_heading_blocks(self.preview_browser)
            if blocks:
                QTimer.singleShot(0, lambda b=blocks[0]: scroll_browser_to_block(self.preview_browser, b))
            return

        lines: List[str] = []
        raw_atts = it.get("attachments") or []
        real_atts, inline_imgs = partition_email_attachments(raw_atts)
        if real_atts and can_outlook:
            enc_id = urllib.parse.quote(it["id"], safe="")
            att_links = [
                f"[📎 {a}](outlook://att/{enc_id}/{urllib.parse.quote(str(a), safe='')})"
                for a in real_atts
            ]
            att_line = f"**Attachments:** {' · '.join(att_links)}"
            if inline_imgs:
                att_line += f" *(+{len(inline_imgs)} inline)*"
            lines.append(att_line + "\n")

        raw_body = str(it.get("body_text") or "")
        if it.get("html_body"):
            raw_body = inject_html_tables_into_body(raw_body, str(it["html_body"]))
        body = format_emeeting_body_with_callouts(
            raw_body,
            primary_tdoc=primary_tdoc,
            interactive_links=True,
            enable_tdoc_links=bool(self._open_tdoc_callback),
            deduplicate_quotes=False,
        )
        lines.append(body if body else "*(Empty message body)*")
        self.preview_browser.setMarkdown("\n\n".join(lines))
