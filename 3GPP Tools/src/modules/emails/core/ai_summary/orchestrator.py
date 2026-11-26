"""Pure-Python orchestrator for incremental leaf, selection, and 2-stage Per-AI email summaries."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from core.ai.cli_launch import load_standing_cli_context
from core.ai.llm_runner import (
    CancellationFlag,
    ChunkCallback,
    ModelSpec,
    resolve_effective_model_spec,
    run_llm_text,
)
from core.utils.paths import get_project_root

from .formatting import (
    extract_deterministic_facts,
    format_messages_for_llm,
    get_message_timestamp_str,
    parse_summary_output,
    sort_messages_chronologically,
)
from .models import EmailSummaryRecord, SummaryScope
from .store import EmailSummaryStore, utc_now_iso

logger = logging.getLogger(__name__)

PROMPTS_DIR = get_project_root() / "config" / "prompts"


def load_summary_prompt(filename: str, fallback: str) -> str:
    path = PROMPTS_DIR / filename
    if path.is_file():
        try:
            return path.read_text(encoding="utf-8").strip()
        except Exception as exc:
            logger.debug("Could not load prompt %s: %s", filename, exc)
    return fallback.strip()


def _max_message_timestamp(messages: Sequence[Dict[str, Any]]) -> str:
    stamps = [get_message_timestamp_str(m) for m in messages if get_message_timestamp_str(m)]
    return max(stamps) if stamps else utc_now_iso()


class EmailSummaryOrchestrator:
    """Coordinates incremental leaf summaries, ad-hoc selections, and 2-stage Per-AI synthesis."""

    def __init__(
        self,
        store: EmailSummaryStore,
        meeting_dir: Optional[Path] = None,
        *,
        wg: str = "",
        meeting_id: str = "",
    ) -> None:
        self.store = store
        self.meeting_dir = Path(meeting_dir) if meeting_dir else store.db_path.parent.parent
        self.wg = str(wg or "").strip().upper()
        self.meeting_id = str(meeting_id or "").strip()

    def summarize_leaf(
        self,
        scope: SummaryScope,
        messages: Sequence[Dict[str, Any]],
        *,
        mode: str = "auto",
        model_spec: Optional[ModelSpec] = None,
        tdoc_meta: Optional[Dict[str, Any]] = None,
        cancel_flag: Optional[CancellationFlag] = None,
        on_progress: Optional[Callable[[str], None]] = None,
        on_chunk: Optional[ChunkCallback] = None,
    ) -> Tuple[Optional[EmailSummaryRecord], str]:
        """Summarize a Thread, TDoc family, or Selection (full or incremental delta merge)."""
        norm_scope = scope.normalized()
        sorted_msgs = sort_messages_chronologically(messages)
        if not sorted_msgs:
            return None, "No messages available in this scope to summarize."

        eff_spec = resolve_effective_model_spec("leaf" if norm_scope.scope_type in ("thread", "tdoc") else "synthesis", model_spec)
        if not eff_spec.model and eff_spec.provider == "ollama":
            return None, "No Ollama model or external CLI available. Configure a summary model first."

        existing = self.store.get_summary(norm_scope.scope_type, norm_scope.scope_key)
        freshness = self.store.compute_freshness(norm_scope.scope_type, norm_scope.scope_key, sorted_msgs)

        req_mode = (mode or "auto").strip().lower()
        use_incremental = (
            existing is not None
            and norm_scope.scope_type in ("thread", "tdoc")
            and freshness.new_count > 0
            and req_mode != "full"
            and (req_mode == "incremental" or not freshness.recommend_full_rebuild)
        )

        if existing is not None and freshness.new_count == 0 and req_mode == "auto":
            if on_progress:
                on_progress("Summary is already up to date.")
            return existing, ""

        primary_tdoc = norm_scope.scope_key if norm_scope.scope_type == "tdoc" else ""
        facts = extract_deterministic_facts(sorted_msgs, primary_tdoc=primary_tdoc, tdoc_meta=tdoc_meta)

        if use_incremental and existing is not None:
            new_id_set = set(freshness.new_message_ids)
            delta_msgs = []
            for idx, m in enumerate(sorted_msgs, start=1):
                if str(m.get("id") or m.get("email_id") or "").strip() in new_id_set:
                    mc = dict(m)
                    mc["_msg_idx"] = idx
                    delta_msgs.append(mc)
            if on_progress:
                on_progress(f"Updating summary incrementally (+{len(delta_msgs)} new message(s)) via {eff_spec.short_badge()}...")
            sys_prompt = load_summary_prompt(
                "email_summary_incremental_system.txt",
                "Update the previous 3GPP email summary using the new messages. Output ### 🆕 What Changed Since Last Cutoff followed by the full updated 4-section summary.",
            )
            delta_block, _ = format_messages_for_llm(delta_msgs)
            user_prompt = (
                f"Scope: {norm_scope.scope_type.upper()} `{norm_scope.scope_key}` — {norm_scope.title}\n"
                f"Previous Cutoff: {existing.cutoff_utc or 'unknown'}\n"
                f"Total Messages Now: {len(sorted_msgs)} (New messages: {len(delta_msgs)})\n\n"
                f"=== PREVIOUS SUMMARY ===\n{existing.summary_md}\n\n"
                f"=== NEW MESSAGES ===\n{delta_block}"
            )
            next_hops = existing.incremental_hops + 1
            covered_ids = list(
                dict.fromkeys(
                    list(existing.covered_message_ids)
                    + [str(m.get("id") or m.get("email_id") or "").strip() for m in sorted_msgs]
                )
            )
        else:
            if on_progress:
                on_progress(f"Synthesizing {len(sorted_msgs)} message(s) via {eff_spec.short_badge()}...")
            if norm_scope.scope_type == "selection":
                sys_prompt = load_summary_prompt(
                    "email_summary_selection_system.txt",
                    "Synthesize across the selected 3GPP email discussions and address the focus prompt if provided.",
                )
            else:
                sys_prompt = load_summary_prompt(
                    "email_summary_thread_system.txt",
                    "Summarize the 3GPP email discussion into Core Contention, Company Positions, Revisions/Open Issues, and Consensus/Next Steps.",
                )
            msgs_block, _ = format_messages_for_llm(sorted_msgs)
            meta_lines = [
                f"Scope: {norm_scope.scope_type.upper()} `{norm_scope.scope_key}`",
                f"Title: {norm_scope.title}",
                f"Meeting: {self.meeting_id or 'N/A'} ({self.wg or 'N/A'})",
                f"Total Messages: {len(sorted_msgs)}",
                f"Participating Companies: {', '.join(facts.get('companies') or []) or 'Unknown'}",
            ]
            if facts.get("tdoc_title"):
                meta_lines.append(f"TDoc Title: {facts['tdoc_title']} (Author: {facts.get('tdoc_author', 'N/A')}, Status: {facts.get('tdoc_status', 'N/A')})")
            if norm_scope.focus_prompt:
                meta_lines.append(f"Focus Prompt: {norm_scope.focus_prompt}")
            user_prompt = "\n".join(meta_lines) + "\n\nCleaned Chronological Messages:\n\n" + msgs_block
            next_hops = 0
            covered_ids = [str(m.get("id") or m.get("email_id") or "").strip() for m in sorted_msgs]

        res = run_llm_text(
            sys_prompt,
            user_prompt,
            eff_spec,
            cancel_flag=cancel_flag,
            on_chunk=on_chunk,
        )
        if not res.ok:
            return None, res.error or "Summary generation failed."

        summary_md, delta_md, structured_meta = parse_summary_output(res.text)
        if not summary_md:
            return None, "Model returned an empty summary."

        rec = EmailSummaryRecord(
            scope_type=norm_scope.scope_type,
            scope_key=norm_scope.scope_key,
            title=norm_scope.title,
            summary_md=summary_md,
            delta_md=delta_md,
            structured_meta=structured_meta,
            deterministic_facts=facts,
            cutoff_utc=_max_message_timestamp(sorted_msgs),
            covered_message_ids=[x for x in covered_ids if x],
            message_count=len(sorted_msgs),
            incremental_hops=next_hops,
            provider=res.provider,
            model=res.model,
            reasoning_effort=res.reasoning_effort,
            focus_prompt=norm_scope.focus_prompt,
        )
        saved = self.store.upsert_summary(rec, save_history=True)
        return saved, ""

    def summarize_ai(
        self,
        ai_code: str,
        ai_label: str,
        leaf_buckets: Dict[str, Dict[str, Any]],
        *,
        mode: str = "auto",
        leaf_model_spec: Optional[ModelSpec] = None,
        synthesis_model_spec: Optional[ModelSpec] = None,
        tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
        focus_prompt: str = "",
        cancel_flag: Optional[CancellationFlag] = None,
        on_progress: Optional[Callable[[str], None]] = None,
        on_chunk: Optional[ChunkCallback] = None,
    ) -> Tuple[Optional[EmailSummaryRecord], str]:
        """
        2-Stage Map-Reduce cross-thinking synthesis for an Agenda Item (`ai_code`):
        - Stage 1 (Map): Reuse fresh leaf summaries; incrementally update only stale/missing leaves.
        - Stage 2 (Reduce): Cross-think across all leaf summaries + TDoc metadata + standing company context.
        """
        clean_ai = str(ai_code or "").strip()
        if not clean_ai:
            return None, "Agenda Item code is required."
        if not leaf_buckets:
            return None, f"No email discussions found under AI {clean_ai}."

        eff_leaf = resolve_effective_model_spec("leaf", leaf_model_spec)
        eff_synth = resolve_effective_model_spec("synthesis", synthesis_model_spec)
        req_mode = (mode or "auto").strip().lower()

        # Stage 1: Identify stale or missing leaves
        leaf_records: List[Tuple[str, Dict[str, Any], EmailSummaryRecord]] = []
        stale_keys: List[str] = []
        for bkey, bucket in leaf_buckets.items():
            msgs = bucket.get("messages") or []
            if not msgs:
                continue
            stype = str(bucket.get("scope_type") or "thread").lower()
            skey = str(bucket.get("scope_key") or bkey)
            fresh = self.store.compute_freshness(stype, skey, msgs)
            if req_mode == "full" or not fresh.has_summary or fresh.is_stale:
                stale_keys.append(bkey)

        total_refresh = len(stale_keys)
        refreshed = 0
        all_msgs_for_ai: List[Dict[str, Any]] = []

        for bkey, bucket in leaf_buckets.items():
            if cancel_flag and cancel_flag.is_cancelled():
                return None, "Cancelled by user."
            msgs = list(bucket.get("messages") or [])
            if not msgs:
                continue
            all_msgs_for_ai.extend(msgs)
            stype = str(bucket.get("scope_type") or "thread").lower()
            skey = str(bucket.get("scope_key") or bkey)
            btitle = str(bucket.get("title") or skey)
            tmeta = bucket.get("tdoc_meta") or ((tdoc_meta_map or {}).get(skey.upper()) if stype == "tdoc" else None)

            if bkey in stale_keys:
                refreshed += 1
                if on_progress:
                    on_progress(f"Stage 1/2: Summarizing {refreshed}/{total_refresh} ({skey}) via {eff_leaf.short_badge()}...")
                leaf_mode = "full" if req_mode == "full" else "auto"
                rec, err = self.summarize_leaf(
                    SummaryScope(scope_type=stype, scope_key=skey, title=btitle),
                    msgs,
                    mode=leaf_mode,
                    model_spec=eff_leaf,
                    tdoc_meta=tmeta,
                    cancel_flag=cancel_flag,
                )
                if err or rec is None:
                    if cancel_flag and cancel_flag.is_cancelled():
                        return None, "Cancelled by user."
                    logger.warning("Leaf summary skipped for %s (%s): %s", skey, stype, err)
                    continue
                leaf_records.append((bkey, bucket, rec))
            else:
                existing_leaf = self.store.get_summary(stype, skey)
                if existing_leaf is not None:
                    leaf_records.append((bkey, bucket, existing_leaf))

        if not leaf_records:
            return None, f"Could not obtain any leaf summaries for AI {clean_ai}."

        if cancel_flag and cancel_flag.is_cancelled():
            return None, "Cancelled by user."

        # Stage 2: Cross-thinking Reduce pass
        if on_progress:
            on_progress(f"Stage 2/2: Cross-thinking across {len(leaf_records)} topic(s) via {eff_synth.short_badge()}...")

        standing_ctx = load_standing_cli_context().strip()
        sys_prompt = load_summary_prompt(
            "email_summary_ai_reduce_system.txt",
            "Perform cross-TDoc strategic synthesis across the provided leaf summaries under this Agenda Item.",
        )

        leaf_sections: List[str] = []
        for idx, (_bkey, bucket, rec) in enumerate(leaf_records, 1):
            stype = rec.scope_type
            skey = rec.scope_key
            link_tag = f"[{skey}](tdoc:{skey})" if stype == "tdoc" else f"[{rec.title}](thread:{skey})"
            facts = rec.deterministic_facts or {}
            comp_str = ", ".join(facts.get("companies") or []) or "Unknown"
            status_bits = []
            if facts.get("tdoc_status"):
                status_bits.append(f"Official Status: {facts['tdoc_status']}")
            if facts.get("my_status") and facts.get("my_status") != "⚪ Neutral":
                status_bits.append(f"My Status: {facts['my_status']}")
            if facts.get("max_rev_tag"):
                status_bits.append(f"Latest Rev: {facts['max_rev_tag']}")
            if facts.get("is_objected"):
                status_bits.append("Objected/Parked")
            status_line = " | ".join(status_bits)
            section = [
                f"## Item {idx}: {link_tag} — {rec.title}",
                f"- **Scope:** `{stype}:{skey}` | **Messages:** {rec.message_count} | **Companies:** {comp_str}",
            ]
            if status_line:
                section.append(f"- **Signals & Status:** {status_line}")
            section.append("")
            section.append(rec.summary_md)
            leaf_sections.append("\n".join(section))

        prompt_parts = [
            f"Agenda Item: {ai_label or clean_ai}",
            f"Meeting: {self.meeting_id or 'N/A'} ({self.wg or 'N/A'})",
            f"Total Leaf Discussions Analyzed: {len(leaf_records)}",
        ]
        if focus_prompt.strip():
            prompt_parts.append(f"Delegate Focus Prompt: {focus_prompt.strip()}")
        if standing_ctx:
            prompt_parts.append(f"\n=== STANDING COMPANY / DELEGATE CONTEXT ===\n{standing_ctx[:3000]}")
        prompt_parts.append("\n=== LEAF SUMMARIES ===\n\n" + "\n\n---\n\n".join(leaf_sections))

        res = run_llm_text(
            sys_prompt,
            "\n".join(prompt_parts),
            eff_synth,
            cancel_flag=cancel_flag,
            on_chunk=on_chunk,
        )
        if not res.ok:
            return None, res.error or "AI cross-thinking synthesis failed."

        summary_md, delta_md, structured_meta = parse_summary_output(res.text)
        if not summary_md:
            return None, "Model returned an empty AI briefing."

        dedup_msgs: Dict[str, Dict[str, Any]] = {}
        for m in all_msgs_for_ai:
            mid = str(m.get("id") or m.get("email_id") or "").strip()
            if mid:
                dedup_msgs[mid] = m
        ai_msgs = list(dedup_msgs.values()) if dedup_msgs else all_msgs_for_ai
        ai_facts = extract_deterministic_facts(ai_msgs)
        ai_facts["leaf_count"] = len(leaf_records)
        ai_facts["leaf_keys"] = [f"{r.scope_type}:{r.scope_key}" for _, _, r in leaf_records]

        rec = EmailSummaryRecord(
            scope_type="ai",
            scope_key=clean_ai,
            title=ai_label or f"AI {clean_ai}",
            summary_md=summary_md,
            delta_md=delta_md,
            structured_meta=structured_meta,
            deterministic_facts=ai_facts,
            cutoff_utc=_max_message_timestamp(ai_msgs),
            covered_message_ids=list(dedup_msgs.keys()),
            message_count=len(ai_msgs),
            incremental_hops=0,
            provider=res.provider,
            model=res.model,
            reasoning_effort=res.reasoning_effort,
            focus_prompt=focus_prompt.strip(),
        )
        saved = self.store.upsert_summary(rec, save_history=True)
        return saved, ""
