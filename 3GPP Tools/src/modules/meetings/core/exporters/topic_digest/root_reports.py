# --- File: src/modules/meetings/core/exporters/topic_digest/root_reports.py ---
"""Root-level topic dossier reports: E2E prompt/checklist, master index.md, consolidated briefing."""

import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.export_manifest import ExportManifest
from modules.meetings.core.topic_decision import TopicDecision
from modules.meetings.core.topic_manager import MetaTopicDefinition

from .constants import STAGE2_REPORT_FILENAME, STAGE2_SYNTHESIS_FILENAME
from .status_and_outcomes import (
    _resolve_effective_synthesis_state,
    _resolve_rec_analysis_status,
    format_analysis_badge,
    format_lifecycle_badge,
    format_stage2_summary_badge,
)


def write_e2e_synthesis_prompt(
    target_dir: Path,
    topic: MetaTopicDefinition,
    meeting_records: List[Dict[str, Any]],
    synthesis_status: Optional[str] = None,
    synthesis_state: Optional[Dict[str, Any]] = None,
) -> Path:
    """
    Writes `company_positions_e2e_prompt.md` at the topic root containing copy-pasteable
    Stage-1 (per-meeting low-cost map + token-minimal delta merge) and Stage-2 (cross-WG E2E reduce)
    prompts + live status checklist + dated Stage-2 baseline chaining metadata.
    """
    out_p = target_dir / "company_positions_e2e_prompt.md"
    s2_state = _resolve_effective_synthesis_state(target_dir, synthesis_status, synthesis_state)
    eff_s2_status = str(s2_state.get("status") or "pending")
    s2_exists = bool(s2_state.get("exists"))
    s2_report_rel = str(s2_state.get("file") or STAGE2_SYNTHESIS_FILENAME).replace("\\", "/")
    s2_folder = str(s2_state.get("folder_name") or "").strip()
    s2_pkg_id = str(s2_state.get("package_id") or "").strip()
    s2_as_of = str(s2_state.get("as_of") or "").strip()
    s2_rev = str(s2_state.get("revision") or "").strip()
    changed_meetings: List[str] = list(s2_state.get("changed_meetings") or [])
    s2_summary = format_stage2_summary_badge(eff_s2_status, s2_state, as_markdown_link=True)

    pending_rows = []
    stale_rows = []
    done_rows = []
    skip_rows = []
    rec_by_key: Dict[str, Dict[str, Any]] = {}
    stale_keys_fallback: List[str] = []
    for rec in meeting_records:
        st, exists, a_rel = _resolve_rec_analysis_status(rec, target_dir)
        m_fid = rec["meeting_folder_id"]
        wg = rec.get("wg", "")
        if wg and m_fid:
            rec_by_key[f"{wg}/{m_fid}"] = rec
        if st == "no_activity":
            skip_rows.append(m_fid)
        elif st == "stale":
            stale_rows.append(m_fid)
            if wg and m_fid:
                stale_keys_fallback.append(f"{wg}/{m_fid}")
        elif st == "analyzed":
            done_rows.append(m_fid)
        else:
            pending_rows.append(m_fid)

    if not changed_meetings and eff_s2_status == "stale" and stale_keys_fallback:
        changed_meetings = stale_keys_fallback

    lines = [
        f"# 🌐 E2E Cross-WG Company Position Synthesis Guide: {topic.name}",
        "",
        "This workspace is structured for **token-efficient, resumable 2-stage AI analysis** across Working Groups:",
        "1. **Stage 1 (Map — Per-Meeting, Low/Medium Cost Model):** Analyze active meetings (`⬜ Pending`) once, and use token-minimal delta merges (`🔄 Stale`) whenever an already-analyzed meeting gains post-meeting outcomes or new revisions.",
        "2. **Stage 2 (Reduce — Dated Cross-WG Synthesis Package → `stage2_<YYYY-MM-DD>_r1/`):** Synthesize E2E company positions across WGs using only active `*/_company_positions_analysis.md` outputs + `agreements_timeline.md` + `ls_dialogue.md` (and chain `--baseline` from the prior dated Stage-2 package on updates).",
        "",
        f"**Current Status:** ✅ {len(done_rows)} Analyzed | 🔄 {len(stale_rows)} Stale (Delta Ready) | ⬜ {len(pending_rows)} Pending | ⚪ {len(skip_rows)} Skipped (0 Inputs)",
        f"**Latest Stage-2 Synthesis:** {s2_summary}",
    ]

    if s2_exists and s2_pkg_id:
        base_id = str(s2_state.get("baseline_package_id") or "none supplied")
        claims_note = f" | **Claims:** {s2_state['claims']}" if s2_state.get("claims") is not None else ""
        words_note = f" | **Words:** {s2_state['synthesis_words']}" if s2_state.get("synthesis_words") is not None else ""
        lines.append(f"**Package Metadata:** `{s2_pkg_id}` | **Baseline:** `{base_id}`{claims_note}{words_note}")

    lines.extend([
        "",
        "---",
        "",
        "## 🔹 Stage 1 Prompt (Run Once Per Active Meeting — Supports Full & Delta Merge)",
        "```text",
        "Check each active meeting's status in the checklist below (skip rows marked '⚪ Skip (0 Inputs)'):",
        "",
        "- If Stage-1 Output is '⬜ Pending':",
        "  Read `<WG>/<MeetingFolder>/_meeting_brief.md` and `<WG>/<MeetingFolder>/_company_positions_input*.md`.",
        "  Write `<WG>/<MeetingFolder>/_company_positions_analysis.md` containing:",
        "  1. Key Technical Decisions & Outcomes in this meeting (and which proposals won/lost).",
        "  2. Company Stance Table: | Company | Architectural Position | Concrete Proposals (TDoc IDs) | Confirmed Allies | Outcome |",
        "     - Strict rule: ignore companies flagged as carry-over unconfirmed co-sponsors.",
        "  3. Open Issues / Contested Topics carrying forward.",
        "",
        "- If Stage-1 Output is '🔄 Stale (Delta Ready)':",
        "  TOKEN-SAVING RULE: Do NOT re-read `_company_positions_input*.md`.",
        "  Read ONLY `<WG>/<MeetingFolder>/_company_positions_analysis.md` and the `## 🔄 Delta Since Last Analysis` section in `<WG>/<MeetingFolder>/_meeting_brief.md`.",
        "  Merge the new decisions, revision transitions, and newly added/un-stubbed TDocs directly into `<WG>/<MeetingFolder>/_company_positions_analysis.md` in place (update stale stances/statuses and Lifecycle State; keep the single clean analysis file).",
        "",
        "Be concise, factual, and cite TDoc IDs for every claim.",
        "```",
        "",
        "### Meeting Analysis Checklist (Chronological)",
        "",
        "| Date | WG | Meeting | State | Active Docs | Decisions | Brief | Input Packet(s) | Stage-1 Output |",
        "| :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :--- |",
    ])

    for rec in meeting_records:
        wg = rec["wg"]
        m_fid = rec["meeting_folder_id"]
        d_str = rec.get("start_date", "") or "N/A"
        lc = rec.get("lifecycle") or ("concluded" if rec.get("decision_count", 0) else "upcoming")
        lc_badge = format_lifecycle_badge(lc)
        brief_rel = f"{wg}/{m_fid}/_meeting_brief.md"
        digests = rec.get("digest_rel_paths", [])
        digest_links = " · ".join(f"[{Path(p).name}]({p})" for p in digests) if digests else "—"
        st, exists, analysis_rel = _resolve_rec_analysis_status(rec, target_dir)
        badge = format_analysis_badge(st)
        if st == "no_activity":
            stage1_cell = f"{badge}" + (f" *([old file]({analysis_rel}))*" if exists else "")
        elif exists:
            stage1_cell = f"[{badge}]({analysis_rel})"
        else:
            stage1_cell = f"{badge} (`{analysis_rel}`)"
        lines.append(
            f"| {d_str} | **{wg}** | {m_fid} | {lc_badge} | {rec['tdoc_count']} | {rec['decision_count']} | [Brief]({brief_rel}) | {digest_links} | {stage1_cell} |"
        )

    if changed_meetings:
        lines.extend([
            "",
            f"### 🔄 Meetings Updated Since Baseline (`{s2_pkg_id or s2_report_rel}`)",
            "",
            "| Meeting Key | Stage-1 Status | Active Docs | Decisions | Links |",
            "| :--- | :--- | :---: | :---: | :--- |",
        ])
        for m_key in changed_meetings:
            rec = rec_by_key.get(m_key, {})
            st, exists, analysis_rel = _resolve_rec_analysis_status(rec, target_dir) if rec else ("stale", False, f"{m_key}/_company_positions_analysis.md")
            brief_rel = rec.get("brief_rel_path") or f"{m_key}/_meeting_brief.md"
            a_link = f"[{format_analysis_badge(st)}]({analysis_rel})" if exists else format_analysis_badge(st)
            lines.append(
                f"| `{m_key}` | {a_link} | {rec.get('tdoc_count', '—')} | {rec.get('decision_count', '—')} | [Brief]({brief_rel}) |"
            )

    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    next_rev = "r2" if (s2_folder and s2_as_of == today_str and s2_rev == "r1") else "r1"
    target_pkg_dir = f"stage2_{today_str}_{next_rev}/"

    if s2_exists:
        baseline_lines = [
            f"- **Prior Baseline Package:** `{s2_report_rel}` (Package ID: `{s2_pkg_id or s2_folder or 'legacy-root'}`).",
            f"- **Baseline Chaining Rule:** Do NOT overwrite `{s2_folder + '/' if s2_folder else s2_report_rel}`. Read the prior synthesis `{s2_report_rel}` as `--baseline`, preserve unchanged sections/claims, and merge updates into a NEW dated folder `{target_pkg_dir}`.",
        ]
        if changed_meetings:
            baseline_lines.append(
                f"- **Updated Meetings Since Baseline ({len(changed_meetings)}):** {', '.join(changed_meetings)} — prioritize reading their updated `_company_positions_analysis.md` + `_meeting_brief.md` delta sections."
            )
        elif eff_s2_status == "stale":
            baseline_lines.append(
                "- **Current Status:** Stage-1 inputs changed since this baseline — merge updated meetings cleanly into the new dated folder instead of regenerating unchanged sections from scratch."
            )
        else:
            baseline_lines.append(
                "- **Current Status:** All active Stage-1 analyses match this baseline. Run a new dated subfolder only when a subsequent meeting round or fresh revision arrives."
            )
    else:
        baseline_lines = [
            "- Only active `*/*/_company_positions_analysis.md` files (skip any row marked '⚪ Skip (0 Inputs)' above; do NOT re-read raw TDocs unless verifying a specific quote).",
        ]

    lines.extend([
        "",
        "---",
        "",
        f"## 🔸 Stage 2 Prompt (Cross-WG End-to-End Reduce → `{target_pkg_dir}{STAGE2_REPORT_FILENAME}`)",
        "```text",
        f"You are analyzing the cross-WG evolution of '{topic.name}' across {', '.join(sorted(topic.target_wgs))}.",
        "Read:",
        "- `index.md`",
        "- `agreements_timeline.md`",
        "- `ls_dialogue.md`",
        *baseline_lines,
        "",
        f"Write the new dated synthesis package into `{target_pkg_dir}` (`{STAGE2_REPORT_FILENAME}`, `COVERAGE.md`, `working/package_state.json`) covering:",
        "1. Per-Company E2E Strategy Matrix (how each major vendor/operator aligns across WGs — e.g. consistency vs. contradictions between SA2 and RAN2 stances).",
        "2. Key Architectural Fault Lines & Alliances over time (chronological trajectory across meetings).",
        "3. Outcomes & Impact on Baseline Specifications (which company's vision shaped the final agreements/pCRs).",
        "4. Strategic Outlook & Open Conflicts entering the next meeting cycle.",
        "```",
        "",
    ])

    with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_p


def write_master_index_markdown(
    target_dir: Path,
    topic: MetaTopicDefinition,
    manifest: ExportManifest,
    total_decisions: int,
    total_active_docs: int,
    stub_count: int,
    round_count: int,
    session_id: int,
    scope_mode: str,
    meeting_records: List[Dict[str, Any]],
    synthesis_status: Optional[str] = None,
    synthesis_state: Optional[Dict[str, Any]] = None,
) -> Path:
    """
    Generates root `index.md` using a single shared template for both full exports
    and fast offline checkup/sync passes.
    """
    now_str = datetime.datetime.now().strftime("%Y-%m-%d")
    s2_state = _resolve_effective_synthesis_state(target_dir, synthesis_status, synthesis_state)
    eff_s2_status = str(s2_state.get("status") or "pending")
    s2_exists = bool(s2_state.get("exists"))
    s2_report_rel = str(s2_state.get("file") or STAGE2_SYNTHESIS_FILENAME).replace("\\", "/")
    s2_packages: List[Dict[str, Any]] = list(s2_state.get("packages") or [])
    s2_kpi = format_stage2_summary_badge(eff_s2_status, s2_state, as_markdown_link=True)

    concluded_cnt = sum(1 for r in meeting_records if r.get("lifecycle") == "concluded")
    upcoming_cnt = sum(1 for r in meeting_records if r.get("lifecycle") == "upcoming")
    no_act_cnt = sum(1 for r in meeting_records if r.get("lifecycle") == "no_activity")

    active_mtgs = 0
    analyzed_mtgs = 0
    stale_mtgs = 0
    pending_mtgs = 0
    for rec in meeting_records:
        st, _, _ = _resolve_rec_analysis_status(rec, target_dir)
        if st == "no_activity":
            continue
        active_mtgs += 1
        if st == "analyzed":
            analyzed_mtgs += 1
        elif st == "stale":
            stale_mtgs += 1
        else:
            pending_mtgs += 1

    stage1_summary = f"{analyzed_mtgs}/{active_mtgs} Active Meetings Analyzed"
    extra_s1 = []
    if stale_mtgs:
        extra_s1.append(f"🔄 {stale_mtgs} Stale (Delta Ready)")
    if pending_mtgs:
        extra_s1.append(f"⬜ {pending_mtgs} Pending")
    if no_act_cnt:
        extra_s1.append(f"⚪ {no_act_cnt} Empty Skipped")
    if extra_s1:
        stage1_summary += " (" + ", ".join(extra_s1) + ")"

    if s2_exists:
        s2_nav = f"- 🏆 **[Cross-WG E2E Company Position Synthesis]({s2_report_rel})** — {s2_kpi}."
    elif eff_s2_status == "in_progress":
        s2_nav = f"- 🏆 **Cross-WG E2E Company Position Synthesis** — {s2_kpi}."
    else:
        s2_nav = f"- 🏆 **Cross-WG E2E Company Position Synthesis** — ⬜ Pending (generate `stage2_<YYYY-MM-DD>_r1/{STAGE2_REPORT_FILENAME}` using Stage 2 prompt)."

    contrib_kpi = f"{total_active_docs} documents"
    if stub_count > 0:
        contrib_kpi += f" *(including {stub_count} reserved/missing placeholder stub(s))*"

    delta_link_line = (
        f"- 🆕 **[Latest LLM Delta Feed](deltas/delta_session_{session_id}.md)** — Incremental additions in this session."
        if (target_dir / "deltas" / f"delta_session_{session_id}.md").is_file()
        else ""
    )

    lines = [
        f"# 🌐 Cross-WG Topic Dossier: {topic.name}",
        f"**Last Updated:** {now_str} (Session #{session_id}) | **Scope Mode:** `{scope_mode}`",
        "",
        "> [!NOTE]",
        f"> {topic.description or 'Consolidated multi-WG tracking organized by Working Group and Meeting.'}",
        "",
        "## 📊 Executive KPI Strip",
        "",
        f"- **Participating Working Groups:** {', '.join(sorted(topic.target_wgs))}",
        (
            f"- **Analyzed Meetings:** {len(meeting_records)} meetings "
            f"({concluded_cnt} Concluded, {upcoming_cnt} Upcoming, {no_act_cnt} No Activity) across {round_count} synchronized round(s)"
            if meeting_records
            else "- **Analyzed Meetings:** None"
        ),
        f"- **Unified Technical Decisions:** {total_decisions} (Agreements, WAs, pCRs, CRs, WFs)",
        f"- **Active Scoped Contributions:** {contrib_kpi}",
        f"- **Stage-1 Per-Meeting Analyses:** {stage1_summary}",
        f"- **Stage-2 Cross-WG Synthesis:** {s2_kpi}",
        "",
        "## 🧭 Navigation & LLM Workflows",
        "",
        "- ⚖️ **[Unified Technical Decisions Baseline](agreements_timeline.md)** — Chronological matrix of all RAN agreements and SA/CT pCRs/CRs.",
        "- 📬 **[Cross-WG Liaison Statement Dialogue](ls_dialogue.md)** — Inter-group question and reply threads.",
        "- 🧠 **[2-Stage E2E Company Position Prompt & Checklist](company_positions_e2e_prompt.md)** — Resumable, per-meeting LLM instructions designed for standard/non-frontier models.",
        s2_nav,
    ]
    if delta_link_line:
        lines.append(delta_link_line)

    lines.extend([
        "",
        "## 🗓️ Cross-WG Chronological Meeting Timeline",
        "",
        "| Date | Synchronized Round | WG | Meeting | State | Decisions | Active Inputs | LS (In/Out) | Meeting Brief & LLM Packet |",
        "| :--- | :--- | :--- | :--- | :--- | :---: | :---: | :---: | :--- |",
    ])

    for rec in meeting_records:
        d_str = rec.get("start_date") or "N/A"
        r_lbl = rec.get("round_label") or rec.get("cohort_key", "")
        wg = rec["wg"]
        m_fid = rec["meeting_folder_id"]
        lc = rec.get("lifecycle") or ("concluded" if rec.get("decision_count", 0) else "upcoming")
        lc_badge = format_lifecycle_badge(lc)
        brief_rel = rec["brief_rel_path"]
        digests = rec.get("digest_rel_paths", [])
        digest_str = " · ".join(f"[{Path(p).name}]({p})" for p in digests) if digests else "—"
        st, exists, analysis_rel = _resolve_rec_analysis_status(rec, target_dir)
        a_badge = format_analysis_badge(st)
        if st == "no_activity":
            analysis_cell = "⚪ 0 Inputs"
        elif exists:
            analysis_cell = f"[{a_badge}]({analysis_rel})"
        else:
            analysis_cell = a_badge
        ls_cell = f"{rec.get('ls_in_count', 0)} / {rec.get('ls_out_count', 0)}"
        lines.append(
            f"| `{d_str}` | {r_lbl} | **{wg}** | [{m_fid}]({brief_rel}) | {lc_badge} | {rec['decision_count']} | {rec['tdoc_count']} | {ls_cell} | [Brief]({brief_rel}) · {digest_str} · {analysis_cell} |"
        )

    lines.extend([
        "",
        "## 🏛️ Per-Working-Group Evolution",
        "",
        "| Working Group | Chronological Meetings (Click to Open Brief) | Decisions | Active Docs | Stage-1 Analysis |",
        "| :--- | :--- | :---: | :---: | :--- |",
    ])

    by_wg: Dict[str, List[Dict[str, Any]]] = {}
    for rec in meeting_records:
        by_wg.setdefault(rec["wg"], []).append(rec)

    for wg in sorted(by_wg.keys()):
        recs = by_wg[wg]
        links = " → ".join(f"[{r['meeting_folder_id']}]({r['brief_rel_path']})" for r in recs)
        tot_d = sum(r["decision_count"] for r in recs)
        tot_c = sum(r["tdoc_count"] for r in recs)
        wg_statuses = [_resolve_rec_analysis_status(r, target_dir)[0] for r in recs]
        wg_active = [s for s in wg_statuses if s != "no_activity"]
        wg_done = sum(1 for s in wg_active if s == "analyzed")
        wg_stale = sum(1 for s in wg_active if s == "stale")
        if not wg_active:
            wg_s1 = f"⚪ 0 Active ({len(recs)} Empty)"
        else:
            wg_s1 = f"✅ {wg_done}/{len(wg_active)}"
            if wg_stale:
                wg_s1 += f" · 🔄 {wg_stale} Stale"
        lines.append(f"| **{wg}** | {links} | {tot_d} | {tot_c} | {wg_s1} |")

    if s2_packages:
        lines.extend([
            "",
            "---",
            "### 📦 Stage-2 Cross-WG Synthesis Packages",
            "",
            "| # | Package Folder | Cut-Off Date | Revision | Package ID | Baseline ID | Validation | Artifacts |",
            "| :---: | :--- | :--- | :---: | :--- | :--- | :--- | :--- |",
        ])
        for idx, pkg in enumerate(s2_packages, start=1):
            p_folder = pkg.get("folder_name") or "(root)"
            p_as_of = pkg.get("as_of") or "N/A"
            p_rev = pkg.get("revision") or "—"
            p_id = pkg.get("package_id") or "—"
            p_base = pkg.get("baseline_package_id") or "none supplied"
            v_pass = pkg.get("validation_passed")
            if pkg.get("status") == "in_progress":
                val_str = "🏗️ In Progress"
            elif v_pass is True:
                extra_v = []
                if pkg.get("claims") is not None:
                    extra_v.append(f"{pkg['claims']} claims")
                if pkg.get("synthesis_words") is not None:
                    extra_v.append(f"{pkg['synthesis_words']}w")
                val_str = "✅ Passed" + (f" ({', '.join(extra_v)})" if extra_v else "")
            elif v_pass is False:
                val_str = "⚠️ Validation Issues"
            else:
                val_str = "—"

            arts = []
            if pkg.get("status") == "completed" and pkg.get("report_rel_path"):
                arts.append(f"[Report]({pkg['report_rel_path']})")
            if pkg.get("coverage_rel_path"):
                arts.append(f"[COVERAGE.md]({pkg['coverage_rel_path']})")
            if pkg.get("handoff_zip_rel_path"):
                arts.append(f"[Handoff ZIP]({pkg['handoff_zip_rel_path']})")
            art_str = " · ".join(arts) if arts else "—"
            lines.append(
                f"| {idx} | `{p_folder}` | `{p_as_of}` | `{p_rev}` | `{p_id}` | `{p_base}` | {val_str} | {art_str} |"
            )

    lines.extend([
        "",
        "---",
        "### 🔄 Incremental Export History",
        "",
        "| Session | Timestamp | Scope Mode | New Documents | Total Registered |",
        "| :--- | :--- | :--- | :--- | :--- |",
    ])
    for sess in manifest.sessions:
        s_id = sess.get("session_id")
        s_time = sess.get("timestamp")
        s_mode = sess.get("scope_mode")
        added_cnt = len(sess.get("added_tdocs", []))
        tot_cnt = sess.get("total_tdocs", 0)
        lines.append(f"| #{s_id} | {s_time} | `{s_mode}` | +{added_cnt} | {tot_cnt} |")

    lines.append("")
    out_p = target_dir / "index.md"
    with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_p


def write_consolidated_briefing_markdown(
    target_dir: Path,
    topic: MetaTopicDefinition,
    all_decisions: List[TopicDecision],
    meeting_records: List[Dict[str, Any]],
    round_count: int,
) -> Path:
    """Generates single-file `consolidated_briefing.md` linking decisions, briefs, digests, and Stage-1 analyses."""
    lines = [
        f"# Consolidated Executive Briefing: {topic.name}",
        "",
        "## 1. Executive Summary & Scope",
        f"This briefing consolidates decisions and contribution briefs across {', '.join(sorted(topic.target_wgs))} "
        f"spanning {len(meeting_records)} meetings ({round_count} synchronized round(s)).",
        "",
        "## 2. Chronological Technical Decisions",
        "",
    ]
    for dec in sorted(all_decisions, key=lambda d: (d.date or d.cohort_key, d.wg, d.tdoc_id or "")):
        m_lbl = dec.meeting_folder_id or f"{dec.wg}_{dec.meeting_number}"
        link_suffix = f" ([{dec.tdoc_id or 'Agreement'}]({dec.dossier_rel_path}))" if dec.dossier_rel_path else ""
        lines.append(f"- **[{dec.date or dec.cohort_key} | {m_lbl}]** ({dec.decision_type}){link_suffix}: {dec.title_or_summary}")

    if meeting_records:
        lines.extend(["", "## 3. Per-Meeting Briefs, Digests & Stage-1 Analyses", ""])
        for rec in meeting_records:
            digests = rec.get("digest_rel_paths", [])
            digest_links = ", ".join(f"[{Path(p).name}]({p})" for p in digests) if digests else "—"
            st, exists, analysis_rel = _resolve_rec_analysis_status(rec, target_dir)
            a_suffix = f" · Analysis: [{format_analysis_badge(st)}]({analysis_rel})" if (exists and st != "no_activity") else ""
            lines.append(
                f"- **{rec.get('start_date') or 'N/A'} · {rec['meeting_folder_id']}**: "
                f"[{Path(rec['brief_rel_path']).name}]({rec['brief_rel_path']}) · Digests: {digest_links}{a_suffix}"
            )

    lines.append("")
    out_p = target_dir / "consolidated_briefing.md"
    with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_p
