# --- File: src/modules/meetings/core/exporters/post_meeting/mixins/agreements.py ---
"""
RAN Chairman's Notes agreement attachment (agreements.md, AI -> WI inheritance)
and Markdown rendering.

Mixin of PostMeetingExporterThread: relies on attributes set in __init__.
"""
from typing import List
from modules.meetings.ui.tdocs_models import TDocsTableModel


class AgreementsMixin:
    """RAN Chairman's Notes agreement attachment (agreements.md, AI -> WI inheritance)"""


    def _attach_chair_notes_agreements(self, tree: dict, tdoc_map: dict, wi_db_info: dict):
        """Loads cached Chairman's Notes agreements and attaches matching items to `tree[release][wi]['agreements']`."""
        from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser

        _, raw_agreements, _ = ChairNotesAgreementsParser.scan_and_cache_chair_notes(
            self.meeting_dir, mtg_info=self.mtg_info, agenda_map=self.agenda_map
        )

        if not raw_agreements:
            return

        for ag in raw_agreements:
            if ag.get("is_previous_meeting_recap"):
                continue
            ai_num = str(ag.get("ai", "")).strip() or "Unassigned"
            ai_title = str(ag.get("ai_title", "")).strip()
            tdoc_id = str(ag.get("tdoc_id", "")).strip().upper()
            raw_wi = str(ag.get("wi", "")).strip()

            t_row = tdoc_map.get(tdoc_id, {}) if tdoc_id else {}
            associated_wis = []
            if raw_wi:
                for part in TDocsTableModel._split_related_wis(raw_wi):
                    can = self._canonicalize_wi_name(part, wi_db_info)
                    if can and can not in associated_wis:
                        associated_wis.append(can)
            if not associated_wis and t_row:
                associated_wis = self._resolve_tdoc_wis(t_row, wi_db_info)
            if not associated_wis and ai_num:
                # In RAN/3GPP meetings, agreements under AI headings inherit WIs associated with that AI
                if hasattr(self, "ai_to_wis") and ai_num in self.ai_to_wis:
                    for w in self.ai_to_wis[ai_num]:
                        can = self._canonicalize_wi_name(w, wi_db_info)
                        if can and can != "General" and can not in associated_wis:
                            associated_wis.append(can)
                elif self.agenda_map and ai_num in self.agenda_map:
                    item = self.agenda_map[ai_num]
                    acronym = getattr(item, "acronym", None)
                    if acronym:
                        can = self._canonicalize_wi_name(acronym, wi_db_info)
                        if can and can not in associated_wis:
                            associated_wis.append(can)
            if not associated_wis:
                associated_wis = ["General"]

            matching_wis = []
            for wi in associated_wis:
                if wi == "General":
                    wg_name = str(self.mtg_info.get("wg_name", "")).strip().upper()
                    is_ran = wg_name.startswith("RAN") or "RAN" in wg_name
                    if self.scope_mode in ("whole", "visible"):
                        matching_wis.append(wi)
                    elif self.scope_mode == "custom" and ("general" in self.custom_wis or ai_num in self.custom_ais):
                        matching_wis.append(wi)
                    elif getattr(self, "include_general_ls", False) or is_ran:
                        # Include general non-WI agreements for RAN meetings
                        matching_wis.append(wi)
                    continue
                wi_tokens = TDocsTableModel._extract_base_tokens(wi)
                if self.scope_mode == "followed" and (wi_tokens & wi_db_info["followed_tokens"]):
                    matching_wis.append(wi)
                elif self.scope_mode == "starred" and (wi_tokens & wi_db_info["starred_tokens"]):
                    matching_wis.append(wi)
                elif self.scope_mode == "tracked" and (wi_tokens & wi_db_info["tracked_tokens"]):
                    matching_wis.append(wi)
                elif self.scope_mode == "custom":
                    if any(t in self.custom_wis for t in wi_tokens) or (wi.lower() in self.custom_wis) or (ai_num in self.custom_ais):
                        matching_wis.append(wi)
                elif self.scope_mode in ("whole", "visible"):
                    matching_wis.append(wi)

            if not matching_wis:
                continue

            release = str(ag.get("release", "")).strip() or self._resolve_tdoc_release(t_row, matching_wis, wi_db_info)
            for target_wi in matching_wis:
                if release not in tree:
                    tree[release] = {}
                if target_wi not in tree[release]:
                    tree[release][target_wi] = {
                        "info": self._get_wi_display_info(target_wi, wi_db_info, release),
                        "ais": set(),
                        "ai_titles": {},
                        "categories": {"LS_IN": [], "pCR": [], "CR": [], "LS_OUT": [], "DISC": []},
                        "agreements": [],
                    }
                tree[release][target_wi]["ais"].add(ai_num)
                if ai_num and ai_title and ai_num not in tree[release][target_wi]["ai_titles"]:
                    tree[release][target_wi]["ai_titles"][ai_num] = ai_title
                tree[release][target_wi]["agreements"].append(ag)

    def _render_agreements_markdown(self, agreements_list: List[dict]) -> str:
        """Renders Chairman's Notes agreements grouped by Agenda Item in clean Markdown callouts."""
        if not agreements_list:
            return ""
        lines = []
        for ag in agreements_list:
            kind = str(ag.get("kind", "Agreement")).strip()
            ai = str(ag.get("ai", "")).strip()
            ai_title = str(ag.get("ai_title", "")).strip()
            tdoc_id = str(ag.get("tdoc_id", "")).strip()
            text = str(ag.get("text", "")).strip()

            icon = "🟢"
            if "ffs" in kind.lower() or "open issue" in kind.lower():
                icon = "🔵"
            elif "working assumption" in kind.lower():
                icon = "🟡"
            elif "agreed tdoc" in kind.lower():
                icon = "📄"

            header_parts = [f"{icon} **[{kind}]**"]
            if ai:
                header_parts.append(f"**AI {ai}**" + (f" (*{ai_title}*)" if ai_title else ""))
            if tdoc_id:
                header_parts.append(f"(Related TDoc: `{tdoc_id}`)")
            lines.append("- " + " — ".join(header_parts))
            for subline in text.split("\n"):
                lines.append(f"  > {subline}")
            lines.append("")
        return "\n".join(lines)
