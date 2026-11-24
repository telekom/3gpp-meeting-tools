"""Per-meeting agreement enrichment and Agreed-TDoc synthesis processor."""

from __future__ import annotations

import os
from typing import List, Set

from core.utils.paths import to_extended_path
from modules.meetings.core.tdocs.chair_notes_agreements_parser import (
    _resolve_inherited_field,
)


class AgreementProcessorMixin:
    """Mixin aggregating per-meeting agreements from Chair Notes and Agreed Excel TDocs."""

    def _process_single_meeting(self, mtg: dict) -> List[dict]:
        if self.is_cancelled:
            return []

        mtg_num = str(mtg.get("meeting_number", "Unknown")).strip()
        wg_name = str(mtg.get("wg_name", "")).strip()
        tag = f"{wg_name} #{mtg_num}"
        folder_name = mtg.get("folder_name") or mtg_num

        local_mtg_dir = self.cache_dir / folder_name
        agenda_dir = local_mtg_dir / "Agenda"
        os.makedirs(to_extended_path(agenda_dir), exist_ok=True)

        # 1. Load or download TDocs Excel list
        tdocs_by_id, ai_wi_from_excel, ai_title_from_excel = self._load_or_fetch_tdocs_excel(
            mtg, local_mtg_dir, agenda_dir, tag
        )
        if self.is_cancelled:
            return []

        # 2. Also load agenda.csv if present in Agenda/
        ai_wi_from_csv, ai_title_from_csv = self._load_agenda_csv_map(agenda_dir)
        combined_ai_wis = {**ai_wi_from_excel, **ai_wi_from_csv}
        combined_ai_titles = {**ai_title_from_excel, **ai_title_from_csv}

        # 3. Load or download/parse Chairman's Notes agreements
        chair_agenda_data, chair_agreements = self._load_or_fetch_chair_notes(
            mtg, local_mtg_dir, agenda_dir, tag
        )
        if self.is_cancelled:
            return []

        # Merge any AI -> WI / Title discovered from chair_agenda_data
        for tinfo in (chair_agenda_data or {}).values():
            ai_val = str(tinfo.get("Agenda Item", "")).strip()
            wi_val = str(tinfo.get("Related WIs", "")).strip()
            sub_val = str(tinfo.get("Sub-Item", "")).strip()
            if ai_val and ai_val != "N/A":
                if wi_val and ai_val not in combined_ai_wis:
                    combined_ai_wis[ai_val] = wi_val
                if sub_val and ai_val not in combined_ai_titles:
                    combined_ai_titles[ai_val] = sub_val

        # 4. Enrich Chair Notes agreements with Excel TDoc metadata & inherited AI metadata
        candidates: List[dict] = []
        seen_tdoc_agreements: Set[str] = set()

        for ag in (chair_agreements or []):
            is_recap = bool(ag.get("is_previous_meeting_recap", False))
            if is_recap and not self.include_prev_recaps:
                continue

            kind = str(ag.get("kind", "Agreement")).strip()
            if not self._kind_enabled(kind):
                continue

            ai_num = str(ag.get("ai", "")).strip()
            tdoc_id = str(ag.get("tdoc_id", "")).strip().upper()
            wi = str(ag.get("wi", "")).strip()
            ai_title = str(ag.get("ai_title", "")).strip()
            release = str(ag.get("release", "")).strip()
            text = str(ag.get("text", "")).strip()
            source_file = str(ag.get("source_file", "")).strip()

            tdoc_row = tdocs_by_id.get(tdoc_id, {}) if tdoc_id else {}
            chair_tdoc_row = chair_agenda_data.get(tdoc_id, {}) if tdoc_id else {}

            tdoc_title = str(tdoc_row.get("Title") or chair_tdoc_row.get("Title") or "").strip()
            tdoc_source = str(tdoc_row.get("Source") or chair_tdoc_row.get("Source") or "").strip()
            tdoc_type = str(tdoc_row.get("Type") or chair_tdoc_row.get("Type") or "").strip()
            tdoc_status = str(tdoc_row.get("TDoc Status") or chair_tdoc_row.get("Result") or "").strip()

            if not ai_num and tdoc_row.get("Agenda item"):
                ai_num = str(tdoc_row.get("Agenda item", "")).strip()
            if not wi:
                wi = (
                    str(tdoc_row.get("Related WIs", "")).strip()
                    or _resolve_inherited_field(ai_num, combined_ai_wis)
                )
            if not ai_title:
                ai_title = (
                    _resolve_inherited_field(ai_num, combined_ai_titles)
                    or str(tdoc_row.get("Agenda item description", "")).strip()
                )
            if not release:
                release = str(tdoc_row.get("Release", "")).strip()

            if tdoc_id:
                seen_tdoc_agreements.add(tdoc_id)

            record = {
                "WG": wg_name,
                "Meeting": tag,
                "meeting_number": mtg_num,
                "start_date": str(mtg.get("start_date", "")).strip(),
                "end_date": str(mtg.get("end_date", "")).strip(),
                "folder_name": folder_name,
                "url_key": mtg.get("url_key", ""),
                "docs_folder_url": mtg.get("docs_folder_url", ""),
                "is_electronic": mtg.get("is_electronic", 0),
                "wg_name": wg_name,
                "AI": ai_num,
                "AI Title": ai_title,
                "Work Item": wi,
                "Release": release,
                "Kind": f"{kind} (Prev Mtg Recap)" if is_recap else kind,
                "Agreement Text": text,
                "TDoc": tdoc_id,
                "TDoc Title": tdoc_title,
                "Source": tdoc_source,
                "TDoc Type": tdoc_type,
                "TDoc Status": tdoc_status,
                "Chair Notes File": source_file,
                "is_previous_meeting_recap": is_recap,
            }
            if self._matches_filter(record):
                candidates.append(record)

        # 5. Also include Agreed pCRs / CRs / TPs from official Excel TDoc list
        if self.include_agreed_tdocs:
            for tid, row in tdocs_by_id.items():
                if tid in seen_tdoc_agreements:
                    continue
                status_raw = str(row.get("TDoc Status", "") or row.get("Result", "")).strip()
                type_raw = str(row.get("Type", "")).strip()
                for_raw = str(row.get("For", "")).strip()
                low_status = status_raw.lower()
                low_type = type_raw.lower()
                low_for = for_raw.lower()

                is_agreed_status = any(
                    s in low_status for s in ("agreed", "approved", "endorsed", "block approve")
                ) and not any(
                    neg in low_status for neg in ("not agreed", "not approved", "not endorsed")
                )
                if not is_agreed_status:
                    continue

                is_substantive_type = (
                    any(t in low_type for t in ("pcr", "cr", "tp", "draft", "tr", "ts", "ls out", "lsout", "sid", "wid"))
                    or any(f in low_for for f in ("agreement", "approval", "decision"))
                    or wg_name.upper().startswith("SA")
                )
                if not is_substantive_type:
                    continue

                ai_num = str(row.get("Agenda item", "") or row.get("Agenda Item", "")).strip()
                wi = str(row.get("Related WIs", "")).strip() or _resolve_inherited_field(ai_num, combined_ai_wis)
                ai_title = (
                    str(row.get("Agenda item description", "")).strip()
                    or _resolve_inherited_field(ai_num, combined_ai_titles)
                )
                tdoc_title = str(row.get("Title", "")).strip()
                tdoc_source = str(row.get("Source", "")).strip()
                release = str(row.get("Release", "")).strip()
                chair_comm = ""
                if tid in chair_agenda_data:
                    chair_comm = str(chair_agenda_data[tid].get("Comments", "")).strip()

                text_summary = f"[{status_raw}] {tid}: {tdoc_title}"
                if chair_comm:
                    text_summary += f"\n\nChair Notes Remarks:\n{chair_comm}"

                record = {
                    "WG": wg_name,
                    "Meeting": tag,
                    "meeting_number": mtg_num,
                    "start_date": str(mtg.get("start_date", "")).strip(),
                    "end_date": str(mtg.get("end_date", "")).strip(),
                    "folder_name": folder_name,
                    "url_key": mtg.get("url_key", ""),
                    "docs_folder_url": mtg.get("docs_folder_url", ""),
                    "is_electronic": mtg.get("is_electronic", 0),
                    "wg_name": wg_name,
                    "AI": ai_num,
                    "AI Title": ai_title,
                    "Work Item": wi,
                    "Release": release,
                    "Kind": "Agreed TDoc / TP / pCR",
                    "Agreement Text": text_summary,
                    "TDoc": tid,
                    "TDoc Title": tdoc_title,
                    "Source": tdoc_source,
                    "TDoc Type": type_raw,
                    "TDoc Status": status_raw,
                    "Chair Notes File": "TDoc_List_Meeting.xlsx",
                    "is_previous_meeting_recap": False,
                }
                if self._matches_filter(record):
                    candidates.append(record)

        self._log(f"   ↳ [{tag}] {len(candidates)} matching agreement(s)/item(s).")
        return candidates
