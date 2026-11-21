"""Background thread collecting matching topic decisions, contributions, and liaisons."""

import logging
from typing import Any, Dict, List

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.cohort_resolver import MeetingCohort
from modules.meetings.core.cross_wg_ls_tracker import CrossWgLiaisonTracker
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdocs.chair_notes_agreements_parser import clean_chair_notes_text
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.core.topic_decision import (
    TopicDecision,
    classify_tdoc_decision_type,
    normalize_chair_notes_decision_type,
)
from modules.meetings.core.topic_manager import MetaTopicDefinition

logger = logging.getLogger("modules.meetings.ui.dialogs.topic_inputs_dialog")


class InputsCollectorThread(QThread):
    """Background worker resolving matching decisions, contributions, and liaison statements."""
    progress_status = pyqtSignal(str, int, int)
    collection_finished = pyqtSignal(list, list, object)  # decisions, contributions, ls_tracker
    collection_failed = pyqtSignal(str)

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        topic: MetaTopicDefinition,
        rounds: List[MeetingCohort],
        scope_mode: str = "full_dossier",
        parent=None,
    ):
        super().__init__(parent)
        self.meetings_db = meetings_db
        self.topic = topic
        self.rounds = list(rounds)
        self.scope_mode = scope_mode
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True
        self.requestInterruption()

    def run(self):
        try:
            from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread
            helper_thread = TopicDossierExporterThread(
                meetings_db=self.meetings_db,
                topic=self.topic,
                selected_cohorts=self.rounds,
                export_scope_mode=self.scope_mode,
                allow_network_download=True,
            )

            all_decisions: List[TopicDecision] = []
            all_contributions: List[Dict[str, Any]] = []
            ls_tracker = CrossWgLiaisonTracker()

            total_rounds = len(self.rounds)
            for r_idx, m_round in enumerate(self.rounds, start=1):
                if self._is_cancelled or self.isInterruptionRequested():
                    return

                self.progress_status.emit(f"Scanning {m_round.label} ({r_idx}/{total_rounds})...", r_idx, total_rounds)

                for mtg in m_round.meetings:
                    wg_name = str(mtg.get("wg_name", "")).strip().upper()
                    mtg_num = str(mtg.get("meeting_number", "")).strip()
                    folder_name = str(mtg.get("folder_name", "")).strip()
                    mtg_date = str(mtg.get("start_date", "")).strip()

                    # 1. Parse Chairman's Notes text agreements
                    chair_agreements = helper_thread._load_meeting_agreements(mtg)
                    for ca in chair_agreements:
                        wi_str = str(ca.get("wi", "") or ca.get("study_item", "")).strip()
                        ai_str = str(ca.get("ai", "") or ca.get("agenda_item", "")).strip()
                        text_str = clean_chair_notes_text(ca.get("text", "") or ca.get("summary", ""))

                        if helper_thread._matches_topic_rule(wg_name, ai_str, text_str, wi_str):
                            dec_type = normalize_chair_notes_decision_type(ca.get("decision_type", "Agreement"))
                            all_decisions.append(TopicDecision(
                                cohort_key=m_round.key,
                                wg=wg_name,
                                meeting_number=mtg_num,
                                date=mtg_date,
                                decision_type=dec_type,
                                title_or_summary=text_str,
                                tdoc_id=ca.get("tdoc_id"),
                                source=ca.get("source") or f"Chairman's Notes AI {ai_str}",
                                target_spec=ca.get("spec", ""),
                                ai=ai_str,
                                wi=wi_str,
                                status=ca.get("status", "Agreed"),
                                notes=ca.get("notes", ""),
                            ))

                    # 2. Parse Meeting TDocs
                    tdocs_rows = helper_thread._load_meeting_tdocs(mtg)
                    for row in tdocs_rows:
                        tid = str(row.get("TDoc", "") or row.get("tdoc_id", "") or row.get("TD#", "")).strip().upper()
                        if not tid:
                            continue
                        title = str(row.get("Title", "") or row.get("title", "")).strip()
                        ai_val = str(row.get("Agenda Item", "") or row.get("AI", "") or "").strip()
                        wi_val = str(row.get("Related WIs", "") or row.get("wi", "") or "").strip()
                        t_type = str(row.get("Type", "") or row.get("type", "") or "").strip()
                        status = str(row.get("TDoc Status", "") or row.get("Status", "") or "").strip()
                        source = str(row.get("Source", "") or row.get("source", "") or "").strip()
                        ai_desc = str(row.get("Agenda item description", "") or row.get("AI Description", "") or "").strip()

                        if not self.topic.matches_tdoc(tid, wg_name, ai=ai_val, title=title, related_wis=wi_val, ai_description=ai_desc):
                            continue

                        # Check decision classification
                        dec_type = classify_tdoc_decision_type(t_type, title, status)
                        if dec_type:
                            all_decisions.append(TopicDecision(
                                cohort_key=m_round.key,
                                wg=wg_name,
                                meeting_number=mtg_num,
                                date=mtg_date,
                                decision_type=dec_type,
                                title_or_summary=title,
                                tdoc_id=tid,
                                source=source,
                                target_spec=str(row.get("Spec", "") or row.get("spec", "")).strip(),
                                ai=ai_val,
                                wi=wi_val,
                                status=status,
                                source_url=resolve_tdoc_download_url(
                                    tid,
                                    docs_url=mtg.get("docs_folder_url"),
                                    main_url=mtg.get("url_key"),
                                    meta=mtg,
                                ),
                            ))

                        is_ls = "ls" in t_type.lower() or "liaison" in t_type.lower()
                        if is_ls:
                            ls_tracker.ingest_ls_tdoc(
                                tdoc_id=tid,
                                wg=wg_name,
                                meeting_number=mtg_num,
                                cycle=m_round.key,
                                tdoc_type=t_type,
                                title=title,
                                source=source,
                                status=status,
                                reply_to=row.get("Reply to", ""),
                                reply_in=row.get("Reply in", ""),
                            )

                        clean_src = CompanySanitizer.get_confirmed_source_text(source) or source
                        all_contributions.append({
                            "tdoc_id": tid,
                            "wg": wg_name,
                            "meeting_number": mtg_num,
                            "folder_name": folder_name,
                            "round_key": m_round.key,
                            "round_label": m_round.label,
                            "date": mtg_date,
                            "title": title,
                            "ai": ai_val,
                            "ai_desc": ai_desc,
                            "wi": wi_val,
                            "type": t_type,
                            "status": status,
                            "source": source,
                            "clean_source": clean_src,
                            "spec": str(row.get("Spec", "") or row.get("spec", "")).strip(),
                            "is_decision": bool(dec_type),
                            "decision_type": dec_type or "",
                            "source_url": resolve_tdoc_download_url(tid, folder_name),
                            "raw_row": row,
                        })

            if not self._is_cancelled:
                self.collection_finished.emit(all_decisions, all_contributions, ls_tracker)
        except Exception as e:
            logger.exception(f"[InputsCollectorThread] Collection failed: {e}")
            if not self._is_cancelled:
                self.collection_failed.emit(str(e))
