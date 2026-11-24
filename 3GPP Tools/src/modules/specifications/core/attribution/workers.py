# --- File: src/modules/specifications/core/attribution/workers.py ---
"""Background QThread workers for non-blocking clause attribution, verification, and origin CR ranking."""

import logging
from typing import Any, Dict, List, Optional, Sequence

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import get_cache_root, get_project_root
from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener

from .engine import SpecClauseAttributionEngine

logger = logging.getLogger(__name__)


class SpecClauseAttributionThread(QThread):
    """Background worker executing clause attribution without freezing the Qt event loop."""

    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(
        self,
        engine: SpecClauseAttributionEngine,
        spec_number: str,
        clause_number: str,
        target_version: str,
        base_version: Optional[str] = None,
        clause_title: str = "",
        diff_text: str = "",
        clause_content: str = "",
        is_new_clause: Optional[bool] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.engine = engine
        self.spec_number = spec_number
        self.clause_number = clause_number
        self.target_version = target_version
        self.base_version = base_version
        self.clause_title = clause_title
        self.diff_text = diff_text
        self.clause_content = clause_content

    def run(self):
        try:
            results = self.engine.attribute_clause(
                spec_number=self.spec_number,
                clause_number=self.clause_number,
                target_version=self.target_version,
                base_version=self.base_version,
                clause_title=self.clause_title,
                diff_text=self.diff_text,
                clause_content=self.clause_content,
                is_new_clause=getattr(self, "is_new_clause", None),
                progress_cb=self.progress.emit,
            )
            self.finished.emit(results)
        except Exception as e:
            logger.error(f"Error during clause attribution thread: {e}", exc_info=True)
            self.error.emit(str(e))


class SpecClauseDeepVerifyThread(QThread):
    """Background worker executing targeted Stage 2 content fingerprinting on top candidates."""

    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(
        self,
        engine: SpecClauseAttributionEngine,
        candidates: List[Dict[str, Any]],
        diff_text: str,
        parent=None,
    ):
        super().__init__(parent)
        self.engine = engine
        self.candidates = [dict(c) for c in candidates]
        self.diff_text = diff_text

    def run(self):
        try:
            top_k = self.candidates[:3]
            total = len(top_k)
            for idx, cand in enumerate(top_k):
                if self.isInterruptionRequested():
                    return

                tdoc_id = cand.get("tdoc", "")
                self.progress.emit(
                    int((idx / max(1, total)) * 100),
                    f"Verifying candidate {idx + 1}/{total} ({tdoc_id}) in DOCX...",
                )
                local_path = SpecTDocOpener.resolve_local_tdoc_file(tdoc_id)
                if not local_path or not local_path.exists():
                    from core.network.session import NetworkSession
                    from modules.meetings.core.meetings_db import MeetingsDatabase
                    from modules.meetings.core.tdocs.tdoc_url_resolver import (
                        normalize_3gpp_url,
                        resolve_tdoc_download_url,
                    )

                    data_db = get_project_root() / "db" / "3gpp_data.db"
                    meetings_db = MeetingsDatabase(data_db)
                    meeting = meetings_db.find_meeting_by_tdoc(tdoc_id)
                    download_url = ""
                    target_dir = get_cache_root() / "tdocs"
                    if meeting:
                        docs_url = meeting.get("docs_folder_url") or ""
                        folder_name = meeting.get("folder_name") or meeting.get("meeting_number", "unknown")
                        target_dir = get_cache_root() / folder_name / "TDocs"
                        if docs_url:
                            norm_docs = normalize_3gpp_url(docs_url)
                            download_url = f"{norm_docs.rstrip('/')}/{tdoc_id}.zip"
                    if not download_url:
                        download_url = resolve_tdoc_download_url(tdoc_id)

                    if download_url:
                        try:
                            target_dir.mkdir(parents=True, exist_ok=True)
                            zip_dest = target_dir / f"{tdoc_id}.zip"
                            NetworkSession.download_file(
                                url=download_url,
                                dest_path=zip_dest,
                                cancel_cb=self.isInterruptionRequested,
                                atomic=True,
                            )
                            local_path = zip_dest
                        except Exception as dl_err:
                            logger.debug(f"Could not download {tdoc_id} for deep verify: {dl_err}")

                if local_path and local_path.exists():
                    matched, pct, snip = self.engine.verify_candidate_in_docx(
                        tdoc=tdoc_id,
                        diff_text=self.diff_text,
                        doc_path=local_path,
                        cr_number=cand.get("cr_number", ""),
                    )
                    if matched:
                        cand["is_verified"] = True
                        cand["verified_match_pct"] = pct
                        cand["verified_snippet"] = snip
                        cand["score"] = max(cand["score"], min(100, cand["score"] + 20))
                        cand["confidence"] = "Verified"
                        badge = f"✅ Verified in DOCX ({pct}% Text Match)"
                        if badge not in cand.get("evidence_badges", []):
                            cand.setdefault("evidence_badges", []).insert(0, badge)
                        cand["local_doc_path"] = str(local_path)

            self.candidates.sort(key=lambda x: (x.get("is_verified", False), x.get("score", 0)), reverse=True)
            self.progress.emit(100, "Deep verification complete.")
            self.finished.emit(self.candidates)
        except Exception as e:
            logger.error(f"Error during deep verification: {e}", exc_info=True)
            self.error.emit(str(e))


class SpecFeatureOriginRankingThread(QThread):
    """Background worker executing feature-level Origin CR ranking across specification releases."""

    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(
        self,
        engine: SpecClauseAttributionEngine,
        spec_number: str,
        query: str,
        target_version: Optional[str] = None,
        affected_clauses_hint: Optional[Sequence[str]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.engine = engine
        self.spec_number = spec_number
        self.query = query
        self.target_version = target_version
        self.affected_clauses_hint = list(affected_clauses_hint) if affected_clauses_hint else None

    def run(self):
        try:
            results = self.engine.rank_origin_crs_for_query(
                spec_number=self.spec_number,
                query=self.query,
                target_version=self.target_version,
                affected_clauses_hint=self.affected_clauses_hint,
                progress_cb=self.progress.emit,
            )
            self.finished.emit(results)
        except Exception as e:
            logger.error(f"Error during feature origin ranking thread: {e}", exc_info=True)
            self.error.emit(str(e))
