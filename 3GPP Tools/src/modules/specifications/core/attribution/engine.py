# --- File: src/modules/specifications/core/attribution/engine.py ---
"""Composed SpecClauseAttributionEngine combining text helpers, matching, verification, and ranking."""

from typing import Any, Optional

from core.utils.paths import get_project_root
from modules.specifications.core.spec_history_db import SpecHistoryDatabase
from modules.specifications.core.spec_tdoc_report import SpecTDocReportGenerator

from .clause_matcher import ClauseMatcherMixin
from .docx_verifier import DocxVerifierMixin
from .origin_ranker import OriginRankerMixin
from .text_helpers import AttributionTextMixin


class SpecClauseAttributionEngine(
    ClauseMatcherMixin,
    DocxVerifierMixin,
    OriginRankerMixin,
    AttributionTextMixin,
):
    """Multi-stage attribution engine linking 3GPP clause modifications to originating CRs/p-CRs."""

    def __init__(
        self,
        history_db: Optional[SpecHistoryDatabase] = None,
        search_db: Optional[Any] = None,
        report_gen: Optional[SpecTDocReportGenerator] = None,
    ):
        self.history_db = history_db or SpecHistoryDatabase(get_project_root() / "db" / "3gpp_spec_history.db")
        self.search_db = search_db
        self.report_gen = report_gen or SpecTDocReportGenerator(history_db=self.history_db)
