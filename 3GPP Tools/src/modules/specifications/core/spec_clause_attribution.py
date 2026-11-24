# --- File: src/modules/specifications/core/spec_clause_attribution.py ---
"""Backward-compatible re-export shim for 3GPP Clause-to-TDoc Attribution Engine."""

import logging

from modules.specifications.core.attribution import (
    BOILERPLATE_STOPWORDS,
    SpecClauseAttributionEngine,
    SpecClauseAttributionThread,
    SpecClauseDeepVerifyThread,
    SpecFeatureOriginRankingThread,
)

logger = logging.getLogger(__name__)

__all__ = [
    "BOILERPLATE_STOPWORDS",
    "SpecClauseAttributionEngine",
    "SpecClauseAttributionThread",
    "SpecClauseDeepVerifyThread",
    "SpecFeatureOriginRankingThread",
]
