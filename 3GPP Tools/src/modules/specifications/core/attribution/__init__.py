# --- File: src/modules/specifications/core/attribution/__init__.py ---
"""Modular 3GPP Clause-to-TDoc Attribution package."""

from .engine import SpecClauseAttributionEngine
from .text_helpers import BOILERPLATE_STOPWORDS
from .workers import (
    SpecClauseAttributionThread,
    SpecClauseDeepVerifyThread,
    SpecFeatureOriginRankingThread,
)

__all__ = [
    "BOILERPLATE_STOPWORDS",
    "SpecClauseAttributionEngine",
    "SpecClauseAttributionThread",
    "SpecClauseDeepVerifyThread",
    "SpecFeatureOriginRankingThread",
]
