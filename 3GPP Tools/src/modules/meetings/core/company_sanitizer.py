# --- File: src/modules/meetings/core/company_sanitizer.py ---
import re

from core.config.source_companies import SIGNATURE_SYNONYMS_REGEX, get_matching_contributors


class CompanySanitizer:
    @classmethod
    def get_matching_contributors(cls, original_source: str) -> list:
        return get_matching_contributors(original_source)