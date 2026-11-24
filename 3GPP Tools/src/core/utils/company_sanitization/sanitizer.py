# --- File: src/core/utils/company_sanitization/sanitizer.py ---
"""Composed CompanySanitizer class preserving single-source-of-truth class state."""

import re
from typing import Callable, Dict, List, Optional, Pattern, Tuple

from .defaults import (
    AFTER_Q_DELIM_RE,
    BUILTIN_SYNONYMS_PATTERNS,
    COMMA_SEMICOLON_SPLIT_RE,
    CONFIRMED_NORM_SPLIT_RE,
    CORP_SUFFIX_RE,
    INLINE_Q_RE,
    LEADING_WS_RE,
    PRIMARY_SPLIT_RE,
    ROLE_BRACKET_RE,
    SOURCE_REPLACE_REGEX,
    TENTATIVE_INNER_SPLIT_RE,
    TENTATIVE_OUTER_STRIP_RE,
    WORD_TOKEN_RE,
)
from .formatters import CompanyFormatterMixin
from .matcher import CompanyMatcherMixin
from .tokenizer import CompanyTokenizerMixin


class CompanySanitizer(CompanyMatcherMixin, CompanyTokenizerMixin, CompanyFormatterMixin):
    """Normalizes 3GPP company names and separates confirmed vs tentative co-signers."""

    _BUILTIN_SYNONYMS_PATTERNS: Dict[str, str] = BUILTIN_SYNONYMS_PATTERNS

    SIGNATURE_SYNONYMS_REGEX: Dict[str, Pattern] = {
        name: re.compile(pattern)
        for name, pattern in _BUILTIN_SYNONYMS_PATTERNS.items()
    }

    # Rich descriptive metadata loaded from src/config/company_synonyms.json (+ _user.json)
    # Structure: {canonical_name: {"pattern": str, "full_name": str, "category": str, "description": str}}
    COMPANY_METADATA: Dict[str, Dict[str, str]] = {}

    # Precompiled fast-matching structures rebuilt by `_compile_fast_matchers()`
    _WORD_TOKEN_RE: Pattern = WORD_TOKEN_RE
    _FAST_WORD_INDEX: Dict[str, List[Tuple[int, str]]] = {}
    _FAST_PLAIN_SUB_RULES: List[Tuple[int, str, str]] = []
    _FAST_ANCHORED_REGEX_RULES: List[Tuple[int, str, str, Callable]] = []
    _FAST_UNANCHORED_REGEX_RULES: List[Tuple[int, str, Callable]] = []
    _COMPILED_DICT_SNAPSHOT: Optional[Tuple[Tuple[str, Pattern], ...]] = None

    SOURCE_REPLACE_REGEX = SOURCE_REPLACE_REGEX
    ROLE_BRACKET_RE = ROLE_BRACKET_RE
    CORP_SUFFIX_RE = CORP_SUFFIX_RE

    _INLINE_Q_RE = INLINE_Q_RE
    _AFTER_Q_DELIM_RE = AFTER_Q_DELIM_RE
    _LEADING_WS_RE = LEADING_WS_RE
    _PRIMARY_SPLIT_RE = PRIMARY_SPLIT_RE
    _COMMA_SEMICOLON_SPLIT_RE = COMMA_SEMICOLON_SPLIT_RE
    _CONFIRMED_NORM_SPLIT_RE = CONFIRMED_NORM_SPLIT_RE
    _TENTATIVE_OUTER_STRIP_RE = TENTATIVE_OUTER_STRIP_RE
    _TENTATIVE_INNER_SPLIT_RE = TENTATIVE_INNER_SPLIT_RE


# Load externalized rules and descriptive metadata at module initialization
CompanySanitizer.reload_rules()
