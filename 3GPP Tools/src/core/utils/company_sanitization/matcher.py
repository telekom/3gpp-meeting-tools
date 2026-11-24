# --- File: src/core/utils/company_sanitization/matcher.py ---
"""External JSON rule loading, C-accelerated regex compilation, and company matching."""

import json
import logging
import re
from pathlib import Path
from typing import Callable, Dict, List, Optional, Pattern, Tuple

logger = logging.getLogger(__name__)


class CompanyMatcherMixin:
    """Mixin compiling company synonym rules and matching text fragments."""

    @classmethod
    def get_config_paths(cls) -> Tuple[Path, Path]:
        """Returns (default_config_path, user_config_path) for externalized company synonym rules."""
        config_dir = Path(__file__).resolve().parent.parent.parent.parent / "config"
        return config_dir / "company_synonyms.json", config_dir / "company_synonyms_user.json"

    @staticmethod
    def _split_top_level_branches(pattern: str) -> List[str]:
        """Splits a regex pattern string on top-level `|` operators (outside parentheses/character classes)."""
        branches: List[str] = []
        depth_paren = 0
        in_char_class = False
        curr: List[str] = []
        i = 0
        n = len(pattern)
        while i < n:
            ch = pattern[i]
            if ch == '\\' and i + 1 < n:
                curr.append(pattern[i:i + 2])
                i += 2
                continue
            if ch == '[' and not in_char_class:
                in_char_class = True
            elif ch == ']' and in_char_class:
                in_char_class = False
            elif not in_char_class:
                if ch == '(':
                    depth_paren += 1
                elif ch == ')':
                    depth_paren = max(0, depth_paren - 1)
                elif ch == '|' and depth_paren == 0:
                    branches.append(''.join(curr))
                    curr = []
                    i += 1
                    continue
            curr.append(ch)
            i += 1
        branches.append(''.join(curr))
        return branches

    @classmethod
    def _compile_fast_matchers(cls) -> None:
        """
        Compiles `SIGNATURE_SYNONYMS_REGEX` into a 3-stage C-accelerated execution plan:
        1. `_FAST_WORD_INDEX`: O(1) hash lookup for pure `\\b<word>\\b` branches (`\\bkt\\b`, `\\bnec\\b`, `\\bbupt\\b`, etc.)
        2. `_FAST_PLAIN_SUB_RULES`: C-level `sub in clean` substring checks for literal branches (`nokia`, `samsung`, etc.)
        3. `_FAST_ANCHORED_REGEX_RULES`: C-level `anchor in clean` pre-filter guarding `regex.search(clean)` for complex branches
        4. `_FAST_UNANCHORED_REGEX_RULES`: Fallback `regex.search(clean)` for any custom regex without a fixed literal anchor
        """
        word_index: Dict[str, List[Tuple[int, str]]] = {}
        plain_sub_rules: List[Tuple[int, str, str]] = []
        anchored_rx: List[Tuple[int, str, str, Callable]] = []
        unanchored_rx: List[Tuple[int, str, Callable]] = []

        items = list(cls.SIGNATURE_SYNONYMS_REGEX.items())
        for idx, (name, rx) in enumerate(items):
            if rx.flags & ~re.UNICODE:
                unanchored_rx.append((idx, name, rx.search))
                continue

            pat = rx.pattern
            branches = cls._split_top_level_branches(pat)
            for b in branches:
                b_s = b.strip()
                while b_s.startswith('(') and b_s.endswith(')') and not b_s.startswith('(?'):
                    inner = b_s[1:-1]
                    d = 0
                    in_cc = False
                    ok = True
                    k = 0
                    while k < len(inner):
                        c = inner[k]
                        if c == '\\':
                            k += 2
                            continue
                        if c == '[' and not in_cc:
                            in_cc = True
                        elif c == ']' and in_cc:
                            in_cc = False
                        elif not in_cc:
                            if c == '(':
                                d += 1
                            elif c == ')':
                                d -= 1
                                if d < 0:
                                    ok = False
                                    break
                        k += 1
                    if ok and d == 0 and not in_cc:
                        b_s = inner.strip()
                    else:
                        break

                if not b_s:
                    continue

                m_wb = re.fullmatch(r'\\b([a-z0-9_]+)\\b', b_s)
                if m_wb:
                    word_index.setdefault(m_wb.group(1), []).append((idx, name))
                    continue

                if all(ch not in b_s for ch in r'.^$*+?{}[]\|()') and b_s == b_s.lower():
                    plain_sub_rules.append((idx, name, b_s))
                    continue

                try:
                    b_rx = re.compile(b_s)
                except re.error:
                    b_rx = rx
                b_no_wb = b_s[2:] if b_s.startswith(r'\b') else b_s
                b_required = re.sub(r'(\[[^\]]*\]|\([^\)]*\)|\\.|[^\\])[\?\*]', ' ', b_no_wb)
                b_required = re.sub(r'\[[^\]]*\]', ' ', b_required)
                b_required = re.sub(r'\([^\)]*\)', ' ', b_required)
                b_required = re.sub(r'\\[bBsSwWdD]', ' ', b_required)
                b_required = b_required.replace(r'\.', '.')
                anchor = None
                if '|' not in b_required and '{' not in b_required and '+' not in b_required and '\\' not in b_required:
                    lits = re.findall(r'[a-z0-9_@.-]{2,}', b_required)
                    if lits:
                        candidate = max(lits, key=len)
                        if candidate == candidate.lower():
                            anchor = candidate
                if anchor:
                    anchored_rx.append((idx, name, anchor, b_rx.search))
                else:
                    unanchored_rx.append((idx, name, b_rx.search))

        cls._FAST_WORD_INDEX = word_index
        cls._FAST_PLAIN_SUB_RULES = plain_sub_rules
        cls._FAST_ANCHORED_REGEX_RULES = anchored_rx
        cls._FAST_UNANCHORED_REGEX_RULES = unanchored_rx
        cls._COMPILED_DICT_SNAPSHOT = tuple(items)

    @classmethod
    def clear_caches(cls) -> None:
        """Clears all internal LRU caches so dynamic rule changes take effect immediately."""
        for method in (
            cls.tokenize_source_segments,
            cls.parse_contributors_detailed,
            cls._get_primary_contributors_cached,
        ):
            if hasattr(method, "cache_clear"):
                method.cache_clear()

    @classmethod
    def reload_rules(
        cls,
        default_config_path: Optional[Path] = None,
        user_config_path: Optional[Path] = None,
    ) -> None:
        """
        Loads or reloads company synonym regex rules and descriptive metadata from:
        1. Built-in `_BUILTIN_SYNONYMS_PATTERNS` (guaranteed zero-I/O baseline)
        2. `src/config/company_synonyms.json` (tracked catalog with full_name, category, description)
        3. `src/config/company_synonyms_user.json` (optional user overrides/additions)

        Mutates `SIGNATURE_SYNONYMS_REGEX` and `COMPANY_METADATA` in-place to preserve
        any external references, rebuilds the C-accelerated fast matcher index, and clears LRU caches.
        """
        def_path, usr_path = cls.get_config_paths()
        if default_config_path is not None:
            def_path = Path(default_config_path)
        if user_config_path is not None:
            usr_path = Path(user_config_path)

        new_regex: Dict[str, Pattern] = {
            name: re.compile(pattern)
            for name, pattern in cls._BUILTIN_SYNONYMS_PATTERNS.items()
        }
        new_meta: Dict[str, Dict[str, str]] = {
            name: {
                "pattern": pattern,
                "full_name": name,
                "category": "Contributor",
                "description": "",
            }
            for name, pattern in cls._BUILTIN_SYNONYMS_PATTERNS.items()
        }

        for cfg_path in (def_path, usr_path):
            if not cfg_path or not cfg_path.is_file():
                continue
            try:
                raw = json.loads(cfg_path.read_text(encoding="utf-8"))
                companies_block = raw.get("companies", raw) if isinstance(raw, dict) else {}
                if not isinstance(companies_block, dict):
                    continue
                for canonical_name, entry in companies_block.items():
                    if not canonical_name or str(canonical_name).startswith("_"):
                        continue
                    if isinstance(entry, str):
                        pattern_str = entry
                        meta_entry = {
                            "pattern": pattern_str,
                            "full_name": new_meta.get(canonical_name, {}).get("full_name", canonical_name),
                            "category": new_meta.get(canonical_name, {}).get("category", "Contributor"),
                            "description": new_meta.get(canonical_name, {}).get("description", ""),
                        }
                    elif isinstance(entry, dict):
                        pattern_str = entry.get("pattern") or new_meta.get(canonical_name, {}).get("pattern", "")
                        if not pattern_str:
                            continue
                        prev = new_meta.get(canonical_name, {})
                        meta_entry = {
                            "pattern": str(pattern_str),
                            "full_name": str(entry.get("full_name") or prev.get("full_name") or canonical_name),
                            "category": str(entry.get("category") or prev.get("category") or "Contributor"),
                            "description": str(entry.get("description") or prev.get("description") or ""),
                        }
                    else:
                        continue

                    try:
                        compiled = re.compile(pattern_str)
                        new_regex[canonical_name] = compiled
                        new_meta[canonical_name] = meta_entry
                    except re.error as exc:
                        logger.warning(
                            "Invalid regex pattern for company %r in %s: %s",
                            canonical_name,
                            cfg_path,
                            exc,
                        )
            except Exception as exc:
                logger.warning("Failed to load company synonyms from %s: %s", cfg_path, exc)

        cls.SIGNATURE_SYNONYMS_REGEX.clear()
        cls.SIGNATURE_SYNONYMS_REGEX.update(new_regex)
        cls.COMPANY_METADATA.clear()
        cls.COMPANY_METADATA.update(new_meta)

        cls._compile_fast_matchers()
        cls.clear_caches()

    @classmethod
    def get_company_metadata(cls, canonical_name: str) -> Dict[str, str]:
        """
        Returns descriptive metadata for a canonical company name:
        `{"pattern": ..., "full_name": ..., "category": ..., "description": ...}`.
        """
        return dict(
            cls.COMPANY_METADATA.get(
                canonical_name,
                {
                    "pattern": cls._BUILTIN_SYNONYMS_PATTERNS.get(canonical_name, ""),
                    "full_name": canonical_name,
                    "category": "Contributor",
                    "description": "",
                },
            )
        )

    @classmethod
    def _match_text_to_companies(cls, text_fragment: str) -> list:
        if not text_fragment:
            return []
        clean = cls.SOURCE_REPLACE_REGEX.sub(' ', str(text_fragment)).strip().lower()
        if not clean:
            return []

        if cls._COMPILED_DICT_SNAPSHOT is None or len(cls.SIGNATURE_SYNONYMS_REGEX) != len(cls._COMPILED_DICT_SNAPSHOT):
            cls._compile_fast_matchers()

        hits: Dict[int, str] = {}
        word_index = cls._FAST_WORD_INDEX
        if word_index:
            for w in cls._WORD_TOKEN_RE.findall(clean):
                entries = word_index.get(w)
                if entries:
                    for idx, name in entries:
                        hits[idx] = name

        for idx, name, sub in cls._FAST_PLAIN_SUB_RULES:
            if sub in clean:
                hits[idx] = name

        for idx, name, anchor, fn in cls._FAST_ANCHORED_REGEX_RULES:
            if anchor in clean and fn(clean):
                hits[idx] = name

        for idx, name, fn in cls._FAST_UNANCHORED_REGEX_RULES:
            if fn(clean):
                hits[idx] = name

        if len(hits) <= 1:
            return list(hits.values())
        return [hits[k] for k in sorted(hits)]
