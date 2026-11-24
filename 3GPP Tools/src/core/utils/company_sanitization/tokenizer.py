# --- File: src/core/utils/company_sanitization/tokenizer.py ---
"""Source string segmentation into confirmed vs tentative co-signers and contributor extraction."""

from functools import lru_cache


class CompanyTokenizerMixin:
    """Mixin tokenizing raw 3GPP Source fields and resolving confirmed/tentative contributors."""

    @classmethod
    @lru_cache(maxsize=65536)
    def tokenize_source_segments(cls, original_source: str) -> tuple:
        """
        Splits a raw 3GPP TDoc 'Source' string into a tuple of (raw_text, is_tentative) segments
        preserving exact original characters so concatenating all `raw_text` reproduces `original_source`.

        A segment is marked `is_tentative=True` when:
        - It is enclosed in square brackets `[...]` (excluding role tags like `[pen holder]` or `[Rapporteur]`), or
        - It carries a question mark `?` or `(?)` (indicating an unconfirmed carry-over co-signer in CR/pCR revisions).
        """
        if not original_source:
            return ()

        s = str(original_source)
        # Fast path when neither '?' nor '[' nor ']' is present
        if '?' not in s and '[' not in s and ']' not in s:
            return ((s, False),)

        # Identify character ranges of role brackets like [pen holder] so they are not treated as tentative companies
        role_spans = [m.span() for m in cls.ROLE_BRACKET_RE.finditer(s)]

        def _is_in_role_bracket(pos: int) -> bool:
            return any(start <= pos < end for start, end in role_spans)

        # Step 1: Partition top-level vs bracketed `[...]` regions
        regions = []
        n = len(s)
        i = 0
        while i < n:
            if s[i] == '[' and not _is_in_role_bracket(i):
                j = i + 1
                depth = 1
                while j < n and depth > 0:
                    if s[j] == '[' and not _is_in_role_bracket(j):
                        depth += 1
                    elif s[j] == ']':
                        depth -= 1
                    j += 1
                while j < n and s[j] in ' \t':
                    if j + 1 < n and s[j + 1] == '?':
                        j += 2
                    elif s[j:].startswith('(?)'):
                        j += 3
                    else:
                        break
                while j < n and (s[j] == '?' or s[j:].startswith('(?)')):
                    j += 3 if s[j:].startswith('(?)') else 1
                regions.append((s[i:j], True))
                i = j
            else:
                j = i
                while j < n and not (s[j] == '[' and not _is_in_role_bracket(j)):
                    j += 1
                regions.append((s[i:j], False))
                i = j

        # Step 2: Within non-bracketed regions, split into individual company tokens while preserving delimiters
        raw_tokens = []
        for reg_text, is_bracketed in regions:
            if is_bracketed:
                raw_tokens.append([reg_text, True])
                continue

            pos = 0
            reg_len = len(reg_text)
            curr_start = 0
            while pos < reg_len:
                ch = reg_text[pos]
                if ch in ',;+':
                    suffix_match = cls.CORP_SUFFIX_RE.match(reg_text[pos:]) if ch == ',' else None
                    if suffix_match and pos > curr_start:
                        pos += suffix_match.end()
                        continue

                    seg = reg_text[curr_start:pos]
                    delim = ch
                    pos += 1
                    while pos < reg_len and reg_text[pos] in ' \t':
                        delim += reg_text[pos]
                        pos += 1
                    if seg:
                        has_q = ('?' in seg) or (']' in seg and '[' not in seg)
                        raw_tokens.append([seg, has_q])
                    raw_tokens.append([delim, False])
                    curr_start = pos
                    continue

                if ch in '?( ' and ('?' in reg_text[pos:pos + 6]):
                    q_match = cls._INLINE_Q_RE.match(reg_text[pos:])
                    if q_match:
                        q_end = pos + q_match.end()
                        rest = reg_text[q_end:]
                        if rest and not cls._AFTER_Q_DELIM_RE.match(rest):
                            seg = reg_text[curr_start:q_end]
                            raw_tokens.append([seg, True])
                            ws_match = cls._LEADING_WS_RE.match(rest)
                            if ws_match:
                                raw_tokens.append([ws_match.group(0), False])
                                q_end += ws_match.end()
                            pos = q_end
                            curr_start = pos
                            continue

                pos += 1

            if curr_start < reg_len:
                seg = reg_text[curr_start:]
                has_q = ('?' in seg) or (']' in seg and '[' not in seg)
                raw_tokens.append([seg, has_q])

        # Step 3: Merge adjacent segments that have the same tentative state
        if not raw_tokens:
            return ((s, False),)

        for idx in range(1, len(raw_tokens) - 1):
            tok_text, tok_tent = raw_tokens[idx]
            if not tok_tent and not tok_text.strip(',;+ \t'):
                if raw_tokens[idx - 1][1] and raw_tokens[idx + 1][1]:
                    raw_tokens[idx][1] = True

        merged = []
        for tok_text, tok_tent in raw_tokens:
            if not tok_text:
                continue
            if merged and merged[-1][1] == tok_tent:
                merged[-1] = (merged[-1][0] + tok_text, tok_tent)
            else:
                merged.append((tok_text, tok_tent))

        return tuple(merged)

    @classmethod
    @lru_cache(maxsize=65536)
    def parse_contributors_detailed(cls, original_source: str) -> tuple:
        """
        Parses a raw Source string into (confirmed_companies, tentative_companies).
        - `confirmed_companies`: tuple of canonical company names that are confirmed contributors.
        - `tentative_companies`: tuple of canonical company names that only appear with `?`, `(?)`, or inside `[...]`.
        If a company appears as both confirmed and tentative in the same string, confirmed takes precedence.
        """
        if not original_source:
            return (), ()

        s = str(original_source)
        if '?' not in s and '[' not in s and ']' not in s:
            confirmed = tuple(cls._match_text_to_companies(s))
            return confirmed, ()

        segments = cls.tokenize_source_segments(s)
        confirmed_text_parts = []
        tentative_text_parts = []

        for seg_text, is_tent in segments:
            if is_tent:
                tentative_text_parts.append(seg_text)
            else:
                confirmed_text_parts.append(seg_text)

        confirmed_list = cls._match_text_to_companies(" ".join(confirmed_text_parts))
        confirmed_set = set(confirmed_list)

        tentative_raw = cls._match_text_to_companies(" ".join(tentative_text_parts))
        tentative_list = [c for c in tentative_raw if c not in confirmed_set]

        return tuple(confirmed_list), tuple(tentative_list)

    @classmethod
    def get_matching_contributors(cls, original_source: str, include_tentative: bool = False) -> list:
        """
        Returns canonical contributor names matched in `original_source`.
        By default (`include_tentative=False`), excludes unconfirmed/tentative sources
        marked with `?`, `(?)`, or enclosed in square brackets `[...]`.
        """
        if not original_source:
            return []

        confirmed, tentative = cls.parse_contributors_detailed(str(original_source))
        if include_tentative:
            return list(confirmed) + [c for c in tentative if c not in confirmed]
        return list(confirmed)

    @classmethod
    @lru_cache(maxsize=65536)
    def _get_primary_contributors_cached(cls, original_source: str, include_tentative: bool = False) -> tuple:
        s = str(original_source).strip()
        if not s:
            return ()
        segments = cls.tokenize_source_segments(s)
        if not segments:
            return ()

        first_seg_text, first_is_tent = segments[0]
        if first_is_tent and not include_tentative:
            return ()

        parts = cls._PRIMARY_SPLIT_RE.split(first_seg_text, maxsplit=1)
        first_author = parts[0] if parts else first_seg_text
        return tuple(cls.get_matching_contributors(first_author, include_tentative=include_tentative))

    @classmethod
    def get_primary_contributors(cls, original_source: str, include_tentative: bool = False) -> list:
        """
        Extracts the primary (first-listed) contributor from `original_source`.
        Respects unconfirmed/tentative exclusions unless `include_tentative=True`.
        """
        if not original_source:
            return []
        return list(cls._get_primary_contributors_cached(str(original_source), bool(include_tentative)))

    @classmethod
    def has_tentative_sources(cls, original_source: str) -> bool:
        """Returns True if `original_source` contains any tentative (`?` or `[...]`) segments."""
        if not original_source:
            return False
        s = str(original_source)
        if '?' not in s and '[' not in s and ']' not in s:
            return False
        return any(is_tent for _, is_tent in cls.tokenize_source_segments(s))
