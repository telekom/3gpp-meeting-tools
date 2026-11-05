# --- File: src/core/utils/company_sanitizer.py ---
import html
import re
from functools import lru_cache


class CompanySanitizer:
    SIGNATURE_SYNONYMS_REGEX = {
        'Deutsche Telekom': re.compile(r'(\bdt\b)|(deutsche tele[kc]om)|(@magenta\.at)|(@t-mobile\.pl)|(@telekom\.de)'),
        'KT': re.compile(r'(\bkt\b)'),
        'Nokia': re.compile(r'nokia'),
        'Qualcomm': re.compile(r'qualcom[m]?'),
        'Huawei': re.compile(r'(huawei)|(h[i]?s[i]?l[l]?icon)'),

        # ---> THE FIX: Added a \b (word boundary) so it won't trigger on cic[tmobile]
        'T-Mobile USA': re.compile(r'(\bt[-]?mobile(?!\.pl))|(@t-mobile\.com)'),

        # ---> NEW: Explicit entry for CICT
        'CICT': re.compile(r'cict'),

        'Verizon': re.compile(r'verizon'),
        'ZTE': re.compile(r'zte'),
        'NTT DoCoMo': re.compile(r'(ntt)|(docomo)'),
        'Samsung': re.compile(r'samsung'),
        'Blackberry': re.compile(r'blackberry'),
        'Toyota': re.compile(r'toyota'),
        'LG': re.compile(r'lg[e]?( electronics)?'),
        'Cisco': re.compile(r'cisco'),
        'Oppo': re.compile(r'oppo'),
        'Interdigital': re.compile(r'inter[ ]?digital|interdigtial|intedigital'),
        'Mediatek': re.compile(r'mediatek'),
        'NEC': re.compile(r'nec'),
        'Orange': re.compile(r'orange'),
        'Continental': re.compile(r'continental'),
        'Thales': re.compile(r'thales'),
        'China Telecom': re.compile(r'china telecom'),
        'TIM': re.compile(r'(telecom italia)|(tim)'),
        'Ericsson': re.compile(r'eric[s]?sson'),
        'Convida Wireless': re.compile(r'convida'),
        'China Mobile': re.compile(r'china mobile'),
        'China Unicom': re.compile(r'china unicom'),
        'Comtech': re.compile(r'comtech'),
        'Gemalto': re.compile(r'gemalto'),
        'Intel': re.compile(r'intel'),
        'KDDI': re.compile(r'kddi'),
        'KPN': re.compile(r'kpn'),
        'Oracle': re.compile(r'oracle'),
        'FirstNet': re.compile(r'firstnet'),
        'Telstra': re.compile(r'telstra'),
        'Vodafone': re.compile(r'vodafone'),
        'Tencent': re.compile(r'tencent'),
        'Sprint': re.compile(r'sprint'),
        'Sony': re.compile(r'sony'),
        'CableLabs': re.compile(r'cablelabs'),

        # Word boundaries prevent collision with CATT
        'AT&T': re.compile(r'\bat[&]?t\b'),
        'Charter': re.compile(r'charter'),
        'Lenovo': re.compile(r'(motorola mobility)|(lenovo)'),
        'SK Telecom': re.compile(r'sk telecom'),
        'TNO': re.compile(r'tno'),
        'Telefonica': re.compile(r'telefonica'),
        'Softbank': re.compile(r'softbank'),
        'Volkwagen': re.compile(r'volkswagen'),
        'Vivo': re.compile(r'vivo'),
        'Xiaomi': re.compile(r'xiaomi'),
        'Alibaba': re.compile(r'alibaba'),
        'Apple': re.compile(r'apple'),
        'Broadcom': re.compile(r'broadcom'),
        'CAICT': re.compile(r'caict'),
        'CATR': re.compile(r'catr'),
        'CATT': re.compile(r'\bcatt\b'),
        'CMCC': re.compile(r'cmcc'),
        'UK HO': re.compile(r'uk ho'),
        'Affirmed': re.compile(r'affirmed'),
        'Expway': re.compile(r'expway'),
        'HPE': re.compile(r'hewlett packard'),

        'RAN WG1': re.compile(r'ran wg1'),
        'RAN WG2': re.compile(r'ran wg2'),
        'RAN WG3': re.compile(r'ran wg3'),
        'SA WG1': re.compile(r'sa wg1'),
        'SA WG2': re.compile(r'sa wg2'),
        'SA WG3': re.compile(r'sa wg3'),
        'SA WG4': re.compile(r'sa wg4'),
        'SA WG5': re.compile(r'sa wg5'),
        'SA WG6': re.compile(r'sa wg6'),
        'TSG SA': re.compile(r'tsg sa'),
        'TSG CT': re.compile(r'tsg ct'),
        'TSG RAN': re.compile(r'tsg ran'),
        'CT WG1': re.compile(r'ct wg1'),
        'CT WG2': re.compile(r'ct wg2'),
        'CT WG3': re.compile(r'ct wg3'),
        'CT WG4': re.compile(r'ct wg4'),

        'IETF': re.compile(r'ietf'),
        'IEEE': re.compile(r'ieee'),
        'BBF': re.compile(r'bbf'),
        'Swisscom': re.compile(r'swisscom'),
        'Sharp': re.compile(r'sharp'),
        'Korea Telecom': re.compile(r'korea telecom'),
        'Siemens': re.compile(r'siemens'),
        'Sierra Wireless': re.compile(r'sierra wireless'),
        'Fraunhofer HHI': re.compile(r'fraunhofer hhi'),
        'GSMA': re.compile(r'gsma'),
        'Broadband Forum': re.compile(r'broadband forum'),
        'OneM2M': re.compile(r'onem2m'),
        'Juniper Networks': re.compile(r'juniper'),
        'Spirent': re.compile(r'spirent'),
        'Airbus': re.compile(r'airbus'),
        'Mavenir': re.compile(r'mavenir'),
        'Sennheiser': re.compile(r'sennheiser'),
        'Sandvine': re.compile(r'sandvine'),
        'Philips': re.compile(r'philips'),
        'Google': re.compile(r'google'),
        'Matrixx': re.compile(r'matrixx'),
        'Comcast': re.compile(r'comcast'),
        'Turkcell': re.compile(r'turkcell'),
        'Airtel': re.compile(r'airtel'),
        'Bosch': re.compile(r'bosch'),
        'Sequans': re.compile(r'sequans'),
        'Rakuten': re.compile(r'rakute'),
        'Futurewei': re.compile(r'futurewei'),
        'Microsoft': re.compile(r'microsoft'),
        'Facebook': re.compile(r'facebook|meta'),
        'DISH': re.compile(r'dish|boost mobile'),
        'ETRI': re.compile(r'etri'),
        'INSPUR': re.compile(r'inspur'),

        # Added at SA2#175AHE
        'HONOR': re.compile(r'honor'),
        'JIO': re.compile(r'jio platforms'),
        'Kyocera': re.compile(r'kyocera'),
        'NVIDIA': re.compile(r'nvidia'),
        'Ofinno': re.compile(r'ofinno'),
        'Skylo': re.compile(r'skylo technologies'),
        'Peraton Labs': re.compile(r'peraton labs'),
        'CEWiT': re.compile(r'cewit'),
        'CANON': re.compile(r'canon'),
        'Eutelsat': re.compile(r'eutelsat'),
        'Panasonic': re.compile(r'panasonic'),
        'Quectel': re.compile(r'quectel'),
        'SpaceX': re.compile(r'spacex'),
        'Tejas Networks': re.compile(r'tejas network'),
        'ViaSat': re.compile(r'viasat'),
        'VIAVI': re.compile(r'viavi Solutions'),
        'ASUSTeK': re.compile(r'asustek'),
        'BDBOS': re.compile(r'bdbos'),
        'Keysight Technologies': re.compile(r'keysight technologies'),
        'Sateliot': re.compile(r'sateliot'),
        'Novamint': re.compile(r'novamint'),
    }

    SOURCE_REPLACE_REGEX = re.compile(
        r'\(Rapporteur\)|\(\?\)|[\[\]\?\(\)]|(([\w\d]{2,3})-(\d\d)([\d]+))',
        re.IGNORECASE,
    )

    # Role or administrative annotations placed inside [...] that are NOT unconfirmed companies
    ROLE_BRACKET_RE = re.compile(
        r'\[\s*(?:pen\s*holder|rapporteur|editor|moderator|convener|chair(?:man)?|wi\s*rapporteur|'
        r'ad-?hoc\s*chair|contact\s*person|lead|tbd|ffs|to\s+be\s+[\w\s]+)\s*\]',
        re.IGNORECASE,
    )

    # Corporate legal suffixes that sometimes follow a comma (e.g. "Rakuten Mobile, Inc?")
    CORP_SUFFIX_RE = re.compile(
        r'^,\s*(?:inc\.?|ltd\.?|llc\.?|gmbh\.?|corp\.?|corporation|co\.?|n\.?v\.?|s\.?a\.?|plc\.?|ag|kg)\b'
        r'(\s*\(\s*\?\s*\)|\s*\?+)?',
        re.IGNORECASE,
    )

    @classmethod
    def _match_text_to_companies(cls, text_fragment: str) -> list:
        if not text_fragment:
            return []
        clean = cls.SOURCE_REPLACE_REGEX.sub(' ', str(text_fragment)).strip().lower()
        if not clean:
            return []
        return [key for key, regex in cls.SIGNATURE_SYNONYMS_REGEX.items() if regex.search(clean)]

    @classmethod
    @lru_cache(maxsize=16384)
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
        regions = []  # list of (text, in_company_bracket)
        n = len(s)
        i = 0
        while i < n:
            if s[i] == '[' and not _is_in_role_bracket(i):
                # Find matching closing bracket ']' (or end of string if unclosed, e.g. ",[vivo, Huawei")
                j = i + 1
                depth = 1
                while j < n and depth > 0:
                    if s[j] == '[' and not _is_in_role_bracket(j):
                        depth += 1
                    elif s[j] == ']':
                        depth -= 1
                    j += 1
                # Include any trailing '?' or '(?)' attached to the closing bracket, e.g. "[...]?"
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
        raw_tokens = []  # list of [text, is_tentative]
        for reg_text, is_bracketed in regions:
            if is_bracketed:
                raw_tokens.append([reg_text, True])
                continue

            # Split around standard delimiters (',' ';' '+') and also after '?' / '(?)' when used as an implicit separator
            pos = 0
            reg_len = len(reg_text)
            curr_start = 0
            while pos < reg_len:
                ch = reg_text[pos]
                # Standard comma / semicolon / plus separator
                if ch in ',;+':
                    # Check if what follows the comma is a corporate suffix like ", Inc?" or ", Ltd."
                    suffix_match = cls.CORP_SUFFIX_RE.match(reg_text[pos:])
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

                # Check for "(?)" or "?" acting as an inline terminator when the author forgot a comma
                # e.g. "Huawei? Samsung? TNO?" or "ETRI?NEC?"
                q_match = re.match(r'^(\(\s*\?\s*\)|\?+)', reg_text[pos:])
                if q_match:
                    q_end = pos + q_match.end()
                    # Look ahead: if followed by whitespace + word or directly by a word (and not a comma/semicolon)
                    rest = reg_text[q_end:]
                    if rest and not re.match(r'^\s*[,;+]', rest):
                        seg = reg_text[curr_start:q_end]
                        raw_tokens.append([seg, True])
                        ws_match = re.match(r'^[ \t]+', rest)
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

        # Step 3: Merge adjacent segments that have the same tentative state (including delimiters between two tentative segments)
        if not raw_tokens:
            return ((s, False),)

        # If a pure delimiter (like ", ") sits between two tentative tokens, make the delimiter tentative too so
        # the greyed-out span is visually continuous (e.g. "Ericsson?, Samsung?, Nokia?" -> single italic grey span).
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
    @lru_cache(maxsize=16384)
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
    def get_primary_contributors(cls, original_source: str, include_tentative: bool = False) -> list:
        """
        Extracts the primary (first-listed) contributor from `original_source`.
        Respects unconfirmed/tentative exclusions unless `include_tentative=True`.
        """
        if not original_source:
            return []

        s = str(original_source).strip()
        segments = cls.tokenize_source_segments(s)
        if not segments:
            return []

        # Take the first non-delimiter segment
        first_seg_text, first_is_tent = segments[0]
        if first_is_tent and not include_tentative:
            return []

        parts = re.split(r'[,;]|\band\b', first_seg_text, maxsplit=1)
        first_author = parts[0] if parts else first_seg_text
        return cls.get_matching_contributors(first_author, include_tentative=include_tentative)

    @classmethod
    def has_tentative_sources(cls, original_source: str) -> bool:
        """Returns True if `original_source` contains any tentative (`?` or `[...]`) segments."""
        if not original_source:
            return False
        s = str(original_source)
        if '?' not in s and '[' not in s and ']' not in s:
            return False
        return any(is_tent for _, is_tent in cls.tokenize_source_segments(s))

    @classmethod
    def format_source_html(cls, original_source: str) -> str:
        """
        Renders `original_source` as rich HTML where confirmed contributors appear in standard text
        and unconfirmed/tentative contributors (`?`, `(?)`, `[...]`) are rendered greyed-out and italicized.
        """
        if not original_source:
            return ""
        s = str(original_source).strip()
        if not cls.has_tentative_sources(s):
            return html.escape(s)

        parts = []
        for seg_text, is_tent in cls.tokenize_source_segments(s):
            esc = html.escape(seg_text)
            if is_tent:
                parts.append(f'<span style="color: #94A3B8; font-style: italic;">{esc}</span>')
            else:
                parts.append(esc)
        return "".join(parts)

    @classmethod
    def build_source_tooltip_html(cls, original_source: str) -> str:
        """
        Builds a rich, self-explanatory HTML tooltip for a TDoc Source string.
        When tentative sources (`?` or `[...]`) are present, breaks down Confirmed vs. Unconfirmed
        companies and explains the 3GPP/SA2 convention and statistical exclusion rule.
        """
        if not original_source:
            return ""
        s = str(original_source).strip()
        if not cls.has_tentative_sources(s):
            return f"<div style='max-width: 420px; white-space: pre-wrap;'>{html.escape(s)}</div>"

        confirmed, tentative = cls.parse_contributors_detailed(s)
        formatted_inline = cls.format_source_html(s)

        conf_str = ", ".join(confirmed) if confirmed else "<i>None recognized</i>"
        tent_str = ", ".join(tentative) if tentative else "<i>Unconfirmed / bracketed segment(s)</i>"

        return (
            f"<div style='max-width: 440px;'>"
            f"<b>Source Field:</b><br>{formatted_inline}"
            f"<hr style='border: none; border-top: 1px solid #CBD5E1; margin: 6px 0;'>"
            f"<b>✅ Confirmed Contributors ({len(confirmed)}):</b> {conf_str}<br>"
            f"<b>❔ Unconfirmed / Carry-over ({len(tentative)}):</b> "
            f"<span style='color: #64748B; font-style: italic;'>{tent_str}</span>"
            f"<hr style='border: none; border-top: 1px solid #CBD5E1; margin: 6px 0;'>"
            f"<span style='color: #475569; font-size: 11px;'>"
            f"ℹ️ <b>How to read <code>?</code> and <code>[...]</code> in Source:</b><br>"
            f"In 3GPP revisions (especially SA2 CRs/pCRs), the source list is carried over from the "
            f"original contribution, but companies marked with <code>?</code> or <code>[...]</code> have "
            f"<b>not confirmed</b> co-signing this revision (and may even oppose the change).<br>"
            f"• <i>Unconfirmed companies are excluded from contribution counts and statistics.</i>"
            f"</span>"
            f"</div>"
        )

    @classmethod
    def get_confirmed_source_text(cls, original_source: str) -> str:
        """
        Returns clean comma-separated confirmed source segments (preserving original
        casing, qualifiers, and role brackets like '[pen holder]'), excluding tentative
        ('?' or '[...]') segments.
        """
        if not original_source:
            return ""
        if not cls.has_tentative_sources(original_source):
            return str(original_source).strip()

        confirmed_parts = []
        for seg_text, is_tent in cls.tokenize_source_segments(original_source):
            if not is_tent:
                for sub in re.split(r'[,;]+', seg_text):
                    sub_s = sub.strip()
                    if sub_s:
                        confirmed_parts.append(sub_s)
        return ", ".join(confirmed_parts)

    @classmethod
    def get_tentative_source_text(cls, original_source: str) -> str:
        """
        Returns clean comma-separated unconfirmed/tentative source segments,
        stripped of outer brackets ('[...]'), question marks ('?'), and parenthetical question marks.
        Deduplicates against confirmed sources to prevent repeated names.
        """
        if not original_source or not cls.has_tentative_sources(original_source):
            return ""

        confirmed_norm = {c.lower() for c in re.split(r'[,;\s]+', cls.get_confirmed_source_text(original_source)) if c}
        tentative_parts = []
        seen = set()

        for seg_text, is_tent in cls.tokenize_source_segments(original_source):
            if is_tent:
                cleaned = re.sub(r'^[\[\(\s]+|[\]\)\s\?？]+$', '', seg_text)
                for sub in re.split(r'[,;?？\(\)\[\]]+', cleaned):
                    s = sub.strip()
                    if s and s.lower() not in confirmed_norm and s.lower() not in seen:
                        seen.add(s.lower())
                        tentative_parts.append(s)
        return ", ".join(tentative_parts)

    @classmethod
    def format_source_markdown_header_lines(cls, original_source: str) -> List[str]:
        """
        Formats source lines for Markdown metadata headers using the Two-Line Header pattern (Approach B):
        - '- **Source:** <confirmed>'
        - If tentative sources exist, adds:
          '- **Unconfirmed / Carry-over Sources:** <tentative> *(carry-over from earlier revision; unconfirmed co-signers)*'
        """
        if not original_source:
            return ["- **Source:** Unknown"]

        if not cls.has_tentative_sources(original_source):
            return [f"- **Source:** {str(original_source).strip()}"]

        confirmed = cls.get_confirmed_source_text(original_source)
        tentative = cls.get_tentative_source_text(original_source)

        lines = []
        if confirmed:
            lines.append(f"- **Source:** {confirmed}")
        else:
            lines.append("- **Source:** *(none confirmed)*")

        if tentative:
            lines.append(
                f"- **Unconfirmed / Carry-over Sources:** {tentative} "
                f"*(carry-over from earlier revision; unconfirmed co-signers)*"
            )
        return lines

    @classmethod
    def format_source_markdown_table_cell(cls, original_source: str) -> str:
        """
        Formats a clean, single-line representation of the source for Markdown tables
        (e.g. in index.md and Company summary reports).
        - If tentative sources exist: 'Confirmed *(carry-over: Tentative)*'
        - Escapes markdown table pipes ('|') and removes newlines.
        """
        if not original_source:
            return ""

        if not cls.has_tentative_sources(original_source):
            return str(original_source).replace('\n', ' ').replace('\r', '').replace('|', '\\|').strip()

        confirmed = cls.get_confirmed_source_text(original_source)
        tentative = cls.get_tentative_source_text(original_source)

        if confirmed and tentative:
            cell = f"{confirmed} *(carry-over: {tentative})*"
        elif confirmed:
            cell = confirmed
        elif tentative:
            cell = f"*(carry-over only: {tentative})*"
        else:
            cell = str(original_source).strip()

        return cell.replace('\n', ' ').replace('\r', '').replace('|', '\\|')

    @classmethod
    def format_source_triage_context(cls, original_source: str) -> str:
        """
        Formats the Source parameter passed to the LLM for TDoc technical triage.
        Prevents LLMs from falsely attributing proposals to unconfirmed carry-over co-sponsors.
        """
        if not original_source:
            return "Unknown"

        if not cls.has_tentative_sources(original_source):
            return str(original_source).strip()

        confirmed = cls.get_confirmed_source_text(original_source)
        tentative = cls.get_tentative_source_text(original_source)

        if confirmed and tentative:
            return (
                f"{confirmed} "
                f"[Note: {tentative} are carry-over co-sponsors from an earlier draft and have not confirmed co-signing this revision]"
            )
        elif confirmed:
            return confirmed
        elif tentative:
            return f"None confirmed (carry-over co-sponsors: {tentative})"
        return str(original_source).strip()

