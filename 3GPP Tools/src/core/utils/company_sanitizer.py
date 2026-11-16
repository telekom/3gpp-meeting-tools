# --- File: src/core/utils/company_sanitizer.py ---
import html
import json
import logging
import re
from functools import lru_cache
from pathlib import Path
from typing import Callable, Dict, List, Optional, Pattern, Set, Tuple

logger = logging.getLogger(__name__)


class CompanySanitizer:
    # Built-in zero-I/O baseline catalog of canonical 3GPP contributors and regex rules.
    # Externalized descriptions, metadata, and optional user overrides are loaded from:
    #   - src/config/company_synonyms.json (tracked baseline with full_name, category, description)
    #   - src/config/company_synonyms_user.json (optional local user overrides/additions, preserved across updates)
    _BUILTIN_SYNONYMS_PATTERNS: Dict[str, str] = {
        'Deutsche Telekom': r'(\bdt\b)|(deutsche tele[kc]om)|(@magenta\.at)|(@t-mobile\.pl)|(@telekom\.de)',
        'KT': r'(\bkt\b)',
        'Nokia': r'nokia',
        'Qualcomm': r'qualcom[m]?',
        'Huawei': r'(huawei)|(h[i]?s[i]?l[l]?icon)',

        # Word boundary prevents triggering on cic[tmobile]
        'T-Mobile USA': r'(\bt[-]?mobile(?!\.pl))|(@t-mobile\.com)',

        'CICT': r'cict',
        'Verizon': r'verizon',
        'ZTE': r'\bzte\b',
        'NTT DoCoMo': r'(\bntt\b)|(docomo)',
        'Samsung': r'samsung',
        'Blackberry': r'blackberry',
        'Toyota': r'toyota',
        'LG': r'lg[e]?( electronics)?',
        'Cisco': r'cisco',
        'Oppo': r'oppo',
        'Interdigital': r'inter[ ]?digital|interdigtial|intedigital',
        'Mediatek': r'mediatek',
        'NEC': r'\bnec\b',
        'Orange': r'orange',
        'Continental': r'continental',
        'Thales': r'thales',
        'China Telecom': r'china\s*telecom',
        'TIM': r'(telecom italia)|(\btim\b)',
        'Ericsson': r'eric[s]?sson',
        'Convida Wireless': r'convida',
        'China Mobile': r'china mobile',
        'China Unicom': r'china unicom',
        'Comtech': r'comtech',
        'Gemalto': r'gemalto',
        'Intel': r'intel',
        'KDDI': r'kddi',
        'KPN': r'kpn',
        'Oracle': r'oracle',
        'FirstNet': r'firstnet',
        'Telstra': r'telstra',
        'Vodafone': r'vodafone',
        'Tencent': r'tencent',
        'Sprint': r'sprint',
        'Sony': r'sony',
        'CableLabs': r'cablelabs',

        # Word boundaries prevent collision with CATT
        'AT&T': r'\bat[&]?t\b',
        'Charter': r'charter',
        'Lenovo': r'(motorola mobility)|(lenovo)',
        'SK Telecom': r'sk telecom',
        'TNO': r'tno',
        'Telefonica': r'telefonica',
        'Softbank': r'softbank',
        'Volkwagen': r'volkswagen',
        'Vivo': r'vivo',
        'Xiaomi': r'xiaomi',
        'Alibaba': r'alibaba',
        'Apple': r'apple',
        'Broadcom': r'broadcom',
        'CAICT': r'caict',
        'CATR': r'catr',
        'CATT': r'\bcatt\b',
        'CMCC': r'cmcc',
        'UK HO': r'uk ho',
        'Affirmed': r'affirmed',
        'Expway': r'expway',
        'HPE': r'hewlett packard',

        'RAN WG1': r'ran wg1',
        'RAN WG2': r'ran wg2',
        'RAN WG3': r'ran wg3',
        'SA WG1': r'sa wg1',
        'SA WG2': r'sa wg2',
        'SA WG3': r'sa wg3',
        'SA WG4': r'sa wg4',
        'SA WG5': r'sa wg5',
        'SA WG6': r'sa wg6',
        'TSG SA': r'tsg sa',
        'TSG CT': r'tsg ct',
        'TSG RAN': r'tsg ran',
        'CT WG1': r'ct wg1',
        'CT WG2': r'ct wg2',
        'CT WG3': r'ct wg3',
        'CT WG4': r'ct wg4',

        'IETF': r'ietf',
        'IEEE': r'ieee',
        'BBF': r'bbf',
        'Swisscom': r'swisscom',
        'Sharp': r'sharp',
        'Korea Telecom': r'korea telecom',
        'Siemens': r'siemens',
        'Sierra Wireless': r'sierra wireless',
        'Fraunhofer HHI': r'fraunhofer hhi',
        'GSMA': r'gsma',
        'Broadband Forum': r'broadband forum',
        'OneM2M': r'onem2m',
        'Juniper Networks': r'juniper',
        'Spirent': r'spirent',
        'Airbus': r'airbus',
        'Mavenir': r'mavenir',
        'Sennheiser': r'sennheiser',
        'Sandvine': r'sandvine',
        'Philips': r'philips',
        'Google': r'google',
        'Matrixx': r'matrixx',
        'Comcast': r'comcast',
        'Turkcell': r'turkcell',
        'Airtel': r'airtel',
        'Bosch': r'bosch',
        'Sequans': r'sequans',
        'Rakuten': r'rakute',
        'Futurewei': r'futurewei',
        'Microsoft': r'microsoft',
        'Facebook': r'facebook|meta',
        'DISH': r'dish|boost mobile',
        'ETRI': r'etri',
        'INSPUR': r'inspur',

        # Added at SA2#175AHE
        'HONOR': r'honor',
        'JIO': r'jio platforms',
        'Kyocera': r'kyocera',
        'NVIDIA': r'nvidia',
        'Ofinno': r'ofinno',
        'Skylo': r'skylo technologies',
        'Peraton Labs': r'peraton labs',
        'CEWiT': r'cewit',
        'CANON': r'canon',
        'Eutelsat': r'eutelsat',
        'Panasonic': r'panasonic',
        'Quectel': r'quectel',
        'SpaceX': r'spacex',
        'Tejas Networks': r'tejas network',
        'ViaSat': r'viasat',
        'VIAVI': r'\bviavi\b',
        'ASUSTeK': r'asustek',
        'BDBOS': r'bdbos',
        'Keysight Technologies': r'keysight technologies',
        'Sateliot': r'sateliot',
        'Novamint': r'novamint',

        # Added at SA2#178
        'ATIS': r'\batis\b|ims911',
        'BSI': r'\bbsi\b',
        'BUPT': r'\bbupt\b|beijing university of posts',
        'CSCN': r'\bcscn\b|china satcom|china satellite network',
        'ETSI': r'\betsi\b',
        'Free Stream Technologies': r'free\s*stream\s*tech',
        'HFCL': r'\bhfcl\b',
        'HIPE IP': r'\bhipe\s*ip\b',
        'IIT Bombay': r'\biit\s*bombay\b|\biitb\b',
        'IIT Kharagpur': r'\biit\s*kharagpur\b|\biitkgp\b',
        'NCSC': r'\bncsc\b',
        'NPL': r'\bnpl\b|national physical laboratory',
        'NSA': r'\bus\s+nsa\b|\bnsa\b',
        'Pengcheng Laboratory': r'\bpeng\s*cheng\s*lab|\bpcl\b',
        'PML': r'\bpml\b|purple mountain lab',
        'Seoul National University': r'seoul national university|\bsnu\b',
        'SKY Perfect JSAT': r'sky\s*perfect|\bjsat\b',
        'Transsion': r'\btranssion\b|\binfinix\b|\btecno\b|\bitel\b',
        'UNISOC': r'\bunisoc\b|\bspreadtrum\b',
        'ZGC Institute': r'\bzgc\b|ubiquitous-x',
    }

    SIGNATURE_SYNONYMS_REGEX: Dict[str, Pattern] = {
        name: re.compile(pattern)
        for name, pattern in _BUILTIN_SYNONYMS_PATTERNS.items()
    }

    # Rich descriptive metadata loaded from src/config/company_synonyms.json (+ _user.json)
    # Structure: {canonical_name: {"pattern": str, "full_name": str, "category": str, "description": str}}
    COMPANY_METADATA: Dict[str, Dict[str, str]] = {}

    # Precompiled fast-matching structures rebuilt by `_compile_fast_matchers()`
    _WORD_TOKEN_RE: Pattern = re.compile(r'\w+')
    _FAST_WORD_INDEX: Dict[str, List[Tuple[int, str]]] = {}
    _FAST_PLAIN_SUB_RULES: List[Tuple[int, str, str]] = []
    _FAST_ANCHORED_REGEX_RULES: List[Tuple[int, str, str, Callable]] = []
    _FAST_UNANCHORED_REGEX_RULES: List[Tuple[int, str, Callable]] = []
    _COMPILED_DICT_SNAPSHOT: Optional[Tuple[Tuple[str, Pattern], ...]] = None

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

    # Precompiled helper regexes used inside hot loops (`tokenize_source_segments`, `get_primary_contributors`, etc.)
    _INLINE_Q_RE = re.compile(r'^(\(\s*\?\s*\)|\?+)')
    _AFTER_Q_DELIM_RE = re.compile(r'^\s*[,;+]')
    _LEADING_WS_RE = re.compile(r'^[ \t]+')
    _PRIMARY_SPLIT_RE = re.compile(r'[,;]|\band\b')
    _COMMA_SEMICOLON_SPLIT_RE = re.compile(r'[,;]+')
    _CONFIRMED_NORM_SPLIT_RE = re.compile(r'[,;\s]+')
    _TENTATIVE_OUTER_STRIP_RE = re.compile(r'^[\[\(\s]+|[\]\)\s\?？]+$')
    _TENTATIVE_INNER_SPLIT_RE = re.compile(r'[,;?？\(\)\[\]]+')

    @classmethod
    def get_config_paths(cls) -> Tuple[Path, Path]:
        """Returns (default_config_path, user_config_path) for externalized company synonym rules."""
        config_dir = Path(__file__).resolve().parent.parent.parent / "config"
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
            # If a pattern was compiled with special flags (other than default), run it directly
            if rx.flags & ~re.UNICODE:
                unanchored_rx.append((idx, name, rx.search))
                continue

            pat = rx.pattern
            branches = cls._split_top_level_branches(pat)
            for b in branches:
                b_s = b.strip()
                # Strip redundant outer parentheses if balanced
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

                # Case 1: Pure word-bounded token `\b<word>\b`
                m_wb = re.fullmatch(r'\\b([a-z0-9_]+)\\b', b_s)
                if m_wb:
                    word_index.setdefault(m_wb.group(1), []).append((idx, name))
                    continue

                # Case 2: Pure literal substring without regex metacharacters
                if all(ch not in b_s for ch in r'.^$*+?{}[]\|()') and b_s == b_s.lower():
                    plain_sub_rules.append((idx, name, b_s))
                    continue

                # Case 3: Regex branch — extract longest guaranteed required literal substring as a fast C `in` pre-filter
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
        return dict(cls.COMPANY_METADATA.get(canonical_name, {
            "pattern": cls._BUILTIN_SYNONYMS_PATTERNS.get(canonical_name, ""),
            "full_name": canonical_name,
            "category": "Contributor",
            "description": "",
        }))

    @classmethod
    def _match_text_to_companies(cls, text_fragment: str) -> list:
        if not text_fragment:
            return []
        clean = cls.SOURCE_REPLACE_REGEX.sub(' ', str(text_fragment)).strip().lower()
        if not clean:
            return []

        # Recompile fast plan automatically if a caller mutated SIGNATURE_SYNONYMS_REGEX directly
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

                # Check for "(?)" or "?" acting as an inline terminator when the author forgot a comma
                # e.g. "Huawei? Samsung? TNO?" or "ETRI?NEC?"
                if ch in '?( ' and ('?' in reg_text[pos:pos + 6]):
                    q_match = cls._INLINE_Q_RE.match(reg_text[pos:])
                    if q_match:
                        q_end = pos + q_match.end()
                        # Look ahead: if followed by whitespace + word or directly by a word (and not a comma/semicolon)
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
                for sub in cls._COMMA_SEMICOLON_SPLIT_RE.split(seg_text):
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

        confirmed_norm = {
            c.lower()
            for c in cls._CONFIRMED_NORM_SPLIT_RE.split(cls.get_confirmed_source_text(original_source))
            if c
        }
        tentative_parts = []
        seen = set()

        for seg_text, is_tent in cls.tokenize_source_segments(original_source):
            if is_tent:
                cleaned = cls._TENTATIVE_OUTER_STRIP_RE.sub('', seg_text)
                for sub in cls._TENTATIVE_INNER_SPLIT_RE.split(cleaned):
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


# Load externalized rules and descriptive metadata at module initialization
CompanySanitizer.reload_rules()
