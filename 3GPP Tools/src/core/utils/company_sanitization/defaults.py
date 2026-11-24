# --- File: src/core/utils/company_sanitization/defaults.py ---
"""Built-in baseline company synonym regex catalog and precompiled helper regexes."""

import re
from typing import Dict, Pattern

# Built-in zero-I/O baseline catalog of canonical 3GPP contributors and regex rules.
# Externalized descriptions, metadata, and optional user overrides are loaded from:
#   - src/config/company_synonyms.json (tracked baseline with full_name, category, description)
#   - src/config/company_synonyms_user.json (optional local user overrides/additions, preserved across updates)
BUILTIN_SYNONYMS_PATTERNS: Dict[str, str] = {
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

WORD_TOKEN_RE: Pattern = re.compile(r'\w+')

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

INLINE_Q_RE = re.compile(r'^(\(\s*\?\s*\)|\?+)')
AFTER_Q_DELIM_RE = re.compile(r'^\s*[,;+]')
LEADING_WS_RE = re.compile(r'^[ \t]+')
PRIMARY_SPLIT_RE = re.compile(r'[,;]|\band\b')
COMMA_SEMICOLON_SPLIT_RE = re.compile(r'[,;]+')
CONFIRMED_NORM_SPLIT_RE = re.compile(r'[,;\s]+')
TENTATIVE_OUTER_STRIP_RE = re.compile(r'^[\[\(\s]+|[\]\)\s\?？]+$')
TENTATIVE_INNER_SPLIT_RE = re.compile(r'[,;?？\(\)\[\]]+')
