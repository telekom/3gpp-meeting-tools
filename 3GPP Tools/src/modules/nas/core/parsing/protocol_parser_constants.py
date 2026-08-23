import re

# WordprocessingML XML Namespaces & Tags
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TAG_BODY = f"{W_NS}body"
TAG_P = f"{W_NS}p"
TAG_TBL = f"{W_NS}tbl"
TAG_TR = f"{W_NS}tr"
TAG_TC = f"{W_NS}tc"
TAG_T = f"{W_NS}t"
TAG_TAB = f"{W_NS}tab"
TAG_BR = f"{W_NS}br"
TAG_CR = f"{W_NS}cr"
TAG_HYPHEN = f"{W_NS}noBreakHyphen"
TAG_TCPR = f"{W_NS}tcPr"
TAG_GRIDSPAN = f"{W_NS}gridSpan"
TAG_VMERGE = f"{W_NS}vMerge"
TAG_PPR = f"{W_NS}pPr"
TAG_JC = f"{W_NS}jc"

# Path & Specification Number Regular Expressions
RE_PART_INDEX = re.compile(r"_(\d+)_")
RE_SPEC_NUMBER = re.compile(r"(24|25|36|37|38)[._]?(301|501|331|413|423|412|473|463)")
RE_VERSION_STEM = re.compile(r"-([a-zA-Z0-9]{3})(?:_\d+.*)?$")

# Headings & Captions
RE_CLAUSE_HEADER = re.compile(r"^((?:6|9|D\.6)(?:\.[0-9A-Za-z]+)+)\s*(.*)$")
RE_DESC_TABLE = re.compile(r"([A-Za-z0-9\-_]+)\s+field\s+descriptions", re.IGNORECASE)
RE_CAPTION = re.compile(
    r"^Table\s+([8D]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+?)(?:\s+message\s+content)?$",
    re.IGNORECASE,
)
RE_IE_HEADING = re.compile(r"^((?:9\.[2-9]|9\.1[0-9]|D\.6)(?:\.[0-9A-Za-z]+)*)\s+(.+)$")
RE_MAJOR_BOUNDARY = re.compile(r"^(?:[1-8]|10|11|12|Annex\s+[A-Z])\b")

# ASN.1 Declarations & IOS Sets
RE_TYPE_DECL = re.compile(r"([A-Za-z0-9\-]+)(?:\s*\{[^}]*\})?\s*::=\s*", re.MULTILINE)
RE_TYPE_KIND = re.compile(r"^(SEQUENCE|CHOICE|ENUMERATED|BIT STRING|OCTET STRING|INTEGER|BOOLEAN)", re.IGNORECASE)
RE_FIELD_LINE = re.compile(r"^([A-Za-z0-9\-]+)\s+(.+)$", re.DOTALL)
RE_STRIP_KEYWORDS = re.compile(r"\s+(?:OPTIONAL|MANDATORY|DEFAULT\s+[^,\s]+).*", re.IGNORECASE)
RE_COND_NEED = re.compile(r"--\s*(Cond\s+[A-Za-z0-9\-]+|Need\s+[MNRS])", re.IGNORECASE)

# RAN3 Information Object Sets
RE_IE_ID_CONST = re.compile(r"^(id-[A-Za-z0-9\-]+)\s+(?:ProtocolIE-ID|INTEGER)\s*::=\s*(\d+)", re.MULTILINE)
RE_OBJECT_SET_ITEM = re.compile(
    r"\{\s*ID\s+([A-Za-z0-9\-]+)\s+CRITICALITY\s+([A-Za-z0-9\-]+)\s+(?:TYPE|EXTENSION)\s+([A-Za-z0-9\-]+(?:\s*\{[^}]*\})?)\s+PRESENCE\s+([A-Za-z0-9\-]+)\s*\}",
    re.DOTALL
)
RE_CONTAINER_REF = re.compile(
    r"Protocol(?:IE|Extension)-(?:Container|SingleContainer|ContainerList|ContainerPair)\s*\{[^}]*\{\s*([A-Za-z0-9\-]+)\s*\}\s*\}",
    re.IGNORECASE
)
RE_ELEM_PROC_MSG = re.compile(
    r"(?:INITIATING MESSAGE|SUCCESSFUL OUTCOME|UNSUCCESSFUL OUTCOME)\s+([A-Za-z0-9\-]+)",
    re.MULTILINE
)

# RRC Structural Unwrapping Patterns
RE_SETUP_RELEASE = re.compile(r"SetupRelease\s*\{\s*([A-Za-z0-9\-]+)\s*\}", re.IGNORECASE)
RE_SEQ_OF = re.compile(r"SEQUENCE\s*(?:\(SIZE\s*\([^)]*\)\)\s*)?OF\s+([A-Za-z0-9\-]+)", re.IGNORECASE)
RE_OCTET_CONTAINING = re.compile(r"OCTET\s+STRING\s*\(\s*CONTAINING\s+([A-Za-z0-9\-]+)\s*\)", re.IGNORECASE)
RE_CRITICAL_EXT_IES = re.compile(r"\b([A-Za-z0-9\-]+-IEs)\b")
RE_STRIP_EXTRANEOUS = re.compile(r"[\(\{\[].*$")