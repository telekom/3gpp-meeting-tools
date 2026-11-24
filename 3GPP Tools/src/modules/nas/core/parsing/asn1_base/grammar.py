"""ASN.1 brace matching, token splitting, field parsing, and type definition extraction."""

import re
from typing import Any, Dict, List

from modules.nas.core.parsing.protocol_parser_constants import (
    RE_COND_NEED,
    RE_FIELD_LINE,
    RE_STRIP_KEYWORDS,
    RE_TYPE_DECL,
    RE_TYPE_KIND,
)


def find_first_structural_brace(text: str) -> int:
    """Locate the index of the first `{` outside ASN.1 `--` comments."""
    in_comment = False
    i = 0
    n = len(text)
    while i < n:
        c = text[i]
        if not in_comment and c == "-" and i + 1 < n and text[i + 1] == "-":
            in_comment = True
            i += 2
            continue
        elif in_comment:
            if c == "\n":
                in_comment = False
            elif c == "-" and i + 1 < n and text[i + 1] == "-":
                in_comment = False
                i += 2
                continue
            i += 1
            continue
        if c == "{":
            return i
        i += 1
    return -1


def find_matching_brace(text: str, start_idx: int) -> int:
    """Locate the matching `}` for the `{` at `start_idx`, ignoring ASN.1 `--` comments."""
    depth = 0
    in_comment = False
    i = start_idx
    n = len(text)
    while i < n:
        c = text[i]
        if not in_comment and c == "-" and i + 1 < n and text[i + 1] == "-":
            in_comment = True
            i += 2
            continue
        elif in_comment:
            if c == "\n":
                in_comment = False
            elif c == "-" and i + 1 < n and text[i + 1] == "-":
                in_comment = False
                i += 2
                continue
            i += 1
            continue

        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def detect_asn1_format(type_str: str) -> str:
    """Detect intrinsic scalar or container ASN.1 format from a type declaration."""
    s = type_str.strip()
    if re.match(r"^INTEGER\b", s, re.IGNORECASE):
        return "INTEGER"
    if re.match(r"^ENUMERATED\b", s, re.IGNORECASE):
        return "ENUMERATED"
    if re.match(r"^BOOLEAN\b", s, re.IGNORECASE):
        return "BOOLEAN"
    if re.match(r"^BIT\s+STRING\b", s, re.IGNORECASE):
        return "BIT STRING"
    if re.match(r"^OCTET\s+STRING\b", s, re.IGNORECASE):
        return "OCTET STRING"
    if re.match(r"^CHOICE\b", s, re.IGNORECASE):
        return "CHOICE"
    if re.match(r"^SEQUENCE\s+(?:\(SIZE\s*\([^)]*\)\)\s*)?OF\b", s, re.IGNORECASE):
        return "SEQUENCE OF"
    if re.match(r"^SEQUENCE\b", s, re.IGNORECASE):
        return "SEQUENCE"
    return ""


def split_asn1_tokens(text: str) -> List[str]:
    """Split comma-separated ASN.1 fields at depth 0, preserving trailing single-line comments."""
    tokens: List[str] = []
    current: List[str] = []
    depth = p_depth = b_depth = 0
    in_comment = False
    i = 0
    n = len(text)

    while i < n:
        c = text[i]

        if not in_comment and c == "-" and i + 1 < n and text[i + 1] == "-":
            in_comment = True
            current.append(c)
            current.append(text[i + 1])
            i += 2
            continue
        elif in_comment:
            current.append(c)
            if c == "\n":
                in_comment = False
            elif c == "-" and i + 1 < n and text[i + 1] == "-":
                current.append(text[i + 1])
                i += 2
                in_comment = False
                continue
            i += 1
            continue

        if c == "{":
            depth += 1
        elif c == "}":
            depth = max(0, depth - 1)
        elif c == "(":
            p_depth += 1
        elif c == ")":
            p_depth = max(0, p_depth - 1)
        elif c == "[":
            b_depth += 1
        elif c == "]":
            b_depth = max(0, b_depth - 1)

        if c == "," and depth == 0 and p_depth == 0 and b_depth == 0:
            current.append(c)
            j = i + 1
            while j < n and text[j] in (" ", "\t"):
                current.append(text[j])
                j += 1
            if j + 1 < n and text[j] == "-" and text[j + 1] == "-":
                while j < n and text[j] != "\n":
                    current.append(text[j])
                    j += 1
            tokens.append("".join(current).strip())
            current = []
            i = j
            continue

        current.append(c)
        i += 1

    if current:
        tail = "".join(current).strip()
        if tail:
            tokens.append(tail)

    return tokens


def _resolve_presence(fragment: str, parent_kind: str) -> str:
    need_match = RE_COND_NEED.search(fragment)
    if need_match:
        code = need_match.group(1).strip()
        return f"C ({code})" if "Cond" in code else f"O ({code})"
    if "OPTIONAL" in fragment.upper():
        return "O"
    return "M" if parent_kind == "SEQUENCE" else "O"


def parse_sequence_fields(inner_asn1: str, parent_kind: str) -> List[Dict[str, Any]]:
    """Extract field records while preserving inline SEQUENCE/CHOICE bodies and comments."""
    fields: List[Dict[str, Any]] = []
    cleaned = re.sub(r"\[\[|\]\]", " ", inner_asn1)
    tokens = split_asn1_tokens(cleaned)

    for tok in tokens:
        tok = tok.strip()
        if not tok or tok == "..." or tok.startswith("--"):
            continue

        field_match = RE_FIELD_LINE.match(tok)
        if not field_match:
            continue

        fname = field_match.group(1).strip()
        rest = field_match.group(2).strip()

        if rest.startswith(("SEQUENCE", "CHOICE")) and "{" in rest:
            brace_start = find_first_structural_brace(rest)
            brace_end = find_matching_brace(rest, brace_start) if brace_start >= 0 else -1

            if brace_start >= 0 and brace_end > brace_start:
                raw_type_body = rest[:brace_end + 1].strip()
                tail = rest[brace_end + 1:].strip()
                inner_body = rest[brace_start + 1:brace_end]

                presence = _resolve_presence(tail, parent_kind)
                child_kind = "SEQUENCE" if rest.startswith("SEQUENCE") else "CHOICE"
                nested_fields = parse_sequence_fields(inner_body, child_kind)

                fields.append({
                    "name": fname,
                    "type": raw_type_body,
                    "presence": presence,
                    "format": child_kind,
                    "raw_field": tok,
                    "nested_fields": nested_fields,
                    "iei": "",
                })
                continue

        presence = _resolve_presence(rest, parent_kind)
        clean_type_candidate = RE_STRIP_KEYWORDS.sub("", rest.split("--")[0]).strip().rstrip(",").strip()
        detected_format = detect_asn1_format(clean_type_candidate) or detect_asn1_format(rest) or parent_kind

        fields.append({
            "name": fname,
            "type": clean_type_candidate,
            "presence": presence,
            "format": detected_format,
            "raw_field": tok,
            "nested_fields": [],
            "iei": "",
        })

    return fields


def extract_asn1_type_definitions(asn1_text: str) -> Dict[str, Dict[str, Any]]:
    """Parse raw ASN.1 text into structured type definitions retaining full comments."""
    type_defs: Dict[str, Dict[str, Any]] = {}
    matches = list(RE_TYPE_DECL.finditer(asn1_text))

    for i, match in enumerate(matches):
        line_start = asn1_text.rfind("\n", 0, match.start())
        line_start = 0 if line_start == -1 else line_start + 1
        prefix_line = asn1_text[line_start:match.start()]
        if "--" in prefix_line or prefix_line.strip():
            continue

        type_name = match.group(1).strip()
        start_pos = match.end()
        if i + 1 < len(matches):
            next_line_start = asn1_text.rfind("\n", 0, matches[i + 1].start())
            end_pos = (
                next_line_start
                if next_line_start != -1 and next_line_start >= start_pos
                else matches[i + 1].start()
            )
        else:
            end_pos = len(asn1_text)

        raw_block = asn1_text[match.start():end_pos].strip()
        for kw in ("\nEND", "\nDEFINITIONS"):
            kw_idx = raw_block.find(kw)
            if kw_idx != -1:
                raw_block = raw_block[:kw_idx].strip()

        raw_lines = raw_block.splitlines()
        while raw_lines and raw_lines[-1].strip().startswith("--"):
            raw_lines.pop()
        raw_block = "\n".join(raw_lines).strip()

        def_body = asn1_text[start_pos:end_pos].strip()
        def_lines = def_body.splitlines()
        while def_lines and def_lines[-1].strip().startswith("--"):
            def_lines.pop()
        def_body = "\n".join(def_lines).strip()

        kind_match = RE_TYPE_KIND.match(def_body)
        type_kind = kind_match.group(1).upper() if kind_match else "TYPE"
        fields: List[Dict[str, Any]] = []

        brace_start = find_first_structural_brace(def_body)
        if brace_start >= 0:
            brace_end = find_matching_brace(def_body, brace_start)
            if brace_end > brace_start:
                inner_content = def_body[brace_start + 1:brace_end]
                fields = parse_sequence_fields(inner_content, type_kind)

        type_defs[type_name] = {
            "kind": type_kind,
            "fields": fields,
            "raw_asn1": raw_block,
            "clause": type_name,
        }

    return type_defs
