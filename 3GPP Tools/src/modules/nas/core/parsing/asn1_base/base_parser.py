"""Base class for 3GPP ASN.1 DOCX parsing delegating to `scanner` and `grammar`."""

import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from modules.nas.core.parsing.protocol_parser_constants import extract_part_index

from .grammar import (
    detect_asn1_format,
    extract_asn1_type_definitions,
    find_first_structural_brace,
    find_matching_brace,
    parse_sequence_fields,
    split_asn1_tokens,
)
from .scanner import (
    build_inspector_html,
    extract_prose_summary_and_direction,
    fallback_extract_asn1,
    normalize_key,
    scan_xml_documents,
)


class BaseAsn1DocxParser:
    """Base class for 3GPP ASN.1 spec parsing with Clause 9.2/9.3 and Clause 6.2/6.3 extraction."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "38.331"):
        self.docx_paths = sorted(docx_paths, key=self._extract_part_index)
        self.spec_number = spec_number
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(path: Any) -> int:
        return extract_part_index(path)

    @staticmethod
    def _normalize_key(name: str) -> str:
        return normalize_key(name)

    @staticmethod
    def _extract_prose_summary_and_direction(html_str: str) -> Tuple[str, str]:
        return extract_prose_summary_and_direction(html_str)

    def _scan_xml_documents(
        self,
        progress_callback: Optional[Callable[[str, int], None]] = None,
    ) -> Tuple[str, Dict[str, str], Dict[str, Dict[str, str]], Dict[str, str]]:
        return scan_xml_documents(self.docx_paths, progress_callback=progress_callback)

    def _fallback_extract_asn1(self) -> str:
        return fallback_extract_asn1(self.docx_paths)

    def _extract_asn1_type_definitions(self, asn1_text: str) -> Dict[str, Dict[str, Any]]:
        return extract_asn1_type_definitions(asn1_text)

    @staticmethod
    def _find_first_structural_brace(text: str) -> int:
        return find_first_structural_brace(text)

    @staticmethod
    def _find_matching_brace(text: str, start_idx: int) -> int:
        return find_matching_brace(text, start_idx)

    @staticmethod
    def _detect_asn1_format(type_str: str) -> str:
        return detect_asn1_format(type_str)

    def _parse_sequence_fields(self, inner_asn1: str, parent_kind: str) -> List[Dict[str, Any]]:
        return parse_sequence_fields(inner_asn1, parent_kind)

    def _split_asn1_tokens(self, text: str) -> List[str]:
        return split_asn1_tokens(text)

    def _build_inspector_html(
        self,
        type_name: str,
        assigned_clause: str,
        raw_asn1: str,
        desc_table_html: str,
    ) -> str:
        return build_inspector_html(type_name, assigned_clause, raw_asn1, desc_table_html)
