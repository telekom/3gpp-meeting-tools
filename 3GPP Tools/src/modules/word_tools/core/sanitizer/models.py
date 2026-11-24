"""OpenXML namespaces, patterns, language presets, and audit/config dataclasses."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

from modules.word_tools.core.author_renamer import (
    AuthorStats,
    CORE_PROPS_XML,
    CP_NS,
    DCMITYPE_NS,
    DCTERMS_NS,
    DC_NS,
    STANDARD_NAMESPACES,
    W_AUTHOR,
    W_INITIALS,
    W_NS,
    XSI_NS,
    close_word_document_for_update,
    is_file_open_in_word,
)

# Register standard namespaces
for _prefix, _uri in STANDARD_NAMESPACES.items():
    try:
        ET.register_namespace(_prefix, _uri)
    except Exception:
        pass

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
ET.register_namespace("a", A_NS)

W_VAL = f"{{{W_NS}}}val"
W_EAST_ASIA = f"{{{W_NS}}}eastAsia"
W_BIDI = f"{{{W_NS}}}bidi"
W_LANG = f"{{{W_NS}}}lang"
W_NOPROOF = f"{{{W_NS}}}noProof"

STORY_XML_PATTERN = re.compile(
    r"^word/(document|comments|commentsExtended|commentsIds|header\d+|footer\d+|footnotes|endnotes|glossary/document)\.xml$",
    re.IGNORECASE,
)
STYLES_XML_PATTERN = re.compile(r"^word/styles\.xml$", re.IGNORECASE)

# Standard language presets common in 3GPP and international standardization
LANGUAGE_PRESETS = [
    ("en-GB", "English (United Kingdom) [en-GB] — 3GPP Standard"),
    ("en-US", "English (United States) [en-US]"),
    ("de-DE", "German (Germany) [de-DE]"),
    ("fr-FR", "French (France) [fr-FR]"),
    ("zh-CN", "Chinese (Simplified) [zh-CN]"),
    ("ja-JP", "Japanese [ja-JP]"),
    ("ko-KR", "Korean [ko-KR]"),
    ("es-ES", "Spanish (Spain) [es-ES]"),
    ("sv-SE", "Swedish [sv-SE]"),
    ("fi-FI", "Finnish [fi-FI]"),
]


@dataclass
class LanguageStats:
    """Statistics on proofing languages detected in a document."""

    language_counts: Dict[str, int] = field(default_factory=dict)
    no_proof_count: int = 0
    language_change_marks: int = 0
    pure_language_change_marks: int = 0
    compound_language_change_marks: int = 0
    default_language: str = ""

    @property
    def primary_language(self) -> str:
        if self.default_language:
            return self.default_language
        if not self.language_counts:
            return "Unknown"
        return max(self.language_counts.items(), key=lambda kv: kv[1])[0]

    @property
    def summary_text(self) -> str:
        if not self.language_counts:
            return self.default_language or "No language tags detected"
        total = sum(self.language_counts.values())
        parts = []
        for lang, count in sorted(
            self.language_counts.items(), key=lambda kv: kv[1], reverse=True
        )[:3]:
            pct = int(round((count / total) * 100))
            parts.append(f"{lang} ({pct}%)")
        return ", ".join(parts)


@dataclass
class DocumentHygieneAudit:
    """Complete pre-flight health and revision audit of a Word document."""

    file_path: Path
    file_name: str
    file_size_kb: float
    authors: List[AuthorStats] = field(default_factory=list)
    total_insertions: int = 0
    total_deletions: int = 0
    total_formatting: int = 0
    total_comments: int = 0
    language_stats: LanguageStats = field(default_factory=LanguageStats)
    doc_creator: str = ""
    doc_last_modified_by: str = ""

    @property
    def total_authors(self) -> int:
        return len(self.authors)

    @property
    def total_changes(self) -> int:
        return (
            self.total_insertions
            + self.total_deletions
            + self.total_formatting
            + self.total_comments
        )


@dataclass
class SanitizeConfig:
    """User-selectable options for document sanitization."""

    # Language Unification & Language Revision Purge
    unify_language: bool = True
    target_language: str = "en-GB"
    apply_to_all_stories: bool = True
    update_styles_and_defaults: bool = True
    purge_language_changes: bool = True
    clear_no_proof: bool = False

    # Revision Authors & Metadata
    update_authors: bool = False
    author_mapping: Dict[str, str] = field(default_factory=dict)
    new_initials: Optional[str] = None
    update_metadata: bool = True

    # Formatting & Comments Cleanup
    purge_all_formatting_changes: bool = False
    remove_all_comments: bool = False

    # Safety
    backup: bool = True


__all__ = [
    "AuthorStats",
    "W_AUTHOR",
    "W_INITIALS",
    "W_NS",
    "A_NS",
    "CP_NS",
    "DC_NS",
    "DCTERMS_NS",
    "DCMITYPE_NS",
    "XSI_NS",
    "CORE_PROPS_XML",
    "W_VAL",
    "W_EAST_ASIA",
    "W_BIDI",
    "W_LANG",
    "W_NOPROOF",
    "STORY_XML_PATTERN",
    "STYLES_XML_PATTERN",
    "LANGUAGE_PRESETS",
    "LanguageStats",
    "DocumentHygieneAudit",
    "SanitizeConfig",
    "close_word_document_for_update",
    "is_file_open_in_word",
]
