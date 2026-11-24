# --- File: modules/word_tools/core/sensitivity_label.py ---
"""Backward-compatible re-export shim for ``modules.word_tools.core.sensitivity``."""

from __future__ import annotations

import logging
import platform

from .sensitivity import (
    MsoAssignmentMethod,
    _extract_localized_setting,
    _inspect_label_from_com_document,
    _inspect_label_from_docx_file,
    _inspect_label_via_word_com,
    _local_tag,
    _log,
    _parse_clp_syncfile,
    _read_policy_xml,
    batch_set_sensitivity_label,
    discover_windows_sensitivity_policy,
    get_document_sensitivity_label,
    set_sensitivity_label,
)

logger = logging.getLogger(__name__)

__all__ = [
    "logger",
    "platform",
    "_local_tag",
    "_read_policy_xml",
    "_extract_localized_setting",
    "_parse_clp_syncfile",
    "discover_windows_sensitivity_policy",
    "_log",
    "MsoAssignmentMethod",
    "get_document_sensitivity_label",
    "_inspect_label_from_docx_file",
    "_inspect_label_from_com_document",
    "_inspect_label_via_word_com",
    "set_sensitivity_label",
    "batch_set_sensitivity_label",
]
