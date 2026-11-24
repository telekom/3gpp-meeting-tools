"""Microsoft Purview / MIP sensitivity policy discovery and Word labeling package."""

from __future__ import annotations

from .batch_labeler import batch_set_sensitivity_label
from .com_labeler import set_sensitivity_label
from .inspector import (
    MsoAssignmentMethod,
    _inspect_label_from_com_document,
    _inspect_label_from_docx_file,
    _inspect_label_via_word_com,
    _log,
    get_document_sensitivity_label,
)
from .policy_discovery import (
    _extract_localized_setting,
    _local_tag,
    _parse_clp_syncfile,
    _read_policy_xml,
    discover_windows_sensitivity_policy,
)

__all__ = [
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
