"""Backward-compatible re-export shim for 3GPP Word styles and packaging helpers.

Canonical implementation lives in ``modules.word_tools.core.styles``.
"""

from __future__ import annotations

import logging
import time

from modules.word_tools.core.styles import (
    ALL_3GPP_STYLE_NAMES_LOWER,
    LINKED_CHAR_STYLES,
    OFFICIAL_3GPP_DOTX_URL,
    OFFICIAL_CR_TEMPLATE_URL,
    OFFICIAL_PCR_TEMPLATE_ZIP_URL,
    STYLES_3GPP_REGISTRY,
    Style3GPP,
    StyleInfo,
    TABLE_GRID_XML,
    W_NS,
    ensure_3gpp_package_styles,
    get_3gpp_dotx_path,
    get_3gpp_package_part,
    get_installed_3gpp_dotx_path,
    get_pseudo_cr_template_path,
    get_template_cr_path,
    install_3gpp_dotx_template,
    is_3gpp_dotx_installed,
    merge_and_promote_3gpp_styles,
    promote_3gpp_styles_in_word,
    promote_3gpp_styles_in_xml,
    uninstall_3gpp_dotx_template,
)

logger = logging.getLogger(__name__)

__all__ = [
    "ALL_3GPP_STYLE_NAMES_LOWER",
    "LINKED_CHAR_STYLES",
    "OFFICIAL_3GPP_DOTX_URL",
    "OFFICIAL_CR_TEMPLATE_URL",
    "OFFICIAL_PCR_TEMPLATE_ZIP_URL",
    "STYLES_3GPP_REGISTRY",
    "Style3GPP",
    "StyleInfo",
    "TABLE_GRID_XML",
    "W_NS",
    "ensure_3gpp_package_styles",
    "get_3gpp_dotx_path",
    "get_3gpp_package_part",
    "get_installed_3gpp_dotx_path",
    "get_pseudo_cr_template_path",
    "get_template_cr_path",
    "install_3gpp_dotx_template",
    "is_3gpp_dotx_installed",
    "logger",
    "merge_and_promote_3gpp_styles",
    "promote_3gpp_styles_in_word",
    "promote_3gpp_styles_in_xml",
    "time",
    "uninstall_3gpp_dotx_template",
]
