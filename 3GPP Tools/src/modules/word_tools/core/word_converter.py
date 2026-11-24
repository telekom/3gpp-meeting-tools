"""Backward-compatible re-export shim for Word document conversion.

Canonical implementation lives in ``modules.word_tools.core.converter``.
"""

from __future__ import annotations

import logging

from modules.word_tools.core.converter import (
    WD_FORMAT_DOC,
    WD_FORMAT_DOC_DEFAULT,
    WD_FORMAT_HTML,
    WD_FORMAT_PDF,
    WD_FORMAT_RTF,
    WD_FORMAT_TXT,
    WD_FORMAT_XML_DOCX,
    WD_FORMAT_XPS,
    WordConverterThread,
    _configure_unattended_word,
    _export_pdf_with_word_once,
    _is_rpc_disconnect_error,
    _is_valid_pdf,
    _safe_com_cleanup,
    _sanitize_file_attributes,
    _suppress_faulthandler_for_word_com,
    convert_doc_to_docx,
    convert_doc_to_docx_word,
)

logger = logging.getLogger(__name__)

__all__ = [
    "WD_FORMAT_DOC",
    "WD_FORMAT_DOC_DEFAULT",
    "WD_FORMAT_HTML",
    "WD_FORMAT_PDF",
    "WD_FORMAT_RTF",
    "WD_FORMAT_TXT",
    "WD_FORMAT_XML_DOCX",
    "WD_FORMAT_XPS",
    "WordConverterThread",
    "_configure_unattended_word",
    "_export_pdf_with_word_once",
    "_is_rpc_disconnect_error",
    "_is_valid_pdf",
    "_safe_com_cleanup",
    "_sanitize_file_attributes",
    "_suppress_faulthandler_for_word_com",
    "convert_doc_to_docx",
    "convert_doc_to_docx_word",
    "logger",
]
