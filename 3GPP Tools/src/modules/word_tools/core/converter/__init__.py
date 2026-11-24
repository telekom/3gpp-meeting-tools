"""Modular Word document format conversion package."""

from modules.word_tools.core.converter.com_utils import (
    WD_FORMAT_DOC,
    WD_FORMAT_DOC_DEFAULT,
    WD_FORMAT_HTML,
    WD_FORMAT_PDF,
    WD_FORMAT_RTF,
    WD_FORMAT_TXT,
    WD_FORMAT_XML_DOCX,
    WD_FORMAT_XPS,
    _configure_unattended_word,
    _is_rpc_disconnect_error,
    _is_valid_pdf,
    _safe_com_cleanup,
    _sanitize_file_attributes,
    _suppress_faulthandler_for_word_com,
)
from modules.word_tools.core.converter.doc_to_docx import (
    _export_pdf_with_word_once,
    convert_doc_to_docx,
    convert_doc_to_docx_word,
)
from modules.word_tools.core.converter.worker import WordConverterThread

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
]
