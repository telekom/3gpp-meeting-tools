"""Synchronous .doc -> .docx conversion and isolated Word PDF export helpers."""

from __future__ import annotations

import logging
import os
import shutil
import tempfile
import time
from pathlib import Path
from typing import Optional, Union

import pythoncom
import win32com.client

from core.utils.office_printer_guard import temporary_virtual_default_printer
from core.utils.paths import to_extended_path
from modules.word_tools.core.converter.com_utils import (
    WD_FORMAT_DOC_DEFAULT,
    WD_FORMAT_PDF,
    WD_FORMAT_XML_DOCX,
    _configure_unattended_word,
    _is_rpc_disconnect_error,
    _is_valid_pdf,
    _safe_com_cleanup,
    _sanitize_file_attributes,
    _suppress_faulthandler_for_word_com,
)
from modules.word_tools.core.libreoffice_converter import convert_doc_to_docx_libreoffice
from modules.word_tools.core.sensitivity_label import set_sensitivity_label

logger = logging.getLogger(__name__)


def _export_pdf_with_word_once(source: Path, target: Path) -> None:
    """
    Perform one isolated Word PDF export.

    Used by the media-repair diagnostic path.  Each test gets a fresh Word COM
    instance so a failed PDF renderer cannot contaminate the next candidate.
    """
    word = None
    doc = None
    com_server_dead = False
    try:
        pythoncom.CoInitialize()
        if target.exists():
            try:
                target.unlink()
            except OSError:
                pass

        with temporary_virtual_default_printer(log=logger):
            word = win32com.client.DispatchEx("Word.Application")
            _configure_unattended_word(word, log=logger)

            doc = word.Documents.Open(
                FileName=os.path.normpath(str(source)),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument="__3gpp_no_pw__",
                WritePasswordDocument="__3gpp_no_pw__",
                NoEncodingDialog=True,
            )

            try:
                set_sensitivity_label(doc)
            except Exception:
                pass

            with _suppress_faulthandler_for_word_com():
                doc.ExportAsFixedFormat(
                    OutputFileName=os.path.normpath(str(target)),
                    ExportFormat=WD_FORMAT_PDF,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    IncludeDocProps=True,
                    CreateBookmarks=1,
                )

            try:
                while word.BackgroundPrintingStatus > 0:
                    time.sleep(0.05)
            except Exception:
                pass

        if not _is_valid_pdf(target):
            raise RuntimeError("Word returned without producing a valid PDF.")
    except Exception as exc:
        com_server_dead = _is_rpc_disconnect_error(exc)
        raise
    finally:
        _safe_com_cleanup(word=word, doc=doc, com_server_dead=com_server_dead)


def convert_doc_to_docx_word(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """Synchronously converts a legacy .doc to .docx using Microsoft Word COM automation."""
    log = logger or logging.getLogger(__name__)
    source = Path(doc_path).resolve()

    if not source.exists():
        raise FileNotFoundError(f"Source document not found: {source}")

    if source.suffix.lower() == ".docx":
        return source

    target = Path(output_path).resolve() if output_path else source.with_suffix(".docx")
    if target.exists() and target.stat().st_size > 0:
        return target

    _sanitize_file_attributes(source)
    os.makedirs(to_extended_path(target.parent), exist_ok=True)
    if target.exists():
        try:
            _sanitize_file_attributes(target)
            target.unlink()
        except Exception:
            pass

    temp_dir = Path(tempfile.gettempdir())
    temp_target = temp_dir / f"stage_{source.stem}.docx"
    if temp_target.exists():
        try:
            temp_target.unlink()
        except Exception:
            pass

    source_str = os.path.normpath(str(source))
    temp_target_str = os.path.normpath(str(temp_target))

    word = None
    doc = None
    new_doc = None
    saved = False
    com_server_dead = False

    try:
        pythoncom.CoInitialize()

        with temporary_virtual_default_printer(log=log):
            word = win32com.client.DispatchEx("Word.Application")
            _configure_unattended_word(word, log=log)

            try:
                doc = word.Documents.Open(
                    FileName=source_str,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    PasswordDocument="__3gpp_no_pw__",
                    WritePasswordDocument="__3gpp_no_pw__",
                    NoEncodingDialog=True,
                )
            except Exception:
                try:
                    if hasattr(word, "ProtectedViewWindows") and word.ProtectedViewWindows.Count > 0:
                        pv = word.ProtectedViewWindows.Item(1)
                        doc = pv.Edit()
                    else:
                        doc = word.Documents.Open(
                            FileName=source_str,
                            ConfirmConversions=False,
                            ReadOnly=True,
                            AddToRecentFiles=False,
                        )
                except Exception as open_exc:
                    if _is_rpc_disconnect_error(open_exc):
                        com_server_dead = True
                    raise RuntimeError("MS Word failed to open source document.") from open_exc

            if doc is None:
                raise RuntimeError(f"Word failed to acquire a valid document handle for {source.name}")

            try:
                doc.Convert()
            except Exception:
                pass

            try:
                set_sensitivity_label(doc)
            except Exception:
                pass

            try:
                doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_XML_DOCX)
                if temp_target.exists() and temp_target.stat().st_size > 0:
                    saved = True
            except Exception:
                pass

            if not saved:
                try:
                    doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_DOC_DEFAULT)
                    if temp_target.exists() and temp_target.stat().st_size > 0:
                        saved = True
                except Exception:
                    pass

            if not saved or not temp_target.exists() or temp_target.stat().st_size == 0:
                new_doc = word.Documents.Add()
                new_doc.Content.FormattedText = doc.Content.FormattedText
                new_doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_XML_DOCX)
                new_doc.Close(SaveChanges=False)
                new_doc = None
                if temp_target.exists() and temp_target.stat().st_size > 0:
                    saved = True

        if not temp_target.exists() or temp_target.stat().st_size == 0:
            raise RuntimeError(f"MS Word failed to write target .docx for {source.name}")

        shutil.copy2(to_extended_path(temp_target), to_extended_path(target))
        _sanitize_file_attributes(target)
        try:
            temp_target.unlink()
        except Exception:
            pass

        log.info(f"Successfully converted via MS Word: {source.name} -> {target.name}")
        return target

    except Exception as exc:
        if _is_rpc_disconnect_error(exc):
            com_server_dead = True
        raise
    finally:
        _safe_com_cleanup(word=word, doc=doc, new_doc=new_doc, com_server_dead=com_server_dead)
        word = None
        doc = None
        new_doc = None


def convert_doc_to_docx(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """Primary Entry Point: Attempts Word COM conversion first, falling back to LibreOffice."""
    log = logger or logging.getLogger(__name__)
    source = Path(doc_path).resolve()

    try:
        return convert_doc_to_docx_word(source, output_path=output_path, logger=log)
    except Exception as word_err:
        log.warning(f"Word COM conversion failed for '{source.name}' ({word_err}). Falling back to LibreOffice...")
        try:
            return convert_doc_to_docx_libreoffice(source, output_path=output_path, logger=log)
        except Exception as lo_err:
            log.error(f"Both Word and LibreOffice conversions failed for '{source.name}'. LibreOffice error: {lo_err}")
            raise
