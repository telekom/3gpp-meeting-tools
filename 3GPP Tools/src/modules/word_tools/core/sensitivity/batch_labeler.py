"""Batch sensitivity labeling across folders or explicit Word file lists."""

from __future__ import annotations

import gc
import logging
import os
from pathlib import Path
from typing import Callable, List, Optional, Tuple, Union

from modules.word_tools.core.word_config import WordConfig
from .com_labeler import set_sensitivity_label
from .inspector import _inspect_label_from_docx_file

logger = logging.getLogger(__name__)


def batch_set_sensitivity_label(
    folder_path: Optional[Union[str, Path]] = None,
    file_paths: Optional[List[Union[str, Path]]] = None,
    label_id: Optional[str] = None,
    label_name: Optional[str] = None,
    overwrite_existing: bool = False,
    recursive: bool = False,
    progress_callback: Optional[Callable[[int, int, str, str], None]] = None,
) -> dict:
    """
    Batch applies a sensitivity label across a set of Word files or all .doc and .docx files in a directory.

    Args:
        folder_path: Root directory containing Word documents (optional if file_paths given).
        file_paths: Explicit list of file paths to process (optional if folder_path given).
        label_id: Sensitivity Label GUID (or None for active config).
        label_name: Sensitivity Label Name (or None for active config).
        overwrite_existing: If True, re-labels already classified documents.
        recursive: If True, processes subdirectories recursively (when folder_path is used).
        progress_callback: Optional callback(current_idx, total_count, file_name, status_message).

    Returns:
        Summary dict with keys: 'total', 'labeled', 'skipped', 'errors', 'error_files'.
    """
    if file_paths is not None:
        all_files = [
            Path(p).resolve()
            for p in file_paths
            if Path(p).is_file()
            and Path(p).suffix.lower() in {".doc", ".docx"}
            and not Path(p).name.startswith("~$")
        ]
    elif folder_path is not None:
        root = Path(folder_path).resolve()
        if not root.exists() or not root.is_dir():
            raise FileNotFoundError(f"Directory not found: {folder_path}")
        # Discover candidate files
        pattern = "**/*" if recursive else "*"
        all_files = [
            p
            for p in root.glob(pattern)
            if p.is_file()
            and p.suffix.lower() in {".doc", ".docx"}
            and not p.name.startswith("~$")
        ]
    else:
        raise ValueError("Either folder_path or file_paths must be specified.")

    total = len(all_files)
    labeled = 0
    skipped = 0
    errors = 0
    error_files: List[Tuple[str, str]] = []

    if total == 0:
        return {"total": 0, "labeled": 0, "skipped": 0, "errors": 0, "error_files": []}

    cfg = WordConfig.load()
    if not cfg.get("enable_sensitivity_labels", False):
        return {"total": total, "labeled": 0, "skipped": total, "errors": 0, "error_files": []}

    if not label_id or not label_name:
        active_label = WordConfig.get_active_sensitivity_label()
        label_id = label_id or active_label.get("id") or cfg.get("sensitivity_level_label_id")
        label_name = label_name or active_label.get("name") or cfg.get("sensitivity_level_label_name")

    if not label_id or not label_name:
        return {"total": total, "labeled": 0, "skipped": total, "errors": 0, "error_files": []}

    # If overwrite_existing is False, quickly filter out .docx that already have a label headlessly
    candidates: List[Path] = []
    for p in all_files:
        if not overwrite_existing and p.suffix.lower() == ".docx":
            existing = _inspect_label_from_docx_file(p)
            if existing and existing.get("id"):
                skipped += 1
                if progress_callback:
                    progress_callback(
                        skipped,
                        total,
                        p.name,
                        f"Skipped (Already labeled: {existing.get('name')})",
                    )
                continue
        candidates.append(p)

    if not candidates:
        return {"total": total, "labeled": 0, "skipped": skipped, "errors": 0, "error_files": []}

    # Process remaining files via Word COM
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        for idx, doc_path in enumerate(candidates, start=skipped + 1):
            if progress_callback:
                progress_callback(idx, total, doc_path.name, f"Processing with {label_name}...")

            doc = None
            try:
                doc = word.Documents.Open(
                    FileName=os.path.normpath(str(doc_path)),
                    ReadOnly=False,
                    ConfirmConversions=False,
                    AddToRecentFiles=False,
                )
                changed = set_sensitivity_label(
                    doc,
                    label_id=label_id,
                    label_name=label_name,
                    return_changed=True,
                    force_overwrite=overwrite_existing,
                )
                if changed:
                    doc.Save()
                    labeled += 1
                else:
                    skipped += 1
            except Exception as e:
                errors += 1
                error_files.append((doc_path.name, str(e)))
                logger.warning(f"Error labeling {doc_path.name}: {e}")
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass
    finally:
        word = None
        gc.collect()
        pythoncom.CoUninitialize()

    return {
        "total": total,
        "labeled": labeled,
        "skipped": skipped,
        "errors": errors,
        "error_files": error_files,
    }
