# --- File: modules/word_tools/core/sensitivity_label.py ---
import gc
import logging
import os
import platform
import traceback
import uuid
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple, Union

from modules.word_tools.core.word_config import WordConfig

logger = logging.getLogger(__name__)


def _log(message: str, level: int = logging.INFO, ui_logger=None) -> None:
    """Log through Python logging; optionally support legacy signal/logger callers."""
    logger.log(level, message)


class MsoAssignmentMethod(Enum):
    NOT_SET = -1  # The assignment method value is not set.
    STANDARD = 0  # The label is applied by default.
    PRIVILEGED = 1  # The label was manually selected.
    AUTO = 2  # The label is applied automatically.

    @staticmethod
    def from_int(value_int: int):
        try:
            return MsoAssignmentMethod(value_int)
        except Exception as e:
            logger.warning(f"Value does not exist in Enum: {e}")
            return None


def get_document_sensitivity_label(document_or_path: Union[Any, str, Path]) -> Optional[dict]:
    """
    Inspect the applied Microsoft Information Protection (MIP) / Purview sensitivity label.

    Supports:
      1. File path (str or Path) to a .docx or .docm document (fast headless OpenXML inspection).
      2. File path to legacy binary .doc document (via background Word COM inspection on Windows).
      3. Open Word COM Document object.

    Returns:
      dict with keys: 'id', 'name', 'method', 'set_date', 'site_id', 'action_id', 'enabled',
      or None if no label is applied or file is unclassified.
    """
    if isinstance(document_or_path, (str, Path)):
        p = Path(document_or_path)
        if not p.exists():
            return None
        suffix = p.suffix.lower()
        if suffix in {".docx", ".docm"}:
            return _inspect_label_from_docx_file(p)
        elif suffix == ".doc":
            # Try headless zip in case it's actually an OpenXML file named .doc
            res = _inspect_label_from_docx_file(p)
            if res:
                return res
            if platform.system() == "Windows":
                return _inspect_label_via_word_com(p)
        return None

    # COM Document object
    return _inspect_label_from_com_document(document_or_path)


def _inspect_label_from_docx_file(path: Path) -> Optional[dict]:
    """Inspect sensitivity label properties directly from docProps/custom.xml in a .docx/.docm zip package."""
    try:
        with zipfile.ZipFile(path, "r") as z:
            if "docProps/custom.xml" not in z.namelist():
                return None
            xml_bytes = z.read("docProps/custom.xml")
            root = ET.fromstring(xml_bytes.decode("utf-8", errors="replace"))
            props: Dict[str, str] = {}
            guid: Optional[str] = None
            for prop in root:
                prop_name = prop.attrib.get("name", "")
                val = "".join(prop.itertext()).strip()
                if prop_name.startswith("MSIP_Label_"):
                    parts = prop_name.split("_")
                    if len(parts) >= 4:
                        guid = parts[2]
                        field = "_".join(parts[3:])
                        props[field] = val

            if guid:
                return {
                    "id": guid,
                    "name": props.get("Name", ""),
                    "method": props.get("Method", ""),
                    "set_date": props.get("SetDate", ""),
                    "site_id": props.get("SiteId", ""),
                    "action_id": props.get("ActionId", ""),
                    "content_bits": props.get("ContentBits", "0"),
                    "enabled": props.get("Enabled", "").lower() == "true",
                }
    except Exception as e:
        logger.debug(f"Headless sensitivity label inspection failed for {path.name}: {e}")
    return None


def _inspect_label_from_com_document(doc: Any) -> Optional[dict]:
    """Inspect sensitivity label directly from an active Word COM Document object."""
    try:
        if getattr(doc, "SensitivityLabel", None) is None:
            return None
        sl = doc.SensitivityLabel.GetLabel()
        if sl is None:
            return None
        assignment_method = getattr(sl, "AssignmentMethod", -1)
        if assignment_method == -1:
            return None
        method_enum = MsoAssignmentMethod.from_int(assignment_method)
        method_str = method_enum.name if method_enum else str(assignment_method)
        return {
            "id": getattr(sl, "LabelId", ""),
            "name": getattr(sl, "LabelName", ""),
            "method": method_str,
            "assignment_method": assignment_method,
            "set_date": getattr(sl, "SetDate", ""),
            "site_id": getattr(sl, "SiteId", ""),
            "action_id": getattr(sl, "ActionId", ""),
            "enabled": True,
        }
    except Exception as e:
        logger.debug(f"COM SensitivityLabel inspection error: {e}")
        return None


def _inspect_label_via_word_com(path: Path) -> Optional[dict]:
    """Open a legacy .doc file headlessly in Word COM to read its sensitivity label."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(
            FileName=os.path.normpath(str(path)),
            ReadOnly=True,
            ConfirmConversions=False,
            AddToRecentFiles=False,
        )
        return _inspect_label_from_com_document(doc)
    except Exception as e:
        logger.debug(f"Word COM inspection failed for {path.name}: {e}")
        return None
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        doc = word = None
        gc.collect()
        pythoncom.CoUninitialize()


def set_sensitivity_label(
    document: Any,
    label_id: Optional[str] = None,
    label_name: Optional[str] = None,
    ui_logger=None,
    return_changed: bool = False,
    force_overwrite: bool = False,
):
    """
    Applies a corporate sensitivity label to a Word COM Document object to bypass manual save popups.

    Args:
        document: Microsoft Word COM Document instance.
        label_id: Optional GUID of the label. If None, loaded from active WordConfig.
        label_name: Optional display name of the label. If None, loaded from active WordConfig.
        ui_logger: Optional UI logging target.
        return_changed: If True, returns boolean indicating whether the label was modified.
        force_overwrite: If True, overwrites an existing label even if one was already set.
    """
    if platform.system() != "Windows":
        return False if return_changed else document

    config = WordConfig.load()

    # The Gatekeeper check
    if not config.get("enable_sensitivity_labels", False):
        _log("   ➔ Sensitivity labels disabled in config. Skipping...", logging.DEBUG, ui_logger)
        return False if return_changed else document

    # Resolve active label if not explicitly provided
    if not label_id or not label_name:
        active_label = WordConfig.get_active_sensitivity_label()
        label_id = label_id or active_label.get("id") or config.get("sensitivity_level_label_id")
        label_name = label_name or active_label.get("name") or config.get("sensitivity_level_label_name")

    if not label_id or not label_name:
        _log("⚠️ Sensitivity labels enabled, but ID or Name is missing in config!", logging.WARNING, ui_logger)
        return False if return_changed else document

    did_something = False
    try:
        # Check if the document actually supports sensitivity labels
        if getattr(document, "SensitivityLabel", None) is None:
            return False if return_changed else document

        sensitivity_label = document.SensitivityLabel.GetLabel()
        assignment_method = MsoAssignmentMethod.from_int(sensitivity_label.AssignmentMethod)

        needs_setting = (assignment_method == MsoAssignmentMethod.NOT_SET)
        if not needs_setting and force_overwrite:
            current_id = getattr(sensitivity_label, "LabelId", "")
            if current_id != label_id:
                needs_setting = True

        if needs_setting:
            _log(f"   ➔ Applying corporate Sensitivity Label: {label_name} ({label_id})...", logging.INFO, ui_logger)

            new_sl = document.SensitivityLabel.CreateLabelInfo()
            new_sl.ActionId = str(uuid.uuid1())
            new_sl.AssignmentMethod = 1  # Privileged
            new_sl.ContentBits = 0
            new_sl.IsEnabled = True
            new_sl.Justification = ""
            new_sl.LabelId = label_id
            new_sl.LabelName = label_name
            # Dynamically generate current UTC time
            new_sl.SetDate = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            new_sl.SiteId = str(uuid.uuid1())

            document.SensitivityLabel.SetLabel(new_sl, new_sl)
            did_something = True

            if config.get("save_document_after_setting_sensitivity_label"):
                document.Save()
        else:
            _log(f"   ➔ Sensitivity label already set ({getattr(sensitivity_label, 'LabelName', 'Classified')}). Skipping...", logging.DEBUG, ui_logger)

    except Exception as e:
        _log(f"⚠️ Could not set sensitivity label (Feature may be disabled): {e}", logging.WARNING, ui_logger)
        if did_something:
            logger.exception("Sensitivity label operation failed after modifying the label.")

    return did_something if return_changed else document


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
            Path(p).resolve() for p in file_paths
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
            p for p in root.glob(pattern)
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

    # If overwrite_existing is False, quickly filter out .docx that already have a label headlessly
    candidates: List[Path] = []
    for p in all_files:
        if not overwrite_existing and p.suffix.lower() == ".docx":
            existing = _inspect_label_from_docx_file(p)
            if existing and existing.get("id"):
                skipped += 1
                if progress_callback:
                    progress_callback(skipped, total, p.name, f"Skipped (Already labeled: {existing.get('name')})")
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
