"""Headless OpenXML and Word COM sensitivity label inspection helpers."""

from __future__ import annotations

import gc
import logging
import os
import platform
import zipfile
import xml.etree.ElementTree as ET
from enum import Enum
from pathlib import Path
from typing import Any, Dict, Optional, Union

logger = logging.getLogger(__name__)


def _log(message: str, level: int = logging.INFO, ui_logger=None) -> None:
    """Log through Python logging; optionally support legacy signal/logger callers."""
    logger.log(level, message)


class MsoAssignmentMethod(Enum):
    """Office SensitivityLabel assignment method enum."""

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


def get_document_sensitivity_label(
    document_or_path: Union[Any, str, Path],
) -> Optional[dict]:
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
