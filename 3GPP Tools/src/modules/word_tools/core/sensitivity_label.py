# --- File: modules/word_tools/core/sensitivity_label.py ---
import gc
import gzip
import logging
import os
import platform
import uuid
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Set, Tuple, Union

from modules.word_tools.core.word_config import WordConfig

logger = logging.getLogger(__name__)


def _local_tag(tag: Any) -> str:
    """Return local XML element name without namespace."""
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def _read_policy_xml(path: Path) -> Optional[ET.Element]:
    """Read and parse a Windows Office CLP / MIP policy file (gzipped or plain XML)."""
    try:
        raw = path.read_bytes()
        if not raw:
            return None
        if raw[:2] == b"\x1f\x8b" or path.suffix.lower() == ".gz":
            text = gzip.decompress(raw).decode("utf-8", errors="replace")
        else:
            text = raw.decode("utf-8", errors="replace")
        return ET.fromstring(text)
    except Exception as exc:
        logger.debug("Skipping unreadable CLP policy file %s: %s", path, exc)
        return None


def _extract_localized_setting(setting_elem: ET.Element) -> str:
    """Extract preferred text from a <setting type='localizedString'> element."""
    default_val = ""
    en_val = ""
    first_val = ""
    for child in setting_elem:
        if _local_tag(child.tag).lower() != "value":
            continue
        txt = (child.text or "").strip()
        if not txt:
            continue
        if not first_val:
            first_val = txt
        locale = (child.attrib.get("locale") or "").strip().lower()
        if locale == "default":
            default_val = txt
        elif locale.startswith("en"):
            en_val = txt
    return default_val or en_val or first_val


def _parse_clp_syncfile(root: ET.Element, source_file: Path) -> Optional[dict]:
    """Parse a Microsoft Office CLP <SyncFile> XML tree into structured policy metadata."""
    if _local_tag(root.tag) != "SyncFile":
        return None

    sync_type = ""
    tenant_id = ""
    upn = ""
    content_elem: Optional[ET.Element] = None

    for child in root:
        tag = _local_tag(child.tag)
        if tag == "Type":
            sync_type = (child.text or "").strip()
        elif tag == "TenantId":
            tenant_id = (child.text or "").strip()
        elif tag == "Content":
            content_elem = child

    if sync_type and sync_type.lower() not in {"mippolicies", "mippolicy"}:
        return None
    if content_elem is None:
        return None

    labels_container: Optional[ET.Element] = None
    policies_container: Optional[ET.Element] = None

    for child in content_elem:
        tag = _local_tag(child.tag)
        if tag == "Upn":
            upn = (child.text or "").strip()
        elif tag == "labels":
            labels_container = child
        elif tag == "policies":
            policies_container = child

    if labels_container is None:
        return None

    # Policy-level settings
    mandatory = False
    default_label_id = ""
    outlook_default_label_id = ""
    for child in labels_container:
        if _local_tag(child.tag) == "settings":
            for setting in child:
                if _local_tag(setting.tag) != "setting":
                    continue
                k = (setting.attrib.get("key") or "").strip().lower()
                v = (setting.attrib.get("value") or "").strip()
                if k == "mandatory":
                    mandatory = v.lower() == "true"
                elif k == "defaultlabel":
                    default_label_id = v
                elif k == "outlookdefaultlabel":
                    outlook_default_label_id = v

    # Collect label IDs that enforce RMS / protection encryption rules
    encrypted_ids: Set[str] = set()
    if policies_container is not None:
        for rule in policies_container.iter():
            if _local_tag(rule.tag) != "rule":
                continue
            has_protection = False
            rule_label_ids: List[str] = []
            for elem in rule.iter():
                etag = _local_tag(elem.tag)
                if etag == "action":
                    aname = (elem.attrib.get("name") or "").strip().lower()
                    if aname.startswith("rightsprotect") or "encrypt" in aname:
                        disabled = False
                        for arg in elem:
                            if _local_tag(arg.tag) == "argument" and (arg.attrib.get("key") or "").lower() == "disabled":
                                disabled = (arg.attrib.get("value") or "").strip().lower() == "true"
                        if not disabled:
                            has_protection = True
                elif etag == "keyValue":
                    if (elem.attrib.get("key") or "").strip().lower() == "label":
                        val = (elem.attrib.get("value") or "").strip().lower()
                        if val:
                            rule_label_ids.append(val)
            if has_protection:
                encrypted_ids.update(rule_label_ids)

    # Extract all enabled, non-parent File-scoped labels (supporting nested sublabels)
    discovered_labels: List[dict] = []
    seen_ids: Set[str] = set()

    for label_elem in labels_container.iter():
        if _local_tag(label_elem.tag) != "label":
            continue

        label_id = (label_elem.attrib.get("id") or "").strip()
        name = (label_elem.attrib.get("name") or "").strip()
        if not label_id or not name:
            continue

        enabled = (label_elem.attrib.get("enabled") or "true").strip().lower() == "true"
        is_parent_attr = (label_elem.attrib.get("isParentLabel") or "false").strip().lower() == "true"
        if not enabled or is_parent_attr:
            continue

        desc_text = ""
        display_name = ""
        tooltip = ""
        order_val = 999
        is_parent_setting = False
        content_types = ""

        for child in label_elem:
            ctag = _local_tag(child.tag)
            if ctag == "description" and not desc_text:
                desc_text = (child.text or "").strip()
            elif ctag == "settings":
                for setting in child:
                    if _local_tag(setting.tag) != "setting":
                        continue
                    k = (setting.attrib.get("key") or "").strip().lower()
                    v = (setting.attrib.get("value") or "").strip()
                    if k == "displayname":
                        display_name = _extract_localized_setting(setting)
                    elif k == "tooltip":
                        tooltip = _extract_localized_setting(setting)
                    elif k == "order":
                        try:
                            order_val = int(v)
                        except ValueError:
                            pass
                    elif k == "isparent":
                        is_parent_setting = v.lower() == "true"
                    elif k == "contenttype":
                        content_types = v

        if is_parent_setting:
            continue

        # Filter to labels scoped for Files (Word/Office documents) when contenttype is specified
        if content_types:
            ct_tokens = {tok.strip().lower() for tok in content_types.split(",") if tok.strip()}
            if "file" not in ct_tokens:
                continue

        lid_lower = label_id.lower()
        if lid_lower in seen_ids:
            continue
        seen_ids.add(lid_lower)

        discovered_labels.append({
            "id": label_id,
            "name": name,
            "display_name": display_name or name,
            "description": tooltip or desc_text or "",
            "order": order_val,
            "is_encrypted": lid_lower in encrypted_ids,
            "tenant_id": tenant_id,
            "source": "windows_policy",
        })

    discovered_labels.sort(key=lambda item: (item["order"], item["name"].lower()))

    if not discovered_labels:
        return None

    # Default selection rule: lowest-order unencrypted File label (fallback to first label if all encrypted)
    recommended_default_id = ""
    for lbl in discovered_labels:
        if not lbl.get("is_encrypted"):
            recommended_default_id = lbl["id"]
            break
    if not recommended_default_id:
        recommended_default_id = discovered_labels[0]["id"]

    return {
        "tenant_id": tenant_id,
        "upn": upn,
        "mandatory": mandatory,
        "default_label_id": default_label_id,
        "outlook_default_label_id": outlook_default_label_id,
        "recommended_default_id": recommended_default_id,
        "source_file": str(source_file),
        "labels": discovered_labels,
    }


def discover_windows_sensitivity_policy(clp_dir: Optional[Union[str, Path]] = None) -> Optional[dict]:
    """
    Discover the user's Microsoft Purview / MIP sensitivity labels from the local Windows
    Office CLP policy cache (%LOCALAPPDATA%\\Microsoft\\Office\\CLP\\policy*.gz).

    Args:
        clp_dir: Optional explicit directory to scan (primarily for deterministic testing).

    Returns:
        dict with keys 'tenant_id', 'upn', 'mandatory', 'recommended_default_id',
        'source_file', and 'labels', or None if no local policy is found.
    """
    candidate_files: List[Path] = []

    if clp_dir is not None:
        search_root = Path(clp_dir)
        if search_root.is_dir():
            for p in search_root.glob("policy*"):
                if p.is_file() and p.suffix.lower() in {".gz", ".xml", ""}:
                    candidate_files.append(p)
    else:
        if platform.system() != "Windows":
            return None
        local_app = os.environ.get("LOCALAPPDATA", "")
        app_data = os.environ.get("APPDATA", "")
        search_dirs: List[Path] = []
        if local_app:
            search_dirs.append(Path(local_app) / "Microsoft" / "Office" / "CLP")
            search_dirs.append(Path(local_app) / "Microsoft" / "MSIP")
        if app_data:
            search_dirs.append(Path(app_data) / "Microsoft" / "CLP")

        for d in search_dirs:
            if not d.is_dir():
                continue
            for p in d.rglob("policy*"):
                if p.is_file() and p.suffix.lower() in {".gz", ".xml", ""}:
                    candidate_files.append(p)

    if not candidate_files:
        return None

    # Prioritize most recently synced policy file
    candidate_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)

    for policy_path in candidate_files:
        root = _read_policy_xml(policy_path)
        if root is None:
            continue
        parsed = _parse_clp_syncfile(root, policy_path)
        if parsed and parsed.get("labels"):
            return parsed

    return None


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
    active_label = WordConfig.get_active_sensitivity_label()
    if not label_id or not label_name:
        label_id = label_id or active_label.get("id") or config.get("sensitivity_level_label_id")
        label_name = label_name or active_label.get("name") or config.get("sensitivity_level_label_name")

    if not label_id or not label_name:
        _log("   ➔ No sensitivity label configured or detected in user policy. Skipping...", logging.DEBUG, ui_logger)
        return False if return_changed else document

    # Resolve TenantId / SiteId for Purview compliance
    site_id = ""
    for lbl in config.get("sensitivity_labels", []):
        if isinstance(lbl, dict) and (lbl.get("id") or "").lower() == str(label_id).lower():
            site_id = lbl.get("tenant_id", "")
            break
    site_id = site_id or active_label.get("tenant_id") or config.get("sensitivity_tenant_id") or str(uuid.uuid1())

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
            new_sl.SiteId = site_id

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
