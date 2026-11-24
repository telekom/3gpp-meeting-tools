"""Windows Office CLP / Purview MIP sensitivity policy discovery and XML parsing."""

from __future__ import annotations

import gzip
import logging
import os
import platform
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, List, Optional, Set, Union

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
                            if (
                                _local_tag(arg.tag) == "argument"
                                and (arg.attrib.get("key") or "").lower() == "disabled"
                            ):
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

        discovered_labels.append(
            {
                "id": label_id,
                "name": name,
                "display_name": display_name or name,
                "description": tooltip or desc_text or "",
                "order": order_val,
                "is_encrypted": lid_lower in encrypted_ids,
                "tenant_id": tenant_id,
                "source": "windows_policy",
            }
        )

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


def discover_windows_sensitivity_policy(
    clp_dir: Optional[Union[str, Path]] = None,
) -> Optional[dict]:
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
