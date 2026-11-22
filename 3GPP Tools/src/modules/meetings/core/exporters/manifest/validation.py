# --- File: src/modules/meetings/core/exporters/manifest/validation.py ---
"""Constants, JSON sanitization, folder naming, and Markdown completeness checks."""

import datetime
import posixpath
import re
from pathlib import Path
from typing import Any, Dict, Optional, Union

from core.utils.paths import to_extended_path

MANIFEST_FILENAME = ".export_manifest.json"
MANIFEST_VERSION = 2


def _clean_for_json(obj: Any) -> Any:
    """Recursively converts sets, tuples, and Paths into JSON-serializable types."""
    if isinstance(obj, (set, frozenset)):
        return sorted([_clean_for_json(x) for x in obj], key=lambda x: str(x))
    if isinstance(obj, (list, tuple)):
        return [_clean_for_json(x) for x in obj]
    if isinstance(obj, dict):
        return {str(k): _clean_for_json(v) for k, v in obj.items()}
    if isinstance(obj, Path):
        return str(obj)
    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.isoformat()
    return obj


def _json_fallback(obj: Any) -> Any:
    """Fallback handler for json.dump default parameter."""
    if isinstance(obj, (set, frozenset)):
        return sorted(list(obj), key=lambda x: str(x))
    if isinstance(obj, tuple):
        return list(obj)
    if isinstance(obj, Path):
        return str(obj)
    if hasattr(obj, "to_dict"):
        return obj.to_dict()
    if hasattr(obj, "__dict__"):
        return obj.__dict__
    return str(obj)


def sanitize_folder_name(name: str) -> str:
    """Sanitizes directory names, removing illegal Windows/filesystem and URL-breaking characters."""
    clean = re.sub(r'[\\/*?:"<>|#%]', "_", str(name or "")).strip()
    return re.sub(r"_+", "_", clean).strip("._ ") or "General"


def is_export_complete(
    md_path: Union[str, Path],
    doc_info: Optional[Dict[str, Any]] = None,
    source_path: Optional[Union[str, Path]] = None,
) -> bool:
    """
    Validates whether an existing Markdown export file on disk is complete and usable.
    Returns False if:
      - File does not exist or is 0 bytes
      - File size is suspiciously small (< 120 bytes, e.g. truncated header or empty stub)
      - File is an empty placeholder stub ("Document conversion was not available")
      - File is a presentation export (PPTX/PPT) but has no slides ("## Slide ")
      - File is a presentation export referencing slide images that are missing on disk
    """
    p = Path(to_extended_path(md_path))
    if not p.is_file():
        return False
    try:
        st = p.stat()
        if st.st_size == 0:
            return False

        with open(p, "r", encoding="utf-8", errors="replace") as f:
            head = f.read(2048)

        # Minimum content threshold: reject tiny stubs unless stamped by converter
        if st.st_size < 50:
            return False
        if st.st_size < 120 and "<!--- tdoc_converter:" not in head:
            return False

        # Placeholder stub created when source was missing
        if "Document conversion was not available" in head:
            return False

        # Placeholder stub created when document was reserved / not uploaded
        if (
            "Status: Reserved / Not Uploaded" in head
            or "allocated or reserved on the meeting agenda but has not yet been uploaded" in head
            or "⏳ **Status: Reserved" in head
        ):
            return False

        # Corrupted image prefix from previous bug
        if "__-________" in head:
            return False

        # Detect presentation export
        is_pres = False
        if source_path and Path(source_path).suffix.lower() in (".pptx", ".ppt", ".pdf"):
            is_pres = True
        elif doc_info:
            src_file = str(
                doc_info.get("original_file", "") or doc_info.get("file", "") or doc_info.get("filename", "")
            ).lower()
            if src_file.endswith((".pptx", ".ppt")):
                is_pres = True
            doc_type = str(doc_info.get("type", "")).lower()
            if "presentation" in doc_type or "slides" in doc_type:
                is_pres = True

        if not is_pres:
            if "type=presentation" in head or "mode=presentation" in head:
                is_pres = True
            elif re.search(r"\*\*Original file:\*\*.*?\.ppt[x]?", head, re.IGNORECASE):
                is_pres = True
            elif "- **Slides:**" in head:
                is_pres = True

        # Read content to validate completeness and image presence
        with open(p, "r", encoding="utf-8", errors="replace") as f:
            content = f.read(65536)

        if "__-________" in content:
            return False

        if is_pres and "## Slide " not in content:
            return False

        # Check all referenced images exist on disk
        all_imgs = re.findall(r"!\[[^\]]*\]\(([^)]+)\)", content)
        for img_rel in all_imgs:
            img_rel = img_rel.strip().replace("\\", "/")
            if img_rel.startswith(("http://", "https://", "data:")):
                continue
            img_path = (p.parent / img_rel).resolve()
            if not img_path.is_file():
                # Check fallback locations: {stem}_media/ and media/
                img_name = posixpath.basename(img_rel)
                stem_media = (p.parent / f"{p.stem}_media" / img_name).resolve()
                sub_media = (p.parent / "media" / img_name).resolve()
                if not stem_media.is_file() and not sub_media.is_file():
                    return False

        return True
    except Exception:
        return False
