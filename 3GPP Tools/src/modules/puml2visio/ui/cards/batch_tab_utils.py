"""Helper utilities for Batch Convert Tab and operation cards."""

import logging
from pathlib import Path
from typing import Optional
import xml.etree.ElementTree as ET
import zipfile

from modules.puml2visio.utils.utils import strip_watermark

logger = logging.getLogger(__name__)


def is_visio_registered() -> bool:
    """Non-fatal check whether Microsoft Visio COM server is registered in Windows."""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Visio.Application")
        winreg.CloseKey(key)
        return True
    except (FileNotFoundError, OSError, ImportError):
        return False


def is_powerpoint_registered() -> bool:
    """Non-fatal check whether Microsoft PowerPoint COM server is registered in Windows."""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application")
        winreg.CloseKey(key)
        return True
    except (FileNotFoundError, OSError, ImportError):
        return False


def detect_puml_diagram_type(file_path: Path) -> str:
    """Quickly scans the opening lines of a PlantUML file to infer the diagram type."""
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = [f.readline() for _ in range(35)]
        content = " ".join(lines).lower()
        if "sequence" in content or "->" in content or "autonumber" in content:
            return "Sequence"
        elif "class " in content or "interface " in content:
            return "Class"
        elif "start" in content and "stop" in content:
            return "Activity"
        elif "[*]" in content or "state " in content:
            return "State"
        elif "usecase" in content or "actor " in content:
            return "Use Case"
        elif "package " in content or "node " in content or "component " in content:
            return "Component / Arch"
    except Exception:
        pass
    return "PlantUML"


def extract_plantuml_from_vsdx_fast(vsdx_path: Path) -> Optional[str]:
    """Headless OpenXML extraction of embedded PlantUML source from a .vsdx archive.

    Runs in under 5ms without spawning any COM Visio process.
    """
    try:
        with zipfile.ZipFile(vsdx_path, "r") as z:
            names = z.namelist()
            for name in names:
                if name.startswith("visio/pages/page") and name.endswith(".xml"):
                    content = z.read(name)
                    if b"@startuml" in content:
                        root = ET.fromstring(content)
                        text = "".join(root.itertext())
                        if "@startuml" in text:
                            return strip_watermark(text)
    except Exception as e:
        logger.debug(f"Fast vsdx extract failed for {vsdx_path.name}: {e}")
    return None


def format_file_size(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    elif num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    else:
        return f"{num_bytes / (1024 * 1024):.1f} MB"
