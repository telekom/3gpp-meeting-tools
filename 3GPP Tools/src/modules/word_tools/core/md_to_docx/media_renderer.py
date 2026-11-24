"""Image resolution and vector/metafile/PlantUML diagram rendering helpers."""

from __future__ import annotations

import logging
import os
import re
import shutil
import subprocess
import tempfile
import urllib.parse
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QByteArray, QSize
from PyQt5.QtGui import QImage, QPainter
from PyQt5.QtSvg import QSvgRenderer

from core.utils.utils import get_best_java
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.utils.paths import get_puml2visio_asset_path

logger = logging.getLogger(__name__)

# Callout styling configuration: (border_hex, fill_hex, default_title)
CALLOUT_CONFIG = {
    "NOTE": ("#0284C7", "F0F9FF", "ℹ️ NOTE"),
    "TIP": ("#10B981", "ECFDF5", "💡 TIP"),
    "IMPORTANT": ("#8B5CF6", "F5F3FF", "❗ IMPORTANT"),
    "WARNING": ("#F59E0B", "FFFBEB", "⚠️ WARNING"),
    "CAUTION": ("#EF4444", "FEF2F2", "🛑 CAUTION"),
}

PLANTUML_START_DIRECTIVES = (
    "@startuml",
    "@startmindmap",
    "@startwbs",
    "@startgantt",
    "@startjson",
    "@startyaml",
    "@startditaa",
    "@startsalt",
)


def render_svg_to_png(svg_path: Path, output_png_path: Path, max_dimension: int = 1600) -> bool:
    """Renders an SVG file to a high-resolution PNG image using QtSvg."""
    try:
        svg_bytes = svg_path.read_bytes()
        renderer = QSvgRenderer(QByteArray(svg_bytes))
        if not renderer.isValid():
            return False

        default_size = renderer.defaultSize()
        w = default_size.width() if default_size.width() > 0 else 800
        h = default_size.height() if default_size.height() > 0 else 600

        # Scale up proportionally for crisp print resolution if small
        if w < 800 and h < 800:
            scale = min(max_dimension / max(w, 1), 2.0)
            w = int(w * scale)
            h = int(h * scale)

        img = QImage(QSize(w, h), QImage.Format_ARGB32)
        img.fill(0)  # Transparent background

        painter = QPainter(img)
        renderer.render(painter)
        painter.end()

        output_png_path.parent.mkdir(parents=True, exist_ok=True)
        return img.save(str(output_png_path), "PNG")
    except Exception as e:
        logger.debug(f"SVG to PNG rendering failed for '{svg_path}': {e}")
        return False


def render_metafile_to_png(metafile_path: Path, output_png_path: Path) -> bool:
    """Converts a Windows Metafile (.emf / .wmf) to PNG using the application metafile converter."""
    try:
        from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_metafiles

        output_png_path.parent.mkdir(parents=True, exist_ok=True)
        convert_metafiles([(metafile_path, output_png_path)], dpi=150)
        return output_png_path.is_file() and output_png_path.stat().st_size > 0
    except Exception as e:
        logger.debug(f"Metafile to PNG conversion failed for '{metafile_path}': {e}")
        return False


def _extract_plantuml_title(puml_code: str) -> Optional[str]:
    """Extracts an explicit diagram title from PlantUML source code if defined."""
    for line in puml_code.splitlines():
        s = line.strip()
        m_title = re.match(r"^title\s+(.+)$", s, re.IGNORECASE)
        if m_title:
            return m_title.group(1).strip().strip('"')
        m_start = re.match(r"^@startuml\s+(.+)$", s, re.IGNORECASE)
        if m_start:
            cand = m_start.group(1).strip().strip('"')
            if cand and not cand.startswith("-"):
                return cand
    return None


def render_plantuml_to_png(puml_code: str, output_png_path: Path) -> bool:
    """Compiles PlantUML sequence/architecture text into a PNG using the bundled plantuml.jar."""
    java_exe, _ = get_best_java()
    if not java_exe:
        return False

    jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
    if not jar_path.exists():
        return False

    code = puml_code.strip()
    if not code.lower().startswith("@start"):
        code = f"@startuml\n{code}\n@enduml"

    with tempfile.TemporaryDirectory() as td:
        puml_file = Path(td) / "diagram.puml"
        puml_file.write_text(code, encoding="utf-8")

        cmd = [java_exe, "-jar", str(jar_path), "-tpng", str(puml_file)]
        kwargs = {"creationflags": 0x08000000} if os.name == "nt" else {}
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=45, **kwargs)
            png_files = sorted(Path(td).glob("*.png"))
            gen_png = Path(td) / "diagram.png"
            if not gen_png.exists() and png_files:
                gen_png = png_files[0]
            if gen_png.exists() and gen_png.stat().st_size > 0:
                output_png_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(gen_png, output_png_path)
                return True
            else:
                logger.debug(f"PlantUML did not produce image (exit {res.returncode}): {res.stderr}")
        except Exception as e:
            logger.debug(f"PlantUML compilation error: {e}")
            return False
    return False


def resolve_image_path(
    src: str,
    md_path: Path,
    companion_media_dir: Optional[Path] = None,
    auto_detect_media: bool = True,
) -> Optional[Path]:
    """Resolves an image reference from Markdown into a local file path."""
    if not src:
        return None

    raw = src.strip()
    if raw.lower().startswith("file:///"):
        raw = raw[8:] if len(raw) > 9 and raw[9] == ":" else raw[7:]

    clean_src = urllib.parse.unquote(raw.split("?")[0].split("#")[0].strip())
    if not clean_src:
        return None

    p = Path(clean_src)

    # 1. Absolute path check
    try:
        if p.is_Absolute() if hasattr(p, "is_Absolute") else p.is_absolute():
            if p.is_file():
                return p
    except Exception:
        pass

    # 2. Relative to markdown file parent
    try:
        cand = (md_path.parent / clean_src).resolve()
        if cand.is_file():
            return cand
    except Exception:
        pass

    # 3. Explicit companion media directory
    if companion_media_dir and companion_media_dir.is_dir():
        cand = (companion_media_dir / p.name).resolve()
        if cand.is_file():
            return cand

    # 4. Auto-detected companion media directory ({stem}_media/) and standard subdirectories
    if auto_detect_media:
        cand = (md_path.parent / f"{md_path.stem}_media" / p.name).resolve()
        if cand.is_file():
            return cand

        for sub in ("media", "images", "assets", "figures"):
            cand = (md_path.parent / sub / p.name).resolve()
            if cand.is_file():
                return cand

    return None
