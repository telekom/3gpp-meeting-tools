"""OpenXML .docx package style patching with lock-resilient atomic replacement."""

from __future__ import annotations

import io
import logging
import os
import re
import shutil
import time
import zipfile
from pathlib import Path

from core.utils.paths import to_extended_path
from modules.word_tools.core.styles.templates import get_template_cr_path
from modules.word_tools.core.styles.xml_promoter import (
    merge_and_promote_3gpp_styles,
    promote_3gpp_styles_in_xml,
)

logger = logging.getLogger(__name__)


def get_3gpp_package_part(part_name: str) -> bytes:
    """
    Extracts a raw package part (e.g. 'word/styles.xml', 'word/theme/theme1.xml')
    directly from the cached official 3GPP template, with full 3GPP QuickStyle promotion.
    """
    tpl_path = get_template_cr_path()
    with zipfile.ZipFile(to_extended_path(tpl_path), "r") as z:
        data = z.read(part_name)
    if part_name == "word/styles.xml":
        return promote_3gpp_styles_in_xml(data)
    return data


def ensure_3gpp_package_styles(docx_path: Path) -> None:
    """
    Non-destructively ensures that the target .docx file contains all official 3GPP styles
    and the standard TableGrid style with proper borders and lines.
    Preserves all existing document styles, tables, themes, and fonts.
    """
    ext_path = to_extended_path(docx_path)
    if not os.path.exists(ext_path):
        raise FileNotFoundError(f"Target document does not exist: {docx_path}")

    template_styles_xml = get_3gpp_package_part("word/styles.xml")
    template_theme_xml = get_3gpp_package_part("word/theme/theme1.xml")
    template_font_table_xml = get_3gpp_package_part("word/fontTable.xml")

    tmp_path = docx_path.with_name(docx_path.name + ".stylepatch.tmp")
    ext_tmp = to_extended_path(tmp_path)

    try:
        # Retry loop to read the source file in case another process (Word teardown, OneDrive, AV) holds a transient lock
        max_read_attempts = 15
        read_delay = 0.2
        source_bytes = None
        for attempt in range(max_read_attempts):
            try:
                with open(ext_path, "rb") as sf:
                    source_bytes = sf.read()
                break
            except (PermissionError, OSError):
                if attempt == max_read_attempts - 1:
                    raise
                time.sleep(read_delay)

        out_buf = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(source_bytes), "r") as zin, zipfile.ZipFile(
            out_buf, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            namelist = set(zin.namelist())

            # Non-destructive merge: preserve existing document styles, merge 3GPP styles & ensure TableGrid
            if "word/styles.xml" in namelist:
                existing_styles = zin.read("word/styles.xml")
                merged_styles = merge_and_promote_3gpp_styles(existing_styles, template_styles_xml)
            else:
                merged_styles = promote_3gpp_styles_in_xml(template_styles_xml)

            for item in zin.infolist():
                if item.filename == "word/styles.xml":
                    zout.writestr(item, merged_styles)
                elif item.filename == "word/theme/theme1.xml":
                    # Preserve existing theme (colors, fonts, effects)
                    zout.writestr(item, zin.read(item.filename))
                elif item.filename == "word/fontTable.xml":
                    # Preserve existing font table
                    zout.writestr(item, zin.read(item.filename))
                elif item.filename == "word/settings.xml":
                    # Remove <w:linkStyles/> to prevent Microsoft Word from overwriting styles on document open
                    try:
                        settings_text = zin.read(item.filename).decode("utf-8")
                        settings_text = re.sub(r"<(?:\w+:)?linkStyles\b[^>]*/?>", "", settings_text)
                        zout.writestr(item, settings_text.encode("utf-8"))
                    except Exception:
                        zout.writestr(item, zin.read(item.filename))
                else:
                    zout.writestr(item, zin.read(item.filename))

            # Only provide theme/fontTable if completely absent from target document
            if "word/theme/theme1.xml" not in namelist and template_theme_xml:
                zout.writestr("word/theme/theme1.xml", template_theme_xml)
            if "word/fontTable.xml" not in namelist and template_font_table_xml:
                zout.writestr("word/fontTable.xml", template_font_table_xml)

        with open(ext_tmp, "wb") as tf:
            tf.write(out_buf.getvalue())

        # Robust retry loop guards against transient Windows file locking (Word teardown, OneDrive, antivirus/indexers)
        max_replace_attempts = 25
        replace_delay = 0.25
        last_replace_exc = None
        for attempt in range(max_replace_attempts):
            try:
                os.replace(ext_tmp, ext_path)
                last_replace_exc = None
                break
            except (PermissionError, OSError) as exc:
                last_replace_exc = exc
                if attempt % 4 == 0 and attempt > 0:
                    logger.debug(
                        "Target file %s is locked; waiting for release (%d/%d)...",
                        docx_path.name, attempt + 1, max_replace_attempts,
                    )
                # On final attempts, try direct file stream copy fallback
                if attempt >= max_replace_attempts - 3:
                    try:
                        shutil.copyfile(ext_tmp, ext_path)
                        try:
                            os.remove(ext_tmp)
                        except OSError:
                            pass
                        last_replace_exc = None
                        break
                    except Exception:
                        pass
                time.sleep(replace_delay)

        if last_replace_exc is not None:
            raise last_replace_exc

        logger.info("Ensured authentic 3GPP styles and table formatting in %s", docx_path.name)
    finally:
        if os.path.exists(ext_tmp):
            try:
                os.remove(ext_tmp)
            except OSError:
                pass
