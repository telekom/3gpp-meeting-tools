"""Official 3GPP Word template download, caching, and %APPDATA% 3gpp_70.dotx installation."""

from __future__ import annotations

import io
import logging
import os
import shutil
import zipfile
from pathlib import Path
from typing import Optional, Tuple

import requests

from core.utils.paths import to_extended_path
from modules.word_tools.core.styles.registry import (
    OFFICIAL_3GPP_DOTX_URL,
    OFFICIAL_CR_TEMPLATE_URL,
    OFFICIAL_PCR_TEMPLATE_ZIP_URL,
)
from modules.word_tools.core.styles.xml_promoter import promote_3gpp_styles_in_xml

logger = logging.getLogger(__name__)


def _assets_dir() -> Path:
    assets_dir = Path(__file__).resolve().parents[4] / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    return assets_dir


def get_template_cr_path() -> Path:
    """
    Returns the absolute path to the cached Template_3GPP_CR.docx.
    If the template is missing locally, attempts an automatic download from 3GPP.
    """
    assets_dir = _assets_dir()
    tpl_path = assets_dir / "Template_3GPP_CR.docx"

    if tpl_path.exists() and tpl_path.stat().st_size > 5000:
        return tpl_path

    # Fallback auto-download: try corporate-aware NetworkSession first, then direct requests
    try:
        logger.info("Template_3GPP_CR.docx not found in assets; downloading from 3GPP via NetworkSession...")
        from core.network.session import NetworkSession
        NetworkSession.download_file(
            url=OFFICIAL_CR_TEMPLATE_URL,
            dest_path=tpl_path,
            timeout=30,
            atomic=True,
        )
        if tpl_path.exists() and tpl_path.stat().st_size > 5000:
            logger.info("Successfully downloaded Template_3GPP_CR.docx (%d bytes).", tpl_path.stat().st_size)
            return tpl_path
    except Exception as exc:
        logger.warning("NetworkSession download failed for Template_3GPP_CR.docx: %s. Trying direct requests...", exc)

    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        resp = requests.get(OFFICIAL_CR_TEMPLATE_URL, headers=headers, timeout=20)
        if resp.status_code == 200 and len(resp.content) > 5000:
            with open(to_extended_path(tpl_path), "wb") as f:
                f.write(resp.content)
            logger.info("Successfully downloaded Template_3GPP_CR.docx via requests (%d bytes).", len(resp.content))
            return tpl_path
    except Exception as exc:
        logger.warning("Could not auto-download Template_3GPP_CR.docx: %s", exc)

    if tpl_path.exists():
        return tpl_path
    raise FileNotFoundError(f"3GPP CR Template not found at {tpl_path} and download failed.")


def get_pseudo_cr_template_path() -> Path:
    """
    Returns the absolute path to the cached official Pseudo-CR-Form-v0-1.docx.
    If missing, downloads and extracts the official 3GPP template from 3GPP FTP.
    """
    assets_dir = _assets_dir()
    pcr_path = assets_dir / "Pseudo-CR-Form-v0-1.docx"

    if pcr_path.exists() and pcr_path.stat().st_size > 5000:
        return pcr_path

    zip_path = assets_dir / "Pseudo-CR-Form-v0-1.zip"
    try:
        logger.info("Pseudo-CR-Form-v0-1.docx not found; downloading from 3GPP...")
        from core.network.session import NetworkSession
        NetworkSession.download_file(
            url=OFFICIAL_PCR_TEMPLATE_ZIP_URL,
            dest_path=zip_path,
            timeout=30,
            atomic=True,
        )
        if zip_path.exists():
            with zipfile.ZipFile(to_extended_path(zip_path), "r") as z:
                z.extractall(assets_dir)
            if pcr_path.exists():
                logger.info("Successfully extracted official Pseudo-CR-Form-v0-1.docx.")
                return pcr_path
    except Exception as exc:
        logger.warning("Could not auto-download official Pseudo-CR template: %s", exc)
    finally:
        if zip_path.exists():
            try:
                zip_path.unlink()
            except OSError:
                pass

    if pcr_path.exists():
        return pcr_path
    return get_template_cr_path()


def get_3gpp_dotx_path() -> Path:
    """
    Returns the absolute path to the cached 3gpp_70.dotx template.
    """
    assets_dir = _assets_dir()
    dotx_path = assets_dir / "3gpp_70.dotx"

    if dotx_path.exists() and dotx_path.stat().st_size > 5000:
        return dotx_path

    try:
        logger.info("3gpp_70.dotx not found; downloading from 3GPP...")
        from core.network.session import NetworkSession
        NetworkSession.download_file(
            url=OFFICIAL_3GPP_DOTX_URL,
            dest_path=dotx_path,
            timeout=30,
            atomic=True,
        )
        if dotx_path.exists() and dotx_path.stat().st_size > 5000:
            logger.info("Successfully downloaded 3gpp_70.dotx (%d bytes).", dotx_path.stat().st_size)
            return dotx_path
    except Exception as exc:
        logger.warning("Could not auto-download 3gpp_70.dotx: %s", exc)

    return dotx_path


def is_3gpp_dotx_installed() -> bool:
    """Returns True if 3gpp_70.dotx is installed in the user's Word templates folder."""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return False
    target = Path(appdata) / "Microsoft" / "Templates" / "3gpp_70.dotx"
    return target.exists() and target.stat().st_size > 5000


def get_installed_3gpp_dotx_path() -> Optional[Path]:
    """Returns the Path to 3gpp_70.dotx in the user's Word templates folder if installed."""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    target = Path(appdata) / "Microsoft" / "Templates" / "3gpp_70.dotx"
    return target if target.exists() and target.stat().st_size > 5000 else None


def install_3gpp_dotx_template(force: bool = False) -> Tuple[bool, str]:
    """
    Installs 3gpp_70.dotx into the user's Word templates folder (%APPDATA%\\Microsoft\\Templates)
    so that Microsoft Word automatically links and attaches the official 3GPP styles and ribbon.
    Returns (success: bool, message: str).
    """
    try:
        appdata = os.environ.get("APPDATA")
        if not appdata:
            return False, "APPDATA environment variable is not defined."

        word_templates = Path(appdata) / "Microsoft" / "Templates"
        word_templates.mkdir(parents=True, exist_ok=True)
        target = word_templates / "3gpp_70.dotx"

        dotx_source = get_3gpp_dotx_path()
        if not dotx_source.exists() or dotx_source.stat().st_size == 0:
            return False, "Could not acquire official 3gpp_70.dotx template from 3GPP."

        if not force and target.exists() and target.stat().st_size == dotx_source.stat().st_size:
            return True, f"3GPP Word template is already installed at {target}."

        shutil.copy2(to_extended_path(dotx_source), to_extended_path(target))
        # Ensure the installed template itself has all authentic 3GPP styles promoted with QuickStyle flags
        try:
            with zipfile.ZipFile(to_extended_path(target), "r") as zin:
                if "word/styles.xml" in zin.namelist():
                    dotx_styles = zin.read("word/styles.xml")
                    dotx_promoted = promote_3gpp_styles_in_xml(dotx_styles)
                    dotx_buf = io.BytesIO()
                    with zipfile.ZipFile(dotx_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                        for item in zin.infolist():
                            if item.filename == "word/styles.xml":
                                zout.writestr(item, dotx_promoted)
                            else:
                                zout.writestr(item, zin.read(item.filename))
                    with open(to_extended_path(target), "wb") as f:
                        f.write(dotx_buf.getvalue())
        except Exception as promote_err:
            logger.debug("Failed promoting styles in installed 3gpp_70.dotx: %s", promote_err)

        logger.info("Installed and promoted 3gpp_70.dotx into %s (%d bytes)", target, target.stat().st_size)
        return True, f"Successfully installed 3GPP Word template to {target}."
    except Exception as exc:
        msg = f"Failed to install 3GPP Word template: {exc}"
        logger.warning(msg)
        return False, msg


def uninstall_3gpp_dotx_template() -> Tuple[bool, str]:
    """
    Removes 3gpp_70.dotx from the user's Word templates folder (%APPDATA%\\Microsoft\\Templates).
    Returns (success: bool, message: str).
    """
    try:
        appdata = os.environ.get("APPDATA")
        if not appdata:
            return False, "APPDATA environment variable is not defined."

        target = Path(appdata) / "Microsoft" / "Templates" / "3gpp_70.dotx"
        if not target.exists():
            return True, "3GPP Word template is not currently installed."

        try:
            target.unlink()
            logger.info("Removed 3gpp_70.dotx from %s", target)
            return True, "Successfully removed 3GPP Word template from Word templates folder."
        except PermissionError:
            return False, "Could not remove 3gpp_70.dotx: file is locked by a running Microsoft Word process."
    except Exception as exc:
        msg = f"Failed to remove 3GPP Word template: {exc}"
        logger.warning(msg)
        return False, msg
