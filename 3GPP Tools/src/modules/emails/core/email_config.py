# --- File: src/modules/emails/core/email_config.py ---
import json
import logging
from core.utils.paths import get_project_root

logger = logging.getLogger(__name__)

CONFIG_PATH = get_project_root() / "config" / "emails_config.json"


def load_wg_email_config(wg: str) -> list:
    """Loads email folder configuration for a given Working Group."""
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get(wg.upper(), [])
        except Exception as e:
            logger.error(f"Failed to load emails_config.json: {e}")
    return []


def save_wg_email_config(wg: str, folder_list: list):
    """Saves email folder configuration for a given Working Group."""
    data = {}
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
    data[wg.upper()] = folder_list
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        logger.error(f"Failed to save emails_config.json: {e}")
