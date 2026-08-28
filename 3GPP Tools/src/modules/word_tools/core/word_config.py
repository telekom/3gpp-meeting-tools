# --- File: src/modules/word_tools/core/word_config.py ---
import json
import logging
from pathlib import Path

import core.utils.paths

CONFIG_PATH = core.utils.paths.get_project_root() / "word_config.json"


class WordConfig:
    @staticmethod
    def load() -> dict:
        default = {
            "enable_sensitivity_labels": True,
            "sensitivity_level_label_id": "55339bf0-f345-473a-9ec8-6ca7c8197055",
            "sensitivity_level_label_name": "OFFEN",
            "save_document_after_setting_sensitivity_label": False,
            "libreoffice_path": "",
        }
        try:
            if CONFIG_PATH.exists():
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    default.update(json.load(f))
            else:
                WordConfig.save(default)
        except Exception as e:
            logging.error(f"Failed to load Word config: {e}")
        return default

    @staticmethod
    def save(data: dict):
        try:
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            logging.error(f"Failed to save Word config: {e}")

    @staticmethod
    def get_libreoffice_path() -> str:
        return WordConfig.load().get("libreoffice_path", "").strip()

    @staticmethod
    def set_libreoffice_path(path_str: str):
        cfg = WordConfig.load()
        cfg["libreoffice_path"] = path_str.strip()
        WordConfig.save(cfg)