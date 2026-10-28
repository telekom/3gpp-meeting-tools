# --- File: src/modules/word_tools/core/word_config.py ---
import json
import logging
from typing import Dict, List, Optional

import core.utils.paths

CONFIG_PATH = core.utils.paths.get_project_root() / "config" / "word_config.json"

DEFAULT_SENSITIVITY_LABELS = [
    {
        "id": "55339bf0-f345-473a-9ec8-6ca7c8197055",
        "name": "OFFEN",
        "description": "Public / Non-confidential (Default for 3GPP Contributions)",
    },
    {
        "id": "5c6efe59-7e7c-4bd6-ba39-1850710f2efa",
        "name": "INTERN",
        "description": "Internal / Company-wide",
    },
    {
        "id": "5d95d744-d996-4a3c-a104-fbf2486a4c52",
        "name": "VERTRAULICH - unverschlüsselt",
        "description": "Confidential - unencrypted",
    },
]


class WordConfig:
    DEFAULT_SENSITIVITY_LABELS = DEFAULT_SENSITIVITY_LABELS

    @staticmethod
    def load() -> dict:
        default = {
            "enable_sensitivity_labels": True,
            "active_sensitivity_label_id": "55339bf0-f345-473a-9ec8-6ca7c8197055",
            "sensitivity_level_label_id": "55339bf0-f345-473a-9ec8-6ca7c8197055",
            "sensitivity_level_label_name": "OFFEN",
            "sensitivity_labels": [dict(l) for l in DEFAULT_SENSITIVITY_LABELS],
            "save_document_after_setting_sensitivity_label": False,
            "libreoffice_path": "",
            "default_source_company": "",
        }
        try:
            if CONFIG_PATH.exists():
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    default.update(data)
            else:
                WordConfig.save(default)
        except Exception as e:
            logging.error(f"Failed to load Word config: {e}")

        # Ensure sensitivity_labels contains at least the defaults and is a list of dicts
        labels = default.get("sensitivity_labels")
        if not isinstance(labels, list) or not labels:
            default["sensitivity_labels"] = [dict(l) for l in DEFAULT_SENSITIVITY_LABELS]
        else:
            existing_ids = {l.get("id") for l in default["sensitivity_labels"] if isinstance(l, dict)}
            for def_label in DEFAULT_SENSITIVITY_LABELS:
                if def_label["id"] not in existing_ids:
                    default["sensitivity_labels"].append(dict(def_label))

        # Ensure active_sensitivity_label_id is valid
        active_id = default.get("active_sensitivity_label_id") or default.get("sensitivity_level_label_id")
        valid_ids = {l.get("id") for l in default["sensitivity_labels"] if isinstance(l, dict)}
        if not active_id or active_id not in valid_ids:
            active_id = default["sensitivity_labels"][0].get("id")
        default["active_sensitivity_label_id"] = active_id

        # Sync legacy keys to active label
        for l in default["sensitivity_labels"]:
            if isinstance(l, dict) and l.get("id") == active_id:
                default["sensitivity_level_label_id"] = l.get("id")
                default["sensitivity_level_label_name"] = l.get("name")
                break

        return default

    @staticmethod
    def save(data: dict):
        try:
            # Sync active label to legacy keys for backward compatibility
            active_id = data.get("active_sensitivity_label_id")
            labels = data.get("sensitivity_labels", [])
            for l in labels:
                if isinstance(l, dict) and l.get("id") == active_id:
                    data["sensitivity_level_label_id"] = l.get("id")
                    data["sensitivity_level_label_name"] = l.get("name")
                    break

            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logging.error(f"Failed to save Word config: {e}")

    @staticmethod
    def get_sensitivity_labels() -> List[dict]:
        return WordConfig.load().get("sensitivity_labels", [dict(l) for l in DEFAULT_SENSITIVITY_LABELS])

    @staticmethod
    def get_active_sensitivity_label() -> dict:
        cfg = WordConfig.load()
        active_id = cfg.get("active_sensitivity_label_id")
        for l in cfg.get("sensitivity_labels", []):
            if isinstance(l, dict) and l.get("id") == active_id:
                return l
        if cfg.get("sensitivity_labels"):
            return cfg["sensitivity_labels"][0]
        return {
            "id": cfg.get("sensitivity_level_label_id", "55339bf0-f345-473a-9ec8-6ca7c8197055"),
            "name": cfg.get("sensitivity_level_label_name", "OFFEN"),
            "description": "",
        }

    @staticmethod
    def set_active_sensitivity_label_id(label_id: str):
        cfg = WordConfig.load()
        cfg["active_sensitivity_label_id"] = label_id.strip()
        WordConfig.save(cfg)

    @staticmethod
    def add_or_update_sensitivity_label(label_id: str, name: str, description: str = ""):
        cfg = WordConfig.load()
        labels = cfg.get("sensitivity_labels", [])
        found = False
        for l in labels:
            if isinstance(l, dict) and l.get("id") == label_id:
                l["name"] = name.strip()
                l["description"] = description.strip()
                found = True
                break
        if not found:
            labels.append({
                "id": label_id.strip(),
                "name": name.strip(),
                "description": description.strip(),
            })
        cfg["sensitivity_labels"] = labels
        WordConfig.save(cfg)

    @staticmethod
    def remove_sensitivity_label(label_id: str) -> bool:
        cfg = WordConfig.load()
        labels = cfg.get("sensitivity_labels", [])
        new_labels = [l for l in labels if isinstance(l, dict) and l.get("id") != label_id]
        if len(new_labels) == len(labels):
            return False
        cfg["sensitivity_labels"] = new_labels
        if cfg.get("active_sensitivity_label_id") == label_id and new_labels:
            cfg["active_sensitivity_label_id"] = new_labels[0].get("id")
        WordConfig.save(cfg)
        return True

    @staticmethod
    def get_libreoffice_path() -> str:
        return WordConfig.load().get("libreoffice_path", "").strip()

    @staticmethod
    def set_libreoffice_path(path_str: str):
        cfg = WordConfig.load()
        cfg["libreoffice_path"] = path_str.strip()
        WordConfig.save(cfg)
