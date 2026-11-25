# --- File: src/modules/word_tools/core/word_config.py ---
import json
import logging
from typing import List, Optional

import core.utils.paths

CONFIG_PATH = core.utils.paths.get_project_root() / "config" / "word_config.json"

# No hardcoded tenant GUIDs: labels are dynamically discovered from the user's Windows/Office CLP policy.
DEFAULT_SENSITIVITY_LABELS: List[dict] = []

_LEGACY_SEED_LABEL_IDS = {
    "55339bf0-f345-473a-9ec8-6ca7c8197055",
    "5c6efe59-7e7c-4bd6-ba39-1850710f2efa",
    "5d95d744-d996-4a3c-a104-fbf2486a4c52",
}


class WordConfig:
    DEFAULT_SENSITIVITY_LABELS = DEFAULT_SENSITIVITY_LABELS

    @staticmethod
    def _default_payload() -> dict:
        return {
            "enable_sensitivity_labels": True,
            "active_sensitivity_label_id": "",
            "sensitivity_level_label_id": "",
            "sensitivity_level_label_name": "",
            "sensitivity_tenant_id": "",
            "sensitivity_policy_upn": "",
            "sensitivity_policy_initialized": False,
            "sensitivity_labels": [],
            "save_document_after_setting_sensitivity_label": False,
            "libreoffice_path": "",
            "default_source_company": "",
            "excerpt_naming_scheme": "source_clauses_timestamp",
            "cr_naming_scheme": "cr_standard",
        }

    @staticmethod
    def format_label_display(lbl: dict, include_guid: bool = False) -> str:
        """Format a sensitivity label for UI selectors with clear encryption badges."""
        name = (lbl.get("name") or "Unknown").strip()
        display_name = (lbl.get("display_name") or "").strip()
        lid = (lbl.get("id") or "").strip()
        is_encrypted = bool(lbl.get("is_encrypted", False))

        prefix = "🔒 " if is_encrypted else ""
        if display_name and display_name.lower() != name.lower():
            base = f"{name} — {display_name}"
        else:
            base = name
        suffix = " [Encrypted]" if is_encrypted else ""
        guid_part = ""
        if include_guid and lid:
            short_guid = f"...{lid[-8:]}" if len(lid) > 8 else lid
            guid_part = f" ({short_guid})"
        return f"{prefix}{base}{suffix}{guid_part}"

    @staticmethod
    def sync_from_windows_policy(
        cfg: dict,
        overwrite_active: bool = False,
        clp_dir: Optional[str] = None,
    ) -> Optional[dict]:
        """
        Discover sensitivity labels from the local Windows Office CLP policy and merge them into cfg.
        Preserves any user-created custom labels. Returns the parsed policy dict or None if unavailable.
        """
        from modules.word_tools.core.sensitivity_label import discover_windows_sensitivity_policy

        policy = discover_windows_sensitivity_policy(clp_dir=clp_dir)
        existing_labels = [l for l in (cfg.get("sensitivity_labels") or []) if isinstance(l, dict) and l.get("id")]

        if policy and policy.get("labels"):
            policy_labels = [dict(l) for l in policy["labels"]]
            policy_ids_lower = {l["id"].lower() for l in policy_labels}

            # Preserve custom labels (or non-legacy manually added labels) that aren't in the policy
            preserved_custom: List[dict] = []
            for old in existing_labels:
                old_id = (old.get("id") or "").strip()
                old_id_lower = old_id.lower()
                if old_id_lower in policy_ids_lower:
                    continue
                src = old.get("source", "")
                if src == "custom" or (not src and old_id_lower not in _LEGACY_SEED_LABEL_IDS):
                    item = dict(old)
                    item.setdefault("source", "custom")
                    preserved_custom.append(item)

            merged = policy_labels + preserved_custom
            cfg["sensitivity_labels"] = merged
            cfg["sensitivity_tenant_id"] = policy.get("tenant_id", "")
            cfg["sensitivity_policy_upn"] = policy.get("upn", "")
            cfg["sensitivity_policy_initialized"] = True

            valid_ids_lower = {l["id"].lower() for l in merged}
            cur_active = (cfg.get("active_sensitivity_label_id") or cfg.get("sensitivity_level_label_id") or "").strip()
            if overwrite_active or not cur_active or cur_active.lower() not in valid_ids_lower:
                cfg["active_sensitivity_label_id"] = policy.get("recommended_default_id") or merged[0]["id"]
            return policy

        # No Windows policy found
        cfg["sensitivity_policy_initialized"] = True
        if not existing_labels:
            cfg["sensitivity_labels"] = []
            cfg["active_sensitivity_label_id"] = ""
            cfg["sensitivity_level_label_id"] = ""
            cfg["sensitivity_level_label_name"] = ""
        return None

    @staticmethod
    def _normalize_active_label(cfg: dict) -> None:
        """Ensure active_sensitivity_label_id and legacy keys match an existing label (or stay empty)."""
        labels = [l for l in (cfg.get("sensitivity_labels") or []) if isinstance(l, dict) and l.get("id")]
        cfg["sensitivity_labels"] = labels

        if not labels:
            cfg["active_sensitivity_label_id"] = ""
            cfg["sensitivity_level_label_id"] = ""
            cfg["sensitivity_level_label_name"] = ""
            return

        active_id = (cfg.get("active_sensitivity_label_id") or cfg.get("sensitivity_level_label_id") or "").strip()
        matched = None
        if active_id:
            for l in labels:
                if (l.get("id") or "").strip().lower() == active_id.lower():
                    matched = l
                    break

        if matched is None:
            # Prefer lowest-order unencrypted label if available, else first label
            unencrypted = [l for l in labels if not l.get("is_encrypted")]
            matched = unencrypted[0] if unencrypted else labels[0]

        cfg["active_sensitivity_label_id"] = matched.get("id", "")
        cfg["sensitivity_level_label_id"] = matched.get("id", "")
        cfg["sensitivity_level_label_name"] = matched.get("name", "")

    @staticmethod
    def load() -> dict:
        default = WordConfig._default_payload()
        file_existed = CONFIG_PATH.exists()
        needs_save = not file_existed

        try:
            if file_existed:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        default.update(data)
        except Exception as e:
            logging.error(f"Failed to load Word config: {e}")

        # Auto-discover from Windows CLP policy on first run or legacy migration
        if not default.get("sensitivity_policy_initialized", False):
            WordConfig.sync_from_windows_policy(default, overwrite_active=False)
            needs_save = True

        WordConfig._normalize_active_label(default)

        if needs_save:
            WordConfig.save(default)

        return default

    @staticmethod
    def save(data: dict):
        try:
            WordConfig._normalize_active_label(data)
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logging.error(f"Failed to save Word config: {e}")

    @staticmethod
    def get_sensitivity_labels() -> List[dict]:
        return WordConfig.load().get("sensitivity_labels", [])

    @staticmethod
    def get_sensitivity_presets() -> List[dict]:
        """Alias for get_sensitivity_labels() for external/MCP caller compatibility."""
        return WordConfig.get_sensitivity_labels()

    @staticmethod
    def get_active_sensitivity_label() -> dict:
        cfg = WordConfig.load()
        labels = cfg.get("sensitivity_labels", [])
        active_id = (cfg.get("active_sensitivity_label_id") or "").strip().lower()
        for l in labels:
            if isinstance(l, dict) and (l.get("id") or "").strip().lower() == active_id:
                return l
        if labels:
            return labels[0]
        return {}

    @staticmethod
    def set_active_sensitivity_label_id(label_id: str):
        cfg = WordConfig.load()
        cfg["active_sensitivity_label_id"] = label_id.strip()
        WordConfig.save(cfg)

    @staticmethod
    def add_or_update_sensitivity_label(
        label_id: str,
        name: str,
        description: str = "",
        display_name: str = "",
        is_encrypted: bool = False,
        source: str = "custom",
    ):
        cfg = WordConfig.load()
        labels = cfg.get("sensitivity_labels", [])
        clean_id = label_id.strip()
        found = False
        for l in labels:
            if isinstance(l, dict) and (l.get("id") or "").strip().lower() == clean_id.lower():
                l["id"] = clean_id
                l["name"] = name.strip()
                l["description"] = description.strip()
                if display_name:
                    l["display_name"] = display_name.strip()
                l["is_encrypted"] = bool(is_encrypted)
                l.setdefault("source", source)
                found = True
                break
        if not found:
            labels.append({
                "id": clean_id,
                "name": name.strip(),
                "display_name": (display_name or name).strip(),
                "description": description.strip(),
                "is_encrypted": bool(is_encrypted),
                "source": source,
            })
        cfg["sensitivity_labels"] = labels
        WordConfig.save(cfg)

    @staticmethod
    def remove_sensitivity_label(label_id: str) -> bool:
        cfg = WordConfig.load()
        labels = cfg.get("sensitivity_labels", [])
        clean_id = label_id.strip().lower()
        new_labels = [
            l for l in labels
            if isinstance(l, dict) and (l.get("id") or "").strip().lower() != clean_id
        ]
        if len(new_labels) == len(labels):
            return False
        cfg["sensitivity_labels"] = new_labels
        if (cfg.get("active_sensitivity_label_id") or "").strip().lower() == clean_id:
            cfg["active_sensitivity_label_id"] = new_labels[0].get("id", "") if new_labels else ""
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

    @staticmethod
    def get_excerpt_naming_scheme() -> str:
        return WordConfig.load().get("excerpt_naming_scheme", "source_clauses_timestamp")

    @staticmethod
    def set_excerpt_naming_scheme(scheme: str):
        cfg = WordConfig.load()
        cfg["excerpt_naming_scheme"] = scheme
        WordConfig.save(cfg)

    @staticmethod
    def get_cr_naming_scheme() -> str:
        return WordConfig.load().get("cr_naming_scheme", "cr_standard")

    @staticmethod
    def set_cr_naming_scheme(scheme: str):
        cfg = WordConfig.load()
        cfg["cr_naming_scheme"] = scheme
        WordConfig.save(cfg)
