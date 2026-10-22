# --- File: src/modules/work_items/core/wi_settings.py ---
import json
import logging
from typing import Set, Tuple

import core.utils.paths

logger = logging.getLogger(__name__)


class WorkItemsSettings:
    """Handles saving and loading configurations specific to the Work Items tab."""

    def __init__(self):
        # Save to the root directory next to meetings_config.json
        self.config_file = core.utils.paths.get_project_root() / "config" / "work_items_config.json"
        self.config_file.parent.mkdir(parents=True, exist_ok=True)

    def _read_data(self) -> dict:
        """Helper to read the raw dictionary from config_file safely."""
        if not self.config_file.exists():
            return {}
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except json.JSONDecodeError as e:
            logger.warning(f"[WorkItemsSettings] Invalid JSON in {self.config_file}; resetting: {e}")
            return {}
        except Exception as e:
            logger.error(f"[WorkItemsSettings] Error reading {self.config_file}: {e}")
            return {}

    def _write_data(self, data: dict):
        """Helper to write the raw dictionary to config_file safely."""
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            logger.error(f"[WorkItemsSettings] Error saving WI config to {self.config_file}: {e}")

    def save_filters(self, filters: dict):
        """Saves the active UI filters to the JSON configuration file."""
        data = self._read_data()
        data["filters"] = filters
        self._write_data(data)

    def get_filters(self) -> dict:
        """Retrieves the saved filters from the JSON configuration file."""
        data = self._read_data()
        return data.get("filters", {})

    def get_tracking_sets(self) -> Tuple[Set[str], Set[str]]:
        """Returns (starred_set, following_set) containing normalized WI code strings."""
        data = self._read_data()
        raw_starred = data.get("starred", [])
        raw_following = data.get("following", [])

        starred = {str(x).strip() for x in raw_starred if str(x).strip()}
        following = {str(x).strip() for x in raw_following if str(x).strip()}
        return starred, following

    def save_tracking_sets(self, starred: Set[str], following: Set[str]):
        """Persists starred and following WI code sets without overwriting filters."""
        data = self._read_data()
        data["starred"] = sorted(str(x).strip() for x in starred if str(x).strip())
        data["following"] = sorted(str(x).strip() for x in following if str(x).strip())
        self._write_data(data)
