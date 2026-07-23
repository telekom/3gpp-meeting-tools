# --- File: src/modules/work_items/core/wi_settings.py ---
import json
from pathlib import Path
import core.utils.paths


class WorkItemsSettings:
    """Handles saving and loading configurations specific to the Work Items tab."""

    def __init__(self):
        # Save to the root directory next to meetings_config.json
        self.config_file = core.utils.paths.get_project_root() / "work_items_config.json"
        self.config_file.parent.mkdir(parents=True, exist_ok=True)

    def save_filters(self, filters: dict):
        """Saves the active UI filters to the JSON configuration file."""
        try:
            data = {}
            if self.config_file.exists():
                try:
                    with open(self.config_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                except json.JSONDecodeError:
                    # Recover gracefully if the file exists but is empty (0 bytes) or corrupted
                    data = {}

            data['filters'] = filters

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            import logging
            logging.error(f"Error saving WI filters: {e}")

    def get_filters(self) -> dict:
        """Retrieves the saved filters from the JSON configuration file."""
        if not self.config_file.exists():
            return {}
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get("filters", {})
        except json.JSONDecodeError:
            return {}
        except Exception as e:
            import logging
            logging.error(f"Error loading WI filters: {e}")
            return {}