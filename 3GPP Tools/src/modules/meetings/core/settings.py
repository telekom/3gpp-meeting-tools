# --- File: modules/meetings/core/settings.py ---
import json
from pathlib import Path
import core.utils.paths

class MeetingsSettings:
    def __init__(self):
        self.config_file = core.utils.paths.get_project_root() / "meetings_config.json"
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        self.cache_dir = self._load_settings()

    def _load_settings(self) -> str:
        fallback = str(Path.home() / "3GPP_Delegate_Helper" / "cache")
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get('download_dir', fallback)
            except Exception:
                pass
        return fallback

    def save_settings(self, download_dir: str):
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['download_dir'] = download_dir

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            self.cache_dir = download_dir
        except Exception as e:
            print(f"Error saving config: {e}")

    def save_last_meeting(self, mtg_info: dict):
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['last_mtg_id'] = mtg_info.get("mtg_id")
            data['last_mtg_number'] = mtg_info.get("meeting_number")

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving last meeting state: {e}")

    def get_last_meeting(self) -> tuple:
        """Returns (last_id, last_num). Returns (None, None) if missing."""
        if not self.config_file.exists():
            return None, None
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get("last_mtg_id"), data.get("last_mtg_number")
        except Exception:
            return None, None

    def save_filters(self, filters: dict):
        try:
            data = {}
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

            data['filters'] = filters

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving filters: {e}")

    def get_filters(self) -> dict:
        if not self.config_file.exists():
            return {}
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get("filters", {})
        except Exception:
            return {}