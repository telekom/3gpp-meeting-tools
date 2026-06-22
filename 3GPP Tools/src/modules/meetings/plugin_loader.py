# --- File: modules/meetings/plugin_loader.py ---
from pathlib import Path
from core.queue_manager import register_task
from modules.meetings.core.scraper import MeetingsCrawlerThread


def register_meetings_plugin():
    register_task(
        target_format="update_meetings_db",
        display_name="Sync 3GPP Meetings",
        thread_factory=lambda file_path, params, app_context: MeetingsCrawlerThread(
            db_path=params["db_path"],
            target_meetings=params.get("target_meetings", []),
            sync_wg=params.get("sync_wg", True),
            sync_docs=params.get("sync_docs", True),
            sync_dyna=params.get("sync_dyna", True)
        )
    )