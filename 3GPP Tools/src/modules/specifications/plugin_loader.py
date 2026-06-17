# --- File: modules/specs_db/plugin_loader.py ---
from core.queue_manager import register_task
from modules.specifications.core.scraper import SpecsCrawlerThread

def register_specs_plugin():
    register_task(
        target_format="update_specs_db",
        display_name="UPDATE 3GPP DB",
        thread_factory=lambda file_path, params, ctx: SpecsCrawlerThread(
            db_path=params.get('db_path'),
            force_metadata_update=params.get('force_metadata', False)
        )
    )