# --- File: modules/emails/plugin_loader.py ---
import logging

def register_emails_plugin():
    """
    Registers the Email Manager module.
    In the future, we can hook into the global QueueManager here
    for background Outlook syncing.
    """
    logging.info("🔌 Plugin Loaded: eMeeting Email Manager")