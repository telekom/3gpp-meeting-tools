# --- File: src/modules/meetings/core/url_router.py ---
import re
from core.network.network_state import NetworkState


class URLRouter:
    """
    Evaluates the current network state, meeting status, and target document type
    to generate an optimal, ordered priority list of candidate URLs for download threads.
    """

    @staticmethod
    def _get_local_server_base(wg_name: str) -> str:
        """Resolves the flattened directory structure used by the 10.10.10.10 local server."""
        wg_upper = wg_name.upper()
        if wg_upper == "SA3LI":
            return "http://10.10.10.10/ftp/SA3LI"
        elif wg_upper.startswith("SA"):
            return f"http://10.10.10.10/ftp/SA/{wg_upper}"
        elif wg_upper.startswith("RAN"):
            return f"http://10.10.10.10/ftp/RAN/{wg_upper}"
        elif wg_upper.startswith("CT"):
            return f"http://10.10.10.10/ftp/CT/{wg_upper}"

        return f"http://10.10.10.10/ftp/{wg_upper}"

    @staticmethod
    def _get_subfolder_order(wg_name: str, target_filename: str) -> list:
        """
        Determines the optimal folder search order based on whether the document
        is a revision or a base contribution.
        """
        wg_upper = wg_name.upper()
        is_revision = bool(re.search(r'(?:r|rev)\d{1,2}[a-zA-Z]?$', target_filename.strip(), re.IGNORECASE))

        if is_revision:
            # Revisions live in Inbox/Revisions first
            order = []
            if wg_upper == "SA2":
                order.append("Inbox/Revisions")
            order.extend(["Inbox", "Docs"])
            return order
        else:
            # Base documents live in Docs first
            order = ["Docs", "Inbox"]
            if wg_upper == "SA2":
                order.append("Inbox/Revisions")
            return order

    @staticmethod
    def build_priority_url_list(wg_name: str, folder_name: str, main_ftp_url: str, is_active_sync: bool,
                                target_filename: str = "") -> list:
        """
        Builds the fallback ordered list of base folder URLs to search for a TDoc.

        Order of Tiers:
        1. Local Server (10.10.10.10) - Flattened structure, zero external latency.
        2. Live Meeting SYNC Folder (Public Web) - For active meeting updates.
        3. Standard Web Archive - Long-term archive path.
        """
        urls = []
        wg_upper = wg_name.upper()
        main_ftp_clean = main_ftp_url.rstrip('/') if main_ftp_url else ""
        subfolders = URLRouter._get_subfolder_order(wg_name, target_filename)

        # -----------------------------------------------------
        # TIER 1: Local On-Site Server (10.10.10.10)
        # -----------------------------------------------------
        if NetworkState.get_instance().is_local_active():
            local_base = URLRouter._get_local_server_base(wg_name)
            for folder in subfolders:
                urls.append(f"{local_base}/{folder}")

        # -----------------------------------------------------
        # TIER 2: Live Meeting SYNC Folder (Public Web)
        # -----------------------------------------------------
        if is_active_sync:
            sync_wg = "SA3LI" if wg_upper == "SA3LI" else wg_upper
            sync_base = f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}"
            for folder in subfolders:
                urls.append(f"{sync_base}/{folder}")

        # -----------------------------------------------------
        # TIER 3: Standard Web Archive (Fallback)
        # -----------------------------------------------------
        if main_ftp_clean:
            for folder in subfolders:
                # Standard web archive for SA2 uses uppercase INBOX/Revisions
                folder_str = "INBOX/Revisions" if folder == "Inbox/Revisions" else folder
                urls.append(f"{main_ftp_clean}/{folder_str}")

        # Remove any potential duplicates while preserving priority order
        seen = set()
        return [x for x in urls if not (x in seen or seen.add(x))]