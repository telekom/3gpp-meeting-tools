# --- File: src/modules/meetings/core/url_router.py ---
from core.network.network_state import NetworkState


class URLRouter:
    """
    Evaluates the current network state and meeting status to generate an ordered
    priority list of URLs (Local -> SYNC -> Standard) for download threads to attempt.
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
    def build_priority_url_list(wg_name: str, folder_name: str, main_ftp_url: str, is_active_sync: bool) -> list:
        """
        Builds the fallback ordered list of base folder URLs to search for a TDoc.
        """
        urls = []
        wg_upper = wg_name.upper()
        main_ftp_clean = main_ftp_url.rstrip('/') if main_ftp_url else ""
        folder_clean = folder_name.strip('/') if folder_name else ""

        # -----------------------------------------------------
        # TIER 1: The Local Server (10.10.10.10)
        # -----------------------------------------------------
        if NetworkState.get_instance().is_local_active():
            local_base = f"{URLRouter._get_local_server_base(wg_name)}/{folder_clean}"
            if wg_upper == "SA2":
                urls.append(f"{local_base}/Inbox/Revisions")
            urls.append(f"{local_base}/Inbox")
            urls.append(f"{local_base}/Docs")

        # -----------------------------------------------------
        # TIER 2: The Live Meeting SYNC Folder
        # -----------------------------------------------------
        if is_active_sync:
            sync_wg = "SA3LI" if wg_upper == "SA3LI" else wg_upper
            sync_base = f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}"
            if wg_upper == "SA2":
                urls.append(f"{sync_base}/Inbox/Revisions")
            urls.append(f"{sync_base}/Inbox")
            urls.append(f"{sync_base}/Docs")

        # -----------------------------------------------------
        # TIER 3: The Standard Web Archive (Fallback)
        # -----------------------------------------------------
        if main_ftp_clean:
            if wg_upper == "SA2":
                urls.append(f"{main_ftp_clean}/INBOX/Revisions")
            urls.append(f"{main_ftp_clean}/Inbox")
            urls.append(f"{main_ftp_clean}/Docs")

        # Remove any potential duplicates while preserving priority order
        seen = set()
        return [x for x in urls if not (x in seen or seen.add(x))]