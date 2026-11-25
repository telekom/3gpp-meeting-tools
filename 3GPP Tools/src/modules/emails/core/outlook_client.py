# --- File: modules/emails/core/outlook_client.py ---
import platform
import traceback
import logging
import os
from pathlib import Path
from typing import Any, Optional, Sequence, Tuple, Union

if platform.system() == 'Windows':
    import win32com.client


class OutlookClient:
    @staticmethod
    def get_namespace():
        if platform.system() != 'Windows':
            logging.error("Outlook integration is only supported on Windows.")
            return None
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            return outlook.GetNamespace("MAPI")
        except Exception as e:
            logging.error(f"Could not retrieve Outlook instance: {e}")
            return None

    @staticmethod
    def get_folder_by_path(folder_path: str, create_if_missing: bool = False):
        """
        Navigates the Outlook folder tree using a string path.
        If create_if_missing is True, it will dynamically create missing folders in the path.
        """
        namespace = OutlookClient.get_namespace()
        if not namespace or not folder_path:
            return None

        parts = [p for p in folder_path.replace('\\', '/').split('/') if p]

        try:
            # Try to find the root folder (usually the email address)
            current_folder = None
            for folder in namespace.Folders:
                if folder.Name.lower() == parts[0].lower():
                    current_folder = folder
                    break

            # If root wasn't found, assume the first part is a subfolder of the default Inbox
            if not current_folder:
                current_folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            else:
                parts.pop(0)

            # Traverse the rest of the path
            for part in parts:
                found = False
                for subfolder in current_folder.Folders:
                    if subfolder.Name.lower() == part.lower():
                        current_folder = subfolder
                        found = True
                        break

                # ---> THE FIX: Automatically create the folder if it doesn't exist!
                if not found:
                    if create_if_missing:
                        current_folder = current_folder.Folders.Add(part)
                    else:
                        logging.error(f"Outlook folder not found: {part}")
                        return None

            return current_folder
        except Exception as e:
            logging.error(f"Failed to resolve Outlook folder path '{folder_path}': {e}")
            return None

    @staticmethod
    def folder_unique_key(folder, fallback_path: str = "") -> str:
        """Returns a normalized unique key for an Outlook MAPIFolder to prevent duplicate scans."""
        try:
            fpath = str(getattr(folder, "FolderPath", "") or "").strip()
            if fpath and not fpath.startswith("<MagicMock"):
                return fpath.replace("\\", "/").strip("/").lower()
        except Exception:
            pass
        try:
            eid = str(getattr(folder, "EntryID", "") or "").strip()
            if eid and not eid.startswith("<MagicMock"):
                return f"eid:{eid}"
        except Exception:
            pass
        if fallback_path:
            return fallback_path.replace("\\", "/").strip("/").lower()
        return f"obj:{id(folder)}"

    @staticmethod
    def iter_folder_tree(
        folder,
        *,
        include_subfolders: bool = True,
        max_depth: int = 4,
        visited_keys: Optional[set] = None,
        fallback_path: str = "",
    ) -> list:
        """
        Returns `[folder, ...subfolders]` up to `max_depth` when `include_subfolders=True`,
        deduplicating against `visited_keys` so nested configured folders are never scanned twice.
        """
        if folder is None:
            return []
        seen = visited_keys if visited_keys is not None else set()
        result = []

        def _walk(cur_folder, depth: int, cur_path: str):
            if cur_folder is None:
                return
            key = OutlookClient.folder_unique_key(cur_folder, fallback_path=cur_path)
            if key not in seen:
                seen.add(key)
                result.append(cur_folder)
            if not include_subfolders or depth <= 0:
                return
            try:
                subfolders = getattr(cur_folder, "Folders", None)
                if not subfolders:
                    return
                for idx, sub in enumerate(subfolders):
                    if sub is None or sub is cur_folder:
                        continue
                    try:
                        sub_name = str(getattr(sub, "Name", "") or f"sub_{idx}").strip()
                    except Exception:
                        sub_name = f"sub_{idx}"
                    sub_path = f"{cur_path}/{sub_name}".strip("/") if cur_path else sub_name
                    _walk(sub, depth - 1, sub_path)
            except Exception:
                pass

        _walk(folder, max_depth, fallback_path)
        return result

    @staticmethod
    def save_email_to_disk(mail_item, tdoc_id: str, meeting_dir: Path) -> str:
        """
        Saves the physical .msg file to [Meeting Dir]/[TDoc]/email approval/
        Returns the absolute path to the saved file.
        """
        try:
            # Create the strict directory structure you requested
            target_dir = meeting_dir / tdoc_id / "email approval"
            target_dir.mkdir(parents=True, exist_ok=True)

            # Clean the subject to make it a valid Windows filename
            clean_subject = "".join(c for c in mail_item.Subject if c.isalnum() or c in " -_").strip()
            if len(clean_subject) > 100:
                clean_subject = clean_subject[:100]

            file_name = f"{clean_subject}.msg"
            file_path = target_dir / file_name

            # Prevent overwriting if multiple emails have the exact same subject
            counter = 1
            while file_path.exists():
                file_path = target_dir / f"{clean_subject}_{counter}.msg"
                counter += 1

            # 3 = olMSG (Save as Outlook Message Format)
            mail_item.SaveAs(str(file_path.absolute()), 3)
            return str(file_path.absolute())
        except Exception as e:
            logging.error(f"Failed to save .msg file for {tdoc_id}: {e}")
            return ""

    @staticmethod
    def move_email(mail_item, target_folder):
        """Moves an email to a specific Outlook folder object."""
        try:
            mail_item.Move(target_folder)
            return True
        except Exception as e:
            logging.error(f"Failed to move email: {e}")
            return False

    @staticmethod
    def move_email_to_target(entry_id: str, base_target_path: str, ai_folder_name: str) -> str:
        """Moves a specific email to [Target Folder]/[AI]. Returns 'SUCCESS', 'DELETED', or 'ERROR'."""
        namespace = OutlookClient.get_namespace()
        if not namespace or not base_target_path: return "ERROR"

        try:
            # 1. Fetch the exact email
            try:
                mail_item = namespace.GetItemFromID(entry_id)
            except Exception as e:
                # ---> CHANGED: Made this a friendly INFO log instead of a WARNING
                # Also truncated the massive 150-character EntryID so it doesn't spam your console
                short_id = entry_id[-10:] if entry_id else "Unknown"
                logging.info(f"Ghost email detected (ID ends in ...{short_id}). Triggering auto-cleanup.")
                return "DELETED"

            # 2. Fetch the Base Target Folder
            base_folder = OutlookClient.get_folder_by_path(base_target_path, create_if_missing=True)
            if not base_folder: return "ERROR"

            # 3. Clean the AI name
            clean_ai = "".join(c for c in ai_folder_name if c.isalnum() or c in " ._-").strip() or "General"

            # 4. Find or Create the AI Subfolder
            target_folder = None
            for folder in base_folder.Folders:
                if folder.Name.lower() == clean_ai.lower():
                    target_folder = folder
                    break

            if not target_folder:
                target_folder = base_folder.Folders.Add(clean_ai)

            # Prevent Outlook from crashing if the item is already there
            if getattr(mail_item.Parent, "FolderPath", "") == target_folder.FolderPath:
                return "SUCCESS"

            # 5. Execute Move
            mail_item.Move(target_folder)
            return "SUCCESS"

        except Exception as e:
            error_msg = str(e)
            if "Cannot move the items" in error_msg or "-2147352567" in error_msg:
                logging.error(
                    f"Outlook locked the file and refused to move it (Is the email currently open?): {entry_id}")
            else:
                logging.error(f"Explicit move failed for {entry_id}: {e}")
            return "ERROR"

    @staticmethod
    def create_email_draft(
        subject: str,
        attachments: Optional[Sequence[Union[str, Path]]] = None,
    ) -> Tuple[bool, str]:
        """
        Opens a new Classic Outlook compose window with ``subject`` and optional
        file attachments via COM automation. Returns ``(success, error_message)``.
        """
        if platform.system() != 'Windows':
            return False, "Outlook COM automation is only supported on Windows."

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.Subject = str(subject or "")

            for item in attachments or []:
                abs_path = os.path.abspath(str(item))
                if abs_path.startswith("\\\\?\\"):
                    abs_path = abs_path[4:]
                if os.path.isfile(abs_path):
                    mail.Attachments.Add(abs_path)

            mail.Display(False)
            return True, ""
        except Exception as e:
            logging.warning(f"Could not open Outlook draft via COM: {e}")
            return False, str(e)

    @staticmethod
    def _update_messages_property(entry_ids: Sequence[str], attr_name: str, target_value: Any, cast_fn) -> int:
        if platform.system() != 'Windows' or not entry_ids:
            return 0
        namespace = OutlookClient.get_namespace()
        if not namespace:
            return 0
        updated = 0
        for raw_id in entry_ids:
            eid = str(raw_id or "").strip()
            if not eid or eid.startswith("MSG-"):
                continue
            try:
                item = namespace.GetItemFromID(eid)
                if item is not None:
                    current = getattr(item, attr_name, None)
                    if current is None or cast_fn(current) != target_value:
                        setattr(item, attr_name, target_value)
                        item.Save()
                        updated += 1
            except Exception as e:
                logging.debug(f"Could not update Outlook {attr_name} for {eid}: {e}")
        return updated

    @staticmethod
    def set_messages_read_status(entry_ids: Sequence[str], is_read: bool = True) -> int:
        """
        Synchronizes read/unread status to Microsoft Outlook for the given EntryIDs.
        Returns the number of Outlook items whose UnRead state was updated.
        """
        return OutlookClient._update_messages_property(entry_ids, "UnRead", not bool(is_read), bool)

    @staticmethod
    def set_messages_flag_status(entry_ids: Sequence[str], is_flagged: bool = True) -> int:
        """
        Synchronizes 2-state Outlook flag status (olFlagMarked=2, olNoFlag=0) for EntryIDs.
        Returns the number of Outlook items whose FlagStatus was updated.
        """
        return OutlookClient._update_messages_property(entry_ids, "FlagStatus", 2 if is_flagged else 0, int)


PR_FLAG_STATUS_PROP = "http://schemas.microsoft.com/mapi/proptag/0x10900003"


def safe_mapi_row_val(row, *keys: str, default=None):
    """Safely fetches the first available non-None column from an Outlook Table row."""
    for key in keys:
        try:
            val = row(key)
            if val is not None:
                return val
        except Exception:
            continue
    return default


def add_mapi_table_columns(tbl, col_names: Sequence[str]) -> None:
    """Safely adds requested MAPI properties/columns to an Outlook Table without aborting on unsupported keys."""
    cols = getattr(tbl, "Columns", None)
    if cols is None:
        return
    for col_name in col_names:
        try:
            cols.Add(col_name)
        except Exception:
            pass


def extract_table_read_flag(row) -> Optional[Tuple[int, int]]:
    """Extracts (is_read, flag_status) directly from an Outlook Table row when UnRead is exposed."""
    unread_val = safe_mapi_row_val(row, "UnRead")
    if unread_val is None:
        return None
    flag_val = safe_mapi_row_val(row, PR_FLAG_STATUS_PROP, "FlagStatus", default=0)
    try:
        flag_status = 2 if int(flag_val or 0) == 2 else 0
    except Exception:
        flag_status = 0
    try:
        is_read = 1 if (unread_val is False or int(unread_val or 0) == 0) else 0
    except Exception:
        is_read = 1 if not bool(unread_val) else 0
    return is_read, flag_status


def extract_item_read_flag(ns, entry_id: str) -> Optional[Tuple[int, int]]:
    """Lightweight Outlook GetItemFromID state read (UnRead + FlagStatus only; never touches Body/HTMLBody)."""
    if not ns or not entry_id:
        return None
    try:
        item = ns.GetItemFromID(entry_id)
        if item is None:
            return None
        is_read = 1 if getattr(item, "UnRead", False) is False else 0
        try:
            flag_status = 2 if int(getattr(item, "FlagStatus", 0) or 0) == 2 else 0
        except Exception:
            flag_status = 0
        return is_read, flag_status
    except Exception:
        return None