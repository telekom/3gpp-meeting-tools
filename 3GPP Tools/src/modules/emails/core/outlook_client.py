# --- File: modules/emails/core/outlook_client.py ---
import platform
import traceback
import logging
import os
from pathlib import Path

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