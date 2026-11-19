# --- File: src/modules/emails/core/general_email_sync.py ---
import re
import logging
import datetime
from pathlib import Path
from typing import List, Dict, Set
from PyQt5.QtCore import QThread, pyqtSignal

from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.thread_markdown_exporter import (
    inject_html_tables_into_body,
    unwrap_outlook_safelinks,
)
from core.utils.company_sanitizer import CompanySanitizer


class GeneralEmailSyncThread(QThread):
    progress_msg = pyqtSignal(str)
    progress_val = pyqtSignal(int, int)  # current, total
    finished = pyqtSignal(bool, str, int)  # success, message, total_synced

    TDOC_REGEX = re.compile(r'\b([A-Z0-9]{2,4}-\d{6,8})\b', re.IGNORECASE)
    REV_REGEX = re.compile(r'\b([A-Z0-9]{2,4}-\d{6,8})(?:r|rev)\s*0?([1-9]\d*)\b', re.IGNORECASE)

    # Regex identifying the boundary where previous email quotations begin
    QUOTE_SPLIT_REGEX = re.compile(
        r'\n(?:\s*[-_]{5,}\s*|\s*(?:Von|From|De|Gesendet|Sent|Betreff|Subject):\s*.+|\s*On\s+.+wrote:\s*)',
        re.IGNORECASE
    )

    def __init__(self, folder_configs: List[dict], db_path: Path,
                 start_date: str = "", end_date: str = "", days_buffer: int = 3):
        super().__init__()
        self.folder_configs = folder_configs
        self.db = GeneralEmailDatabase(db_path)
        self.start_date = start_date
        self.end_date = end_date
        self.days_buffer = days_buffer

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            filter_start, filter_end = None, None
            if self.start_date and self.end_date:
                s_dt = datetime.datetime.strptime(self.start_date, "%Y-%m-%d")
                e_dt = datetime.datetime.strptime(self.end_date, "%Y-%m-%d")
                filter_start = s_dt - datetime.timedelta(days=self.days_buffer)
                filter_end = e_dt + datetime.timedelta(days=self.days_buffer + 1)

            total_synced = 0
            folder_configs = list(self.folder_configs or [])
            ns = OutlookClient.get_namespace()
            if ns:
                try:
                    sent_path = getattr(ns.GetDefaultFolder(5), "FolderPath", "") or ""
                    known_paths = {str(fc.get("folder_path", "")).strip().lower() for fc in folder_configs}
                    if sent_path and sent_path.lower() not in known_paths:
                        folder_configs.append({"folder_path": sent_path, "tag": "Sent"})
                except Exception:
                    pass

            for folder_cfg in folder_configs:
                if self.isInterruptionRequested():
                    break
                folder_path = folder_cfg.get("folder_path", "")
                folder_tag = folder_cfg.get("tag", "General")
                if not folder_path:
                    continue

                self.progress_msg.emit(f"Connecting to: {folder_path} [{folder_tag}]...")
                folder = OutlookClient.get_folder_by_path(folder_path)
                if not folder:
                    logging.warning(f"Could not open Outlook folder: {folder_path}")
                    continue

                items = folder.Items
                total_items = len(items)
                items.Sort("[ReceivedTime]", True)

                self.progress_msg.emit(f"Scanning {total_items} items in [{folder_tag}]...")

                batch_emails = []
                batch_matches = []

                for i in range(1, total_items + 1):
                    if self.isInterruptionRequested():
                        break
                    if i % 15 == 0:
                        self.progress_val.emit(i, total_items)

                    mail_item = items.Item(i)
                    if mail_item.Class != 43:
                        continue

                    # Date Boundary Filter
                    if filter_start and filter_end:
                        mail_date = getattr(mail_item, "ReceivedTime", None) or getattr(mail_item, "SentOn", None)
                        if mail_date:
                            try:
                                dt = datetime.datetime(mail_date.year, mail_date.month, mail_date.day,
                                                       mail_date.hour, mail_date.minute, mail_date.second)
                                if dt > filter_end:
                                    continue
                                if dt < filter_start:
                                    break
                            except Exception:
                                pass

                    entry_id = getattr(mail_item, "EntryID", "")
                    if not entry_id:
                        continue

                    subject = getattr(mail_item, "Subject", "")
                    body = getattr(mail_item, "Body", "")
                    sender_name = getattr(mail_item, "SenderName", "")
                    try:
                        sender_email = getattr(mail_item, "SenderEmailAddress", "")
                    except Exception:
                        sender_email = ""

                    # Resolve sender and internal Exchange addresses
                    sender_name, sender_email = self._resolve_sender(mail_item, sender_name, sender_email, body)
                    company = self._resolve_company(sender_name, sender_email)

                    # Extract matches separating Direct Body from Quoted Reply content
                    matches = self._extract_tdocs(subject, body)
                    if not matches:
                        continue

                    rich_body = body
                    try:
                        html_body = getattr(mail_item, "HTMLBody", "") or ""
                        if html_body and "<table" in html_body.lower():
                            rich_body = inject_html_tables_into_body(rich_body, html_body)
                    except Exception:
                        pass
                    rich_body = unwrap_outlook_safelinks(rich_body)
                    quote_split = self.QUOTE_SPLIT_REGEX.search(rich_body)
                    clean_body = (rich_body[: quote_split.start()].strip() if quote_split else rich_body.strip())

                    attachments: List[str] = []
                    try:
                        atts = getattr(mail_item, "Attachments", None)
                        if atts:
                            for ai in range(1, atts.Count + 1):
                                fname = getattr(atts.Item(ai), "FileName", "") or ""
                                if fname:
                                    attachments.append(fname)
                    except Exception:
                        pass

                    dt_raw = getattr(mail_item, "ReceivedTime", None) or getattr(mail_item, "SentOn", None) or ""
                    date_str = dt_raw.strftime("%Y-%m-%d %H:%M:%S") if hasattr(dt_raw, "strftime") else str(dt_raw)

                    email_record = {
                        "id": entry_id,
                        "folder_tag": folder_tag,
                        "subject": subject,
                        "sender_name": sender_name,
                        "sender_email": sender_email,
                        "company": company,
                        "date_received": date_str,
                        "body_text": rich_body[:12000],
                        "clean_body": clean_body[:12000],
                        "attachments": attachments,
                        "is_read": 1 if getattr(mail_item, "UnRead", False) is False else 0,
                        "is_ignored": 0,
                        "flag_status": 2 if int(getattr(mail_item, "FlagStatus", 0) or 0) == 2 else 0,
                    }
                    batch_emails.append(email_record)

                    for m in matches:
                        batch_matches.append({
                            "email_id": entry_id,
                            "tdoc_id": m["tdoc_id"],
                            "rev_matched": m["rev_matched"],
                            "match_location": m["location"]
                        })

                    total_synced += 1

                    if len(batch_emails) >= 40:
                        self.db.save_emails_batch(batch_emails, batch_matches)
                        batch_emails.clear()
                        batch_matches.clear()

                if batch_emails:
                    self.db.save_emails_batch(batch_emails, batch_matches)

            self.finished.emit(True, f"Successfully synced {total_synced} related email references.", total_synced)
        except Exception as e:
            logging.error(f"General email sync failed: {e}", exc_info=True)
            self.finished.emit(False, str(e), 0)
        finally:
            pythoncom.CoUninitialize()

    def _extract_tdocs(self, subject: str, body: str) -> List[dict]:
        results = {}

        # 1. Subject extraction (Highest Priority)
        for tdoc in self.TDOC_REGEX.findall(subject):
            tdoc_up = tdoc.upper()
            results[tdoc_up] = {"tdoc_id": tdoc_up, "rev_matched": "", "location": "Subject"}

        for base, rev in self.REV_REGEX.findall(subject):
            base_up = base.upper()
            rev_str = f"{base_up}r{int(rev):02d}"
            results[base_up] = {"tdoc_id": base_up, "rev_matched": rev_str, "location": "Subject"}

        # 2. Split Body into Direct Text vs Quoted History
        split_match = self.QUOTE_SPLIT_REGEX.search(body)
        if split_match:
            direct_body = body[:split_match.start()]
            quoted_body = body[split_match.start():]
        else:
            direct_body = body
            quoted_body = ""

        # Direct Body extraction
        for tdoc in self.TDOC_REGEX.findall(direct_body):
            tdoc_up = tdoc.upper()
            if tdoc_up not in results:
                results[tdoc_up] = {"tdoc_id": tdoc_up, "rev_matched": "", "location": "Body"}

        for base, rev in self.REV_REGEX.findall(direct_body):
            base_up = base.upper()
            rev_str = f"{base_up}r{int(rev):02d}"
            if base_up not in results or results[base_up]["location"] != "Subject":
                results[base_up] = {"tdoc_id": base_up, "rev_matched": rev_str, "location": "Body"}

        # Quoted History extraction
        if quoted_body:
            for tdoc in self.TDOC_REGEX.findall(quoted_body):
                tdoc_up = tdoc.upper()
                if tdoc_up not in results:
                    results[tdoc_up] = {"tdoc_id": tdoc_up, "rev_matched": "", "location": "Quoted"}

            for base, rev in self.REV_REGEX.findall(quoted_body):
                base_up = base.upper()
                rev_str = f"{base_up}r{int(rev):02d}"
                if base_up not in results:
                    results[base_up] = {"tdoc_id": base_up, "rev_matched": rev_str, "location": "Quoted"}

        return list(results.values())

    @staticmethod
    def _resolve_sender(mail_item, sender_name: str, sender_email: str, body: str):
        s_name_lower = sender_name.lower()
        s_email_lower = sender_email.lower()
        is_list = "list.etsi.org" in s_email_lower or "dmarc" in s_name_lower or "on behalf of" in s_name_lower

        if is_list:
            try:
                for rep in mail_item.ReplyRecipients:
                    sender_email = getattr(rep, "Address", sender_email)
                    if sender_email.lower().startswith("/o="):
                        ae = getattr(rep, "AddressEntry", None)
                        if ae and ae.GetExchangeUser():
                            sender_email = ae.GetExchangeUser().PrimarySmtpAddress or sender_email
                    break
                reply_names = getattr(mail_item, "ReplyRecipientNames", "")
                if reply_names:
                    sender_name = reply_names.split(';')[0].strip()
            except Exception:
                pass
        else:
            try:
                sender_email_type = str(getattr(mail_item, "SenderEmailType", "")).upper()
                if sender_email_type == "EX" or sender_email.lower().startswith("/o="):
                    sender_obj = getattr(mail_item, "Sender", None)
                    if sender_obj:
                        eu = sender_obj.GetExchangeUser()
                        if eu:
                            smtp = getattr(eu, "PrimarySmtpAddress", "")
                            if smtp:
                                sender_email = smtp
            except Exception:
                pass

        if any(k in sender_name.lower() for k in ["3gpp", "list", "on behalf of", "dmarc"]) or not sender_email:
            dmarc_match = re.search(
                r'From:\s*([^\n<\[]+?)\s*[<\[](?:mailto:)?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})[>\]]',
                body[:1500], re.IGNORECASE)
            if dmarc_match:
                sender_name = dmarc_match.group(1).strip(' \t"\'')
                if not sender_email or "list.etsi.org" in sender_email.lower() or sender_email.lower().startswith(
                        "/o="):
                    sender_email = dmarc_match.group(2).strip()

        return sender_name, sender_email

    @staticmethod
    def _resolve_company(sender_name: str, sender_email: str) -> str:
        raw_sender = f"{sender_name} <{sender_email}>"
        comps = CompanySanitizer.get_matching_contributors(raw_sender)
        if comps:
            return comps[0]
        if sender_email and "@" in sender_email:
            domain = sender_email.split("@")[-1].split(".")[0]
            if domain.lower() not in ["gmail", "yahoo", "hotmail", "outlook"]:
                return domain.title()
        return "Unknown"