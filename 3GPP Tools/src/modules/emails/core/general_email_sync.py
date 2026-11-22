# --- File: src/modules/emails/core/general_email_sync.py ---
import re
import logging
import datetime
from pathlib import Path
from typing import List, Dict, Optional, Set
from PyQt5.QtCore import QThread, pyqtSignal

from modules.emails.core.outlook_client import (
    PR_FLAG_STATUS_PROP,
    OutlookClient,
    add_mapi_table_columns,
    extract_item_read_flag,
    extract_table_read_flag,
    safe_mapi_row_val,
)
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.sync_watermark import (
    apply_mapi_since_restriction,
    local_naive_to_utc,
    utc_to_local_naive,
)
from modules.emails.core.thread_markdown_exporter import (
    inject_html_tables_into_body,
    split_direct_and_quoted_body,
    strip_trailing_quotes,
    unwrap_outlook_safelinks,
)
from core.utils.company_sanitizer import CompanySanitizer


class GeneralEmailSyncThread(QThread):
    progress_msg = pyqtSignal(str)
    progress_val = pyqtSignal(int, int)  # current, total
    finished = pyqtSignal(bool, str, int)  # success, message, total_synced

    TDOC_REGEX = re.compile(
        r'(?<![A-Za-z0-9])((?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,8})(?![A-Za-z0-9])',
        re.IGNORECASE,
    )
    REV_REGEX = re.compile(
        r'(?<![A-Za-z0-9])((?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,8})(?:r|rev)\s*0?([1-9]\d*)[a-z]?(?![A-Za-z0-9])',
        re.IGNORECASE,
    )

    # Regex identifying the boundary where previous email quotations begin
    QUOTE_SPLIT_REGEX = re.compile(
        r'\n(?:\s*[-_]{5,}\s*|\s*(?:Von|From|De|Gesendet|Sent|Betreff|Subject):\s*.+|\s*On\s+.+wrote:\s*)',
        re.IGNORECASE
    )

    def __init__(self, folder_configs: List[dict], db_path: Path,
                 start_date: str = "", end_date: str = "", days_buffer: int = 3,
                 since_dt_utc: Optional[datetime.datetime] = None):
        super().__init__()
        self.folder_configs = folder_configs
        self.db = GeneralEmailDatabase(db_path)
        self.start_date = start_date
        self.end_date = end_date
        self.days_buffer = days_buffer
        self.since_dt_utc = since_dt_utc

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            total_synced = self.sync_folders(
                progress_cb=lambda msg: self.progress_msg.emit(msg),
                progress_val_cb=lambda cur, tot: self.progress_val.emit(cur, tot),
                cancel_check=self.isInterruptionRequested,
            )
            self.finished.emit(True, f"Successfully synced {total_synced} related email references.", total_synced)
        except Exception as e:
            logging.error(f"General email sync failed: {e}", exc_info=True)
            self.finished.emit(False, str(e), 0)
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _to_naive_dt(raw_dt):
        if raw_dt and hasattr(raw_dt, "year"):
            try:
                return datetime.datetime(
                    raw_dt.year, raw_dt.month, raw_dt.day,
                    getattr(raw_dt, "hour", 0), getattr(raw_dt, "minute", 0), getattr(raw_dt, "second", 0),
                )
            except Exception:
                return None
        return None

    def _build_email_record(
        self,
        mail_item,
        folder_tag: str,
        entry_id: str = "",
        require_tdoc: bool = True,
        fallback_tdocs: List[str] = None,
    ):
        """Extracts email record + TDoc matches from an Outlook MailItem, short-circuiting before heavy COM reads."""
        eid = entry_id or getattr(mail_item, "EntryID", "")
        if not eid:
            return None, []
        subject = getattr(mail_item, "Subject", "") or ""
        body = getattr(mail_item, "Body", "") or ""
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

        matches = self._extract_tdocs(subject, body, attachments=attachments)
        if not matches and fallback_tdocs:
            matches = [
                {"tdoc_id": str(t).upper(), "rev_matched": "", "location": "Thread"}
                for t in fallback_tdocs
                if str(t or "").strip()
            ]
        if require_tdoc and not matches:
            return None, []

        sender_name = getattr(mail_item, "SenderName", "") or ""
        try:
            sender_email = getattr(mail_item, "SenderEmailAddress", "") or ""
        except Exception:
            sender_email = ""
        sender_name, sender_email = self._resolve_sender(mail_item, sender_name, sender_email, body)
        company = self._resolve_company(sender_name, sender_email)

        rich_body = body
        try:
            html_body = getattr(mail_item, "HTMLBody", "") or ""
            if html_body and "<table" in html_body.lower():
                rich_body = inject_html_tables_into_body(rich_body, html_body)
        except Exception:
            pass
        rich_body = unwrap_outlook_safelinks(rich_body)
        clean_body = strip_trailing_quotes(rich_body)

        dt_raw = getattr(mail_item, "ReceivedTime", None) or getattr(mail_item, "SentOn", None) or ""
        date_str = dt_raw.strftime("%Y-%m-%d %H:%M:%S") if hasattr(dt_raw, "strftime") else str(dt_raw)

        email_record = {
            "id": eid,
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
        email_matches = [
            {
                "email_id": eid,
                "tdoc_id": m["tdoc_id"],
                "rev_matched": m["rev_matched"],
                "match_location": m["location"],
            }
            for m in matches
        ]
        return email_record, email_matches

    def sync_folders(self, progress_cb=None, progress_val_cb=None, cancel_check=None) -> int:
        filter_start, filter_end = None, None
        if self.start_date and self.end_date:
            s_dt = datetime.datetime.strptime(self.start_date, "%Y-%m-%d")
            e_dt = datetime.datetime.strptime(self.end_date, "%Y-%m-%d")
            filter_start = s_dt - datetime.timedelta(days=self.days_buffer)
            filter_end = e_dt + datetime.timedelta(days=self.days_buffer + 1)

        if self.since_dt_utc is not None:
            since_local = utc_to_local_naive(self.since_dt_utc)
            if filter_start is None or since_local > filter_start:
                filter_start = since_local

        total_synced = 0
        total_flags_updated = 0
        cached_states = self.db.get_cached_email_states()
        folder_configs = list(self.folder_configs or [])
        ns = OutlookClient.get_namespace()
        sent_path_lower = ""
        if ns:
            try:
                sent_path = getattr(ns.GetDefaultFolder(5), "FolderPath", "") or ""
                sent_path_lower = sent_path.strip().lower()
                known_paths = {str(fc.get("folder_path", "")).strip().lower() for fc in folder_configs}
                if sent_path and sent_path_lower not in known_paths:
                    folder_configs.append({"folder_path": sent_path, "tag": "Sent", "include_subfolders": False})
            except Exception:
                pass

        # Expand configured folders + subfolders (most-specific path first so specific tags win over parent folders)
        ordered_configs = sorted(
            folder_configs,
            key=lambda fc: len([p for p in str(fc.get("folder_path", "")).replace("\\", "/").split("/") if p]),
            reverse=True,
        )
        visited_folder_keys: Set[str] = set()
        resolved_targets = []
        for folder_cfg in ordered_configs:
            if cancel_check and cancel_check():
                break
            folder_path = folder_cfg.get("folder_path", "")
            folder_tag = folder_cfg.get("tag", "General")
            include_sub = bool(folder_cfg.get("include_subfolders", True))
            if not folder_path:
                continue

            if progress_cb:
                progress_cb(f"Connecting to: {folder_path} [{folder_tag}]...")
            root_folder = OutlookClient.get_folder_by_path(folder_path)
            if not root_folder:
                logging.warning(f"Could not open Outlook folder: {folder_path}")
                continue
            for sub_f in OutlookClient.iter_folder_tree(
                root_folder,
                include_subfolders=include_sub,
                visited_keys=visited_folder_keys,
                fallback_path=folder_path,
            ):
                resolved_targets.append((sub_f, folder_tag, folder_path))

        for folder, folder_tag, folder_path in resolved_targets:
            if cancel_check and cancel_check():
                break

            batch_emails = []
            batch_matches = []
            flag_updates = []

            tbl = None
            if hasattr(folder, "GetTable"):
                try:
                    tbl = folder.GetTable()
                except Exception:
                    tbl = None

            if tbl is not None and ns is not None:
                cur_folder_path = str(getattr(folder, "FolderPath", folder_path) or "").strip().lower()
                is_sent = (str(folder_tag or "").strip().lower() == "sent") or (
                    bool(sent_path_lower) and cur_folder_path == sent_path_lower
                )
                restrict_utc = self.since_dt_utc
                if restrict_utc is None and filter_start is not None:
                    restrict_utc = local_naive_to_utc(filter_start)
                tbl, restricted = apply_mapi_since_restriction(tbl, restrict_utc, is_sent_folder=is_sent)
                try:
                    tbl.Sort("[ReceivedTime]", True)
                except Exception:
                    try:
                        tbl.Sort("[SentOn]", True)
                    except Exception:
                        pass
                add_mapi_table_columns(
                    tbl,
                    ("EntryID", "MessageClass", "ReceivedTime", "SentOn", "UnRead", "FlagStatus", PR_FLAG_STATUS_PROP),
                )
                row_count = getattr(tbl, "GetRowCount", lambda: 0)() or 0
                if progress_cb:
                    if restricted:
                        progress_cb(f"Scanning {row_count} recent item(s) in [{folder_tag}]...")
                    else:
                        progress_cb(f"Scanning {row_count or 'recent'} items in [{folder_tag}]...")

                idx = 0
                while not tbl.EndOfTable:
                    if cancel_check and cancel_check():
                        break
                    idx += 1
                    if progress_val_cb and row_count and idx % 20 == 0:
                        progress_val_cb(idx, row_count)
                    try:
                        row = tbl.GetNextRow()
                    except Exception:
                        break
                    msg_cls = str(safe_mapi_row_val(row, "MessageClass", default="") or "")
                    if msg_cls and not msg_cls.startswith("IPM.Note"):
                        continue

                    if filter_start or filter_end:
                        dt = self._to_naive_dt(safe_mapi_row_val(row, "ReceivedTime", "SentOn"))
                        if dt:
                            if filter_end and dt > filter_end:
                                continue
                            if filter_start and dt < filter_start:
                                break

                    entry_id = str(safe_mapi_row_val(row, "EntryID", default="") or "")
                    if not entry_id:
                        continue

                    cached = cached_states.get(entry_id)
                    if cached and cached.get("has_body"):
                        state = extract_table_read_flag(row)
                        if state is None:
                            state = extract_item_read_flag(ns, entry_id)
                        if state is not None and (state[0] != cached["is_read"] or state[1] != cached["flag_status"]):
                            flag_updates.append((entry_id, state[0], state[1]))
                            cached["is_read"], cached["flag_status"] = state
                        continue

                    try:
                        mail_item = ns.GetItemFromID(entry_id)
                    except Exception:
                        continue
                    if not mail_item or getattr(mail_item, "Class", 43) != 43:
                        continue
                    rec, rec_matches = self._build_email_record(mail_item, folder_tag, entry_id=entry_id, require_tdoc=True)
                    if rec is None:
                        continue
                    cached_states[entry_id] = {
                        "is_read": rec["is_read"],
                        "flag_status": rec["flag_status"],
                        "has_body": bool(rec["body_text"]),
                    }
                    batch_emails.append(rec)
                    batch_matches.extend(rec_matches)
                    total_synced += 1
                    if len(batch_emails) >= 40:
                        self.db.save_emails_batch(batch_emails, batch_matches, sync_unread=True)
                        batch_emails.clear()
                        batch_matches.clear()
            else:
                items = folder.Items
                total_items = len(items)
                items.Sort("[ReceivedTime]", True)
                if progress_cb:
                    progress_cb(f"Scanning {total_items} items in [{folder_tag}]...")

                for i in range(1, total_items + 1):
                    if cancel_check and cancel_check():
                        break
                    if progress_val_cb and i % 15 == 0:
                        progress_val_cb(i, total_items)

                    mail_item = items.Item(i)
                    if getattr(mail_item, "Class", 43) != 43:
                        continue

                    if filter_start or filter_end:
                        dt = self._to_naive_dt(getattr(mail_item, "ReceivedTime", None) or getattr(mail_item, "SentOn", None))
                        if dt:
                            if filter_end and dt > filter_end:
                                continue
                            if filter_start and dt < filter_start:
                                break

                    entry_id = getattr(mail_item, "EntryID", "")
                    if not entry_id:
                        continue

                    cached = cached_states.get(entry_id)
                    if cached and cached.get("has_body"):
                        is_read = 1 if getattr(mail_item, "UnRead", False) is False else 0
                        flag_status = 2 if int(getattr(mail_item, "FlagStatus", 0) or 0) == 2 else 0
                        if is_read != cached["is_read"] or flag_status != cached["flag_status"]:
                            flag_updates.append((entry_id, is_read, flag_status))
                            cached["is_read"], cached["flag_status"] = is_read, flag_status
                        continue

                    rec, rec_matches = self._build_email_record(mail_item, folder_tag, entry_id=entry_id, require_tdoc=True)
                    if rec is None:
                        continue
                    cached_states[entry_id] = {
                        "is_read": rec["is_read"],
                        "flag_status": rec["flag_status"],
                        "has_body": bool(rec["body_text"]),
                    }
                    batch_emails.append(rec)
                    batch_matches.extend(rec_matches)
                    total_synced += 1
                    if len(batch_emails) >= 40:
                        self.db.save_emails_batch(batch_emails, batch_matches, sync_unread=True)
                        batch_emails.clear()
                        batch_matches.clear()

            if batch_emails:
                self.db.save_emails_batch(batch_emails, batch_matches, sync_unread=True)
            if flag_updates:
                total_flags_updated += self.db.update_email_flags_batch(flag_updates)

        should_cache_drafts = (self.since_dt_utc is None) or (total_synced > 0)
        if should_cache_drafts and not (cancel_check and cancel_check()) and self.db and str(self.db.db_path) != ":memory:":
            try:
                from modules.emails.core.tracked_threads_service import cache_drafts_for_messages
                all_msgs = self.db.get_all_emails_with_tdocs(show_ignored=True)
                cache_drafts_for_messages(
                    all_msgs,
                    Path(self.db.db_path).parent,
                    progress_cb=progress_cb,
                    cancel_check=cancel_check,
                )
            except Exception as exc:
                logging.debug(f"Auto-caching /Drafts/ URLs skipped cleanly: {exc}")

        return total_synced + total_flags_updated

    @classmethod
    def extract_tdocs_from_parts(
        cls,
        subject: str,
        body: str,
        attachments: Optional[List[str]] = None,
    ) -> List[dict]:
        """Extracts TDoc matches from subject, direct/quoted body, and attachment filenames."""
        results = {}
        subj_s = str(subject or "")
        body_s = str(body or "")

        # 1. Subject extraction (Highest Priority)
        for tdoc in cls.TDOC_REGEX.findall(subj_s):
            tdoc_up = tdoc.upper()
            results[tdoc_up] = {"tdoc_id": tdoc_up, "rev_matched": "", "location": "Subject"}

        for base, rev in cls.REV_REGEX.findall(subj_s):
            base_up = base.upper()
            rev_str = f"{base_up}r{int(rev):02d}"
            results[base_up] = {"tdoc_id": base_up, "rev_matched": rev_str, "location": "Subject"}

        # 2. Attachment filenames (treated as Body priority)
        att_text = "\n".join(str(a or "") for a in (attachments or []) if str(a or "").strip())
        if att_text:
            for tdoc in cls.TDOC_REGEX.findall(att_text):
                tdoc_up = tdoc.upper()
                if tdoc_up not in results:
                    results[tdoc_up] = {"tdoc_id": tdoc_up, "rev_matched": "", "location": "Body"}
            for base, rev in cls.REV_REGEX.findall(att_text):
                base_up = base.upper()
                rev_str = f"{base_up}r{int(rev):02d}"
                if base_up not in results or results[base_up]["location"] != "Subject":
                    results[base_up] = {"tdoc_id": base_up, "rev_matched": rev_str, "location": "Body"}

        # 3. Split Body into Direct Text vs Quoted History
        direct_body, quoted_body = split_direct_and_quoted_body(body_s)

        # Direct Body extraction
        for tdoc in cls.TDOC_REGEX.findall(direct_body):
            tdoc_up = tdoc.upper()
            if tdoc_up not in results:
                results[tdoc_up] = {"tdoc_id": tdoc_up, "rev_matched": "", "location": "Body"}

        for base, rev in cls.REV_REGEX.findall(direct_body):
            base_up = base.upper()
            rev_str = f"{base_up}r{int(rev):02d}"
            if base_up not in results or results[base_up]["location"] != "Subject":
                results[base_up] = {"tdoc_id": base_up, "rev_matched": rev_str, "location": "Body"}

        # Quoted History extraction
        if quoted_body:
            for tdoc in cls.TDOC_REGEX.findall(quoted_body):
                tdoc_up = tdoc.upper()
                if tdoc_up not in results:
                    results[tdoc_up] = {"tdoc_id": tdoc_up, "rev_matched": "", "location": "Quoted"}

            for base, rev in cls.REV_REGEX.findall(quoted_body):
                base_up = base.upper()
                rev_str = f"{base_up}r{int(rev):02d}"
                if base_up not in results:
                    results[base_up] = {"tdoc_id": base_up, "rev_matched": rev_str, "location": "Quoted"}

        return list(results.values())

    def _extract_tdocs(self, subject: str, body: str, attachments: Optional[List[str]] = None) -> List[dict]:
        return self.extract_tdocs_from_parts(subject, body, attachments=attachments)

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