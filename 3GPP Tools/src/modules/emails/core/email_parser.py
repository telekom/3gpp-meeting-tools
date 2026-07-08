# --- File: modules/emails/core/email_parser.py ---
import re
import logging
from typing import Dict

from core.utils.company_sanitizer import CompanySanitizer


class EmailParser:
    # Accommodates 6, 7, or 8 digit TDocs correctly
    TDOC_REGEX = re.compile(r'(S2-\d{6,8})', re.IGNORECASE)
    START_END_REGEX = re.compile(r'<<START>>(.*?)<<END>>', re.DOTALL | re.IGNORECASE)
    REVISION_REGEX = re.compile(r'\b(?:r|rev\s*)0?([1-9])\b', re.IGNORECASE)

    @classmethod
    def parse_outlook_item(cls, mail_item, ai_lookup_dict: dict) -> Dict:
        try:
            subject = getattr(mail_item, "Subject", "")
            body = getattr(mail_item, "Body", "")
            sender_name = getattr(mail_item, "SenderName", "")

            try:
                sender_email = getattr(mail_item, "SenderEmailAddress", "")
            except Exception:
                sender_email = ""

            sender_name_lower = sender_name.lower()
            sender_email_lower = sender_email.lower()

            # 1. Broadened Listserv / DMARC Intercept
            is_list_email = "list.etsi.org" in sender_email_lower or "dmarc" in sender_name_lower or "on behalf of" in sender_name_lower

            if is_list_email:
                try:
                    for recipient in mail_item.ReplyRecipients:
                        sender_email = getattr(recipient, "Address", sender_email)

                        # ---> NEW: If the Reply-To is an internal Exchange user (like yourself), resolve it!
                        if sender_email.lower().startswith("/o="):
                            ae = getattr(recipient, "AddressEntry", None)
                            if ae:
                                eu = ae.GetExchangeUser()
                                if eu:
                                    smtp = getattr(eu, "PrimarySmtpAddress", "")
                                    if smtp: sender_email = smtp
                        break  # We only need the primary reply recipient

                    # Fix the Sender Name
                    reply_names = getattr(mail_item, "ReplyRecipientNames", "")
                    if reply_names:
                        sender_name = reply_names.split(';')[0].strip()

                except Exception as e:
                    logging.warning(f"Could not extract ReplyRecipient for listserv bypass: {e}")

            else:
                # 2. Direct Internal Emails (Bypassing the listserv)
                # If a colleague emailed you directly, we still need to resolve their Exchange address
                try:
                    sender_email_type = str(getattr(mail_item, "SenderEmailType", "")).upper()
                    if sender_email_type == "EX" or sender_email.lower().startswith("/o="):
                        sender_obj = getattr(mail_item, "Sender", None)
                        if sender_obj:
                            eu = sender_obj.GetExchangeUser()
                            if eu:
                                smtp = getattr(eu, "PrimarySmtpAddress", "")
                                if smtp: sender_email = smtp
                except Exception as e:
                    logging.warning(f"Error extracting EX sender email address: {e}")

            # 3. TDoc Extraction & Strict Meeting Enforcement
            tdoc_match = cls.TDOC_REGEX.search(subject)
            if not tdoc_match:
                return {}

            base_tdoc = tdoc_match.group(1).upper()

            # ---> STRICT MEETING FILTER: Ignore TDocs not in this meeting's Excel list
            if base_tdoc not in ai_lookup_dict:
                return {}
            agenda_item = ai_lookup_dict.get(base_tdoc, "Unknown AI")

            # ---> IMPROVED REGEX FALLBACK:
            # If the name is still polluted with listserv artifacts, parse the body for the clean From: header
            if any(keyword in sender_name.lower() for keyword in
                   ["3gpp", "list", "emeet", "on behalf of", "dmarc"]) or not sender_email:
                body_head = body[:1500]  # Check a larger chunk at the top of the email

                dmarc_match = re.search(
                    r'From:\s*([^\n<\[]+?)\s*[<\[](?:mailto:)?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})[>\]]',
                    body_head, re.IGNORECASE)

                if dmarc_match:
                    sender_name = dmarc_match.group(1).strip(' \t"\'')
                    if not sender_email or "list.etsi.org" in sender_email.lower() or sender_email.lower().startswith(
                            "/o="):
                        sender_email = dmarc_match.group(2).strip()

            # 4. Company Extraction
            raw_sender_str = f"{sender_name} <{sender_email}>"
            companies = CompanySanitizer.get_matching_contributors(raw_sender_str)
            company = companies[0] if companies else cls._fallback_domain_extract(sender_email)

            # 5. Body Slicing
            short_text = ""
            free_text = ""
            block_match = cls.START_END_REGEX.search(body)

            if block_match:
                short_text = block_match.group(1).strip()
                after_end = body[block_match.end():]
                free_text = cls._clean_free_text(after_end)

            # 6. Revision Hunting
            rev_mentions = []
            rev_matches = cls.REVISION_REGEX.findall(short_text)
            for rev_num in rev_matches:
                normalized = f"{base_tdoc}r{int(rev_num):02d}"
                rev_mentions.append(normalized)

            return {
                "id": getattr(mail_item, "EntryID", ""),
                "tdoc_id": base_tdoc,
                "agenda_item": agenda_item,
                "sender_name": sender_name,
                "sender_email": sender_email,
                "company": company,
                "date_received": str(getattr(mail_item, "ReceivedTime", "")),
                "subject": subject,
                "revisions_mentioned": ", ".join(list(set(rev_mentions))),
                "short_text": short_text,
                "free_text": free_text,
                "msg_path": ""
            }
        except Exception as e:
            logging.error(f"Error parsing email '{getattr(mail_item, 'Subject', 'Unknown')}': {e}")
            return {}

    @staticmethod
    def _fallback_domain_extract(email: str) -> str:
        if not email or "@" not in email:
            return "Unknown"
        domain_parts = email.split('@')[-1].split('.')
        for part in domain_parts:
            if part.lower() not in ["com", "co", "cn", "jp", "uk", "net", "org", "edu"]:
                return part.title()
        return "Unknown"

    @staticmethod
    def _clean_free_text(text: str) -> str:
        cut_markers = ["From: 3GPP_TSG_SA", "________________________________", "On behalf of", "> "]
        lines = text.split('\n')
        clean_lines = []
        for line in lines:
            if any(marker in line for marker in cut_markers): break
            clean_lines.append(line)
        return '\n'.join(clean_lines).strip()