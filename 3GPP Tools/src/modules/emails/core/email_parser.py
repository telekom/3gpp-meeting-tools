# --- File: modules/emails/core/email_parser.py ---
import re
import logging
from typing import Dict
from modules.emails.core.contributor_names import get_matching_contributors


class EmailParser:
    # Captures standard TDocs (e.g., S2-260123) regardless of year
    TDOC_REGEX = re.compile(r'(S2-\d{6})', re.IGNORECASE)

    # Captures <<START>> ... <<END>> block safely across multiple lines
    START_END_REGEX = re.compile(r'<<START>>(.*?)<<END>>', re.DOTALL | re.IGNORECASE)

    # Finds shorthand revisions like "r02", "r2", "rev3", "rev 04"
    REVISION_REGEX = re.compile(r'\b(?:r|rev\s*)0?([1-9])\b', re.IGNORECASE)

    @classmethod
    def parse_outlook_item(cls, mail_item, ai_lookup_dict: dict) -> Dict:
        """
        Parses a raw Outlook COM mail item into a clean dictionary for our SQLite DB.
        ai_lookup_dict is simply { 'S2-260123': '9.1.1' } passed from the main TDocs list.
        """
        try:
            subject = getattr(mail_item, "Subject", "")
            body = getattr(mail_item, "Body", "")
            sender_name = getattr(mail_item, "SenderName", "")

            # Outlook can sometimes throw errors getting the raw address if the user is internal Exchange
            try:
                sender_address = getattr(mail_item, "SenderEmailAddress", "")
            except Exception:
                sender_address = ""

            # 1. TDoc Extraction
            tdoc_match = cls.TDOC_REGEX.search(subject)
            if not tdoc_match:
                return {}  # Not a valid 3GPP eMeeting email

            base_tdoc = tdoc_match.group(1).upper()
            agenda_item = ai_lookup_dict.get(base_tdoc, "Unknown AI")

            # 2. Company Extraction (using your brilliant contributor_names.py)
            raw_sender_str = f"{sender_name} <{sender_address}>"
            # We use dummy sets because get_matching_contributors expects them
            companies = get_matching_contributors(raw_sender_str, set(), set())
            company = companies[0] if companies else cls._fallback_domain_extract(sender_address)

            # 3. Body Slicing
            short_text = ""
            free_text = ""
            block_match = cls.START_END_REGEX.search(body)

            if block_match:
                short_text = block_match.group(1).strip()
                # Free text is everything AFTER the <<END>> tag
                after_end = body[block_match.end():]
                free_text = cls._clean_free_text(after_end)

            # 4. Revision Hunting
            rev_mentions = []
            rev_matches = cls.REVISION_REGEX.findall(short_text)
            for rev_num in rev_matches:
                # Normalizes shorthand "r2" into "S2-260123r02"
                normalized = f"{base_tdoc}r{int(rev_num):02d}"
                rev_mentions.append(normalized)

            return {
                "id": getattr(mail_item, "EntryID", ""),  # Unique Outlook ID
                "tdoc_id": base_tdoc,
                "agenda_item": agenda_item,
                "sender_name": sender_name,
                "company": company,
                "date_received": str(getattr(mail_item, "ReceivedTime", "")),
                "subject": subject,
                "revisions_mentioned": ", ".join(list(set(rev_mentions))),
                "short_text": short_text,
                "free_text": free_text,
                "msg_path": ""  # Will be populated when we save the .msg file to the hard drive
            }
        except Exception as e:
            logging.error(f"Error parsing email '{getattr(mail_item, 'Subject', 'Unknown')}': {e}")
            return {}

    @staticmethod
    def _fallback_domain_extract(email: str) -> str:
        """If the company isn't in your python list, gracefully extract the domain."""
        if not email or "@" not in email:
            return "Unknown"
        domain_parts = email.split('@')[-1].split('.')
        # Avoid picking "com", "co", "cn" by grabbing the most substantial part
        for part in domain_parts:
            if part.lower() not in ["com", "co", "cn", "jp", "uk", "net", "org", "edu"]:
                return part.title()
        return "Unknown"

    @staticmethod
    def _clean_free_text(text: str) -> str:
        """Cuts off historical quotes so the database doesn't bloat."""
        cut_markers = [
            "From: 3GPP_TSG_SA",
            "________________________________",
            "On behalf of",
            "> "
        ]
        lines = text.split('\n')
        clean_lines = []
        for line in lines:
            if any(marker in line for marker in cut_markers):
                break
            clean_lines.append(line)
        return '\n'.join(clean_lines).strip()