# --- File: modules/emails/core/email_parser.py ---
import re
import logging
from typing import Dict
from modules.emails.core.contributor_names import get_matching_contributors


class EmailParser:
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

            # 1. TDoc Extraction & Meeting Enforcement
            tdoc_match = cls.TDOC_REGEX.search(subject)
            if not tdoc_match:
                return {}

            base_tdoc = tdoc_match.group(1).upper()

            # ---> STRICT MEETING FILTER: Ignore TDocs not in this meeting
            if base_tdoc not in ai_lookup_dict:
                return {}
            agenda_item = ai_lookup_dict.get(base_tdoc, "Unknown AI")

            # ---> DMARC BYPASS: Extract actual sender from body if listserv rewritten
            sender_name_lower = sender_name.lower()
            if "3gpp" in sender_name_lower or "list" in sender_name_lower or "emeet" in sender_name_lower or not sender_email:
                body_head = body[:500]  # Check just the top of the email
                # Matches: From: Real Name <real.email@domain.com> OR From: Real Name [mailto:real.email@domain.com]
                dmarc_match = re.search(r'^From:\s*"?([^<\["\n]+)"?\s*[<\[](?:mailto:)?([^>\]\n]+)[>\]]', body_head,
                                        re.MULTILINE | re.IGNORECASE)
                if dmarc_match:
                    sender_name = dmarc_match.group(1).strip()
                    sender_email = dmarc_match.group(2).strip()

            # 2. Company Extraction
            raw_sender_str = f"{sender_name} <{sender_email}>"
            companies = get_matching_contributors(raw_sender_str, set(), set())
            company = companies[0] if companies else cls._fallback_domain_extract(sender_email)

            # 3. Body Slicing
            short_text = ""
            free_text = ""
            block_match = cls.START_END_REGEX.search(body)

            if block_match:
                short_text = block_match.group(1).strip()
                after_end = body[block_match.end():]
                free_text = cls._clean_free_text(after_end)

            # 4. Revision Hunting
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