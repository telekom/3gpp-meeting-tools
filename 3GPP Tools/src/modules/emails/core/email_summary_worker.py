# --- File: src/modules/emails/core/email_summary_worker.py ---
import html
import logging
import re
from pathlib import Path
from typing import List, Dict, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

from core.ai.ollama_client import OllamaClient
from core.utils.paths import get_project_root

PROMPTS_DIR = get_project_root() / "config" / "prompts"


def load_prompt_file(filename: str, fallback_default: str) -> str:
    """Loads a prompt template from config/prompts/ with timestamp hot-reloading."""
    prompt_path = PROMPTS_DIR / filename
    if prompt_path.exists():
        try:
            with open(prompt_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception as e:
            logging.warning(f"[AI Email Summary] Failed to read {filename}: {e}")
    return fallback_default.strip()


def clean_email_body_for_summary(raw_body: str) -> Tuple[str, int]:
    """
    Strips email replies, quoted chains, inline quotation characters,
    and listserv footers to isolate the author's direct message content.
    Returns (cleaned_text, characters_stripped).
    """
    if not raw_body:
        return "", 0

    original_len = len(raw_body)
    text = raw_body.replace("\x00", "").replace("\r\n", "\n").replace("\r", "\n")

    # 1. Truncate at reply headers and historical boundaries
    quote_boundary_regex = re.compile(
        r'\n(?:\s*[-_]{5,}\s*|'
        r'\s*(?:Von|From|De|Gesendet|Sent|Betreff|Subject):\s*.+|'
        r'\s*On\s+.+wrote:\s*|'
        r'\s*Am\s+.+schrieb.*:\s*|'
        r'\s*<<START>>|\s*<<END>>)',
        re.IGNORECASE
    )
    split_match = quote_boundary_regex.search(text)
    if split_match:
        text = text[:split_match.start()]

    # 2. Line-by-line filtering: remove quote prefixes (">", "|", "&gt;")
    lines = text.split("\n")
    cleaned_lines = []
    disclaimer_markers = (
        "to unsubscribe from the 3gpp",
        "confidentiality notice",
        "this e-mail and any files transmitted",
        "etats-unis / united states",
        "list.etsi.org",
    )

    for line in lines:
        stripped = line.strip()
        if not stripped:
            cleaned_lines.append("")
            continue

        # Check for listserv footers and disclaimers
        lower = stripped.lower()
        if any(d in lower for d in disclaimer_markers):
            break

        # Strip standard quotation prefixes
        if stripped.startswith(">") or stripped.startswith("|") or stripped.startswith("&gt;"):
            continue

        cleaned_lines.append(line.rstrip())

    # 3. Compress consecutive blank lines into a single blank line
    condensed = []
    empty_streak = 0
    for line in cleaned_lines:
        if not line.strip():
            empty_streak += 1
            if empty_streak <= 1:
                condensed.append("")
        else:
            empty_streak = 0
            condensed.append(line)

    result = "\n".join(condensed).strip()
    chars_stripped = max(0, original_len - len(result))
    return result, chars_stripped


class EmailSummaryWorker(QThread):
    """Background worker that sanitizes email chains and streams LLM synthesis."""

    progress_status = pyqtSignal(str)
    token_received = pyqtSignal(str)
    generation_completed = pyqtSignal(str)
    generation_failed = pyqtSignal(str)

    def __init__(
        self,
        client: OllamaClient,
        model: str,
        tdoc_id: str,
        family_tdocs: List[str],
        emails_data: List[Dict],
        parent=None
    ):
        super().__init__(parent)
        self.client = client
        self.model = model
        self.tdoc_id = tdoc_id
        self.family_tdocs = family_tdocs
        self.emails_data = emails_data
        self._is_cancelled = False

    def cancel(self):
        """Flag the worker thread to abort processing."""
        self._is_cancelled = True

    def run(self):
        try:
            if not self.emails_data:
                self.generation_failed.emit("No emails available to summarize.")
                return

            self.progress_status.emit("🧹 Sanitizing and chronologically ordering emails...")

            # 1. Chronological sort (oldest to newest)
            sorted_emails = sorted(
                self.emails_data,
                key=lambda x: str(x.get("date_received", ""))
            )

            total_raw_chars = 0
            total_clean_chars = 0
            formatted_messages = []

            # 2. Extract direct message content with verbose logging
            for idx, e in enumerate(sorted_emails, start=1):
                if self._is_cancelled:
                    self.generation_failed.emit("Summary cancelled by user.")
                    return

                sender_name = e.get("sender_name") or "Unknown Sender"
                company = e.get("company") or "Unknown"
                date_str = str(e.get("date_received", ""))[:16]
                subject = e.get("subject", "")
                raw_body = e.get("body_text", "")

                clean_body, stripped_count = clean_email_body_for_summary(raw_body)
                raw_len = len(raw_body)
                clean_len = len(clean_body)
                total_raw_chars += raw_len
                total_clean_chars += clean_len

                pct_stripped = (stripped_count / raw_len * 100) if raw_len > 0 else 0.0

                logging.info(
                    f"[Email Summary] Email #{idx}/{len(sorted_emails)}: "
                    f"Sender='{sender_name}' ({company}) Date='{date_str}' | "
                    f"Raw: {raw_len} chars -> Clean: {clean_len} chars "
                    f"(Stripped: {stripped_count} chars, {pct_stripped:.1f}%)"
                )

                if clean_len > 1500 and stripped_count == 0:
                    logging.warning(
                        f"[Email Summary] Email #{idx} exceeds 1,500 chars with 0 quotes stripped. "
                        f"Check for non-standard quote headers. Lead text: {clean_body[:100]!r}"
                    )

                if clean_body:
                    msg_block = (
                        f"--- Message [{idx}/{len(sorted_emails)}] ---\n"
                        f"From: {sender_name} ({company})\n"
                        f"Date: {date_str}\n"
                        f"Subject: {subject}\n\n"
                        f"{clean_body}"
                    )
                    formatted_messages.append(msg_block)

            pct_total_reduction = (
                ((total_raw_chars - total_clean_chars) / total_raw_chars * 100)
                if total_raw_chars > 0 else 0.0
            )

            logging.info(
                f"[Email Summary] Thread Compression Stats: "
                f"Emails={len(sorted_emails)} | Raw Chars={total_raw_chars} -> "
                f"Clean Chars={total_clean_chars} | Reduced by {pct_total_reduction:.1f}%"
            )

            if not formatted_messages:
                self.generation_failed.emit("All selected emails were empty after removing quotations.")
                return

            # 3. Assemble prompts
            sys_prompt = load_prompt_file(
                "email_summary_system.txt",
                "You are an expert 3GPP standards delegate. Summarize the following email debate into: "
                "1. Discussion Topic & Core Contention, 2. Company Positions & Alliances, "
                "3. Proposed Solutions & Revisions Mentioned, 4. Current Consensus & Next Steps."
            )
            usr_template = load_prompt_file(
                "email_summary_user.txt",
                "TDoc: {tdoc_id}\nFamily: {family_tdocs}\nCount: {email_count}\n\nEmails:\n{emails_content}"
            )

            emails_payload = "\n\n".join(formatted_messages)

            formatted_user = usr_template.format(
                tdoc_id=self.tdoc_id,
                family_tdocs=", ".join(self.family_tdocs),
                email_count=len(formatted_messages),
                emails_content=emails_payload
            )

            messages = [
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": formatted_user}
            ]

            # 4. Stream chat response from Ollama
            self.progress_status.emit("🦙 Streaming email summary tokens...")
            accumulated = []
            stream_options = {"temperature": 0.1, "num_ctx": 8192}

            for delta in self.client.stream_chat(
                model=self.model,
                messages=messages,
                options=stream_options
            ):
                if self._is_cancelled:
                    self.generation_failed.emit("Summary cancelled by user.")
                    return
                accumulated.append(delta)
                self.token_received.emit(delta)

            full_summary = "".join(accumulated).strip()
            self.generation_completed.emit(full_summary)

        except Exception as e:
            if not self._is_cancelled:
                logging.error(f"[AI Email Summary] Generation error: {e}", exc_info=True)
                self.generation_failed.emit(str(e))