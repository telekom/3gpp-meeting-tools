# --- File: src/modules/emails/core/email_deduplicator.py ---
"""
Fast, lossless email deduplication engine for `GeneralEmailDatabase` (`emails.db`)
and `TrackedThreadsDatabase` (`tracked_threads.db`).

Prevents and automatically cleans up duplicate copies of the same logical email that
have different MAPI `EntryID`s (e.g., emails present in both Inbox and a WG subfolder,
emails moved between folders, listserv echoes, or `.msg` imports alongside Outlook syncs).
"""

from __future__ import annotations

import hashlib
import json
import logging
import re
import sqlite3
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

_TRAIL_COMP_RE = re.compile(r"\s*\([^()]{1,48}\)\s*$")
_NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")
_SUBJ_PREFIX_RE = re.compile(r"^(?:\s*(?:re|fwd?|aw|wg|sv|antw|回复|答复|转发)\s*[:：]\s*)+", re.IGNORECASE)
_LOC_PRIORITY = {"Subject": 3, "Body": 2, "Thread": 2, "Quoted": 1}


def _normalize_sender_key(sender_name: Any, sender_email: Any) -> str:
    """Produces a stable sender key resilient to listserv/DMARC vs direct Exchange addresses."""
    name = str(sender_name or "").strip()
    if name:
        if " on behalf of " in name.lower():
            name = re.split(r"\bon behalf of\b", name, flags=re.IGNORECASE)[-1].strip()
        prev = None
        while prev != name:
            prev = name
            name = _TRAIL_COMP_RE.sub("", name).strip()
        norm_name = _NON_ALNUM_RE.sub("", name.lower())
        if norm_name and norm_name not in ("unknown", "3gpp", "list", "dmarc"):
            return f"n:{norm_name}"

    email = str(sender_email or "").strip().lower()
    if email and "@" in email and "list.etsi.org" not in email:
        return f"e:{email}"
    raw_fallback = _NON_ALNUM_RE.sub("", (name or email).lower())
    return f"r:{raw_fallback}" if raw_fallback else ""


def _normalize_date_minute(date_raw: Any) -> str:
    """Normalizes timestamps to `YYYY-MM-DD HH:MM` precision for cross-folder matching."""
    txt = str(date_raw or "").strip()
    if len(txt) < 16:
        return ""
    norm = txt.replace("T", " ")
    if re.match(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}", norm):
        return norm[:16]
    return ""


def _normalize_body_fingerprint(clean_body: Any, body_text: Any, subject: Any) -> str:
    """Hashes the opening direct body text (or normalized subject when body is empty)."""
    raw = str(clean_body or "").strip() or str(body_text or "").strip()
    if raw:
        alnum = _NON_ALNUM_RE.sub("", raw.lower())[:240]
        if alnum:
            return hashlib.sha1(alnum.encode("utf-8", errors="ignore")).hexdigest()[:16]
    subj = _SUBJ_PREFIX_RE.sub("", str(subject or "").strip()).strip().lower()
    alnum_s = _NON_ALNUM_RE.sub("", subj)[:160]
    if alnum_s:
        return "s:" + hashlib.sha1(alnum_s.encode("utf-8", errors="ignore")).hexdigest()[:14]
    return ""


def compute_message_fingerprint(
    sender_name: Any,
    sender_email: Any,
    date_raw: Any,
    subject: Any,
    clean_body: Any = "",
    body_text: Any = "",
) -> Optional[str]:
    """
    Computes a canonical logical fingerprint `(sender_key | YYYY-MM-DD HH:MM | content_hash)`.
    Returns `None` if timestamp or content is missing so unrelated undated rows are never merged.
    """
    date_min = _normalize_date_minute(date_raw)
    if not date_min:
        return None
    sender_key = _normalize_sender_key(sender_name, sender_email)
    content_sig = _normalize_body_fingerprint(clean_body, body_text, subject)
    if not sender_key or not content_sig:
        return None
    return f"{sender_key}|{date_min}|{content_sig}"


def _parse_attachments_json(raw_val: Any) -> List[str]:
    if isinstance(raw_val, list):
        return [str(x).strip() for x in raw_val if str(x).strip()]
    if isinstance(raw_val, str) and raw_val.strip():
        try:
            parsed = json.loads(raw_val)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed if str(x).strip()]
        except Exception:
            pass
    return []


def _winner_sort_key(row: Dict[str, Any]) -> Tuple[int, int, int, int]:
    """Ranks candidate duplicate rows: prefer live MAPI EntryID > attachments > longest body > read."""
    eid = str(row.get("id") or "")
    is_live_outlook = 0 if (not eid or eid.startswith("MSG-")) else 1
    atts = _parse_attachments_json(row.get("attachments"))
    has_atts = 1 if atts else 0
    body_len = int(row.get("body_len") or 0)
    is_read = 1 if int(row.get("is_read") or 0) else 0
    return (is_live_outlook, has_atts, body_len, is_read)


def deduplicate_general_emails(conn: sqlite3.Connection) -> int:
    """
    Fast, single-pass deduplication of `general_emails` + `general_email_tdoc_matches`.
    Preserves `is_read = MAX`, `flag_status = MAX`, union of attachments, and best TDoc matches.
    Returns the number of duplicate rows removed.
    """
    try:
        cursor = conn.cursor()
        rows = cursor.execute(
            """
            SELECT id, subject, sender_name, sender_email, date_received,
                   substr(COALESCE(clean_body, ''), 1, 500),
                   substr(COALESCE(body_text, ''), 1, 500),
                   attachments, is_read, is_ignored, COALESCE(flag_status, 0),
                   length(COALESCE(body_text, ''))
            FROM general_emails
            """
        ).fetchall()
        if len(rows) < 2:
            return 0

        buckets: Dict[str, List[Dict[str, Any]]] = {}
        for r in rows:
            fp = compute_message_fingerprint(r[2], r[3], r[4], r[1], r[5], r[6])
            if not fp:
                continue
            buckets.setdefault(fp, []).append({
                "id": str(r[0]),
                "attachments": r[7],
                "is_read": int(r[8] or 0),
                "is_ignored": int(r[9] or 0),
                "flag_status": 2 if int(r[10] or 0) == 2 else 0,
                "body_len": int(r[11] or 0),
            })

        dup_groups = [grp for grp in buckets.values() if len(grp) > 1]
        if not dup_groups:
            return 0

        removed_total = 0
        for grp in dup_groups:
            grp_sorted = sorted(grp, key=_winner_sort_key, reverse=True)
            winner = grp_sorted[0]
            losers = grp_sorted[1:]
            winner_id = winner["id"]
            loser_ids = [x["id"] for x in losers]

            merged_is_read = max(x["is_read"] for x in grp_sorted)
            merged_flag = 2 if any(x["flag_status"] == 2 for x in grp_sorted) else 0
            merged_ignored = min(x["is_ignored"] for x in grp_sorted)

            merged_atts: List[str] = []
            for x in grp_sorted:
                for a in _parse_attachments_json(x.get("attachments")):
                    if a not in merged_atts:
                        merged_atts.append(a)

            all_ids = [winner_id] + loser_ids
            ph = ",".join(["?"] * len(all_ids))
            match_rows = cursor.execute(
                f"SELECT tdoc_id, rev_matched, match_location FROM general_email_tdoc_matches WHERE email_id IN ({ph})",
                all_ids,
            ).fetchall()
            best_matches: Dict[str, Tuple[str, str]] = {}
            for td, rev, loc in match_rows:
                td_up = str(td or "").strip().upper()
                if not td_up:
                    continue
                rev_s = str(rev or "").strip()
                loc_s = str(loc or "Body").strip()
                prev = best_matches.get(td_up)
                if prev is None:
                    best_matches[td_up] = (rev_s, loc_s)
                else:
                    best_loc = loc_s if _LOC_PRIORITY.get(loc_s, 0) > _LOC_PRIORITY.get(prev[1], 0) else prev[1]
                    best_rev = rev_s if (rev_s and not prev[0]) else prev[0]
                    best_matches[td_up] = (best_rev, best_loc)

            cursor.execute(
                """
                UPDATE general_emails
                SET is_read = ?, is_ignored = ?, flag_status = ?, attachments = ?
                WHERE id = ?
                """,
                (merged_is_read, merged_ignored, merged_flag, json.dumps(merged_atts), winner_id),
            )
            cursor.execute("DELETE FROM general_email_tdoc_matches WHERE email_id = ?", (winner_id,))
            if best_matches:
                cursor.executemany(
                    """
                    INSERT INTO general_email_tdoc_matches (email_id, tdoc_id, rev_matched, match_location)
                    VALUES (?, ?, ?, ?)
                    """,
                    [(winner_id, td, rev, loc) for td, (rev, loc) in best_matches.items()],
                )
            loser_ph = ",".join(["?"] * len(loser_ids))
            cursor.execute(f"DELETE FROM general_email_tdoc_matches WHERE email_id IN ({loser_ph})", loser_ids)
            cursor.execute(f"DELETE FROM general_emails WHERE id IN ({loser_ph})", loser_ids)
            removed_total += len(loser_ids)

        if removed_total:
            conn.commit()
            logger.info("Deduplicated %d duplicate email(s) in general_emails.", removed_total)
        return removed_total
    except Exception as exc:
        logger.warning("General email deduplication skipped cleanly: %s", exc)
        return 0


def deduplicate_tracked_thread_messages(conn: sqlite3.Connection) -> Tuple[int, List[str]]:
    """
    Fast, single-pass deduplication of `tracked_thread_messages` per thread_id.
    Returns `(removed_count, affected_thread_ids)`.
    """
    try:
        cursor = conn.cursor()
        rows = cursor.execute(
            """
            SELECT id, thread_id, subject, sender_name, sender_email, date_sent,
                   substr(COALESCE(clean_body, ''), 1, 500),
                   substr(COALESCE(body_text, ''), 1, 500),
                   attachments, is_read, COALESCE(flag_status, 0),
                   length(COALESCE(body_text, ''))
            FROM tracked_thread_messages
            """
        ).fetchall()
        if len(rows) < 2:
            return 0, []

        buckets: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
        for r in rows:
            fp = compute_message_fingerprint(r[3], r[4], r[5], r[2], r[6], r[7])
            if not fp:
                continue
            tid = str(r[1] or "")
            buckets.setdefault((tid, fp), []).append({
                "id": str(r[0]),
                "thread_id": tid,
                "attachments": r[8],
                "is_read": int(r[9] or 0),
                "flag_status": 2 if int(r[10] or 0) == 2 else 0,
                "body_len": int(r[11] or 0),
            })

        dup_groups = [grp for grp in buckets.values() if len(grp) > 1]
        if not dup_groups:
            return 0, []

        removed_total = 0
        affected_tids: set = set()
        for grp in dup_groups:
            grp_sorted = sorted(grp, key=_winner_sort_key, reverse=True)
            winner = grp_sorted[0]
            losers = grp_sorted[1:]
            winner_id = winner["id"]
            loser_ids = [x["id"] for x in losers]
            affected_tids.add(winner["thread_id"])

            merged_is_read = max(x["is_read"] for x in grp_sorted)
            merged_flag = 2 if any(x["flag_status"] == 2 for x in grp_sorted) else 0
            merged_atts: List[str] = []
            for x in grp_sorted:
                for a in _parse_attachments_json(x.get("attachments")):
                    if a not in merged_atts:
                        merged_atts.append(a)

            cursor.execute(
                """
                UPDATE tracked_thread_messages
                SET is_read = ?, flag_status = ?, attachments = ?
                WHERE id = ?
                """,
                (merged_is_read, merged_flag, json.dumps(merged_atts), winner_id),
            )
            loser_ph = ",".join(["?"] * len(loser_ids))
            cursor.execute(f"DELETE FROM tracked_thread_messages WHERE id IN ({loser_ph})", loser_ids)
            removed_total += len(loser_ids)

        if removed_total:
            conn.commit()
            logger.info(
                "Deduplicated %d duplicate message(s) across %d tracked thread(s).",
                removed_total,
                len(affected_tids),
            )
        return removed_total, sorted(affected_tids)
    except Exception as exc:
        logger.warning("Tracked thread message deduplication skipped cleanly: %s", exc)
        return 0, []


def deduplicate_tracked_threads_table(db: Any) -> int:
    """
    Merges duplicate `tracked_threads` rows that share the same `(meeting_id, normalized_subject)`.
    Returns the number of redundant threads merged away.
    """
    try:
        with db._get_connection() as conn:
            rows = conn.execute(
                """
                SELECT thread_id, meeting_id, normalized_subject,
                       COALESCE(message_count, 0), COALESCE(last_message_date, '')
                FROM tracked_threads
                """
            ).fetchall()
        if len(rows) < 2:
            return 0

        buckets: Dict[Tuple[str, str], List[Tuple[int, str, str]]] = {}
        for r in rows:
            tid = str(r[0] or "").strip()
            mtg = str(r[1] or "").strip().lower()
            norm = str(r[2] or "").strip().lower()
            if not tid or not norm:
                continue
            msg_cnt = int(r[3] or 0)
            last_dt = str(r[4] or "")
            buckets.setdefault((mtg, norm), []).append((msg_cnt, last_dt, tid))

        merged_count = 0
        for grp in buckets.values():
            if len(grp) <= 1:
                continue
            grp_sorted = sorted(grp, key=lambda x: (x[0], x[1]), reverse=True)
            primary_id = grp_sorted[0][2]
            other_ids = [x[2] for x in grp_sorted[1:]]
            db.merge_threads(primary_id, other_ids)
            merged_count += len(other_ids)
        return merged_count
    except Exception as exc:
        logger.warning("Tracked threads table deduplication skipped cleanly: %s", exc)
        return 0
