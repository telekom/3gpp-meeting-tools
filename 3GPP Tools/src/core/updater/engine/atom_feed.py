# --- File: src/core/updater/engine/atom_feed.py ---
"""GitHub commits Atom feed fetching and XML parsing."""

from __future__ import annotations

import datetime
import html
import re
import xml.etree.ElementTree as ET
from typing import List

import requests

from core.config.config import GITHUB_COMMITS_ATOM_URL
from core.network.session import (
    HttpError,
    NetworkError,
    NetworkSession,
    NetworkTimeoutError,
)

from .models import _ATOM_NS, _SHA_RE, _TAG_RE, CommitInfo, _facade


def _parse_iso_timestamp(iso_str: str) -> float:
    raw = (iso_str or "").strip()
    if not raw:
        return 0.0
    if raw.endswith("Z"):
        raw = raw[:-1] + "+00:00"
    try:
        dt = datetime.datetime.fromisoformat(raw)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=datetime.timezone.utc)
        return dt.timestamp()
    except ValueError:
        return 0.0


def _strip_html_content(raw_html: str) -> str:
    if not raw_html:
        return ""
    text = re.sub(r"<br\s*/?>", "\n", raw_html, flags=re.IGNORECASE)
    text = _TAG_RE.sub("", text)
    text = html.unescape(text)
    lines = [line.rstrip() for line in text.strip().splitlines()]
    return "\n".join(lines).strip()


def parse_commits_atom(xml_bytes: bytes) -> List[CommitInfo]:
    """Parse GitHub's `commits/*.atom` XML feed into chronological CommitInfo items."""
    if not xml_bytes:
        return []
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        raise ValueError(f"Invalid GitHub commits Atom XML: {exc}") from exc

    commits: List[CommitInfo] = []
    for entry in root.findall("a:entry", _ATOM_NS):
        entry_id = (entry.findtext("a:id", "", _ATOM_NS) or "").strip()
        sha_match = _SHA_RE.search(entry_id)
        if not sha_match:
            continue
        sha = sha_match.group(1).lower()

        raw_title = (entry.findtext("a:title", "", _ATOM_NS) or "").strip()
        raw_content = entry.findtext("a:content", "", _ATOM_NS) or ""
        full_message = _strip_html_content(raw_content) or raw_title
        first_line = full_message.splitlines()[0].strip() if full_message else raw_title
        title = first_line or raw_title or sha[:7]

        author_name = (entry.findtext("a:author/a:name", "", _ATOM_NS) or "").strip()
        author_email = (entry.findtext("a:author/a:email", "", _ATOM_NS) or "").strip()
        author = author_name or author_email or "Maintainer"

        updated_iso = (entry.findtext("a:updated", "", _ATOM_NS) or "").strip()
        updated_ts = _parse_iso_timestamp(updated_iso)

        url = ""
        for link in entry.findall("a:link", _ATOM_NS):
            if (link.attrib.get("rel") or "alternate") == "alternate":
                url = (link.attrib.get("href") or "").strip()
                break
        if not url:
            url = f"https://github.com/telekom/3gpp-meeting-tools/commit/{sha}"

        commits.append(
            CommitInfo(
                sha=sha,
                short_sha=sha[:7],
                title=title,
                message=full_message,
                author=author,
                updated_iso=updated_iso,
                updated_ts=updated_ts,
                url=url,
            )
        )
    return commits


def fetch_remote_commits(
    timeout: int = 15,
    feed_url: str = GITHUB_COMMITS_ATOM_URL,
) -> List[CommitInfo]:
    """Fetch and parse recent commits from GitHub using NetworkSession proxies."""
    net_cls = _facade("NetworkSession", NetworkSession)
    session = net_cls.get_instance()
    headers = net_cls.get_request_headers()
    headers["Accept"] = "application/atom+xml, application/xml, text/xml;q=0.9, */*;q=0.8"

    try:
        with session.get(feed_url, headers=headers, timeout=timeout) as resp:
            resp.raise_for_status()
            return parse_commits_atom(resp.content)
    except requests.exceptions.Timeout as exc:
        raise NetworkTimeoutError(f"Timed out checking GitHub updates after {timeout}s.") from exc
    except requests.exceptions.HTTPError as exc:
        code = exc.response.status_code if exc.response is not None else None
        raise HttpError(f"GitHub returned HTTP {code} while checking updates.", status_code=code, url=feed_url) from exc
    except requests.exceptions.RequestException as exc:
        raise NetworkError(f"Network error checking GitHub updates: {exc}") from exc
