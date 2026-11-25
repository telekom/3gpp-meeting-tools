"""FTP WG folder listing and individual Docs/ directory scanning."""

from __future__ import annotations

import logging
import re
from typing import Optional
from urllib.parse import unquote, urljoin

from bs4 import BeautifulSoup

from core.network.session import NetworkSession
from .sources import (
    FILE_EXTENSION_PATTERN,
    HREF_PATTERN,
    MEETING_DENY_LIST,
    TDOC_LINK_PATTERN,
    extract_fallback_date_from_folder,
    extract_meeting_number,
    is_meeting_name,
)

logger = logging.getLogger(__name__)


def fetch_wg_directories(
    wg_name: str,
    ftp_base_url: str,
    is_ah_folder: bool = False,
    target_meetings: Optional[list] = None,
    href_pattern=HREF_PATTERN,
    deny_list=MEETING_DENY_LIST,
    is_meeting_fn=is_meeting_name,
    extract_num_fn=extract_meeting_number,
) -> list:
    """Fetch and filter candidate meeting directory tasks for a WG FTP URL."""
    meeting_tasks = []
    try:
        logger.info(f"🌐 Parsing {ftp_base_url}")
        html = NetworkSession.get_html(ftp_base_url)

        hrefs = href_pattern.findall(html)
        if not hrefs:
            logger.warning(f"⚠️ [Debug] NO links found for {ftp_base_url.split('/')[-2]}.")
            return meeting_tasks

        for href in hrefs:
            if ".." in href or "?" in href:
                continue
            folder_name = href.strip("/").split("/")[-1]
            if folder_name in ["..", ".", ""]:
                continue

            if folder_name.lower() in deny_list:
                continue
            if FILE_EXTENSION_PATTERN.search(folder_name):
                continue

            if not is_ah_folder and not is_meeting_fn(folder_name):
                continue

            absolute_url = urljoin(ftp_base_url, href)
            meeting_num = extract_num_fn(folder_name)

            if target_meetings and not any(
                t["wg"] == wg_name and t["meeting"] == meeting_num for t in target_meetings
            ):
                continue

            url_key = absolute_url.split("ftp/", 1)[-1] if "ftp/" in absolute_url else absolute_url
            docs_url = urljoin(absolute_url, "Docs/") if absolute_url.endswith("/") else f"{absolute_url}/Docs/"

            meeting_tasks.append(
                {
                    "wg_name": wg_name,
                    "folder_name": folder_name,
                    "meeting_num": meeting_num,
                    "url_key": url_key,
                    "absolute_url": absolute_url,
                    "is_ad_hoc": is_ah_folder,
                    "docs_url": docs_url,
                    "fallback_start_date": extract_fallback_date_from_folder(folder_name),
                }
            )

        logger.info(f"✅ {wg_name}: Found {len(meeting_tasks)} meeting folders in {ftp_base_url.split('/')[-2]}.")

    except Exception as e:
        logger.warning(f"⚠️ Directory Fetch Error for {ftp_base_url}: {e}")
    return meeting_tasks


def _fetch_and_parse_docs_listing(url: str, task: dict, tdoc_pattern=TDOC_LINK_PATTERN) -> list:
    """Download one Docs/ directory HTML page and extract sorted TDoc descriptors."""
    logger.info(
        f"🔎 [Docs Scan] {task.get('wg_name', '')} #{task.get('meeting_num', '')}: {url}"
    )
    html = NetworkSession.get_html(url, timeout=8)

    # Prefer the actual href target. The current 3GPP directory renderer does
    # not guarantee that the visible anchor text is the literal filename.
    soup = BeautifulSoup(html, "html.parser")
    filenames = []
    for anchor in soup.find_all("a", href=True):
        href = unquote(str(anchor.get("href", ""))).replace("\\", "/")
        filename = href.split("?")[0].rstrip("/").split("/")[-1].strip()
        if re.search(r"\.(?:zip|doc|docx|pdf)$", filename, re.IGNORECASE):
            filenames.append(filename)

    # Compatibility fallback for older/simple directory listings where the
    # filename is exposed only as anchor text.
    if not filenames:
        filenames = tdoc_pattern.findall(html)

    parsed_by_name = {}
    for filename in filenames:
        clean_name = re.sub(r"\.(?:zip|doc|docx|pdf)$", "", filename.strip(), flags=re.IGNORECASE)
        match = re.match(r"^([A-Za-z0-9]+)-?(\d+)", clean_name)
        if not match:
            continue
        parsed_by_name[clean_name.upper()] = {
            "clean": clean_name,
            "prefix": match.group(1).upper(),
            "num": int(match.group(2)),
        }

    parsed = list(parsed_by_name.values())
    parsed.sort(key=lambda x: (x["num"], x["clean"].lower()))

    if parsed:
        logger.info(
            f"✅ [Docs Scan] {task.get('wg_name', '')} #{task.get('meeting_num', '')}: "
            f"found {len(parsed)} TDoc file(s), range {parsed[0]['clean']} -> {parsed[-1]['clean']}"
        )
    else:
        logger.warning(
            f"⚠️ [Docs Scan] {task.get('wg_name', '')} #{task.get('meeting_num', '')}: "
            f"directory loaded but no TDoc filenames were recognized at {url}"
        )
    return parsed


def process_individual_meeting(task: dict, tdoc_pattern=TDOC_LINK_PATTERN) -> tuple:
    """Scan a single meeting Docs/ directory and return `(docs_data, tdoc_count)`."""
    # IMPORTANT: do not use urljoin(base, "Docs/") here. Phase-2 tasks loaded
    # from the DB provide absolute_url without a trailing slash. urljoin would
    # then treat the meeting folder as a file and replace it, e.g.
    # .../WG3_Security/TSGS3_130_Prague + Docs/ -> .../WG3_Security/Docs/.
    # Append Docs/ explicitly so the meeting-specific folder is preserved.
    explicit_docs_url = str(task.get("docs_url") or "").strip()
    if explicit_docs_url:
        base_docs_url = explicit_docs_url
    else:
        base_docs_url = task["absolute_url"].rstrip("/") + "/Docs/"
    first_tdoc, first_pfx, first_num = "", "", 0
    last_tdoc, last_pfx, last_num = "", "", 0
    tdoc_count = 0
    final_docs_url = base_docs_url

    parsed_list = []
    try:
        parsed_list = _fetch_and_parse_docs_listing(base_docs_url, task, tdoc_pattern=tdoc_pattern)
    except Exception as exc:
        logger.warning(
            f"⚠️ [Docs Scan] {task.get('wg_name', '')} #{task.get('meeting_num', '')}: "
            f"failed to read {base_docs_url}: {exc}"
        )

    # Some historical 3GPP folders use lowercase 'docs'. Try that form only
    # when the canonical URL did not yield a usable listing.
    if not parsed_list and re.search(r"/Docs/?$", base_docs_url, re.IGNORECASE):
        fallback_url = re.sub(r"/Docs/?$", "/docs/", base_docs_url, flags=re.IGNORECASE)
        if fallback_url != base_docs_url:
            try:
                fallback_list = _fetch_and_parse_docs_listing(
                    fallback_url, task, tdoc_pattern=tdoc_pattern
                )
                if fallback_list:
                    parsed_list = fallback_list
                    final_docs_url = fallback_url
            except Exception as exc:
                logger.warning(
                    f"⚠️ [Docs Scan] {task.get('wg_name', '')} #{task.get('meeting_num', '')}: "
                    f"lowercase fallback failed at {fallback_url}: {exc}"
                )

    if parsed_list:
        tdoc_count = len(parsed_list)
        first_tdoc, first_pfx, first_num = (
            parsed_list[0]["clean"],
            parsed_list[0]["prefix"],
            parsed_list[0]["num"],
        )
        last_tdoc, last_pfx, last_num = (
            parsed_list[-1]["clean"],
            parsed_list[-1]["prefix"],
            parsed_list[-1]["num"],
        )

    docs_data = (
        final_docs_url,
        first_tdoc,
        first_pfx,
        first_num,
        last_tdoc,
        last_pfx,
        last_num,
        task["url_key"],
    )
    return docs_data, tdoc_count
