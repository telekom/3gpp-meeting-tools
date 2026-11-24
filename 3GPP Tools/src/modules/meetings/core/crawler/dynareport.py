"""3GPP DynaReport HTML parser for meeting metadata enrichment."""

from __future__ import annotations

import logging
import re
from typing import Optional
from urllib.parse import unquote

from bs4 import BeautifulSoup

from core.network.session import NetworkSession

logger = logging.getLogger(__name__)


def process_dynareport(
    wg_name: str,
    dyna_url: str,
    target_meetings: Optional[list] = None,
) -> list:
    """Parse meeting metadata rows from a 3GPP DynaReport page for one WG."""
    results = []
    try:
        html = NetworkSession.get_html(dyna_url)
        soup = BeautifulSoup(html, "html.parser")

        for row in soup.find_all("tr"):
            cols = row.find_all(["td", "th"])
            if len(cols) < 2:
                continue

            col_texts = []
            for td in cols:
                text = td.get_text(separator=" ", strip=True)
                text = re.sub(r"[\u2010-\u2015\u2212]", "-", text)
                col_texts.append(text)

            m_name = col_texts[0].strip()
            if not m_name or m_name.lower() == "meeting":
                continue

            mtg_id = ""
            url_key = ""
            for a in row.find_all("a", href=True):
                href = a["href"].replace("\\", "/")
                match_id = re.search(r"MtgId=(\d+)", href, re.IGNORECASE)
                if match_id:
                    mtg_id = match_id.group(1)

                if "/ftp/" in href.lower() and "tsg_" in href.lower():
                    parts = re.split(r"/ftp/", href, flags=re.IGNORECASE)
                    if len(parts) > 1:
                        path_after_ftp = unquote(parts[1].split("?")[0].strip("/"))
                        # Prefer a meeting-specific FTP path over a WG-root link.
                        # Some DynaReport rows expose more than one FTP link; the
                        # meeting folder is the deeper path and is what Phase 2 needs.
                        path_depth = len([p for p in path_after_ftp.split("/") if p])
                        current_depth = len([p for p in url_key.split("/") if p]) if url_key else -1
                        if not url_key or path_depth > current_depth:
                            url_key = path_after_ftp

            start_d, end_d, town = "", "", ""
            all_dates = []
            date_idx = -1

            for i in range(1, len(col_texts)):
                found_dates = re.findall(r"(?:19|20)\d{2}[-/]\d{2}[-/]\d{2}", col_texts[i])
                if found_dates:
                    found_dates = [d.replace("/", "-") for d in found_dates]
                    all_dates.extend(found_dates)
                    if date_idx == -1:
                        date_idx = i

            if all_dates:
                start_d = min(all_dates)
                end_d = max(all_dates)

            if date_idx > 0:
                candidate = col_texts[date_idx - 1].strip()
                if candidate and not re.search(
                    r"3GPP|\bRAN[1-6]?\b|\bSA[1-6]?\b|\bCT[1-6]?\b",
                    candidate,
                    re.IGNORECASE,
                ):
                    town = candidate

            full_num = ""
            search_text = m_name + " " + (col_texts[1] if len(col_texts) > 1 else "")
            match_explicit = re.search(r"#(\d+[a-z0-9\-]*)", search_text, re.IGNORECASE)

            if match_explicit:
                full_num = match_explicit.group(1).replace("-", "").strip().upper()
            else:
                tokens = m_name.split()
                found_token = ""
                for token in reversed(tokens):
                    if re.search(r"\d|AH", token, re.IGNORECASE):
                        found_token = token
                        break

                if not found_token and len(col_texts) > 1:
                    if not re.search(r"(?:19|20)\d{2}", col_texts[1]):
                        found_token = col_texts[1]

                if found_token:
                    clean_token = re.sub(
                        r"^(?:3GPP)?(?:R|S|C|RAN|SA|CT|SP|RP|CP)[P0-6]?\s*-?",
                        "",
                        found_token,
                        flags=re.IGNORECASE,
                    )

                    suffix = ""
                    if date_idx >= 4 and len(col_texts) > 2:
                        potential_suffix = col_texts[2].strip()
                        if len(potential_suffix) <= 6 and not re.search(
                            r"(?:19|20)\d{2}", potential_suffix
                        ):
                            suffix = potential_suffix

                    full_num = f"{clean_token}{suffix}".replace("-", "").strip().upper()

            full_num = re.sub(r"(?<!\d)0+(\d)", r"\1", full_num)

            if target_meetings and not any(
                t["wg"] == wg_name and t["meeting"] == full_num for t in target_meetings
            ):
                continue

            new_m_num = ""
            if wg_name.startswith("RAN"):
                is_adhoc = False
                if re.search(r"(?:AH|Ad\s*Hoc|Workshop|Release|Evolution)", m_name, re.IGNORECASE):
                    is_adhoc = True
                elif re.search(r"^R[P1-4]-\d", m_name, re.IGNORECASE):
                    is_adhoc = True
                elif "AH" in url_key.upper() or "AH" in full_num:
                    is_adhoc = True

                if is_adhoc:
                    new_m_num = m_name

            results.append(
                (wg_name, full_num, url_key, mtg_id, m_name, town, start_d, end_d, new_m_num)
            )

    except Exception as e:
        logger.warning(f"⚠️ DynaReport Error for {wg_name}: {e}")

    return results
