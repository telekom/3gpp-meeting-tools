"""
3GPP TDoc Server URL Resolver.

Constructs fully qualified, normalized HTTPS URLs to official 3GPP server archives
for regular TDocs and revisions.
"""
from __future__ import annotations

import re
from typing import Any, Dict, Optional


def normalize_3gpp_url(url: Optional[str]) -> str:
    """Normalizes any 3GPP FTP/HTTP URL to modern HTTPS format."""
    if not url:
        return ""
    clean = str(url).strip()
    if not clean:
        return ""
    # Convert ftp://ftp.3gpp.org to https://www.3gpp.org/ftp
    clean = re.sub(r"^ftp://ftp\.3gpp\.org/?", "https://www.3gpp.org/ftp/", clean, flags=re.IGNORECASE)
    clean = re.sub(r"^http://www\.3gpp\.org/", "https://www.3gpp.org/", clean, flags=re.IGNORECASE)
    if not clean.startswith("http"):
        clean = "https://www.3gpp.org/ftp/" + clean.lstrip("/")
    return clean.rstrip("/")


def is_tdoc_revision(tdoc_id: str) -> bool:
    """Checks whether a TDoc identifier represents a revision (e.g. S2-2401234r01, rev2, r1)."""
    if not tdoc_id:
        return False
    return bool(re.search(r"(?:r|rev)\d{1,2}[a-zA-Z]?$", str(tdoc_id).strip(), re.IGNORECASE))


def resolve_tdoc_download_url(
    tdoc_id: str,
    docs_url: Optional[str] = None,
    revisions_url: Optional[str] = None,
    main_url: Optional[str] = None,
    meta: Optional[Dict[str, Any]] = None,
) -> str:
    """Constructs the fully qualified 3GPP HTTPS download URL for a TDoc ZIP archive.

    Intelligently routes between Docs and INBOX/Revisions folders depending on whether
    the TDoc is a base contribution or revision.

    :param tdoc_id: 3GPP TDoc identifier (e.g. 'S2-2407123', 'S2-2407123r01').
    :param docs_url: Meeting documents base URL (e.g. '.../TSGS2_164_Maastricht/Docs').
    :param revisions_url: Meeting revisions base URL (e.g. '.../INBOX/Revisions').
    :param main_url: Meeting root URL (e.g. '.../TSGS2_164_Maastricht').
    :param meta: Optional row metadata dict from which URLs can be inferred.
    :return: Direct HTTPS download URL to the .zip archive on 3GPP servers, or '' if unresolvable.
    """
    clean_id = str(tdoc_id or "").strip()
    if not clean_id:
        return ""

    # 1. If meta contains an explicit download URL already, validate and normalize it
    if meta:
        for k in ("3gpp_url", "3GPP URL", "tdoc_url", "download_url"):
            val = str(meta.get(k) or "").strip()
            if val and val.startswith("http"):
                return val

        # Infer candidate folder URLs from meta fields if not explicitly supplied
        if not docs_url:
            docs_url = meta.get("docs_folder_url") or meta.get("docs_ftp_url") or meta.get("docs_url")
        if not revisions_url:
            revisions_url = meta.get("revisions_url")
        if not main_url:
            main_url = meta.get("main_ftp_url") or meta.get("url_key")

    norm_docs = normalize_3gpp_url(docs_url)
    norm_revs = normalize_3gpp_url(revisions_url)
    norm_main = normalize_3gpp_url(main_url)

    is_rev = is_tdoc_revision(clean_id)

    # 2. Determine appropriate target folder
    target_folder = ""
    if is_rev:
        if norm_revs:
            target_folder = norm_revs
        elif norm_main:
            target_folder = f"{norm_main}/INBOX/Revisions"
        elif norm_docs:
            if re.search(r"/[Dd]ocs/?$", norm_docs):
                meeting_base = re.sub(r"/[Dd]ocs/?$", "", norm_docs)
                target_folder = f"{meeting_base}/INBOX/Revisions"
            else:
                target_folder = norm_docs
    else:
        if norm_docs:
            target_folder = norm_docs
        elif norm_main:
            target_folder = f"{norm_main}/Docs"
        elif norm_revs:
            if re.search(r"/[Ii][Nn][Bb][Oo][Xx]/[Rr]evisions/?$", norm_revs):
                meeting_base = re.sub(r"/[Ii][Nn][Bb][Oo][Xx]/[Rr]evisions/?$", "", norm_revs)
                target_folder = f"{meeting_base}/Docs"
            else:
                target_folder = norm_revs

    if not target_folder:
        return ""

    # Strip existing .zip if accidentally included in clean_id
    base_id = clean_id[:-4] if clean_id.lower().endswith(".zip") else clean_id
    return f"{target_folder.rstrip('/')}/{base_id}.zip"
