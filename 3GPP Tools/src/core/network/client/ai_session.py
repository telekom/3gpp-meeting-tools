"""Local-network detection and proxy-bypass session builder for Ollama / LLM traffic."""

from __future__ import annotations

import ipaddress
import urllib.parse

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from core.network.client.proxy_and_config import ProxyProfileManager
from core.utils.utils import get_proxies


def is_local_address(url_or_host: str) -> bool:
    """
    Returns True if the destination URL or hostname is localhost,
    a loopback IP, or a private RFC 1918 / LAN address.
    """
    if not url_or_host:
        return True
    try:
        if "://" in url_or_host:
            hostname = urllib.parse.urlsplit(url_or_host).hostname or ""
        else:
            hostname = url_or_host.split(":")[0]

        hostname = hostname.strip().lower()

        if hostname in ("localhost", "127.0.0.1", "::1", ""):
            return True
        if hostname.endswith(".local"):
            return True

        ip = ipaddress.ip_address(hostname)
        return ip.is_loopback or ip.is_private or ip.is_link_local
    except ValueError:
        return False


def _make_proxy_uri(raw_host: str, user: str, password: str) -> str:
    if not raw_host:
        return ""
    scheme, netloc = raw_host.split("://", 1) if "://" in raw_host else ("http", raw_host)
    if user:
        safe_u = urllib.parse.quote(user, safe="")
        auth = f"{safe_u}:{urllib.parse.quote(password, safe='')}@" if password else f"{safe_u}@"
        return f"{scheme}://{auth}{netloc}"
    return f"{scheme}://{netloc}"


def get_ai_session(target_url: str = "http://localhost:11434", proxy_mode: str = "direct") -> requests.Session:
    """
    Creates a dedicated requests.Session tailored for Ollama / LLM tasks.
    Bypasses scraper humanness delays and corporate proxy locks for local traffic.

    Args:
        target_url: The Ollama server endpoint (e.g. http://localhost:11434)
        proxy_mode: 'direct' (bypass proxy), 'auto' (direct for LAN, proxy for WAN),
                    or 'app_proxy' (enforce app proxy settings).
    """
    session = requests.Session()

    adapter = HTTPAdapter(
        pool_connections=5,
        pool_maxsize=10,
        max_retries=Retry(total=1, backoff_factor=0.2, allowed_methods=["GET", "POST"]),
    )
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    should_bypass_proxy = True
    if proxy_mode == "app_proxy":
        should_bypass_proxy = False
    elif proxy_mode == "auto":
        should_bypass_proxy = is_local_address(target_url)
    else:
        should_bypass_proxy = True

    if should_bypass_proxy:
        session.trust_env = False
        session.proxies = {"http": None, "https": None}
    else:
        profile = ProxyProfileManager.load_profile()
        if profile.get("enabled"):
            http_host = profile.get("http_host", "").strip()
            https_host = profile.get("https_host", "").strip()
            user = profile.get("username", "").strip()
            password = profile.get("password", "")
            active_proxies = {}
            if http_host:
                active_proxies["http"] = _make_proxy_uri(http_host, user, password)
            if https_host:
                active_proxies["https"] = _make_proxy_uri(https_host, user, password)
            session.proxies = active_proxies
        else:
            session.proxies = dict(get_proxies())

    return session
