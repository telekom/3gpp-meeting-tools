# --- File: src/core/network/session.py ---
import json
import logging
import random
import threading
import time
import urllib.parse
from pathlib import Path
from typing import Any, Callable, Dict, Optional, Tuple, Union

import requests
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFormLayout,
    QLineEdit,
    QVBoxLayout,
)
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from core.utils.dpapi import (
    SecureCredentialStore,
    redact_sensitive_urls,
)
from core.utils.paths import get_project_root
from core.utils.utils import get_proxies


# ==========================================
# --- NETWORK EXCEPTIONS ---
# ==========================================


class NetworkError(Exception):
    """Base exception for all NetworkSession operations."""

    pass


class HttpError(NetworkError):
    """Raised when an HTTP error status (4xx, 5xx) is received."""

    def __init__(
        self,
        message: str,
        status_code: Optional[int] = None,
        url: Optional[str] = None,
    ):
        super().__init__(message)
        self.status_code = status_code
        self.url = url


class NetworkTimeoutError(NetworkError):
    """Raised when an HTTP request or connection times out."""

    pass


class DownloadCancelledError(NetworkError):
    """Raised when a download stream is cancelled via cancellation callback."""

    pass


# ==========================================
# --- PROXY PROFILE MANAGER ---
# ==========================================


class ProxyProfileManager:
    """Manages sanitized, fully encrypted proxy profile persistence in network_config.json."""

    # Plaintext keys strictly banned from on-disk JSON storage
    FORBIDDEN_PLAINTEXT_KEYS = {
        "password", "passwd", "plain_password", "secret", "pass",
        "http_host", "https_host", "username", "user", "proxy_url",
        "http_proxy", "https_proxy", "encrypted_password"
    }

    @classmethod
    def load_profile(cls) -> dict:
        """
        Loads the proxy profile into memory.
        Decrypts all parameters (hosts, user, password) if available.
        Automatically migrates legacy plaintext configurations to encrypted storage.
        """
        cfg = HumannessConfig.load()
        profile = cfg.get("proxy_profile", {})
        enabled = bool(profile.get("enabled", False))

        # 1. Primary path: fully encrypted payload blob
        encrypted_payload = str(profile.get("encrypted_payload", "")).strip()
        if encrypted_payload:
            decrypted = SecureCredentialStore.decrypt_dict(encrypted_payload)
            return {
                "enabled": enabled,
                "http_host": str(decrypted.get("http_host", "")).strip(),
                "https_host": str(decrypted.get("https_host", "")).strip(),
                "sync_https": bool(decrypted.get("sync_https", True)),
                "username": str(decrypted.get("username", "")).strip(),
                "password": str(decrypted.get("password", "")),
            }

        # 2. Legacy Migration Path: detect older plaintext hosts/usernames
        has_legacy_keys = any(
            k in profile for k in ("http_host", "https_host", "username", "encrypted_password")
        )
        if has_legacy_keys:
            legacy_http = str(profile.get("http_host", "")).strip()
            legacy_https = str(profile.get("https_host", "")).strip()
            legacy_sync = bool(profile.get("sync_https", True))
            legacy_user = str(profile.get("username", "")).strip()
            legacy_enc_pass = str(profile.get("encrypted_password", "")).strip()
            legacy_pass = SecureCredentialStore.decrypt(legacy_enc_pass) if legacy_enc_pass else ""

            # Encrypt everything into the new payload format and wipe plaintext immediately
            migrated_dict = {
                "http_host": legacy_http,
                "https_host": legacy_https,
                "sync_https": legacy_sync,
                "username": legacy_user,
                "password": legacy_pass,
            }
            success, enc_blob = SecureCredentialStore.encrypt_dict(migrated_dict)
            if success:
                cls.save_profile({"enabled": enabled, "encrypted_payload": enc_blob})
                logging.info("🔒 Migrated legacy proxy configuration to fully encrypted storage.")

            return {
                "enabled": enabled,
                "http_host": legacy_http,
                "https_host": legacy_https,
                "sync_https": legacy_sync,
                "username": legacy_user,
                "password": legacy_pass,
            }

        # Default empty profile
        return {
            "enabled": False,
            "http_host": "",
            "https_host": "",
            "sync_https": True,
            "username": "",
            "password": "",
        }

    @classmethod
    def save_profile(cls, raw_data: dict) -> None:
        """
        Enforces a strict schema whitelist. Discards any plaintext proxy parameters
        (hosts, ports, usernames, passwords) to guarantee that only the encrypted
        payload blob is stored on disk.
        """
        clean_profile = {
            "enabled": bool(raw_data.get("enabled", False)),
            "encrypted_payload": str(raw_data.get("encrypted_payload", "")).strip(),
        }

        cfg = HumannessConfig.load()
        cfg["proxy_profile"] = clean_profile
        HumannessConfig.save(cfg)

    @classmethod
    def set_enabled(cls, enabled: bool) -> None:
        """Toggles the profile's enabled status without modifying the encrypted payload."""
        cfg = HumannessConfig.load()
        if "proxy_profile" in cfg and isinstance(cfg["proxy_profile"], dict):
            cfg["proxy_profile"]["enabled"] = bool(enabled)
            HumannessConfig.save(cfg)

    @classmethod
    def get_decrypted_password(cls) -> str:
        """Backward-compatible helper returning decrypted password string."""
        return cls.load_profile().get("password", "")


# ==========================================
# --- HUMANNESS CONFIGURATION ---
# ==========================================
CONFIG_PATH = get_project_root() / "network_config.json"

DEFAULT_UAS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15",
]


class HumannessConfig:
    _cached_cfg = None
    _cfg_lock = threading.RLock()

    @classmethod
    def load(cls) -> dict:
        with cls._cfg_lock:
            if cls._cached_cfg is not None:
                return cls._cached_cfg.copy()

            default = {
                "min_delay": 0.3,
                "max_delay": 1.2,
                "randomize_ua": True,
                "custom_ua": DEFAULT_UAS[0],
            }
            try:
                if CONFIG_PATH.exists():
                    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                        default.update(json.load(f))
            except Exception as e:
                logging.error(f"Failed to load network config: {e}")

            cls._cached_cfg = default
            return cls._cached_cfg.copy()

    @classmethod
    def save(cls, data: dict):
        with cls._cfg_lock:
            try:
                # Top-level safeguard: remove forbidden plaintext keys
                for k in list(data.keys()):
                    if k in ProxyProfileManager.FORBIDDEN_PLAINTEXT_KEYS:
                        data.pop(k, None)

                # Nested safeguard: purge any plaintext keys inside proxy_profile
                if "proxy_profile" in data and isinstance(data["proxy_profile"], dict):
                    for forbidden_key in ProxyProfileManager.FORBIDDEN_PLAINTEXT_KEYS:
                        data["proxy_profile"].pop(forbidden_key, None)

                CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
                with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=4)
                cls._cached_cfg = data.copy()
            except Exception as e:
                logging.error(f"Failed to save network config: {e}")


class NetworkConfigDialog(QDialog):
    """A UI Dialog to configure Humanness rules for network requests."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Humanness & Network Rules")
        self.setMinimumWidth(400)
        self.cfg = HumannessConfig.load()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.min_delay = QDoubleSpinBox()
        self.min_delay.setRange(0.0, 10.0)
        self.min_delay.setSingleStep(0.1)
        self.min_delay.setValue(self.cfg["min_delay"])

        self.max_delay = QDoubleSpinBox()
        self.max_delay.setRange(0.0, 20.0)
        self.max_delay.setSingleStep(0.1)
        self.max_delay.setValue(self.cfg["max_delay"])

        self.randomize_ua = QCheckBox("Rotate Modern User-Agents")
        self.randomize_ua.setChecked(self.cfg["randomize_ua"])
        self.randomize_ua.toggled.connect(self._toggle_custom_ua)

        self.custom_ua = QLineEdit(self.cfg["custom_ua"])
        self.custom_ua.setEnabled(not self.cfg["randomize_ua"])

        form.addRow("Min Delay (s):", self.min_delay)
        form.addRow("Max Delay (s):", self.max_delay)
        form.addRow("", self.randomize_ua)
        form.addRow("Custom User-Agent:", self.custom_ua)
        layout.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.save_and_accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _toggle_custom_ua(self, is_checked):
        self.custom_ua.setEnabled(not is_checked)

    def save_and_accept(self):
        HumannessConfig.save({
            "min_delay": self.min_delay.value(),
            "max_delay": self.max_delay.value(),
            "randomize_ua": self.randomize_ua.isChecked(),
            "custom_ua": self.custom_ua.text().strip(),
        })
        self.accept()


# ==========================================
# --- NETWORK SESSION ---
# ==========================================


class NetworkSession:
    """
    Centralized, thread-safe network manager for all HTTP operations,
    connection pooling, proxy management, and streaming downloads.
    """

    _instance: Optional[requests.Session] = None
    _lock = threading.RLock()

    @classmethod
    def get_instance(cls) -> requests.Session:
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = cls._create_session()
        return cls._instance

    @classmethod
    def get_request_headers(cls) -> Dict[str, str]:
        """Generates thread-isolated request headers without mutating the global session."""
        cfg = HumannessConfig.load()
        user_agent = (
            random.choice(DEFAULT_UAS)
            if cfg.get("randomize_ua", True)
            else cfg.get("custom_ua", DEFAULT_UAS[0])
        )
        return {"User-Agent": user_agent}

    @classmethod
    def apply_delay(cls) -> None:
        """Applies configured sleep delays safely on the calling thread."""
        cfg = HumannessConfig.load()
        min_d = cfg.get("min_delay", 0.0)
        max_d = cfg.get("max_delay", 0.0)
        if max_d > 0:
            delay = random.uniform(min_d, max(min_d, max_d))
            if delay > 0:
                time.sleep(delay)

    @classmethod
    def apply_humanness(cls, session: Optional[requests.Session] = None) -> None:
        """Applies delay and configures randomized headers for scraper compatibility."""
        cls.apply_delay()
        if session is not None:
            headers = cls.get_request_headers()
            with cls._lock:
                session.headers.update(headers)

    @staticmethod
    def _sanitize_proxy_url(url: str) -> str:
        """Removes user credentials from a proxy URL for safe logging."""
        if not url:
            return ""
        try:
            parsed = urllib.parse.urlsplit(url)
            if parsed.username or parsed.password:
                netloc = parsed.hostname or ""
                if parsed.port:
                    netloc += f":{parsed.port}"
                return urllib.parse.urlunsplit(
                    (parsed.scheme, netloc, parsed.path, parsed.query, parsed.fragment)
                )
            return url
        except Exception:
            return "<sanitized-proxy>"

    @classmethod
    def update_proxies(cls, proxies: Dict[str, str]) -> None:
        """
        Atomically updates proxies for the global session with sanitized logging.
        Retrieves session instance before acquiring the lock to avoid re-entrant deadlocks.
        """
        session = cls.get_instance()
        with cls._lock:
            session.proxies = dict(proxies)

        safe_proxies = {
            proto: cls._sanitize_proxy_url(url)
            for proto, url in proxies.items()
        }
        logging.info(f"🌐 Global Network Session proxies updated: {safe_proxies}")

    @staticmethod
    def test_connection(
        proxies: Dict[str, str], test_url: str = "https://www.3gpp.org"
    ) -> Tuple[bool, str]:
        """
        Tests a proxy configuration using a temporary throwaway session.
        Returns: (success: bool, sanitized_message: str)
        """
        try:
            with requests.Session() as test_session:
                test_session.trust_env = False
                test_session.proxies = dict(proxies)
                response = test_session.get(test_url, timeout=10)
                if response.status_code == 200:
                    return True, "Connection to 3GPP server verified successfully!"
                return False, f"Server returned HTTP status code {response.status_code}."
        except Exception as e:
            clean_error = redact_sensitive_urls(str(e))
            logging.warning(f"Proxy test failed: {clean_error}")
            return False, clean_error

    @classmethod
    def get_html(
        cls,
        url: str,
        timeout: int = 20,
        headers: Optional[Dict[str, str]] = None,
    ) -> str:
        session = cls.get_instance()
        cls.apply_delay()

        req_headers = cls.get_request_headers()
        if headers:
            req_headers.update(headers)

        try:
            with session.get(
                url, headers=req_headers, timeout=timeout
            ) as response:
                response.raise_for_status()
                return response.text
        except requests.exceptions.Timeout as e:
            raise NetworkTimeoutError(
                f"Request timed out after {timeout}s: {url}"
            ) from e
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            raise HttpError(
                f"HTTP {status} for {url}: {e}", status_code=status, url=url
            ) from e
        except requests.exceptions.RequestException as e:
            raise NetworkError(f"Network error fetching {url}: {e}") from e

    @classmethod
    def get_bytes(
        cls,
        url: str,
        timeout: int = 20,
        headers: Optional[Dict[str, str]] = None,
    ) -> bytes:
        session = cls.get_instance()
        cls.apply_delay()

        req_headers = cls.get_request_headers()
        if headers:
            req_headers.update(headers)

        try:
            with session.get(
                url, headers=req_headers, timeout=timeout
            ) as response:
                response.raise_for_status()
                return response.content
        except requests.exceptions.Timeout as e:
            raise NetworkTimeoutError(
                f"Request timed out after {timeout}s: {url}"
            ) from e
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            raise HttpError(
                f"HTTP {status} for {url}: {e}", status_code=status, url=url
            ) from e
        except requests.exceptions.RequestException as e:
            raise NetworkError(f"Network error fetching {url}: {e}") from e

    @classmethod
    def get_json(
        cls,
        url: str,
        timeout: int = 20,
        headers: Optional[Dict[str, str]] = None,
    ) -> Any:
        session = cls.get_instance()
        cls.apply_delay()

        req_headers = cls.get_request_headers()
        if headers:
            req_headers.update(headers)

        try:
            with session.get(
                url, headers=req_headers, timeout=timeout
            ) as response:
                response.raise_for_status()
                return response.json()
        except requests.exceptions.Timeout as e:
            raise NetworkTimeoutError(
                f"Request timed out after {timeout}s: {url}"
            ) from e
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            raise HttpError(
                f"HTTP {status} for {url}: {e}", status_code=status, url=url
            ) from e
        except requests.exceptions.RequestException as e:
            raise NetworkError(f"Network error fetching {url}: {e}") from e

    @classmethod
    def head(
        cls,
        url: str,
        timeout: int = 10,
        headers: Optional[Dict[str, str]] = None,
    ) -> Dict[str, str]:
        session = cls.get_instance()
        req_headers = cls.get_request_headers()
        if headers:
            req_headers.update(headers)

        try:
            with session.head(
                url, headers=req_headers, timeout=timeout
            ) as response:
                response.raise_for_status()
                return dict(response.headers)
        except requests.exceptions.Timeout as e:
            raise NetworkTimeoutError(
                f"HEAD request timed out after {timeout}s: {url}"
            ) from e
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            raise HttpError(
                f"HTTP {status} for {url}: {e}", status_code=status, url=url
            ) from e
        except requests.exceptions.RequestException as e:
            raise NetworkError(f"Network error on HEAD {url}: {e}") from e

    @classmethod
    def download_file(
        cls,
        url: str,
        dest_path: Union[str, Path],
        timeout: int = 30,
        progress_cb: Optional[Callable[[int], None]] = None,
        cancel_cb: Optional[Callable[[], bool]] = None,
        chunk_size: int = 16384,
        atomic: bool = True,
    ) -> int:
        dest = Path(dest_path)
        dest.parent.mkdir(parents=True, exist_ok=True)
        part_path = dest.with_suffix(dest.suffix + ".part") if atomic else dest

        if cancel_cb and cancel_cb():
            raise DownloadCancelledError(f"Download cancelled for {dest.name}")

        session = cls.get_instance()
        cls.apply_delay()
        headers = cls.get_request_headers()

        downloaded_bytes = 0

        try:
            with session.get(
                url, headers=headers, stream=True, timeout=timeout
            ) as response:
                response.raise_for_status()

                total_length = response.headers.get("content-length")
                total_bytes = (
                    int(total_length)
                    if total_length and total_length.isdigit()
                    else 0
                )

                with open(part_path, "wb") as f:
                    for chunk in response.iter_content(chunk_size=chunk_size):
                        if cancel_cb and cancel_cb():
                            f.close()
                            if atomic and part_path.exists():
                                part_path.unlink()
                            raise DownloadCancelledError(
                                f"Download cancelled: {dest.name}"
                            )

                        if chunk:
                            f.write(chunk)
                            downloaded_bytes += len(chunk)
                            if progress_cb and total_bytes > 0:
                                percent = int(
                                    (downloaded_bytes / total_bytes) * 100
                                )
                                progress_cb(percent)

            if atomic and part_path.exists():
                part_path.replace(dest)

            return downloaded_bytes

        except DownloadCancelledError:
            raise
        except requests.exceptions.Timeout as e:
            if atomic and part_path.exists():
                try:
                    part_path.unlink()
                except Exception:
                    pass
            raise NetworkTimeoutError(
                f"Download timed out after {timeout}s: {url}"
            ) from e
        except requests.exceptions.HTTPError as e:
            if atomic and part_path.exists():
                try:
                    part_path.unlink()
                except Exception:
                    pass
            status = e.response.status_code if e.response is not None else None
            raise HttpError(
                f"HTTP {status} downloading {dest.name}: {e}",
                status_code=status,
                url=url,
            ) from e
        except requests.exceptions.RequestException as e:
            if atomic and part_path.exists():
                try:
                    part_path.unlink()
                except Exception:
                    pass
            raise NetworkError(
                f"Network error downloading {dest.name}: {e}"
            ) from e
        except Exception as e:
            if atomic and part_path.exists():
                try:
                    part_path.unlink()
                except Exception:
                    pass
            raise NetworkError(
                f"Unexpected error downloading {dest.name}: {e}"
            ) from e

    @staticmethod
    def _create_session() -> requests.Session:
        session = requests.Session()
        session.trust_env = False

        profile = ProxyProfileManager.load_profile()
        if profile.get("enabled"):
            http_host = profile.get("http_host", "").strip()
            https_host = profile.get("https_host", "").strip()
            user = profile.get("username", "").strip()
            password = profile.get("password", "")

            def _make_uri(raw_host: str) -> str:
                if not raw_host:
                    return ""
                scheme, netloc = (
                    raw_host.split("://", 1)
                    if "://" in raw_host
                    else ("http", raw_host)
                )
                if user:
                    safe_u = urllib.parse.quote(user, safe="")
                    auth = (
                        f"{safe_u}:{urllib.parse.quote(password, safe='')}@"
                        if password
                        else f"{safe_u}@"
                    )
                    return f"{scheme}://{auth}{netloc}"
                return f"{scheme}://{netloc}"

            active_proxies = {}
            if http_host:
                active_proxies["http"] = _make_uri(http_host)
            if https_host:
                active_proxies["https"] = _make_uri(https_host)

            session.proxies = active_proxies
        else:
            session.proxies = dict(get_proxies())

        session.headers.update({
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        })

        retry_strategy = Retry(
            total=2,
            backoff_factor=0.5,
            status_forcelist=[429, 500, 502, 504],
            allowed_methods=["GET", "HEAD"],
        )

        adapter = HTTPAdapter(
            pool_connections=50,
            pool_maxsize=100,
            pool_block=False,
            max_retries=retry_strategy,
        )
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        return session

# ==========================================
# --- OLLAMA / AI NETWORK UTILITIES ---
# ==========================================
import ipaddress


def is_local_address(url_or_host: str) -> bool:
    """
    Returns True if the destination URL or hostname is localhost,
    a loopback IP, or a private RFC 1918 / LAN address.
    """
    if not url_or_host:
        return True
    try:
        # Extract hostname if full URL was passed
        if "://" in url_or_host:
            hostname = urllib.parse.urlsplit(url_or_host).hostname or ""
        else:
            hostname = url_or_host.split(":")[0]

        hostname = hostname.strip().lower()

        # Check standard loopback hostnames
        if hostname in ("localhost", "127.0.0.1", "::1", ""):
            return True
        if hostname.endswith(".local"):
            return True

        # Check IP ranges (127.0.0.0/8, 10.0.0.0/8, 192.168.0.0/16, 172.16.0.0/12)
        ip = ipaddress.ip_address(hostname)
        return ip.is_loopback or ip.is_private or ip.is_link_local
    except ValueError:
        # Non-IP hostname (e.g. corporate internal domain)
        return False


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

    # Mount clean adapter with no humanness delays and single-retry failover
    adapter = HTTPAdapter(
        pool_connections=5,
        pool_maxsize=10,
        max_retries=Retry(total=1, backoff_factor=0.2, allowed_methods=["GET", "POST"])
    )
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    should_bypass_proxy = True
    if proxy_mode == "app_proxy":
        should_bypass_proxy = False
    elif proxy_mode == "auto":
        should_bypass_proxy = is_local_address(target_url)
    else:  # "direct" default
        should_bypass_proxy = True

    if should_bypass_proxy:
        session.trust_env = False
        session.proxies = {"http": None, "https": None}
    else:
        # Route through application proxy profile if enabled
        profile = ProxyProfileManager.load_profile()
        if profile.get("enabled"):
            http_host = profile.get("http_host", "").strip()
            https_host = profile.get("https_host", "").strip()
            user = profile.get("username", "").strip()
            password = profile.get("password", "")

            def _make_uri(raw_host: str) -> str:
                if not raw_host:
                    return ""
                scheme, netloc = raw_host.split("://", 1) if "://" in raw_host else ("http", raw_host)
                if user:
                    safe_u = urllib.parse.quote(user, safe="")
                    auth = f"{safe_u}:{urllib.parse.quote(password, safe='')}@" if password else f"{safe_u}@"
                    return f"{scheme}://{auth}{netloc}"
                return f"{scheme}://{netloc}"

            active_proxies = {}
            if http_host:
                active_proxies["http"] = _make_uri(http_host)
            if https_host:
                active_proxies["https"] = _make_uri(https_host)
            session.proxies = active_proxies
        else:
            session.proxies = dict(get_proxies())

    return session