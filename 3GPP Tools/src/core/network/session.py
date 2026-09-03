# --- File: src/core/network/session.py ---
import json
import logging
import random
import threading
import time
from pathlib import Path
from typing import Any, Callable, Dict, Optional, Union

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFormLayout,
    QLineEdit,
    QVBoxLayout,
)

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
# --- HUMANNESS CONFIGURATION ---
# ==========================================
CONFIG_PATH = Path.home() / "3GPP_Tools" / "network_config.json"

DEFAULT_UAS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15",
]


class HumannessConfig:
    _cached_cfg = None
    _cfg_lock = threading.Lock()

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
    _lock = threading.Lock()

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

    @classmethod
    def update_proxies(cls, proxies: Dict[str, str]) -> None:
        """Atomically updates proxies for the global session."""
        with cls._lock:
            session = cls.get_instance()
            session.proxies = dict(proxies)
        logging.info(f"🌐 Global Network Session proxies updated: {proxies}")

    @staticmethod
    def test_connection(
        proxies: Dict[str, str], test_url: str = "https://www.3gpp.org"
    ) -> bool:
        """Tests a proxy configuration using a temporary throwaway session."""
        try:
            with requests.Session() as test_session:
                test_session.trust_env = False
                test_session.proxies = dict(proxies)
                response = test_session.get(test_url, timeout=10)
                return response.status_code == 200
        except Exception as e:
            logging.warning(f"Proxy test failed: {e}")
            return False

    @classmethod
    def get_html(
        cls,
        url: str,
        timeout: int = 20,
        headers: Optional[Dict[str, str]] = None,
    ) -> str:
        """Fetches and returns the decoded string content of a web page or dynamic report."""
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
        """Fetches and returns raw binary bytes from a URL."""
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
        """Fetches and deserializes JSON content from a URL."""
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
        """Executes a lightweight HEAD request and returns response headers dictionary."""
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
        """
        Streams and downloads a file directly to disk with cancellation checks and progress reporting.

        :param url: Source URL to download from.
        :param dest_path: Final destination path on disk.
        :param timeout: Network timeout in seconds.
        :param progress_cb: Optional callback receiving integer progress percent (0-100).
        :param cancel_cb: Optional callable returning True if the download should be aborted.
        :param chunk_size: Size of byte chunks to stream from the connection.
        :param atomic: If True, writes to a .part file and renames upon success to avoid partial files.
        :return: Total count of downloaded bytes.
        """
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

            # Atomically finalize staging file
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