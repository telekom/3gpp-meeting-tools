# --- File: core/network/session.py ---
import logging
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from typing import Optional, Dict

from core.utils.utils import get_proxies

class NetworkSession:
    _instance: Optional[requests.Session] = None

    @classmethod
    def get_instance(cls) -> requests.Session:
        """Returns a thread-safe, persistent HTTP session with retry logic and proxies."""
        if cls._instance is None:
            cls._instance = cls._create_session()
        return cls._instance

    @classmethod
    def get_html(cls, url: str, timeout: int = 20) -> str:
        """Fetches HTML using the shared requests session."""
        session = cls.get_instance()
        response: requests.Response = session.get(url, timeout=timeout)
        response.raise_for_status()
        return response.text

    @classmethod
    def update_proxies(cls, proxies: Dict[str, str]) -> None:
        """Dynamically updates the proxies for the running global session."""
        session = cls.get_instance()
        session.proxies.clear()
        session.proxies.update(proxies)
        logging.info(f"🌐 Global Network Session proxies updated: {proxies}")

    @staticmethod
    def test_connection(proxies: Dict[str, str], test_url: str = "https://www.3gpp.org") -> bool:
        """
        Tests a proxy configuration using a temporary session.
        This prevents modifying the global session if the test fails.
        """
        try:
            test_session = requests.Session()
            test_session.trust_env = False
            test_session.proxies.update(proxies)
            # Use a quick 10-second timeout to keep the UI responsive
            response: requests.Response = test_session.get(test_url, timeout=10)
            return response.status_code == 200
        except Exception as e:
            logging.warning(f"Proxy test failed: {e}")
            return False

    @staticmethod
    def _create_session() -> requests.Session:
        session = requests.Session()
        session.trust_env = False  # Ignore system env vars to prevent conflicts

        # Inject proxies configured in your app's UI
        session.proxies.update(get_proxies())

        # Spoof a standard browser to avoid firewall blocks
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Connection': 'keep-alive'
        })

        # Robust Retry Strategy for anti-bot protection
        retry_strategy = Retry(
            total=5,
            backoff_factor=1,
            status_forcelist=[403, 429, 500, 502, 503, 504],
            allowed_methods=["GET"]
        )
        # ---> UPGRADE: Allow up to 20 concurrent Keep-Alive connections
        adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=20, pool_maxsize=20)
        session.mount("https://", adapter)
        session.mount("http://", adapter)

        return session