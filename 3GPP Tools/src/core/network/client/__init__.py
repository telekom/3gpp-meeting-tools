"""Modular HTTP client, encrypted proxy manager, and local AI session utilities."""

from core.network.client.ai_session import get_ai_session, is_local_address
from core.network.client.errors import (
    DownloadCancelledError,
    HttpError,
    NetworkError,
    NetworkTimeoutError,
)
from core.network.client.network_session import NetworkSession
from core.network.client.proxy_and_config import (
    CONFIG_PATH,
    DEFAULT_UAS,
    HumannessConfig,
    NetworkConfigDialog,
    ProxyProfileManager,
)

__all__ = [
    "CONFIG_PATH",
    "DEFAULT_UAS",
    "DownloadCancelledError",
    "HttpError",
    "HumannessConfig",
    "NetworkConfigDialog",
    "NetworkError",
    "NetworkSession",
    "NetworkTimeoutError",
    "ProxyProfileManager",
    "get_ai_session",
    "is_local_address",
]
