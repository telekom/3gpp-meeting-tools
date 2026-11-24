"""Backward-compatible re-export shim for centralized HTTP and AI sessions.

Canonical implementation lives in ``core.network.client``.
"""

from core.network.client import (
    CONFIG_PATH,
    DEFAULT_UAS,
    DownloadCancelledError,
    HttpError,
    HumannessConfig,
    NetworkConfigDialog,
    NetworkError,
    NetworkSession,
    NetworkTimeoutError,
    ProxyProfileManager,
    get_ai_session,
    is_local_address,
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
