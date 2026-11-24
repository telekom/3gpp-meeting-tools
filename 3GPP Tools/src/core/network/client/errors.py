"""Network exception hierarchy for NetworkSession operations."""

from __future__ import annotations

from typing import Optional


class NetworkError(Exception):
    """Base exception for all NetworkSession operations."""


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


class DownloadCancelledError(NetworkError):
    """Raised when a download stream is cancelled via cancellation callback."""
