# Core Network Client (`core.network.client`)

Acyclic modular package extracted from `core/network/session.py` (thin backward-compatible re-export shim):

- `errors.py`: `NetworkError`, `HttpError`, `NetworkTimeoutError`, `DownloadCancelledError`.
- `proxy_and_config.py`: `CONFIG_PATH`, `DEFAULT_UAS`, `ProxyProfileManager` (DPAPI encrypted persistence), `HumannessConfig`, and `NetworkConfigDialog`.
- `ai_session.py`: `is_local_address()` (RFC 1918 / loopback detection) and `get_ai_session()` (proxy bypass for Ollama).
- `network_session.py`: `NetworkSession` singleton connection pool, retry strategy, proxy updates, and atomic streaming downloads.
