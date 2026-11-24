"""REST communication client for the local Ollama API (/api/chat, /api/tags, /api/ps)."""

from __future__ import annotations

import json
import logging
import socket
import urllib.parse
from typing import Any, Dict, List, Optional, Tuple

from core.network.session import get_ai_session
from .config import _normalize_ollama_host, load_ollama_config


class OllamaClient:
    """Core communication client for the Ollama REST API."""

    def __init__(self, host: str = None, proxy_mode: str = None):
        cfg = load_ollama_config()
        raw_host = _normalize_ollama_host(host or cfg.get("host", "http://127.0.0.1:11434"))
        self.host = raw_host
        self.proxy_mode = proxy_mode or cfg.get("proxy_mode", "direct")
        self.timeout = cfg.get("timeout", 60)
        self.keep_alive = cfg.get("keep_alive", "5m")
        self.session = get_ai_session(self.host, self.proxy_mode)
        # Optional observer used by diagnostic UIs. Called from the requesting thread.
        self.exchange_callback = None

    def reconfigure(self, host: str, proxy_mode: str):
        """Re-initializes session when user changes host or proxy settings."""
        self.host = _normalize_ollama_host(host)
        self.proxy_mode = proxy_mode
        cfg = load_ollama_config()
        self.timeout = cfg.get("timeout", 60)
        self.keep_alive = cfg.get("keep_alive", "5m")
        self.session = get_ai_session(self.host, self.proxy_mode)

    def chat(
        self,
        model: str,
        messages: List[Dict[str, Any]],
        tools: Optional[List[Dict[str, Any]]] = None,
        options: Optional[Dict[str, Any]] = None,
        keep_alive: Optional[str] = None,
        timeout: Optional[float] = None,
    ) -> Dict[str, Any]:
        """Executes a complete non-streaming /api/chat request.

        Preserves the complete Ollama response, including native tool calls and
        timing/token metadata. Hardware selection belongs to the Ollama daemon,
        not to individual chat requests.
        """
        model = str(model or "").strip()
        if not model:
            raise ValueError("Ollama chat requires a model name.")
        if not isinstance(messages, list) or not messages:
            raise ValueError("Ollama chat requires at least one message.")

        payload: Dict[str, Any] = {
            "model": model,
            "messages": messages,
            "stream": False,
            "options": options or {"temperature": 0.2},
            "keep_alive": keep_alive if keep_alive is not None else self.keep_alive,
        }
        if tools:
            payload["tools"] = tools

        request_timeout = float(timeout if timeout is not None else self.timeout)
        endpoint = f"{self.host}/api/chat"
        trace = self.exchange_callback
        if callable(trace):
            trace(
                {
                    "direction": "request",
                    "method": "POST",
                    "endpoint": endpoint,
                    "payload": payload,
                }
            )

        with self.session.post(
            endpoint,
            json=payload,
            timeout=(5.0, request_timeout),
        ) as resp:
            if callable(trace):
                trace(
                    {
                        "direction": "response_headers",
                        "status_code": resp.status_code,
                        "reason": resp.reason,
                        "headers": dict(resp.headers),
                    }
                )
            resp.raise_for_status()
            data = resp.json()
            if callable(trace):
                trace({"direction": "response", "payload": data})

        if not isinstance(data, dict):
            raise ValueError("Ollama /api/chat returned a non-object JSON response.")
        return data

    @staticmethod
    def _format_bytes(value: Any) -> str:
        """Formats Ollama byte counts for compact diagnostics UI display."""
        try:
            size = float(value or 0)
        except (TypeError, ValueError):
            return "-"
        if size <= 0:
            return "-"
        units = ("B", "KB", "MB", "GB", "TB")
        unit = 0
        while size >= 1024.0 and unit < len(units) - 1:
            size /= 1024.0
            unit += 1
        return f"{size:.1f} {units[unit]}" if unit >= 2 else f"{size:.0f} {units[unit]}"

    @staticmethod
    def _processor_summary(size: Any, size_vram: Any) -> str:
        """Derives Ollama's CPU/GPU split from /api/ps memory placement."""
        try:
            total = float(size or 0)
            vram = max(0.0, float(size_vram or 0))
        except (TypeError, ValueError):
            return "Unknown"
        if total <= 0:
            return "Unknown"
        gpu_pct = max(0, min(100, round((vram / total) * 100)))
        cpu_pct = 100 - gpu_pct
        if gpu_pct >= 99:
            return "100% GPU"
        if gpu_pct <= 1:
            return "100% CPU"
        return f"{gpu_pct}% GPU / {cpu_pct}% CPU"

    def get_diagnostics(self) -> Dict[str, Any]:
        """Returns version, installed-model metadata, and current runtime placement.

        /api/tags supplies installed model metadata. /api/ps supplies models currently
        loaded in memory, including size_vram and context_length. The GPU/CPU split is
        derived from size_vram / size and therefore describes observed runtime placement,
        not merely the configured Vulkan/iGPU preference.
        """
        result: Dict[str, Any] = {
            "online": False,
            "error": "",
            "version": "Unknown",
            "models": [],
            "running_models": [],
        }

        # Reuse the normal connectivity path first so diagnostics fail quickly offline.
        online, _, err = self.ping_and_get_models()
        if not online:
            result["error"] = err
            return result
        result["online"] = True

        # Version is informational: older Ollama builds may not expose this endpoint.
        try:
            resp = self.session.get(f"{self.host}/api/version", timeout=2.0)
            if resp.status_code == 200:
                result["version"] = str(resp.json().get("version") or "Unknown")
        except Exception as exc:
            logging.debug("[Ollama] Version diagnostics unavailable: %s", exc)

        try:
            resp = self.session.get(f"{self.host}/api/tags", timeout=3.0)
            resp.raise_for_status()
            data = resp.json()
            for item in data.get("models", []):
                details = item.get("details") or {}
                result["models"].append(
                    {
                        "name": str(item.get("name") or item.get("model") or ""),
                        "size": self._format_bytes(item.get("size")),
                        "parameter_size": str(details.get("parameter_size") or "-"),
                        "quantization": str(details.get("quantization_level") or "-"),
                        "family": str(details.get("family") or "-"),
                    }
                )
        except Exception as exc:
            result["error"] = f"Installed model query failed: {exc}"
            return result

        # Runtime information is optional so diagnostics still work with older servers.
        try:
            resp = self.session.get(f"{self.host}/api/ps", timeout=3.0)
            if resp.status_code == 200:
                data = resp.json()
                for item in data.get("models", []):
                    result["running_models"].append(
                        {
                            "name": str(item.get("name") or item.get("model") or ""),
                            "processor": self._processor_summary(
                                item.get("size"), item.get("size_vram")
                            ),
                            "size": self._format_bytes(item.get("size")),
                            "vram": self._format_bytes(item.get("size_vram")),
                            "context_length": item.get("context_length") or "-",
                        }
                    )
            else:
                logging.debug("[Ollama] /api/ps unavailable: HTTP %s", resp.status_code)
        except Exception as exc:
            logging.debug("[Ollama] Runtime diagnostics unavailable: %s", exc)

        return result

    def ping_and_get_models(self) -> Tuple[bool, List[str], str]:
        """
        Fast non-blocking probe:
        - Raw socket pre-check rejects offline instances in < 1ms without HTTP overhead.
        - REST request queries /api/tags if the port is open.
        """
        parsed = urllib.parse.urlsplit(self.host)
        host_ip = parsed.hostname or "127.0.0.1"
        port = parsed.port or 11434

        # 1. Pure TCP pre-check (< 1ms when port is closed)
        try:
            with socket.create_connection((host_ip, port), timeout=0.5):
                pass
        except OSError:
            return False, [], "Ollama server is not running."
        except Exception as e:
            return False, [], f"Connection error: {e}"

        # 2. Query /api/tags
        endpoint = f"{self.host}/api/tags"
        try:
            resp = self.session.get(endpoint, timeout=2.0)
            if resp.status_code == 200:
                data = resp.json()
                models = [str(m.get("name", "")) for m in data.get("models", []) if m.get("name")]
                return True, models, ""
            return False, [], f"HTTP {resp.status_code}: {resp.reason}"
        except Exception as e:
            return False, [], f"API error: {e}"

    def stream_chat(
        self,
        model: str,
        messages: List[Dict[str, str]],
        options: Optional[Dict[str, Any]] = None,
    ):
        """
        Executes a streaming request against /api/chat.
        Yields text chunks as they arrive.
        """
        endpoint = f"{self.host}/api/chat"
        payload = {
            "model": model,
            "messages": messages,
            "stream": True,
            "options": options or {"temperature": 0.2},
            "keep_alive": self.keep_alive,
        }

        trace = self.exchange_callback
        if callable(trace):
            trace(
                {
                    "direction": "request",
                    "method": "POST",
                    "endpoint": endpoint,
                    "payload": payload,
                }
            )

        with self.session.post(
            endpoint,
            json=payload,
            stream=True,
            timeout=(5.0, float(self.timeout)),
        ) as resp:
            if callable(trace):
                trace(
                    {
                        "direction": "response_headers",
                        "status_code": resp.status_code,
                        "reason": resp.reason,
                        "headers": dict(resp.headers),
                    }
                )
            resp.raise_for_status()
            for line in resp.iter_lines():
                if not line:
                    continue
                raw_line = line.decode("utf-8", errors="replace")
                if callable(trace):
                    trace({"direction": "response_chunk", "raw": raw_line})
                try:
                    chunk = json.loads(raw_line)
                    delta = chunk.get("message", {}).get("content", "")
                    if delta:
                        yield delta
                    if chunk.get("done", False):
                        break
                except Exception as parse_err:
                    logging.warning(f"[Ollama] Stream chunk parse error: {parse_err}")
