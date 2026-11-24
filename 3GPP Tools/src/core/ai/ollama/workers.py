"""Background QThreads for Ollama service lifecycle, diagnostics, and heartbeat monitoring."""

from __future__ import annotations

import logging
import time
from typing import Optional, Tuple

from PyQt5.QtCore import QObject, QThread, pyqtSignal

from .client import OllamaClient
from .config import load_ollama_config, save_ollama_config
from .service import start_ollama_service, stop_ollama_service


class OllamaServiceWorker(QThread):
    """
    Background worker thread to execute service start/stop/restart operations
    without stalling the Qt GUI event loop.
    """

    finished = pyqtSignal(bool, str)
    progress_status = pyqtSignal(str)

    def __init__(self, action: str, custom_path: str = "", host: str = "", parent: QObject = None):
        super().__init__(parent)
        self.action = action.lower()
        self.custom_path = custom_path
        self.host = host

    def run(self):
        if self.action == "start":
            self.progress_status.emit("⏳ Starting Ollama service...")
            success, msg = start_ollama_service(self.custom_path, self.host)
        elif self.action == "stop":
            self.progress_status.emit("⏳ Stopping Ollama service...")
            success, msg = stop_ollama_service(self.host)
        elif self.action == "restart":
            self.progress_status.emit("⏳ Stopping Ollama service...")
            stop_ollama_service(self.host)
            time.sleep(0.5)
            self.progress_status.emit("⏳ Starting Ollama service...")
            success, msg = start_ollama_service(self.custom_path, self.host)
        else:
            success, msg = False, f"Unknown action: {self.action}"

        self.finished.emit(success, msg)


class OllamaDiagnosticsWorker(QThread):
    """Fetches Ollama runtime/model diagnostics without blocking the Qt GUI."""

    finished = pyqtSignal(dict)

    def __init__(self, client, parent: QObject = None):
        super().__init__(parent)
        self.client = client

    def run(self):
        try:
            self.finished.emit(self.client.get_diagnostics())
        except Exception as exc:
            logging.exception("[Ollama] Diagnostics worker failed")
            self.finished.emit(
                {
                    "online": False,
                    "error": str(exc),
                    "version": "",
                    "models": [],
                    "running_models": [],
                }
            )


class OllamaMonitorThread(QThread):
    """
    Non-blocking background thread that periodically monitors Ollama health
    and notifies the UI only when connectivity or model lists change.
    """

    status_updated = pyqtSignal(bool, str, list, str)

    def __init__(self, client: OllamaClient = None, parent: QObject = None):
        super().__init__(parent)
        self.client = client or OllamaClient()
        self._running = True
        self._force_check = False
        self._last_state: Tuple[Optional[bool], str, int, str] = (None, "", -1, "")

    def trigger_check(self):
        """Forces an immediate heartbeat check without waiting for the sleep interval."""
        self._force_check = True

    def run(self):
        logging.info("🦙 [Ollama Monitor] Worker thread started running.")

        while self._running:
            try:
                cfg = load_ollama_config()
                check_interval = max(5, int(cfg.get("check_interval", 10)))
                selected_model = str(cfg.get("selected_model") or "")

                is_online, models, err = self.client.ping_and_get_models()

                # If selected model is missing but models exist, pick the first
                if is_online and models and not selected_model:
                    selected_model = str(models[0])
                    cfg["selected_model"] = selected_model
                    save_ollama_config(cfg)

                current_state = (bool(is_online), str(selected_model or ""), len(models), str(err or ""))

                if current_state != self._last_state:
                    self._last_state = current_state
                    logging.info(
                        f"🦙 [Ollama Monitor] Status changed: online={is_online}, "
                        f"models={len(models)}, active='{selected_model}', err='{err}'"
                    )

                self.status_updated.emit(
                    bool(is_online),
                    str(selected_model or ""),
                    list(models or []),
                    str(err or ""),
                )

            except Exception as e:
                logging.error(f"❌ [Ollama Monitor] Loop exception: {e}", exc_info=True)
                if self._running:
                    self.status_updated.emit(False, "", [], f"Monitor error: {e}")

            # Sleep in 250ms intervals to allow immediate shutdown or forced update
            for _ in range(check_interval * 4):
                if not self._running or self._force_check:
                    self._force_check = False
                    break
                self.msleep(250)

        logging.info("🦙 [Ollama Monitor] Worker thread cleanly exited run loop.")

    def stop(self):
        """Stops the polling thread gracefully on application exit."""
        self._running = False
        # run() does not use a Qt event loop, so quit() alone cannot stop it.
        # The worker checks _running every 250 ms; allow the current bounded
        # network probe to finish rather than killing native code mid-call.
        self.quit()
        if not self.wait(3500):
            logging.warning(
                "Ollama monitor did not stop within 3.5 s; leaving it to exit "
                "cooperatively instead of force-terminating the QThread."
            )
