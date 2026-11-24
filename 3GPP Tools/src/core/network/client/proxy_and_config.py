"""Encrypted proxy profile persistence, humanness config, and settings dialog."""

from __future__ import annotations

import json
import logging
import threading

from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFormLayout,
    QLineEdit,
    QVBoxLayout,
)

from core.utils.dpapi import SecureCredentialStore
from core.utils.paths import get_project_root

CONFIG_PATH = get_project_root() / "config" / "network_config.json"

DEFAULT_UAS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15",
]


class ProxyProfileManager:
    """Manages sanitized, fully encrypted proxy profile persistence in network_config.json."""

    FORBIDDEN_PLAINTEXT_KEYS = {
        "password",
        "passwd",
        "plain_password",
        "secret",
        "pass",
        "http_host",
        "https_host",
        "username",
        "user",
        "proxy_url",
        "http_proxy",
        "https_proxy",
        "encrypted_password",
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


class HumannessConfig:
    """Thread-safe loader/saver for network_config.json humanness and proxy settings."""

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
                for k in list(data.keys()):
                    if k in ProxyProfileManager.FORBIDDEN_PLAINTEXT_KEYS:
                        data.pop(k, None)

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
        cfg = dict(HumannessConfig.load())
        cfg.update(
            {
                "min_delay": self.min_delay.value(),
                "max_delay": self.max_delay.value(),
                "randomize_ua": self.randomize_ua.isChecked(),
                "custom_ua": self.custom_ua.text().strip(),
            }
        )
        HumannessConfig.save(cfg)
        self.accept()
