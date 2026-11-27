"""Network HTTP/HTTPS proxy profile dialog encrypted via Windows DPAPI."""

import urllib.parse
from typing import Optional, Tuple

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.network.session import NetworkSession, ProxyProfileManager
from core.ui.theme import C, R, T
from .styles import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.utils.dpapi import SecureCredentialStore, redact_sensitive_urls


class ProxyTestWorker(QThread):
    """Executes proxy connection testing in a separate thread to keep the UI fluid."""
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, proxies: dict, parent=None):
        super().__init__(parent)
        self.proxies = proxies

    def run(self):
        try:
            success, message = NetworkSession.test_connection(self.proxies)
            self.finished_signal.emit(success, redact_sensitive_urls(message))
        except Exception as e:
            self.finished_signal.emit(False, f"Test error: {redact_sensitive_urls(str(e))}")


class ProxyDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Network Proxy Configuration")
        self.setModal(True)
        self.resize(540, 370)

        self._test_worker: Optional[ProxyTestWorker] = None

        self._setup_ui()
        self._load_existing_settings()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        title = QLabel("🌐 Proxy Configuration & Profiles")
        title.setStyleSheet("font-size: 15px; font-weight: bold; margin-bottom: 2px;")
        layout.addWidget(title)

        desc = QLabel(
            "Configure corporate HTTP/HTTPS proxies. All proxy parameters (hosts, ports, username, "
            "and password) are encrypted locally via Windows DPAPI and are never stored in plain text."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: {T.SIZE_CAPTION}px; margin-bottom: 8px;")
        layout.addWidget(desc)

        self.enable_checkbox = QCheckBox("Activate this proxy profile")
        self.enable_checkbox.setStyleSheet(f"font-weight: bold; color: {C.TEXT_STRONG}; margin-bottom: 6px;")
        self.enable_checkbox.toggled.connect(self._on_enable_toggled)
        layout.addWidget(self.enable_checkbox)

        self.form_container = QWidget()
        form = QFormLayout(self.form_container)
        form.setContentsMargins(0, 0, 0, 0)
        form.setLabelAlignment(Qt.AlignRight)

        input_style = f"padding: 4px; border: 1px solid {C.BORDER_STRONG}; border-radius: {R.SM}px;"
        self.http_input = QLineEdit()
        self.http_input.setPlaceholderText("e.g., proxy.company.com:8080")
        self.http_input.setStyleSheet(input_style)

        self.https_input = QLineEdit()
        self.https_input.setPlaceholderText("e.g., proxy.company.com:8080")
        self.https_input.setStyleSheet(input_style)

        self.sync_checkbox = QCheckBox("Use the same proxy address for HTTPS")
        self.sync_checkbox.setChecked(True)
        self.sync_checkbox.stateChanged.connect(self._on_sync_changed)
        self.http_input.textChanged.connect(self._on_http_changed)

        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Domain\\Username or Username (Optional)")
        self.user_input.setStyleSheet(input_style)

        self.pass_input = QLineEdit()
        self.pass_input.setEchoMode(QLineEdit.Password)
        self.pass_input.setPlaceholderText("Password (Optional)")
        self.pass_input.setStyleSheet(input_style)

        self.remember_cred_checkbox = QCheckBox("Save proxy profile securely (Windows DPAPI)")
        self.remember_cred_checkbox.setChecked(True)
        self.remember_cred_checkbox.setToolTip(
            "Encrypts all proxy configuration (host, port, username, password) using OS-level DPAPI."
        )

        form.addRow("HTTP Proxy:", self.http_input)
        form.addRow("", self.sync_checkbox)
        form.addRow("HTTPS Proxy:", self.https_input)
        form.addRow("Username:", self.user_input)
        form.addRow("Password:", self.pass_input)
        form.addRow("", self.remember_cred_checkbox)
        layout.addWidget(self.form_container)

        self.status_lbl = QLabel("")
        self.status_lbl.setWordWrap(True)
        self.status_lbl.setAlignment(Qt.AlignCenter)
        self.status_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY}; margin: 4px 0;")
        layout.addWidget(self.status_lbl)

        btn_layout = QHBoxLayout()

        self.deactivate_btn = QPushButton("Deactivate (Direct)")
        self.deactivate_btn.setToolTip("Turn off proxy usage while retaining encrypted settings.")
        self.deactivate_btn.clicked.connect(self.deactivate_profile)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cancel_btn.clicked.connect(self.reject)

        self.test_btn = QPushButton("⚡ Test Connection")
        self.test_btn.clicked.connect(self.test_proxy)

        self.save_btn = QPushButton("Save && Apply")
        self.save_btn.setObjectName("primaryBtn")
        self.save_btn.clicked.connect(self.accept)

        btn_layout.addWidget(self.deactivate_btn)
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.test_btn)
        btn_layout.addWidget(self.save_btn)

        layout.addLayout(btn_layout)

    def _on_enable_toggled(self, checked: bool):
        self.form_container.setEnabled(checked)
        self.test_btn.setEnabled(checked)

    def _on_sync_changed(self, state):
        is_synced = (state == Qt.Checked)
        self.https_input.setEnabled(not is_synced)
        if is_synced:
            self.https_input.setText(self.http_input.text())

    def _on_http_changed(self, text):
        if self.sync_checkbox.isChecked():
            self.https_input.setText(text)

    def _load_existing_settings(self):
        profile = ProxyProfileManager.load_profile()
        self.enable_checkbox.setChecked(profile.get("enabled", False))
        self._on_enable_toggled(self.enable_checkbox.isChecked())

        self.http_input.setText(profile.get("http_host", ""))
        self.https_input.setText(profile.get("https_host", ""))
        self.sync_checkbox.setChecked(profile.get("sync_https", True))
        self.user_input.setText(profile.get("username", ""))
        self.pass_input.setText(profile.get("password", ""))

        has_data = bool(
            profile.get("http_host") or profile.get("https_host") or profile.get("username")
        )
        self.remember_cred_checkbox.setChecked(has_data or True)

    def _build_proxy_uri(self, raw_address: str) -> str:
        raw_address = raw_address.strip()
        if not raw_address:
            return ""

        scheme, netloc = (
            raw_address.split("://", 1)
            if "://" in raw_address
            else ("http", raw_address)
        )
        username = self.user_input.text().strip()
        password = self.pass_input.text().strip()

        if username:
            safe_user = urllib.parse.quote(username, safe="")
            if password:
                safe_pass = urllib.parse.quote(password, safe="")
                auth = f"{safe_user}:{safe_pass}@"
            else:
                auth = f"{safe_user}@"
            return f"{scheme}://{auth}{netloc}"

        return f"{scheme}://{netloc}"

    def build_proxies_dict(self) -> dict:
        proxies = {}
        http_raw = self.http_input.text().strip()
        https_raw = self.https_input.text().strip()

        if http_raw:
            proxies["http"] = self._build_proxy_uri(http_raw)
        if https_raw:
            proxies["https"] = self._build_proxy_uri(https_raw)

        return proxies

    def test_proxy(self):
        proxies = self.build_proxies_dict()

        self.test_btn.setEnabled(False)
        self.test_btn.setText("Testing...")
        self.save_btn.setEnabled(False)
        self.status_lbl.setText("⏳ Connecting to www.3gpp.org through proxy...")
        self.status_lbl.setStyleSheet(f"color: {C.ACCENT_SKY}; font-size: {T.SIZE_CAPTION}px;")

        self._test_worker = ProxyTestWorker(proxies, parent=self)
        self._test_worker.finished_signal.connect(self._on_test_finished)
        self._test_worker.start()

    def _on_test_finished(self, success: bool, message: str):
        self.test_btn.setEnabled(True)
        self.test_btn.setText("⚡ Test Connection")
        self.save_btn.setEnabled(True)

        if success:
            self.status_lbl.setText("✅ Connection verified successfully!")
            self.status_lbl.setStyleSheet(f"color: {C.SUCCESS_DARK}; font-size: {T.SIZE_CAPTION}px;")
            QMessageBox.information(self, "Proxy Success", message)
        else:
            self.status_lbl.setText("❌ Connection failed.")
            self.status_lbl.setStyleSheet(f"color: {C.DANGER_ALT}; font-size: {T.SIZE_CAPTION}px;")
            QMessageBox.warning(self, "Proxy Test Failed", message)

    def deactivate_profile(self):
        """Turns off proxy routing while preserving the encrypted configuration on disk."""
        ProxyProfileManager.set_enabled(False)
        NetworkSession.update_proxies({})
        super().accept()

    def accept(self):
        """
        Validates encryption before persistence.
        Guarantees that unencrypted proxy information (hosts, ports, usernames, passwords)
        is NEVER written to disk, updates NetworkSession, and dismisses the dialog cleanly.
        """
        is_enabled = self.enable_checkbox.isChecked()
        http_host = self.http_input.text().strip()
        https_host = self.https_input.text().strip()
        sync_https = self.sync_checkbox.isChecked()
        username = self.user_input.text().strip()
        password = self.pass_input.text()

        encrypted_payload = ""

        if self.remember_cred_checkbox.isChecked() and (http_host or https_host):
            payload_dict = {
                "http_host": http_host,
                "https_host": https_host,
                "sync_https": sync_https,
                "username": username,
                "password": password,
            }
            success, enc_result = SecureCredentialStore.encrypt_dict(payload_dict)
            if not success:
                QMessageBox.warning(
                    self,
                    "Security Notice",
                    "Windows DPAPI encryption is unavailable or encountered an error.\n\n"
                    "Your proxy settings were NOT written to disk to prevent exposing corporate network data. "
                    "They will remain active in memory for the current session only."
                )
                encrypted_payload = ""
            else:
                encrypted_payload = enc_result

        profile_data = {
            "enabled": is_enabled,
            "encrypted_payload": encrypted_payload,
        }
        ProxyProfileManager.save_profile(profile_data)

        if is_enabled:
            NetworkSession.update_proxies(self.build_proxies_dict())
        else:
            NetworkSession.update_proxies({})

        super().accept()

    def closeEvent(self, event):
        if self._test_worker and self._test_worker.isRunning():
            self._test_worker.requestInterruption()
            self._test_worker.wait(1000)
        super().closeEvent(event)

    def get_proxies(self) -> Tuple[str, str]:
        proxies = self.build_proxies_dict()
        return proxies.get("http", ""), proxies.get("https", "")

