# --- File: src/core/ui/ui_components.py ---
import os
import urllib.parse
from pathlib import Path
from typing import Optional, Tuple

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QFont, QIcon, QPainter, QPen, QPixmap
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

from core.network.session import (
    NetworkSession,
    ProxyProfileManager,
    SecureCredentialStore,
    redact_sensitive_urls,
)

# ==========================================
# --- GLOBAL STYLESHEET (ALL-BLUE THEME) ---
# ==========================================
GLOBAL_STYLE = """
    QWidget {
        font-family: "Segoe UI", Arial, sans-serif;
        font-size: 13px;
        color: #333333;
    }
    QToolTip {
        color: #333333;
        background-color: #F8F8F8;
        border: 1px solid #D0D0D0;
        border-radius: 4px;
        padding: 4px;
    }
    QTabWidget::pane {
        border: 1px solid #D0D0D0;
        border-radius: 8px;
        background: #FFFFFF;
        top: -1px;
    }
    QTabBar::tab {
        background: #EAEAEA;
        border: 1px solid #D0D0D0;
        padding: 8px 16px;
        margin-right: 2px;
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
    }
    QTabBar::tab:selected {
        background: #FFFFFF;
        border-bottom-color: #FFFFFF;
        font-weight: bold;
        color: #395396;
    }
    QTabBar::tab:hover:!selected {
        background: #F0F0F0;
    }
    QPushButton {
        padding: 8px 16px;
        border-radius: 6px;
        border: 1px solid #CCCCCC;
        background-color: #F8F8F8;
        font-weight: bold;
    }
    QPushButton:hover {
        background-color: #EAEAEA;
    }
    QPushButton:disabled {
        background-color: #F0F0F0;
        color: #A0A0A0;
        border: 1px solid #DFDFDF;
    }
    QPushButton:checked {
        background-color: #EBF3FC;
        border: 2px solid #395396;
        color: #395396;
    }
    QPushButton#primaryBtn, QPushButton#pptBtn, QPushButton#svgBtn {
        background-color: #1E5C99; 
        color: white; 
        border: none;
    }
    QPushButton#primaryBtn:hover, QPushButton#pptBtn:hover, QPushButton#svgBtn:hover {
        background-color: #15426E;
    }
    QMenu {
        background-color: #FFFFFF;
        border: 1px solid #CCCCCC;
        border-radius: 6px;
        padding: 4px;
    }
    QMenu::item {
        padding: 8px 24px 8px 12px;
        border-radius: 4px;
        font-size: 13px;
        color: #333333;
    }
    QMenu::item:selected {
        background-color: #EBF3FC;
        color: #395396;
        font-weight: bold;
    }
    QPushButton::menu-indicator {
        width: 0px; 
    }
    QSplitter::handle {
        background-color: #E0E0E0;
        height: 2px;
        margin: 4px 0px;
    }
    QSplitter::handle:hover {
        background-color: #395396;
    }
    QStatusBar {
        background-color: #F0F0F0;
        border-top: 1px solid #D0D0D0;
        color: #333333;
    }
    QComboBox {
        padding: 4px 8px;
        border-radius: 4px;
        border: 1px solid #CCCCCC;
        background-color: #FFFFFF;
        font-weight: bold;
        color: #333333;
    }
    QComboBox::drop-down {
        border: none;
        width: 20px;
    }
    QTextEdit#console, QListWidget#queueList {
        background-color: #1E1E1E; 
        color: #D4D4D4; 
        font-family: Consolas, 'Courier New', monospace; 
        font-size: 13px; 
        border-radius: 8px; 
        padding: 8px;
        border: 1px solid #444444;
    }
    QListWidget#queueList::item {
        padding: 4px;
        border-bottom: 1px solid #333333;
    }
    QListWidget#queueList::item:selected {
        background-color: #264F78;
        color: #FFFFFF;
        border-radius: 4px;
    }
"""

# ==========================================
# --- COMMON REUSABLE UI STYLES ---
# ==========================================

CARD_FRAME_STYLE = """
QFrame#cardFrame {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
}
"""

TABLE_STYLE_CLEAN = """
QTableWidget {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 6px;
    gridline-color: #F1F5F9;
    font-size: 12px;
}
QTableWidget::item {
    padding: 4px 8px;
    border-bottom: 1px solid #F1F5F9;
}
QTableWidget::item:selected {
    background-color: #EBF3FC;
    color: #1E293B;
}
QHeaderView::section {
    background-color: #F8FAFC;
    color: #475569;
    font-weight: bold;
    font-size: 11px;
    border: none;
    border-bottom: 1px solid #CBD5E1;
    padding: 6px 8px;
}
"""

SEARCH_INPUT_STYLE = """
QLineEdit {
    padding: 4px 8px;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    font-size: 12px;
    background-color: #FFFFFF;
}
QLineEdit:focus {
    border: 1px solid #1E5C99;
}
"""

BADGE_STYLE_PRIMARY = """
background-color: #E6F4EA;
color: #137333;
font-weight: bold;
font-size: 11px;
border: 1px solid #CEEAD6;
border-radius: 4px;
padding: 2px 8px;
"""

BADGE_STYLE_MUTED = """
background-color: #F1F5F9;
color: #475569;
font-size: 11px;
border: 1px solid #E2E8F0;
border-radius: 4px;
padding: 2px 8px;
"""

BADGE_STYLE_INFO = """
background-color: #EBF8FF;
color: #2B6CB0;
border: 1px solid #BEE3F8;
border-radius: 4px;
padding: 2px 8px;
font-size: 11px;
font-weight: bold;
"""

BUTTON_STYLE_TOOLBAR_SECONDARY = """
QPushButton {
    background-color: #F8FAFC;
    color: #1E293B;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QPushButton:pressed {
    background-color: #E2E8F0;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

BUTTON_STYLE_TOOLBAR_DANGER = """
QPushButton {
    background-color: #FEF2F2;
    color: #DC2626;
    border: 1px solid #FECACA;
    border-radius: 4px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #FEE2E2;
    border-color: #F87171;
    color: #B91C1C;
}
QPushButton:pressed {
    background-color: #FECACA;
    border-color: #EF4444;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

BUTTON_STYLE_TOOLBAR_WARNING = """
QPushButton {
    background-color: #FFFBEB;
    color: #B45309;
    border: 1px solid #FDE68A;
    border-radius: 4px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #FEF3C7;
    border-color: #F59E0B;
}
QPushButton:pressed {
    background-color: #FDE68A;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

COMBOBOX_STYLE_TOOLBAR = """
QComboBox {
    combobox-popup: 0;
    font-size: 11px;
    font-weight: 600;
    color: #1E293B;
    background-color: #F8FAFC;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    padding: 3px 22px 3px 10px;
    min-height: 22px;
}
QComboBox:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QComboBox:focus, QComboBox:on {
    background-color: #FFFFFF;
    border-color: #1E5C99;
}
QComboBox QAbstractItemView {
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    background-color: #FFFFFF;
    color: #1E293B;
    selection-background-color: #EBF3FC;
    selection-color: #1E5C99;
    padding: 4px;
    outline: none;
}
QComboBox QAbstractItemView::item {
    min-height: 22px;
    padding: 2px 8px;
    border-radius: 3px;
}
QComboBox QAbstractItemView::item:hover {
    background-color: #F1F5F9;
    color: #0F172A;
}
QComboBox QAbstractItemView::item:selected {
    background-color: #EBF3FC;
    color: #1E5C99;
    font-weight: bold;
}
QScrollBar:vertical {
    border: none;
    background: #F8FAFC;
    width: 8px;
    margin: 4px 0;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #CBD5E1;
    min-height: 20px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #94A3B8;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""


def create_app_icon():
    """Generates the geometric network icon, saves it physically, and loads it for Windows."""
    try:
        from core.utils.paths import get_project_root
        icon_path = get_project_root() / "3gpp_icon_cache.png"
    except ImportError:
        icon_path = Path("3gpp_icon_cache.png")

    size = 256
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)

    bg_color = QColor("#1A202C")
    painter.setBrush(bg_color)
    painter.setPen(Qt.NoPen)
    corner_radius = size // 5
    painter.drawRoundedRect(2, 2, size - 4, size - 4, corner_radius, corner_radius)

    pen = QPen(QColor("#3B82F6"))
    pen.setWidth(size // 12)
    pen.setJoinStyle(Qt.RoundJoin)
    pen.setCapStyle(Qt.RoundCap)
    painter.setPen(pen)

    center_x = size / 2
    top_y = size * 0.28
    bl_x = size * 0.25
    bl_y = size * 0.72
    br_x = size * 0.75
    br_y = size * 0.72

    painter.drawLine(int(center_x), int(top_y), int(bl_x), int(bl_y))
    painter.drawLine(int(center_x), int(top_y), int(br_x), int(br_y))
    painter.drawLine(int(bl_x), int(bl_y), int(br_x), int(br_y))

    painter.setBrush(QColor("#FFFFFF"))
    painter.setPen(Qt.NoPen)
    node_radius = size // 10

    painter.drawEllipse(int(center_x - node_radius), int(top_y - node_radius), node_radius * 2, node_radius * 2)
    painter.drawEllipse(int(bl_x - node_radius), int(bl_y - node_radius), node_radius * 2, node_radius * 2)
    painter.drawEllipse(int(br_x - node_radius), int(br_y - node_radius), node_radius * 2, node_radius * 2)

    painter.end()

    pixmap.save(str(icon_path), "PNG")
    return QIcon(str(icon_path))


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
        self.resize(540, 360)

        self._test_worker: Optional[ProxyTestWorker] = None

        self._setup_ui()
        self._load_existing_settings()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        title = QLabel("🌐 Proxy Configuration & Profiles")
        title.setStyleSheet("font-size: 15px; font-weight: bold; margin-bottom: 2px;")
        layout.addWidget(title)

        desc = QLabel(
            "Configure corporate HTTP/HTTPS proxies. Passwords are encrypted locally via Windows DPAPI "
            "and are never saved in plain text."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #64748B; font-size: 11px; margin-bottom: 8px;")
        layout.addWidget(desc)

        self.enable_checkbox = QCheckBox("Activate this proxy profile")
        self.enable_checkbox.setStyleSheet("font-weight: bold; color: #1E293B; margin-bottom: 6px;")
        self.enable_checkbox.toggled.connect(self._on_enable_toggled)
        layout.addWidget(self.enable_checkbox)

        self.form_container = QWidget()
        form = QFormLayout(self.form_container)
        form.setContentsMargins(0, 0, 0, 0)
        form.setLabelAlignment(Qt.AlignRight)

        self.http_input = QLineEdit()
        self.http_input.setPlaceholderText("e.g., proxy.company.com:8080")
        self.http_input.setStyleSheet("padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px;")

        self.https_input = QLineEdit()
        self.https_input.setPlaceholderText("e.g., proxy.company.com:8080")
        self.https_input.setStyleSheet("padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px;")

        self.sync_checkbox = QCheckBox("Use the same proxy address for HTTPS")
        self.sync_checkbox.setChecked(True)
        self.sync_checkbox.stateChanged.connect(self._on_sync_changed)
        self.http_input.textChanged.connect(self._on_http_changed)

        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Domain\\Username or Username (Optional)")
        self.user_input.setStyleSheet("padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px;")

        self.pass_input = QLineEdit()
        self.pass_input.setEchoMode(QLineEdit.Password)
        self.pass_input.setPlaceholderText("Password (Optional)")
        self.pass_input.setStyleSheet("padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px;")

        self.remember_cred_checkbox = QCheckBox("Save password securely using Windows DPAPI")
        self.remember_cred_checkbox.setChecked(True)

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
        self.status_lbl.setStyleSheet("font-size: 11px; color: #475569; margin: 4px 0;")
        layout.addWidget(self.status_lbl)

        btn_layout = QHBoxLayout()

        self.deactivate_btn = QPushButton("Deactivate (Direct)")
        self.deactivate_btn.setToolTip("Turn off proxy usage without deleting saved settings.")
        self.deactivate_btn.clicked.connect(self.deactivate_profile)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)

        self.test_btn = QPushButton("⚡ Test Connection")
        self.test_btn.clicked.connect(self.test_proxy)

        self.save_btn = QPushButton("Save && Apply")
        self.save_btn.setObjectName("primaryBtn")
        self.save_btn.setStyleSheet("background-color: #1E5C99; color: white; font-weight: bold;")
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

        saved_pass = ProxyProfileManager.get_decrypted_password()
        if saved_pass:
            self.pass_input.setText(saved_pass)
            self.remember_cred_checkbox.setChecked(True)
        else:
            self.remember_cred_checkbox.setChecked(False)

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
        self.status_lbl.setStyleSheet("color: #0284C7; font-size: 11px;")

        self._test_worker = ProxyTestWorker(proxies, parent=self)
        self._test_worker.finished_signal.connect(self._on_test_finished)
        self._test_worker.start()

    def _on_test_finished(self, success: bool, message: str):
        self.test_btn.setEnabled(True)
        self.test_btn.setText("⚡ Test Connection")
        self.save_btn.setEnabled(True)

        if success:
            self.status_lbl.setText("✅ Connection verified successfully!")
            self.status_lbl.setStyleSheet("color: #166534; font-size: 11px;")
            QMessageBox.information(self, "Proxy Success", message)
        else:
            self.status_lbl.setText("❌ Connection failed.")
            self.status_lbl.setStyleSheet("color: #DC2626; font-size: 11px;")
            QMessageBox.warning(self, "Proxy Test Failed", message)

    def deactivate_profile(self):
        """Turns off proxy routing while preserving the configuration for future use."""
        profile = ProxyProfileManager.load_profile()
        profile["enabled"] = False
        ProxyProfileManager.save_profile(profile)

        NetworkSession.update_proxies({})
        super().accept()

    def accept(self):
        """
        Validates encryption before persistence.
        Guarantees that unencrypted passwords are NEVER written to disk,
        updates NetworkSession, and cleanly dismisses the dialog.
        """
        is_enabled = self.enable_checkbox.isChecked()
        plain_password = self.pass_input.text()
        encrypted_pass = ""

        if self.remember_cred_checkbox.isChecked() and plain_password:
            success, enc_result = SecureCredentialStore.encrypt(plain_password)
            if not success:
                QMessageBox.warning(
                    self,
                    "Security Notice",
                    "Windows DPAPI encryption is unavailable or encountered an error.\n\n"
                    "Your password was NOT written to disk to prevent security risks. "
                    "It will remain active in memory for the current session only."
                )
                encrypted_pass = ""
            else:
                encrypted_pass = enc_result

        profile_data = {
            "enabled": is_enabled,
            "http_host": self.http_input.text().strip(),
            "https_host": self.https_input.text().strip(),
            "sync_https": self.sync_checkbox.isChecked(),
            "username": self.user_input.text().strip(),
            "encrypted_password": encrypted_pass,
        }
        ProxyProfileManager.save_profile(profile_data)

        if is_enabled:
            NetworkSession.update_proxies(self.build_proxies_dict())
        else:
            NetworkSession.update_proxies({})

        super().accept()

    def closeEvent(self, event):
        if self._test_worker and self._test_worker.isRunning():
            self._test_worker.terminate()
            self._test_worker.wait(500)
        super().closeEvent(event)

    def get_proxies(self) -> Tuple[str, str]:
        proxies = self.build_proxies_dict()
        return proxies.get("http", ""), proxies.get("https", "")


# ==========================================
# --- DRAG & DROP INTERACTION WIDGET ---
# ==========================================


class InteractiveDropLabel(QLabel):
    file_dropped = pyqtSignal(list)

    def __init__(self, text, accepted_extensions):
        super().__init__(text)
        self.accepted_extensions = accepted_extensions
        self.setAlignment(Qt.AlignCenter)
        self.setAcceptDrops(True)

        self.default_style = (
            "QLabel { border: 3px dashed #B0B0B0; border-radius: 10px; font-size: 15px; "
            "font-weight: bold; color: #777; background-color: #FAFAFA; }"
        )
        self.hover_style = (
            "QLabel { border: 3px dashed #395396; border-radius: 10px; font-size: 15px; "
            "font-weight: bold; color: #395396; background-color: #EBF3FC; }"
        )
        self.busy_style = (
            "QLabel { border: 3px dashed #D83B01; border-radius: 10px; font-size: 15px; "
            "font-weight: bold; color: #D83B01; background-color: #FDF4F0; }"
        )
        self.error_style = (
            "QLabel { border: 3px dashed #D32F2F; border-radius: 10px; font-size: 15px; "
            "font-weight: bold; color: #D32F2F; background-color: #FDEDED; }"
        )

        self.setStyleSheet(self.default_style)

    def set_state(self, state, text=None):
        if text:
            self.setText(text)
        if state == "ready":
            self.setStyleSheet(self.default_style)
        elif state == "busy":
            self.setStyleSheet(self.busy_style)
        elif state == "error":
            self.setStyleSheet(self.error_style)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(
                url.toLocalFile().lower().endswith(ext)
                for url in urls
                for ext in self.accepted_extensions
            ):
                self.setStyleSheet(self.hover_style)
                event.accept()
                return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)

    def dropEvent(self, event):
        self.setStyleSheet(self.default_style)
        valid_files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if any(
                file_path.lower().endswith(ext)
                for ext in self.accepted_extensions
            ):
                valid_files.append(file_path)
        if valid_files:
            self.file_dropped.emit(valid_files)