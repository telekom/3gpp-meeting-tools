from pathlib import Path

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QPainter, QColor, QIcon, QPixmap, QFont, QPen
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QLabel, QFormLayout,
                             QLineEdit, QCheckBox, QHBoxLayout, QPushButton,
                             QApplication, QMessageBox)

from core.network.session import NetworkSession

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
    /* Toggled/Checked State for Live View Button */
    QPushButton:checked {
        background-color: #EBF3FC;
        border: 2px solid #395396;
        color: #395396;
    }

    /* Primary Action Buttons */
    QPushButton#primaryBtn, QPushButton#pptBtn, QPushButton#svgBtn {
        background-color: #1E5C99; 
        color: white; 
        border: none;
    }
    QPushButton#primaryBtn:hover, QPushButton#pptBtn:hover, QPushButton#svgBtn:hover {
        background-color: #15426E;
    }

    /* Dropdown Menus (Export Button) */
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

    /* Splitter Handle */
    QSplitter::handle {
        background-color: #E0E0E0;
        height: 2px;
        margin: 4px 0px;
    }
    QSplitter::handle:hover {
        background-color: #395396;
    }

    /* Status Bar */
    QStatusBar {
        background-color: #F0F0F0;
        border-top: 1px solid #D0D0D0;
        color: #333333;
    }

    /* Combobox Styling */
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

    /* Dark Theme for Console & Queue List */
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


def create_app_icon():
    """Generates the geometric network icon, saves it physically, and loads it for Windows."""

    # Define a physical path in the project root to save the icon
    # (Falling back to current directory if get_project_root is not easily importable here)
    try:
        from core.utils.paths import get_project_root
        icon_path = get_project_root() / "3gpp_icon_cache.png"
    except ImportError:
        icon_path = Path("3gpp_icon_cache.png")

    # 1. Draw the highest resolution version (256x256)
    size = 256
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)

    # Background
    bg_color = QColor("#1A202C")
    painter.setBrush(bg_color)
    painter.setPen(Qt.NoPen)
    corner_radius = size // 5
    painter.drawRoundedRect(2, 2, size - 4, size - 4, corner_radius, corner_radius)

    # Connections
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

    # Nodes
    painter.setBrush(QColor("#FFFFFF"))
    painter.setPen(Qt.NoPen)
    node_radius = size // 10

    painter.drawEllipse(int(center_x - node_radius), int(top_y - node_radius), node_radius * 2, node_radius * 2)
    painter.drawEllipse(int(bl_x - node_radius), int(bl_y - node_radius), node_radius * 2, node_radius * 2)
    painter.drawEllipse(int(br_x - node_radius), int(br_y - node_radius), node_radius * 2, node_radius * 2)

    painter.end()

    # 2. Save physically to disk so Windows Taskbar has a hard file to reference
    pixmap.save(str(icon_path), "PNG")

    # 3. Return a QIcon loaded directly from the physical file
    return QIcon(str(icon_path))


class ProxyDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Network Configuration")
        self.setModal(True)
        self.resize(520, 250)

        layout = QVBoxLayout()
        title = QLabel("📡 Proxy Configuration")
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(title)

        desc = QLabel("Leave blank to connect directly. Required only for initial downloads.")
        desc.setStyleSheet("color: #666; margin-bottom: 15px;")
        layout.addWidget(desc)

        form = QFormLayout()
        self.http_input = QLineEdit()
        self.http_input.setStyleSheet("padding: 5px; border: 1px solid #ccc; border-radius: 4px;")
        self.https_input = QLineEdit()
        self.https_input.setStyleSheet("padding: 5px; border: 1px solid #ccc; border-radius: 4px;")
        self.sync_checkbox = QCheckBox("Use the same proxy for HTTPS")
        self.sync_checkbox.stateChanged.connect(self.on_sync_changed)
        self.http_input.textChanged.connect(self.on_http_changed)

        form.addRow("HTTP Proxy:", self.http_input)
        form.addRow("", self.sync_checkbox)
        form.addRow("HTTPS Proxy:", self.https_input)
        layout.addLayout(form)

        self.status_lbl = QLabel("")
        self.status_lbl.setWordWrap(True)
        self.status_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_lbl)

        btn_layout = QHBoxLayout()
        self.skip_btn = QPushButton("Skip")
        self.skip_btn.clicked.connect(self.skip)

        self.test_btn = QPushButton("🔄 Test Connection")
        self.test_btn.setToolTip("Ping GitHub to verify if your proxy settings are working.")
        self.test_btn.clicked.connect(self.test_proxy)

        self.save_btn = QPushButton("Save && Continue")
        self.save_btn.setObjectName("primaryBtn")
        self.save_btn.clicked.connect(self.accept)

        btn_layout.addWidget(self.skip_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.test_btn)
        btn_layout.addWidget(self.save_btn)

        layout.addSpacing(10)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def test_proxy(self) -> None:
        """Tests the proxy using the shared NetworkSession tester."""
        http_val: str = self.http_input.text().strip()
        https_val: str = self.https_input.text().strip()

        proxies: dict = {}
        if http_val: proxies['http'] = http_val
        if https_val: proxies['https'] = https_val

        # Update button to show loading state
        self.test_btn.setText("Testing...")
        self.test_btn.setEnabled(False)
        QApplication.processEvents()  # Force UI to update the button text

        # Test the connection using the NetworkSession static method
        success: bool = NetworkSession.test_connection(proxies)

        # Restore button state
        self.test_btn.setText("Test Connection")
        self.test_btn.setEnabled(True)

        if success:
            QMessageBox.information(self, "Success", "Connection to 3GPP server successful!")
        else:
            QMessageBox.warning(self, "Failed", "Connection failed. Please check your proxy settings or firewall.")

    def accept(self) -> None:
        """Saves the proxy and instantly updates the running application session."""
        # ... (Keep your existing code here that saves the proxy to your config file) ...

        # ---> NEW: Update the running global session so the crawler uses it immediately
        http_val: str = self.http_input.text().strip()
        https_val: str = self.https_input.text().strip()
        proxies: dict = {}
        if http_val: proxies['http'] = http_val
        if https_val: proxies['https'] = https_val

        NetworkSession.update_proxies(proxies)

        super().accept()

    def on_sync_changed(self, state):
        if state == Qt.Checked:
            self.https_input.setEnabled(False)
            self.https_input.setText(self.http_input.text())
        else:
            self.https_input.setEnabled(True)

    def on_http_changed(self, text):
        if self.sync_checkbox.isChecked():
            self.https_input.setText(text)

    def skip(self):
        self.http_input.clear()
        self.https_input.clear()
        self.accept()

    def get_proxies(self):
        return self.http_input.text().strip(), self.https_input.text().strip()


# --- Inside: src/core/ui/ui_components.py ---
class InteractiveDropLabel(QLabel):
    file_dropped = pyqtSignal(list)

    def __init__(self, text, accepted_extensions):
        super().__init__(text)
        self.accepted_extensions = accepted_extensions
        self.setAlignment(Qt.AlignCenter)
        self.setAcceptDrops(True)

        # ---> THE FIX: Added 'QLabel { ... }' to prevent style bleeding into ToolTips
        self.default_style = "QLabel { border: 3px dashed #B0B0B0; border-radius: 10px; font-size: 15px; font-weight: bold; color: #777; background-color: #FAFAFA; }"
        self.hover_style = "QLabel { border: 3px dashed #395396; border-radius: 10px; font-size: 15px; font-weight: bold; color: #395396; background-color: #EBF3FC; }"
        self.busy_style = "QLabel { border: 3px dashed #D83B01; border-radius: 10px; font-size: 15px; font-weight: bold; color: #D83B01; background-color: #FDF4F0; }"
        self.error_style = "QLabel { border: 3px dashed #D32F2F; border-radius: 10px; font-size: 15px; font-weight: bold; color: #D32F2F; background-color: #FDEDED; }"

        self.setStyleSheet(self.default_style)

    def set_state(self, state, text=None):
        if text: self.setText(text)
        if state == "ready":
            self.setStyleSheet(self.default_style)
        elif state == "busy":
            self.setStyleSheet(self.busy_style)
        elif state == "error":
            self.setStyleSheet(self.error_style)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith(ext) for url in urls for ext in self.accepted_extensions):
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
            if any(file_path.lower().endswith(ext) for ext in self.accepted_extensions):
                valid_files.append(file_path)
        if valid_files:
            self.file_dropped.emit(valid_files)