"""Standalone top-level HTML email message viewer window."""

import html
import webbrowser

from PyQt5.QtCore import QSettings, Qt, QUrl, pyqtSignal
from PyQt5.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import bind_reader_zoom_settings
from core.ui.theme import C
from modules.emails.core.thread_markdown_exporter import apply_email_reader_document_style


class StandaloneEmailReaderWindow(QWidget):
    """Independent, fully resizable top-level window for reading emails on laptops/secondary monitors."""

    tdoc_selected = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER
        self._parent_ref = parent
        self.setWindowFlags(Qt.Window)
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass
        self.setWindowTitle("📖 Email Message Viewer")
        self.resize(850, 650)
        self.setStyleSheet(f"QWidget {{ background-color: {C.BG_ALT_LIGHT}; }}")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        top_bar = QHBoxLayout()
        self.lbl_title = QLabel("<b>Message Details</b>")
        self.lbl_title.setStyleSheet(f"font-size: 13px; color: {C.OFFICE_BLUE};")
        top_bar.addWidget(self.lbl_title)
        top_bar.addStretch()

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.close)
        top_bar.addWidget(btn_close)
        layout.addLayout(top_bar)

        self.browser = QTextBrowser()
        self.browser.setOpenLinks(False)
        apply_email_reader_document_style(self.browser)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        self.browser.setStyleSheet(f"background: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 12px;")
        self._reader_settings = QSettings("3GPP_Tools", "StandaloneEmailReaderWindow")
        bind_reader_zoom_settings(self.browser, self._reader_settings, key="zoom_steps")
        layout.addWidget(self.browser)

    def update_content(self, html_content: str, subject: str = ""):
        if subject:
            self.setWindowTitle(f"📖 {subject}")
            self.lbl_title.setText(f"<b>{html.escape(subject[:80])}</b>")
        self.browser.setHtml(html_content)

    def _on_anchor_clicked(self, url: QUrl):
        url_str = url.toString()
        if url_str.startswith("tdoc:"):
            target_tdoc = url_str.split(":", 1)[1].strip()
            self.tdoc_selected.emit(target_tdoc)
        elif url_str.startswith("mailto:") or url_str.startswith("http://") or url_str.startswith("https://"):
            webbrowser.open(url_str)
