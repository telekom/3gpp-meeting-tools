"""TDocs Cockpit Quick Guide & Keyboard Shortcuts Dialog."""

from __future__ import annotations

from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QShortcut,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, SEARCH_INPUT_STYLE
from .tabs_html import (
    _build_search_tab_html,
    _build_shortcuts_tab_html,
    _build_table_tab_html,
    _build_toolbar_tab_html,
)


class TDocsHelpDialog(QDialog):
    """Comprehensive in-app help and keyboard shortcuts dialog for the 3GPP TDocs Cockpit."""

    def __init__(self, parent=None, is_ran: bool = False):
        super().__init__(parent)
        self.is_ran = is_ran
        self.setWindowTitle("❓ 3GPP TDocs Cockpit — Quick Guide & Shortcuts")
        self.resize(920, 680)
        self.setMinimumSize(780, 520)

        # Allow pressing F1 or Esc to close/toggle
        QShortcut(QKeySequence("Esc"), self, self.close)
        QShortcut(QKeySequence("F1"), self, self.close)

        self._init_ui()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 14, 16, 14)
        main_layout.setSpacing(10)

        # Header Ribbon
        header_layout = QHBoxLayout()
        header_layout.setSpacing(12)

        icon_lbl = QLabel("❓")
        icon_lbl.setStyleSheet("font-size: 28px;")
        header_layout.addWidget(icon_lbl)

        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)
        dlg_title = QLabel("3GPP TDocs Cockpit — Quick Guide & Reference")
        dlg_title.setStyleSheet(
            f"font-size: 17px; font-weight: bold; color: {C.TEXT_PRIMARY};"
        )
        dlg_sub = QLabel(
            "Operational tools, double-click actions, search operators, and keyboard shortcuts."
        )
        dlg_sub.setStyleSheet(f"font-size: 12px; color: {C.TEXT_MUTED};")
        title_vbox.addWidget(dlg_title)
        title_vbox.addWidget(dlg_sub)
        header_layout.addLayout(title_vbox)

        header_layout.addStretch()

        # In-dialog live filter
        self.search_filter = QLineEdit()
        self.search_filter.setPlaceholderText("🔍 Filter help topics...")
        self.search_filter.setStyleSheet(SEARCH_INPUT_STYLE)
        self.search_filter.setFixedWidth(220)
        self.search_filter.textChanged.connect(self._on_filter_changed)
        header_layout.addWidget(self.search_filter)

        main_layout.addLayout(header_layout)

        # Divider
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setFrameShadow(QFrame.Plain)
        divider.setStyleSheet(
            f"color: {C.BORDER_SUBTLE}; background-color: {C.BORDER_SUBTLE}; height: 1px;"
        )
        main_layout.addWidget(divider)

        # Tab Widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                background-color: {C.BG_CARD};
                top: -1px;
            }}
            QTabBar::tab {{
                background: {C.BG_CANVAS};
                color: {C.TEXT_SECONDARY};
                padding: 8px 16px;
                margin-right: 4px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                border: 1px solid {C.BORDER_STRONG};
                border-bottom: none;
                font-size: 12.5px;
                font-weight: 500;
            }}
            QTabBar::tab:selected {{
                background: {C.BG_CARD};
                color: {C.TEXT_PRIMARY};
                font-weight: bold;
                border-bottom: 2px solid {C.ACCENT_ALT};
            }}
            QTabBar::tab:hover:!selected {{
                background: {C.BORDER_SUBTLE};
                color: {C.TEXT_STRONG};
            }}
        """)

        self.browser_toolbar = QTextBrowser()
        self.browser_toolbar.setOpenExternalLinks(False)
        self.browser_toolbar.setHtml(_build_toolbar_tab_html())
        self.tab_widget.addTab(self.browser_toolbar, "🎛️ Toolbar && Tools")

        self.browser_table = QTextBrowser()
        self.browser_table.setOpenExternalLinks(False)
        self.browser_table.setHtml(_build_table_tab_html())
        self.tab_widget.addTab(self.browser_table, "🖱️ Table && Double-Clicks")

        self.browser_search = QTextBrowser()
        self.browser_search.setOpenExternalLinks(False)
        self.browser_search.setHtml(_build_search_tab_html())
        self.tab_widget.addTab(self.browser_search, "🔍 Search && Filtering")

        self.browser_shortcuts = QTextBrowser()
        self.browser_shortcuts.setOpenExternalLinks(False)
        self.browser_shortcuts.setHtml(_build_shortcuts_tab_html())
        self.tab_widget.addTab(self.browser_shortcuts, "⌨️ Keyboard Shortcuts")

        main_layout.addWidget(self.tab_widget)

        # Footer
        footer_layout = QHBoxLayout()
        tip_lbl = QLabel(
            f"💡 <b>Quick Access:</b> Press <kbd style='background:{C.BG_CANVAS}; padding:1px 5px; border-radius:3px; border:1px solid {C.BORDER_STRONG};'>F1</kbd> anytime in the TDocs window to open this guide."
        )
        tip_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_SECONDARY};")
        footer_layout.addWidget(tip_lbl)
        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        main_layout.addLayout(footer_layout)

    def _on_filter_changed(self, text: str):
        """Finds and highlights text in the active tab's browser."""
        query = text.strip()
        current_browser: QTextBrowser = self.tab_widget.currentWidget()
        if not current_browser:
            return

        if not query:
            cursor = current_browser.textCursor()
            cursor.clearSelection()
            current_browser.setTextCursor(cursor)
            return

        found = current_browser.find(query)
        if not found:
            cursor = current_browser.textCursor()
            cursor.movePosition(cursor.Start)
            current_browser.setTextCursor(cursor)
            current_browser.find(query)
