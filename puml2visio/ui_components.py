import urllib.request
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QLabel, QFormLayout,
                             QLineEdit, QCheckBox, QHBoxLayout, QPushButton,
                             QTextEdit, QPlainTextEdit, QWidget, QApplication)
from PyQt5.QtCore import Qt, pyqtSignal, QRect, QSize
from PyQt5.QtGui import QPainter, QColor, QTextFormat, QIcon, QPixmap, QFont

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
    pixmap = QPixmap(64, 64)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)

    painter.setBrush(QColor("#1E5C99"))
    painter.setPen(Qt.NoPen)
    painter.drawRoundedRect(4, 4, 56, 56, 12, 12)

    painter.setPen(QColor("#FFFFFF"))
    font = QFont("Segoe UI", 16, QFont.Bold)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "P2V")

    painter.end()
    return QIcon(pixmap)


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

        desc = QLabel("Leave blank to connect directly. Required only for initial JAR download.")
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
        self.test_btn.clicked.connect(self.test_connection)

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

    def test_connection(self):
        self.status_lbl.setText("⏳ Testing connection to GitHub... Please wait.")
        self.status_lbl.setStyleSheet("color: #D83B01; font-weight: bold;")
        QApplication.processEvents()

        http_val, https_val = self.get_proxies()
        proxies = {}
        if http_val: proxies['http'] = http_val
        if https_val: proxies['https'] = https_val

        try:
            if proxies:
                proxy_handler = urllib.request.ProxyHandler(proxies)
                opener = urllib.request.build_opener(proxy_handler)
            else:
                opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

            req = urllib.request.Request("https://github.com", method="HEAD")
            opener.open(req, timeout=5)

            self.status_lbl.setText("✅ Connection Successful! You can now save.")
            self.status_lbl.setStyleSheet("color: #6A9955; font-weight: bold;")
        except Exception as e:
            self.status_lbl.setText(f"❌ Connection Failed: {str(e)}")
            self.status_lbl.setStyleSheet("color: #D32F2F; font-weight: bold;")

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


class LineNumberArea(QWidget):
    def __init__(self, editor):
        super().__init__(editor)
        self.code_editor = editor

    def sizeHint(self):
        return QSize(self.code_editor.line_number_area_width(), 0)

    def paintEvent(self, event):
        self.code_editor.lineNumberAreaPaintEvent(event)


class CodeDropTextEdit(QPlainTextEdit):
    file_dropped = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.line_number_area = LineNumberArea(self)

        self.blockCountChanged.connect(self.update_line_number_area_width)
        self.updateRequest.connect(self.update_line_number_area)
        self.cursorPositionChanged.connect(self.highlight_current_line)

        self.update_line_number_area_width(0)
        self.highlight_current_line()

        self.default_style = """
            QPlainTextEdit {
                font-family: Consolas, Courier New, monospace; 
                font-size: 13px; 
                border: 2px solid #E0E0E0; 
                border-radius: 8px; 
                background-color: #FAFAFA;
            }
            QPlainTextEdit:focus {
                border: 2px solid #395396;
                background-color: #FFFFFF;
            }
        """
        self.hover_style = self.default_style.replace("border: 2px solid #E0E0E0;",
                                                      "border: 2px dashed #395396; background-color: #EBF3FC;")
        self.setStyleSheet(self.default_style)
        self.setLineWrapMode(QPlainTextEdit.NoWrap)

    def line_number_area_width(self):
        digits = 1
        max_val = max(1, self.blockCount())
        while max_val >= 10:
            max_val /= 10
            digits += 1

        fm = self.fontMetrics()
        width_char = fm.horizontalAdvance('9') if hasattr(fm, 'horizontalAdvance') else fm.width('9')
        space = 10 + width_char * digits
        return space

    def update_line_number_area_width(self, _):
        self.setViewportMargins(self.line_number_area_width(), 0, 0, 0)

    def update_line_number_area(self, rect, dy):
        if dy:
            self.line_number_area.scroll(0, dy)
        else:
            self.line_number_area.update(0, rect.y(), self.line_number_area.width(), rect.height())
        if rect.contains(self.viewport().rect()):
            self.update_line_number_area_width(0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        cr = self.contentsRect()
        self.line_number_area.setGeometry(QRect(cr.left(), cr.top(), self.line_number_area_width(), cr.height()))

    def lineNumberAreaPaintEvent(self, event):
        painter = QPainter(self.line_number_area)
        painter.fillRect(event.rect(), QColor("#EAEAEA"))

        block = self.firstVisibleBlock()
        block_number = block.blockNumber()
        top = round(self.blockBoundingGeometry(block).translated(self.contentOffset()).top())
        bottom = top + round(self.blockBoundingRect(block).height())

        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                number = str(block_number + 1)
                painter.setPen(QColor("#888888"))
                painter.drawText(0, top, self.line_number_area.width() - 5, self.fontMetrics().height(),
                                 Qt.AlignRight | Qt.AlignVCenter, number)

            block = block.next()
            top = bottom
            bottom = top + round(self.blockBoundingRect(block).height())
            block_number += 1

    def highlight_current_line(self):
        extra_selections = []
        if not self.isReadOnly():
            selection = QTextEdit.ExtraSelection()
            line_color = QColor("#EBF3FC")
            selection.format.setBackground(line_color)
            selection.format.setProperty(QTextFormat.FullWidthSelection, True)
            selection.cursor = self.textCursor()
            selection.cursor.clearSelection()
            extra_selections.append(selection)
        self.setExtraSelections(extra_selections)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith('.vsdx') for url in urls):
                self.setStyleSheet(self.hover_style)
                event.acceptProposedAction()
                return
        super().dragEnterEvent(event)

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        self.setStyleSheet(self.default_style)
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith('.vsdx'):
                    self.file_dropped.emit(file_path)
                    event.acceptProposedAction()
                    return
        super().dropEvent(event)


class InteractiveDropLabel(QLabel):
    file_dropped = pyqtSignal(list)

    def __init__(self, text, accepted_extensions):
        super().__init__(text)
        self.accepted_extensions = accepted_extensions
        self.setAlignment(Qt.AlignCenter)
        self.setAcceptDrops(True)
        self.default_style = "border: 3px dashed #B0B0B0; border-radius: 10px; font-size: 15px; font-weight: bold; color: #777; background-color: #FAFAFA;"
        self.hover_style = "border: 3px dashed #395396; border-radius: 10px; font-size: 15px; font-weight: bold; color: #395396; background-color: #EBF3FC;"
        self.busy_style = "border: 3px dashed #D83B01; border-radius: 10px; font-size: 15px; font-weight: bold; color: #D83B01; background-color: #FDF4F0;"
        self.error_style = "border: 3px dashed #D32F2F; border-radius: 10px; font-size: 15px; font-weight: bold; color: #D32F2F; background-color: #FDEDED;"
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