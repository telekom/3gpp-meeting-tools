# --- File: src/modules/puml2visio/ui/svg_preview_widget.py ---
"""Embedded Interactive SVG Preview Panel for PlantUML Mini-IDE.

Renders generated PlantUML SVGs side-by-side with the code editor using
PyQt5.QtSvg, supporting smooth wheel zoom, drag-to-pan, 1:1 reset, and fit-to-window.
"""

from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QPointF, QRectF, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QPainter, QWheelEvent
from PyQt5.QtSvg import QGraphicsSvgItem, QSvgRenderer
from PyQt5.QtWidgets import (
    QFrame,
    QGraphicsScene,
    QGraphicsView,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
)


class SvgGraphicsView(QGraphicsView):
    """Zoomable, pannable graphics view tailored for rendered SVG diagrams."""

    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setRenderHint(QPainter.Antialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setViewportUpdateMode(QGraphicsView.FullViewportUpdate)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setStyleSheet("border: none; background-color: #ECECEC;")
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)

    def wheelEvent(self, event: QWheelEvent):
        """Smooth zoom in/out with mouse wheel (Ctrl+Wheel or standard Wheel)."""
        factor = 1.15
        if event.angleDelta().y() > 0:
            self.scale(factor, factor)
        else:
            self.scale(1.0 / factor, 1.0 / factor)
        event.accept()


class EmbeddedSvgPreview(QFrame):
    """Side-by-side live SVG diagram viewer panel with zoom and pan controls."""
    browser_requested = pyqtSignal()
    close_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("cardFrame")
        self.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #F8FAFC; 
                border: 1px solid #CBD5E1; 
                border-radius: 6px;
            }
        """)

        self.svg_item: Optional[QGraphicsSvgItem] = None
        self.current_svg_path: Optional[str] = None
        self._first_load = True

        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        # Header Toolbar
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(4, 2, 4, 2)
        header_layout.setSpacing(4)

        title_lbl = QLabel("👁️ Embedded Preview")
        title_lbl.setStyleSheet("font-weight: bold; font-size: 12px; color: #1E293B;")
        header_layout.addWidget(title_lbl)
        header_layout.addStretch()

        # Zoom Out
        self.zoom_out_btn = QPushButton("➖")
        self.zoom_out_btn.setToolTip("Zoom Out")
        self.zoom_out_btn.setFixedWidth(28)
        self.zoom_out_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.zoom_out_btn.clicked.connect(self.zoom_out)
        header_layout.addWidget(self.zoom_out_btn)

        # Zoom In
        self.zoom_in_btn = QPushButton("➕")
        self.zoom_in_btn.setToolTip("Zoom In")
        self.zoom_in_btn.setFixedWidth(28)
        self.zoom_in_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.zoom_in_btn.clicked.connect(self.zoom_in)
        header_layout.addWidget(self.zoom_in_btn)

        # 1:1 Reset
        self.reset_btn = QPushButton("1:1")
        self.reset_btn.setToolTip("Reset Zoom (100%)")
        self.reset_btn.setFixedWidth(32)
        self.reset_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.reset_btn.clicked.connect(self.reset_zoom)
        header_layout.addWidget(self.reset_btn)

        # Fit to Window
        self.fit_btn = QPushButton("⛶ Fit")
        self.fit_btn.setToolTip("Fit Diagram to Window")
        self.fit_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.fit_btn.clicked.connect(self.fit_view)
        header_layout.addWidget(self.fit_btn)

        # Open in Browser button
        self.browser_btn = QPushButton("↗ Browser")
        self.browser_btn.setToolTip("Open this live preview in your default web browser")
        self.browser_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.browser_btn.clicked.connect(self.browser_requested.emit)
        header_layout.addWidget(self.browser_btn)

        # Close Panel
        self.close_btn = QPushButton("✕")
        self.close_btn.setToolTip("Close Preview Panel")
        self.close_btn.setFixedWidth(26)
        self.close_btn.setStyleSheet("""
            QPushButton {
                border: none; 
                font-weight: bold; 
                color: #64748B; 
                padding: 2px;
            }
            QPushButton:hover {
                color: #DC2626;
            }
        """)
        self.close_btn.clicked.connect(self.close_requested.emit)
        header_layout.addWidget(self.close_btn)

        layout.addLayout(header_layout)

        # Graphics Scene & View
        self.scene = QGraphicsScene(self)
        self.view = SvgGraphicsView(self.scene, self)
        layout.addWidget(self.view)

        self.setLayout(layout)

    def load_svg(self, svg_path: str):
        """Loads and renders the SVG file at svg_path in the graphics view."""
        path = Path(svg_path)
        if not path.exists():
            return

        self.current_svg_path = str(path.resolve())

        # Render SVG into item
        renderer = QSvgRenderer(self.current_svg_path)
        if not renderer.isValid():
            return

        self.scene.clear()
        self.svg_item = QGraphicsSvgItem()
        self.svg_item.setSharedRenderer(renderer)
        self.scene.addItem(self.svg_item)
        self.scene.setSceneRect(self.svg_item.boundingRect())

        # Auto-fit on first load or diagram change
        if self._first_load:
            self.fit_view()
            self._first_load = False

    def zoom_in(self):
        self.view.scale(1.2, 1.2)

    def zoom_out(self):
        self.view.scale(1.0 / 1.2, 1.0 / 1.2)

    def reset_zoom(self):
        self.view.resetTransform()

    def fit_view(self):
        if self.svg_item and not self.scene.sceneRect().isEmpty():
            self.view.resetTransform()
            self.view.fitInView(self.scene.sceneRect(), Qt.KeepAspectRatio)
