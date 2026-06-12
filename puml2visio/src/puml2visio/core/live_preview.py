import tempfile
import shutil
import logging
import webbrowser
import re
from pathlib import Path

from PyQt5.QtCore import QObject, QThread, pyqtSignal, QTimer

from puml2visio.utils.utils import generate_cleaned_svg

# Lightweight HTML wrapper (Status text removed from DOM)
HTML_TEMPLATE = """<!DOCTYPE html>
<html>
<head>
    <title>PlantUML Live Preview</title>
    <style>
        body { 
            font-family: "Segoe UI", sans-serif; 
            text-align: center; 
            background-color: #ECECEC; 
            padding: 40px 20px; 
            margin: 0;
        }
        #preview-container {
            display: inline-block;
            background: white;
            padding: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            border-radius: 8px;
            min-width: 400px;
            min-height: 150px;
            text-align: left; /* Aligns error text cleanly */
        }
        img { max-width: 100%; height: auto; }
    </style>
    <script>
        function refreshImage() {
            const img = document.getElementById('preview');
            // Cache-busting query string forces the browser to pull the newest version
            img.src = 'preview.svg?t=' + new Date().getTime();
        }
        setInterval(refreshImage, 1000);
    </script>
</head>
<body>
    <div id="preview-container">
        <img id="preview" src="preview.svg" alt="Waiting for diagram..." onerror="this.src='preview.svg?t=' + new Date().getTime();" />
    </div>
</body>
</html>
"""

# --- STATE-AWARE FALLBACK SVGs ---
OFFLINE_SVG = """<svg width="500" height="150" xmlns="http://www.w3.org/2000/svg">
    <rect width="100%" height="100%" fill="#FFFFFF" rx="8"/>
    <text x="50%" y="45%" font-family="Segoe UI, Arial, sans-serif" font-size="22" font-weight="bold" fill="#555" text-anchor="middle">⏸️ Live Preview is Paused</text>
    <text x="50%" y="65%" font-family="Segoe UI, Arial, sans-serif" font-size="14" fill="#888" text-anchor="middle">Click "👁️ Live Preview" in the app to resume.</text>
</svg>"""

WAITING_SVG = """<svg width="500" height="150" xmlns="http://www.w3.org/2000/svg">
    <rect width="100%" height="100%" fill="#FFFFFF" rx="8"/>
    <text x="50%" y="55%" font-family="Segoe UI, Arial, sans-serif" font-size="18" fill="#888" text-anchor="middle">⏳ Waiting for PlantUML code...</text>
</svg>"""

# --- NEW: DYNAMIC ERROR SVG ---
ERROR_SVG_TEMPLATE = """<svg width="800" height="400" xmlns="http://www.w3.org/2000/svg">
    <rect width="100%" height="100%" fill="#FDEDED" rx="8" stroke="#D32F2F" stroke-width="2"/>
    <text x="20" y="40" font-family="Segoe UI, Arial, sans-serif" font-size="18" font-weight="bold" fill="#D32F2F">❌ PlantUML Syntax Error</text>
    {error_lines}
</svg>"""


class PreviewGeneratorThread(QThread):
    error_occurred = pyqtSignal(str)

    def __init__(self, puml_text: str, jar_path: Path, temp_dir: Path):
        super().__init__()
        self.puml_text = puml_text
        self.jar_path = jar_path
        self.temp_dir = temp_dir

    def run(self):
        try:
            # Render to a distinct temp file to avoid file-lock contention with the browser
            temp_puml = self.temp_dir / "temp_render.puml"
            temp_puml.write_text(self.puml_text, encoding="utf-8")

            # Silent logging so we don't spam the UI terminal
            def silent_log(msg, level):
                pass

            temp_svg = generate_cleaned_svg(temp_puml, self.jar_path, silent_log)

            # Atomically replace the live file
            live_svg = self.temp_dir / "preview.svg"
            shutil.copy2(temp_svg, live_svg)

        except Exception as e:
            # Broadcast the raw error string back to the manager
            self.error_occurred.emit(str(e))


class LivePreviewManager(QObject):
    log_msg = pyqtSignal(str, int)

    def __init__(self, text_edit, jar_path):
        super().__init__()
        self.text_edit = text_edit
        self.jar_path = jar_path
        self.active = False
        self._pending_update = False

        # Isolate all generation to the OS Temp directory
        self.temp_dir = Path(tempfile.gettempdir()) / "puml2visio_live"
        self.temp_dir.mkdir(parents=True, exist_ok=True)

        self.html_path = self.temp_dir / "index.html"
        self.svg_path = self.temp_dir / "preview.svg"

        if not self.svg_path.exists():
            self.svg_path.write_text(WAITING_SVG, encoding="utf-8")

        self.html_path.write_text(HTML_TEMPLATE, encoding="utf-8")

        self.debounce_timer = QTimer()
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.setInterval(750)
        self.debounce_timer.timeout.connect(self._trigger_generation)

        self.generator_thread = None

    def toggle(self, state: bool):
        self.active = state
        if self.active:
            self.text_edit.textChanged.connect(self._on_text_changed)

            if not self.text_edit.toPlainText().strip():
                self.svg_path.write_text(WAITING_SVG, encoding="utf-8")

            url = f"file://{self.html_path.resolve().as_posix()}"
            webbrowser.open(url)

            self.log_msg.emit("👁️ Live Preview activated. Rendering to default browser.", logging.INFO)
            self._trigger_generation()
        else:
            self.text_edit.textChanged.disconnect(self._on_text_changed)
            self.debounce_timer.stop()
            self._pending_update = False
            self.svg_path.write_text(OFFLINE_SVG, encoding="utf-8")
            self.log_msg.emit("🙈 Live Preview deactivated.", logging.INFO)

    def update_now(self):
        if self.active:
            self.debounce_timer.stop()
            self._trigger_generation()

    def _on_text_changed(self):
        if self.active:
            self.debounce_timer.start()

    def _trigger_generation(self):
        text = self.text_edit.toPlainText().strip()
        if not text:
            self.svg_path.write_text(WAITING_SVG, encoding="utf-8")
            return

        if self.generator_thread and self.generator_thread.isRunning():
            self._pending_update = True
            return

        self._pending_update = False
        self.generator_thread = PreviewGeneratorThread(text, self.jar_path, self.temp_dir)

        # --- NEW: Hook into the error broadcast ---
        self.generator_thread.error_occurred.connect(self._on_error)
        self.generator_thread.finished.connect(self._on_thread_finished)

        self.generator_thread.start()

    def _on_error(self, error_msg: str):
        """Intercepts syntax errors and paints them directly into the Live Preview browser."""
        # 1. Escape any stray HTML brackets so they don't break the SVG
        error_msg = error_msg.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        # 2. Clean up the messy temporary path injected by PlantUML
        # Converts "Error line 28 in file: C:\...\temp_render.puml" to "Error line 28"
        error_msg = re.sub(r'(Error line \d+) in file:? .*temp_render\.puml', r'\1', error_msg, flags=re.IGNORECASE)

        # 3. Build individual <text> rows (SVG doesn't support text wrapping natively)
        lines = error_msg.splitlines()
        svg_lines = []
        y_pos = 80

        for line in lines:
            clean_line = line.strip()
            # Ignore the repetitive wrapper label
            if clean_line and clean_line != "PlantUML Syntax Error:":
                svg_lines.append(
                    f'<text x="20" y="{y_pos}" font-family="Consolas, Courier New, monospace" font-size="14" fill="#B71C1C">{clean_line}</text>')
                y_pos += 22

        # 4. Inject into the master template and safely overwrite the live file
        error_svg = ERROR_SVG_TEMPLATE.replace("{error_lines}", "\n".join(svg_lines))

        try:
            self.svg_path.write_text(error_svg, encoding="utf-8")
        except Exception:
            pass  # Ignore file lock issues during rapid typing

    def _on_thread_finished(self):
        if getattr(self, '_pending_update', False):
            self._trigger_generation()