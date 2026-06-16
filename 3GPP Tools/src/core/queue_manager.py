import logging
import re
from pathlib import Path
from PyQt5.QtCore import QObject, pyqtSignal, QThread

# ==========================================
# --- GLOBAL TASK REGISTRY ---
# ==========================================
# Maps a target_format string to a dictionary containing:
# {
#     "factory": callable(file_path, params, app_context) -> QThread,
#     "display_name": str
# }
_TASK_REGISTRY = {}

def register_task(target_format: str, display_name: str, thread_factory: callable):
    """
    Allows independent modules to register their background tasks.
    :param target_format: The string ID of the task (e.g., 'split_docx').
    :param display_name: How it appears in the UI queue.
    :param thread_factory: A lambda or function that returns an instantiated QThread.
                           Signature: func(file_path: Path, params: dict, app_context: dict) -> QThread
    """
    _TASK_REGISTRY[target_format] = {
        "factory": thread_factory,
        "display_name": display_name
    }


# ==========================================
# --- QUEUE MANAGER (THE MODEL) ---
# ==========================================
class QueueManager(QObject):
    log_msg = pyqtSignal(str, int)
    queue_updated = pyqtSignal(list)
    processing_state_changed = pyqtSignal(bool, str)
    conversion_success = pyqtSignal(str)

    def __init__(self, app_context: dict = None):
        super().__init__()
        # Generic dictionary to hold global data (like jar_path) that plugins might need
        self.app_context = app_context or {}
        self.file_queue = []
        self.is_processing = False
        self.conv_thread = None

    def _get_display_name(self, file_path: Path):
        name = file_path.name
        if re.match(r"^\d{4}\.\d{2}\.\d{2} \d{2}-\d{2}-\d{2} ", name):
            return name[20:]
        return name

    def _broadcast_queue_update(self):
        display_items = []
        for index, task in enumerate(self.file_queue, start=1):
            file_path, target_format, _ = task
            display_name = self._get_display_name(file_path)

            # Fetch the clean UI name from the registry
            registry_entry = _TASK_REGISTRY.get(target_format)
            fmt_display = registry_entry["display_name"] if registry_entry else f".{target_format.upper()}"

            display_items.append(f"{index}. {display_name} → {fmt_display}")

        self.queue_updated.emit(display_items)

    def _update_status(self, current_file=""):
        remaining = len(self.file_queue)
        rem_text = f" | {remaining} items waiting in queue." if remaining > 0 else ""

        if current_file:
            self.processing_state_changed.emit(True, f"⚙️ Processing: {current_file}{rem_text}")
        elif self.is_processing:
            self.processing_state_changed.emit(True, f"⚙️ Processing Queue...{rem_text}")
        else:
            self.processing_state_changed.emit(False, "🟢 System Idle.")

    def _route_log(self, *args):
        if len(args) == 1:
            self.log_msg.emit(args[0], logging.INFO)
        elif len(args) >= 2:
            self.log_msg.emit(args[0], args[1])

    def add_item(self, file_path: Path, target_format: str, params: dict = None):
        if target_format not in _TASK_REGISTRY:
            self.log_msg.emit(f"❌ System Error: Unknown task format '{target_format}'.", logging.ERROR)
            return

        self.file_queue.append((file_path, target_format, params or {}))
        self._broadcast_queue_update()
        if not self.is_processing:
            self.process_next()
        else:
            self._update_status()

    def add_batch(self, file_paths: list, target_format: str = "vsdx"):
        if target_format not in _TASK_REGISTRY:
            self.log_msg.emit(f"❌ System Error: Unknown batch task format '{target_format}'.", logging.ERROR)
            return

        for fp in file_paths:
            self.file_queue.append((Path(fp), target_format, {}))
        self._broadcast_queue_update()
        if not self.is_processing:
            self.process_next()
        else:
            self._update_status()

    def remove_items(self, rows: list):
        for row in sorted(rows, reverse=True):
            if 0 <= row < len(self.file_queue):
                del self.file_queue[row]
        self._broadcast_queue_update()
        self._update_status()

    def clear_queue(self):
        self.file_queue.clear()
        self._broadcast_queue_update()
        self._update_status()

    def process_next(self):
        if not self.file_queue:
            self.is_processing = False
            self._update_status()
            return

        self.is_processing = True

        next_file, target_format, params = self.file_queue.pop(0)
        display_name = self._get_display_name(next_file)
        self._broadcast_queue_update()

        registry_entry = _TASK_REGISTRY.get(target_format)
        if not registry_entry:
            self.log_msg.emit(f"❌ Task '{target_format}' is no longer registered.", logging.ERROR)
            self.process_next()
            return

        fmt_display = registry_entry["display_name"]
        self._update_status(f"{display_name} ({fmt_display})")

        try:
            # --- THE MAGIC HANDOFF ---
            # We call the registered factory function, blindly passing the data.
            # The plugin module decides which thread class to create and how to map these parameters.
            self.conv_thread = registry_entry["factory"](next_file, params, self.app_context)

            # Duck-Typing: Connect standard signals if the thread implements them
            if hasattr(self.conv_thread, 'ui_log_msg'):
                self.conv_thread.ui_log_msg.connect(self._route_log)

            if hasattr(self.conv_thread, 'finished_path'):
                self.conv_thread.finished_path.connect(self.conversion_success.emit)

            self.conv_thread.finished.connect(self.process_next)
            self.conv_thread.start()

        except Exception as e:
            self.log_msg.emit(f"❌ Failed to execute task '{target_format}': {str(e)}", logging.ERROR)
            self.process_next()