# --- File: modules/meetings/core/compare_manager.py ---
from PyQt5.QtCore import QObject, pyqtSignal


class ComparisonManager(QObject):
    _instance = None
    # Emits whenever the cart changes: (slot_a_dict, slot_b_dict)
    cart_updated = pyqtSignal(dict, dict)

    @classmethod
    def get_instance(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        super().__init__()
        self.slot_a = {}
        self.slot_b = {}

    def add_to_cart(self, tdoc_name: str, file_path: str):
        """Intelligently pushes the document into the next available slot."""
        if not self.slot_a:
            self.slot_a = {"name": tdoc_name, "path": file_path}
        else:
            # If A is full, put it in B (overwriting B if it was already full)
            self.slot_b = {"name": tdoc_name, "path": file_path}

        self.cart_updated.emit(self.slot_a, self.slot_b)

    def clear_cart(self):
        self.slot_a = {}
        self.slot_b = {}
        self.cart_updated.emit(self.slot_a, self.slot_b)